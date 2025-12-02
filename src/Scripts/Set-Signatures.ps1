<#
.SYNOPSIS
Script to set or update the signatures of users in Outlook Web App and Desktop App on Active Directory environment.

.NOTES
Author: Pablo Gamarra
Started at: 2024-03-19
Last Update: 2024-12-02
Version: 2.0
Github: https://github.com/pablolgamarra

.DESCRIPTION
This script works with Outlook Online Management PowerShell, allowing Microsoft Office 365 or Exchange managers to set the signatures of the users within an organization. It utilizes an HTML template and configures the properties such as DisplayName, PhoneNumber, Mail, etc., by reading the users properties from the Microsoft organization or the Active Directory.

.PARAMETER UserMail
[Optional] Email of the user that is wanted to set the signature for. If its value is $null, all the users signatures will be configured. 

.PARAMETER BusinessName
Business or Organization name used on the generated signatures naming.

.PARAMETER TemplatePath
Path of the HTML template file.

.PARAMETER SetOWA
[Optional] Switch to enable setting signatures on Outlook Web App. Default: $true

.PARAMETER LogPath
[Optional] Path to save execution logs. If not specified, logs only to console.

.EXAMPLE
.\Set-TenantSignatures.ps1 -BusinessName "Acme Corp" -TemplatePath "C:\templates\signature.html"

.EXAMPLE
.\Set-TenantSignatures.ps1 -BusinessName "Acme Corp" -TemplatePath "C:\templates\signature.html" -UserMail "user@acme.com" -LogPath "C:\logs"
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $BusinessName,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $TemplatePath,

    [Parameter()]
    [string] $UserMail,

    [Parameter()]
    [switch] $SetOWA = $true,

    [Parameter()]
    [string] $LogPath
)

begin {
    # Connection state tracker
    $script:ConnectionState = @{
        MgGraph        = $false
        ExchangeOnline = $false
        AzAccount      = $false
    }

    # Statistics tracker
    $script:Statistics = @{
        TotalUsers      = 0
        SuccessfulUsers = 0
        FailedUsers     = 0
        FailedUsersList = @()
    }

    # Logging function
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [string] $Message,
            
            [Parameter()]
            [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
            [string] $Level = "INFO"
        )
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp][$Level] $Message"
        
        # Color coding for console
        switch ($Level) {
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
            default { Write-Host $logMessage }
        }
        
        # If LogPath is specified, append the log message to the file
        if ($LogPath) {
            try {
                $LogDir = Split-Path $LogPath -Parent
                if ($LogDir -and -not (Test-Path -Path $LogDir)) {
                    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
                }
                $LogFile = Join-Path $LogPath "SignatureScript_$(Get-Date -Format 'yyyyMMdd').log"
                Add-Content -Path $LogFile -Value $LogMessage
            }
            catch {
                Write-Warning "Could not write to log file: $_"
            }
        }
    }

    # Verify and install required modules
    function Test-Prerequisites {
        Write-Log "Checking prerequisites..." -Level INFO
        
        $RequiredModules = @('ExchangeOnlineManagement', 'Microsoft.Graph', 'Az.Accounts')
        $InstalledModules = Get-InstalledModule | Select-Object -ExpandProperty Name
        $MissingModules = $RequiredModules | Where-Object { $_ -notin $InstalledModules }

        if ($MissingModules) {
            Write-Log "Missing modules detected: $($MissingModules -join ', ')" -Level WARNING
            Write-Log "Installing missing modules..." -Level INFO
            
            try {
                Install-Module -Name $MissingModules -Force -Scope CurrentUser -AllowClobber
                Write-Log "Modules installed successfully." -Level SUCCESS
            }
            catch {
                Write-Log "Error installing modules: $_" -Level ERROR
                throw "Failed to install required modules: $_"
            }
        }
        else {
            Write-Log "All required modules are installed." -Level SUCCESS
        }
    }

    # Read and validate the HTML template file
    function Read-Template {
        param (
            [Parameter(Mandatory = $true)]
            [string] $TemplatePath
        )
        
        try {
            Write-Log "Reading HTML template file at path: $TemplatePath" -Level "INFO"
            $HTML_SIGNATURE = Get-Content -Path $TemplatePath -Encoding UTF8 -Raw
        }
        catch {
            Throw "Error while reading the HTML template file at path: $TemplatePath. $_"
        }

        if (-not $HTML_SIGNATURE -or $HTML_SIGNATURE.Trim() -eq "") {
            throw "The HTML template file at path: $TemplatePath is empty."
        }

        # Validate placeholders {0} to {4}
        if ($HTML_SIGNATURE -notmatch '\{0\}|\{1\}|\{2\}|\{3\}|\{4\}') {
            throw "The HTML template does not contain the required placeholders {0} to {4}."
        }
        
        Write-Log "HTML template validated successfully." -Level SUCCESS
        return $HTML_SIGNATURE
    }

    # Connect to Microsoft Graph
    function Connect-ToMgGraph {
        if ($script:ConnectionState.MgGraph) {
            Write-Log "Already connected to Microsoft Graph." -Level INFO
            return
        }

        try {
            Write-Log "Connecting to Microsoft Graph..." -Level "INFO"
            Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop
            $script:ConnectionState.MgGraph = $true
            Write-Log "Microsoft Graph connected successfully." -Level "SUCCESS"
        }
        catch {
            Throw "An error occurred while connecting to Microsoft Graph: $_"
        }
    }

    # Connect to Exchange Online
    function Connect-ToExchangeOnline {
        if ($script:ConnectionState.ExchangeOnline) {
            Write-Log "Already connected to Exchange Online." -Level INFO
            return
        }

        try {
            Write-Log "Connecting to Exchange Online..." -Level INFO
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:ConnectionState.ExchangeOnline = $true
            Write-Log "Exchange Online connected successfully." -Level SUCCESS
        }
        catch {
            throw "Failed to connect to Exchange Online: $_"
        }
    }

    # Generate signatures for user(s)
    function Get-Signatures {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string] $BusinessName,
    
            [Parameter(Mandatory = $true)]
            [string] $HTMLTemplate,
    
            [string] $UserMail
        )
    
        Connect-ToMgGraph
    
        # Query users from Microsoft Graph
        Write-Log "Querying users from Microsoft Graph..." -Level INFO
        
        $Users = @()

        try {
            if ($UserMail) {
                Write-Log "Filtering for specific user: $UserMail" -Level INFO
                $Users = @(Get-MgUser -All -Property DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhones | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhones |  Where-Object { $_.Mail -eq $UserMail -and $_.Mail })
            }
            else {
                $Users = Get-MgUser -All -Property DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhones | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhones | Where-Object { $_.Mail -and $_.Mail -ne "" }
            }
        }
        catch {
            throw "Error querying users from Microsoft Graph: $_"
        }

        if (-not $Users -or $Users.Count -eq 0 -or $null -eq $Users.Count) {
            Write-Host $Users.GetType().FullName
            throw "No users found with the specified criteria."
        }

        Write-Log "Found $($Users.Count) user(s) to process." -Level SUCCESS
        $script:Statistics.TotalUsers = $Users.Count
    
        $FormattedSignatures = [System.Collections.ArrayList]::new()
        $i = 0

        # Iterate over users to generate signatures
        foreach ($User in $Users) {
            $i++
            Write-Progress -Activity "Generating signatures" -Status "Processing $($User.DisplayName) ($i of $($Users.Count))" -PercentComplete (($i / $Users.Count) * 100)
            
            try {
                # Get first business phone or empty string
                $BusinessPhone = if ($User.BusinessPhones -and $User.BusinessPhones.Count -gt 0) { 
                    $User.BusinessPhones[0] 
                }
                else { 
                    "" 
                }

                $FormattedSignature = ($HTMLTemplate -f $User.DisplayName, $User.JobTitle, $BusinessPhone, $User.Mail, $User.Mail)

                $AuxSignature = [PSCustomObject]@{
                    UserName      = $User.DisplayName
                    Mail          = $User.Mail
                    Signature     = $FormattedSignature
                    SignatureName = "Firma-$BusinessName-$($User.DisplayName)"
                }
        
                $FormattedSignatures.Add($AuxSignature)
                Write-Log "Signature generated for: $($User.DisplayName)" -Level INFO
            }
            catch {
                Write-Log "Error generating signature for $($User.DisplayName): $_" -Level ERROR
                $script:Statistics.FailedUsers++
                $script:Statistics.FailedUsersList += $User.DisplayName
            }
        }
        
        Write-Progress -Activity "Generating signatures" -Completed
        return ,$FormattedSignatures.ToArray()
    }

    # Set signatures on Outlook Web App
    function Set-OWASignatures {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Array] $Signatures
        )

        # Connect to Exchange Online
        Connect-ToExchangeOnline

        # Configure roaming signatures
        try {
            Write-Log "Checking roaming signatures configuration..." -Level INFO
            $RoamingConfigured = (Get-OrganizationConfig).PostponeRoamingSignaturesUntilLater
            
            if (-not $RoamingConfigured) {
                Write-Log "Postponing roaming signatures on organization..." -Level INFO
                Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
                Write-Log "Roaming signatures postponed." -Level SUCCESS
            }
            else {
                Write-Log "Roaming signatures already configured." -Level INFO
            }
        }
        catch {
            Write-Log "Warning while configuring roaming signatures: $_" -Level WARNING
        }

        # Set signatures for each user
        $i = 0
        foreach ($Signature in $Signatures) {
            $i++
            Write-Progress -Activity "Setting signatures in OWA" -Status "Processing $($Signature.UserName) ($i of $($Signatures.Count))" -PercentComplete (($i / $Signatures.Count) * 100)
            
            try {
                Write-Log "Setting signature for: $($Signature.UserName) ($($Signature.Mail))" -Level INFO
                
                Set-MailboxMessageConfiguration -Identity $Signature.Mail `
                    -SignatureHTML $Signature.Signature `
                    -AutoAddSignature $true `
                    -AutoAddSignatureOnReply $true `
                    -SignatureName $Signature.SignatureName `
                    -ErrorAction Stop
                
                Write-Log "Signature applied successfully for: $($Signature.UserName)" -Level SUCCESS
                $script:Statistics.SuccessfulUsers++
            }
            catch {
                Write-Log "Error setting signature for $($Signature.UserName): $_" -Level ERROR
                $script:Statistics.FailedUsers++
                $script:Statistics.FailedUsersList += $Signature.UserName
            }
        }
        
        Write-Progress -Activity "Setting signatures in OWA" -Completed
    }

    # Disconnect all services safely
    function Disconnect-Services {
        Write-Log "Disconnecting services..." -Level INFO
        
        if ($script:ConnectionState.MgGraph) {
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                Write-Log "Disconnected from Microsoft Graph." -Level SUCCESS
            }
            catch {
                Write-Log "Error disconnecting from Microsoft Graph: $_" -Level WARNING
            }
        }

        if ($script:ConnectionState.ExchangeOnline) {
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                Write-Log "Disconnected from Exchange Online." -Level SUCCESS
            }
            catch {
                Write-Log "Error disconnecting from Exchange Online: $_" -Level WARNING
            }
        }

        if ($script:ConnectionState.AzAccount) {
            try {
                Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
                Write-Log "Disconnected from Azure Account." -Level SUCCESS
            }
            catch {
                Write-Log "Error disconnecting from Azure Account: $_" -Level WARNING
            }
        }
    }

    # Display execution statistics
    function Show-Statistics {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "         EXECUTION SUMMARY" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Total Users Processed:    $($script:Statistics.TotalUsersProcessed)" -ForegroundColor White
        Write-Host "Successful:               $($script:Statistics.SuccessfulUsers)" -ForegroundColor Green
        Write-Host "Failed:                   $($script:Statistics.FailedUsers)" -ForegroundColor Red
        
        if ($script:Statistics.FailedUsersList.Count -gt 0) {
            Write-Host "`nFailed Users:" -ForegroundColor Yellow
            $script:Statistics.FailedUsersList | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        }
        
        Write-Host "========================================`n" -ForegroundColor Cyan
    }
}

process {
    try {
        Write-Log "========================================" -Level "INFO"
        Write-Log "Script execution started" -Level "INFO"
        Write-Log "========================================" -Level "INFO"
        Write-Log "Organization Name: $BusinessName" -Level "INFO"
        Write-Log "Template Path: $TemplatePath" -Level "INFO"
        Write-Log "User Mail Filter: $(if($UserMail){$UserMail}else{'All users'})" -Level "INFO"
        Write-Log "Set OWA Signatures: $SetOWA" -Level "INFO"
        if ($LogPath) { Write-Log "Log Path: $LogPath" -Level "INFO" }
        Write-Log "========================================`n" -Level "INFO"

        # Check prerequisites
        Test-Prerequisites

        # Read and validate HTML template
        $HTMLTemplate = Read-Template -TemplatePath $TemplatePath

        # Generate signatures
        Write-Log "`nGenerating signatures..." -Level "INFO"
        $GeneratedSignatures = Get-Signatures -BusinessName $BusinessName -HTMLTemplate $HTMLTemplate -UserMail $UserMail
        Write-Log "Generated $($GeneratedSignatures.Count) signature(s)." -Level "SUCCESS"

        Write-Host $GeneratedSignatures
        # Set signatures on OWA if enabled
        if ($SetOWA -and $GeneratedSignatures.Count -gt 0) {
            Write-Log "`nSetting signatures in Outlook Web App..." -Level "INFO"
            Set-OWASignatures -Signatures $GeneratedSignatures
        }
        elseif (-not $SetOWA) {
            Write-Log "Skipping OWA signature configuration (disabled by parameter)." -Level "WARNING"
        }

        Write-Log "`nScript execution completed successfully." -Level "SUCCESS"
    }
    catch {
        Write-Log "CRITICAL ERROR: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
        throw
    }
}

end {
    # Disconnect services
    Disconnect-Services
    
    # Show statistics
    Show-Statistics
    
    Write-Log "Script finished." -Level "INFO"
}
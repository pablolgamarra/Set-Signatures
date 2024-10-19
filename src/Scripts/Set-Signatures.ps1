<#
.SYNOPSIS
Script to set or update the signatures of an user of Outlook Web App and Desktop App on Active Directory environment.

.NOTES
Author: Pablo Gamarra
Started at: 2024-03-19
Last Update: 2024-04-13
Versión: 1.3
Github: https://github.com/pablolgamarra

.DESCRIPTION
This script works with Outlook Online Management PowerShell, allowing Microsoft Office 365 or Exchange managers to set the signatures of the users within an organization. It utilizes a HTML template and configures the properties such as DisplayName, PhoneNumber, Mail, etc., by reading the users properties from the Microsoft organization or the Active Directory.

.PARAMETER UserMail
[Optional] Email of the user that are wanted to set the signature for. If its value is $null, all the users signatures will be configured. 

.PARAMETER BusinessName
Business or Organization name used on the generated signatures naming.

.PARAMETER TemplatePath
Path of the HTML template file.
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $BusinessName,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $TemplatePath,

    [Parameter()]
    [string] $UserMail
)

begin {
    # Connect to Microsoft Graph service using credentials
    function Connect-MgGraphViaCred {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.PSCredential] $Credential,

            [string] $Tenant = $_tenantDomain
        )

        # Connect to Azure using credentials.
        $param = @{
            Credential = $credential
            Force      = $true
        }
        if ($tenant) { $param.tenant = $tenant }
        $null = Connect-AzAccount @param

        # Get token for MSGraph.
        $token = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).token

        # Convert token to SecureString if Connect-MgGraph new version is used
        if ((Get-Help Connect-MgGraph -Parameter accesstoken).type.name -eq "securestring") {
            $token = ConvertTo-SecureString $token -AsPlainText -Force
        }

        # Connect to Microsoft Graph using token.
        $null = Connect-MgGraph -AccessToken $token -ErrorAction Stop
    }

    # Generate the Signatures for the user(s)
    function Get-Signatures {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string] $BusinessName,
    
            [Parameter(Mandatory = $true)]
            [string] $TemplatePath,
    
            [string] $UserMail
        )
    
        $HTML_SIGNATURE = Get-Content -Path $TemplatePath -Encoding UTF8
    
        # Connect to MgGraph and Exchange Online Management module.
        try {
            Write-Host "Connecting to MgGraph."
            Connect-MgGraph -Scopes "User.Read.All"
            Write-Host "MgGraph connected."
        }
        catch {
            Write-Host "An error occurred while connecting to MgGraph"
            Write-Host $_.Exception.Message
        }
    
        # Query users from MgGraph service
        if ($UserMail) {
            $Users = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone | Where-Object -Property Mail -eq $UserMail
        }
        else {
            $Users = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone
        }
    
        $FormattedSignatures = @()
        # Iterate over users to set the user information in the signature
        foreach ($User in $Users) {
            $FormattedSignature = ($HTML_SIGNATURE -f $User.DisplayName, $User.JobTitle, $User.BusinessPhone, $User.Mail, $User.Mail)

            $AuxSignature = [PSCustomObject]@{
                UserName      = $User.DisplayName
                Mail          = $User.Mail
                Signature     = $FormattedSignature
                SignatureName = "Firm-$BusinessName-$($User.DisplayName)"
            }
    
            $FormattedSignatures += $AuxSignature
        }
        return $FormattedSignatures
    }

    # Set the user(s) signature(s) on Outlook Web App (OWA).
    function Set-OWASignatures {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $True)]
            [System.Array] $Signatures
        )

        try {
            Write-Host "Connecting to Exchange Online Management."
            Connect-ExchangeOnline
            Write-Host "Exchange Online Management connected."
        }
        catch {
            Write-Host "An error occurred while connecting Exchange Online Management"
            Write-Host $_
        }

        #Postpone roaming signatures on the organization if it is not configured yet
        try {
            $RoamingConfigured = Get-OrganizationConfig -PostponeRoamingSignaturesUntilLater
            
            if (-not ($RoamingConfigured)) {
                Write-Host "Postponing roaming signatures on organization."
                Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
            }
        }
        catch {
            Write-Host "Error while postponing roaming signatures on the organization"
            Write-Host $_ 
        }

        # Establecer las firmas en las cuentas
        foreach ($Signature in $Signatures) {
            Write-Host "Setting the signature of user mailbox: $($Signature.UserName)"

            Set-MailboxMessageConfiguration -Identity $Signature.Mail -SignatureHTML $Signature.Signature -AutoAddSignature $true -AutoAddSignatureOnReply $true -SignatureName $Signature.SignatureName -SignatureHTMLBody $Signature.Signature
            Write-Host "New signature applied to user mailbox"
        }
    }

    # Set the user(s) signature(s) on Outlook Desktop Application (OWA)
    function Set-ODASignatures {
        [CmdLetBinding()]
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.PSCredential] $Credential
        )
        Throw "Not implemented yet."
    }
}

process {
    try {
        Write-Host "Signature Template Path: $TemplatePath"
        Write-Host "Organization Name: $BusinessName"
        Write-Host "User Mail: $UserMail"


        # Install required modules if they are not installed yet
        $RequiredModules = @('ExchangeOnlineManagement', 'Microsoft.Graph', 'Az.Accounts')
        $InstalledModules = Get-InstalledModule | Select-Object -ExpandProperty Name
        $MissingModules = $RequiredModules | Where-Object { $_ -notin $InstalledModules }

        if ($MissingModules) {
            Install-Module -Name $MissingModules -Force
        }

        # Obtener las firmas HTML de los usuarios
        $GeneratedSignatures = Get-Signatures -TemplatePath $TemplatePath -BusinessName $BusinessName -UserMail $UserMail

        Write-Host "Generated Signatures quantity: $($GeneratedSignatures.Count)"

        #Set-OWASignatures -Signatures $GeneratedSignatures

        Write-Host "Signatures Set successfully."
    }
    catch {
        Write-Error "There is an error occurred: $($_.Exception.Message)"
    }
}

end {
    Disconnect-AzAccount
    Disconnect-MgGraph
    Disconnect-ExchangeOnline
}
# Modificar las firmas de los usuarios de una organización.
# Créditos a https://office365itpros.com/ por su inspiración y código base.

# CmdLet personalizado para conectar a Microsoft Graph usando credenciales.
# Créditos a Ondrej Sebela - https://doitpsway.com/how-to-connect-to-the-microsoft-graph-api-using-saved-user-credentials
function Connect-MgGraphViaCred {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $credential,

        [string] $tenant = $_tenantDomain
    )

    # Conectar a Azure usando credenciales.
    $param = @{
        Credential = $credential
        Force      = $true
    }
    if ($tenant) { $param.tenant = $tenant }
    $null = Connect-AzAccount @param

    # Obtener token para MSGraph.
    $token = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).token

    # Convertir token a SecureString si se usa nueva versión de Connect-MgGraph.
    if ((Get-Help Connect-MgGraph -Parameter accesstoken).type.name -eq "securestring") {
        $token = ConvertTo-SecureString $token -AsPlainText -Force
    }

    # Conectar a Microsoft Graph usando el token.
    $null = Connect-MgGraph -AccessToken $token -ErrorAction Stop
}

function Get-FirmasHTML{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credencial,

        [Parameter(Mandatory = $true)]
        [string] $NombreEmpresa,

        [Parameter(Mandatory = $true)]
        [string] $PathFirmaHtml,

        [string] $CorreoUsuario
    )

    $FIRMA_HTML = Get-Content -Path $PathFirmaHtml 

    # Conectar a servicios MG Graph y Exchange Online.
    try {
        Write-Verbose "Conectando a servicio MgGraph..."
        Connect-MgGraphViaCred -credential $Credencial    
    }
    catch {
        Write-Host "Ocurrio un error al conectar al servicio de MgGraph."
        Write-Host $_ -ErrorAction Stop
    }
    Write-Host "Servicio MgGraph conectado."

    # Obtener usuarios de MG Graph.
    if($CorreoUsuario){
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone | Where-Object -Property Mail -eq $CorreoUsuario
    }else{
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone
    }

    $FirmasFormateadas = @()
    # Iterar sobre usuarios para establecer firmas.
    foreach ($Usuario in $Usuarios) {
        $FirmaFormateada = ($FIRMA_HTML -f $Usuario.DisplayName, $Usuario.JobTitle, $Usuario.BusinessPhone, $Usuario.Mail, $Usuario.Mail)
        $FirmaObj = [PSCustomObject]@{
            NombreUsuario = $Usuario.DisplayName
            Correo = $Usuario.Mail
            FirmaFormateada = $FirmaFormateada
            NombreFirma = "Firma-$NombreEmpresa-$($Usuario.DisplayName)"
        }

        $FirmasFormateadas = $FirmasFormateadas + $FirmaObj
    }
    return $FirmasFormateadas
}
# Funcion para establecer las firmas en la App Web de Outlook (OWA)
function Set-OWASignatures{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [System.Array] $Firmas,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential] $Credencial
    )

    try {
        Write-Verbose "Conectando a servicio Exchange Online..."
        Connect-ExchangeOnline -Credential $Credencial -ShowBanner $false
    }
    catch {
        Write-Host "Ocurrio un error al conectar al servicio de Exchange Online."
        Write-Host $_ -ErrorAction Stop
    }

    Write-Host "Servicio Exchange Online conectado."

    #Hacer la configuracion para desactivar las Roaming Signatures en Exchange Online
    try{
        Write-Verbose "Configurar organizacion para posponer el uso de las firmas en Roaming en OWA"
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
    }catch{
        Write-Host "Error al configurar la organizacion"
        Write-Host $_ -ErrorAction Stop
    }

    # Establecer las firmas en las cuentas
    foreach ($Firma in $Firmas) {
        Write-Host "Configurando firma de: $($Firma.Mail)"
        Set-MailboxMessageConfiguration $Firma.Mail -SignatureHTML $Firma.FirmaFormateada -AutoAddSignature $true -AutoAddSignatureOnReply $true
        Write-Host "Firma configurada correctamente"
    }
}

# Función para establecer firmas.
function Set-Signatures {
    [CmdletBinding()]
    param(
        [Parameter]
        [string] $CorreoUsuario,

        [Parameter(Mandatory = $true)]
        [string] $NombreEmpresa,

        [Parameter(Mandatory = $true)]
        [string] $PathFirmaHtml,

        [Parameter(Mandatory=$true)]
        [boolean] $OWA,

        [Parameter(Mandatory=$true)]
        [boolean] $ODA
    )

    # Instalar módulos necesarios si no están instalados.
    $ModulosRequeridos = @('ExchangeOnlineManagement', 'Microsoft.Graph', 'Az.Accounts')
    $ModulosInstalados = Get-InstalledModule | Select-Object -ExpandProperty Name
    $ModulosNoInstalados = $ModulosRequeridos | Where-Object { $_ -notin $ModulosInstalados }

    if ($ModulosNoInstalados) {
        Install-Module -Name $ModulosNoInstalados -Force
    }

    # Leer credenciales para conectar a los servicios
    $Creds = Get-Credential

    #Obtener las firmas HTML de los usuarios o usuario especifico
    $FirmasFormateadas = Get-FirmasHTML -Credencial $Creds -PathFirmaHtml $PathFirmaHtml -NombreEmpresa $NombreEmpresa -CorreoUsuario $CorreoUsuario

    Set-OWASignatures -Firmas $FirmasFormateadas -Credencial $Creds

    Write-Host "Firmas configuradas correctamente"
}



function Set-ODASignatures{
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential] $Credencial
    )
    $PathFirmaLocal = (Get-Item env:appdata).Value + '\Microsoft\Signatures'

    foreach ($Firma in $FirmasFormateadas) {
        #TODO: Definir bien donde van a ir los archivos
        $CarpetaFirmas = <#$PathFirmasHtml#> "C:\Users\pablogamarra\DESARROLLO\PS\Signature-Set\FirmasGeneradas"
        
        #Crear la carpeta donde se guardan las firmas si no existe
        if(!(Test-Path -Path $CarpetaFirmas)){
            New-Item -ItemType Directory -Path $CarpetaFirmas
        }

        #Crear las firmas .htm para cada usuario
        Add-Content -Path "C:\Users\pablogamarra\DESARROLLO\PS\Signature-Set\FirmasGeneradas\$($firma.NombreFirma).htm" -Value $firma.FirmaFormateada

        #Conectarse a la computadora en la que se detecta el email
        $MailUsuario = $Firma.Correo

        $UsuarioAD = Get-ADUser -Filter {UserPrincipalName -eq $MailUsuario}

        $computers = Get-ADComputer -Filter {LastLogonDate -like "*$($UsuarioAD.Name)*"}
        
        foreach ($computer in $computers) {
            Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name | Select-Object Name, UserName, @{Name="LastLogonDate";Expression={$_.ConvertToDateTime($_.LastBootUpTime)}}
        }
}

        <#To check for the user's last logon date, you can use the Win32_NetworkLoginProfile WMI class and retrieve the LastLogon property which is in the format of a timestamp. Here is an example of how you can do this:

Copy code
$user = 'Username'
$computers = Get-ADComputer -Filter {Enabled -eq 'true' -and SamAccountName -like $Prefix}
foreach ($computer in $computers) {
    $lastLogon = (gwmi -Class Win32_NetworkLoginProfile -ComputerName $computer.Name | Where-Object {$_.Name -eq $user}).LastLogon
    if($lastLogon){
        write-host "$user last logged on $computer on $(($lastLogon).ToLocalTime())"
    }
}
This script uses the Get-ADComputer cmdlet to find all enabled computers that match the specified prefix. Then it uses a foreach loop to iterate through each computer and run the Get-WmiObject cmdlet to retrieve the Win32_NetworkLoginProfile class, and filters it by the username you are searching for. The LastLogon property is retrieved and converted to a readable date format using the .ToLocalTime() method.

It's important to note that the LastLogon property only gives you the last logon date of the user on that specific computer and it's not a global property across the domain. Also, the above script uses Win32_NetworkLoginProfile class which is only available on Windows 7 and Windows Server 2008 R2 or later.

Also, you need to make sure that the necessary PowerShell modules are installed and available in your environment and the user running the script has the permission to run the script.#>

    }

    $PathFirmaNueva = Join-Path -Path $PathFirmaLocal -ChildPath "$NombreFirma.htm"
    $FirmaFormateada | Out-File -FilePath $PathFirmaNueva -Force

        <#
    # Actualizar configuración de firma en Outlook.
    $OutlookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $Profiles = (Get-ChildItem $OutlookProfilePath).PSChildName

    if (!$Profiles -or $Profiles.Count -ne 1) {
        Write-Host "Warning - Applying signature to all Outlook profiles" 
        $OutlookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
    }
    else {
        $OutLookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$($Profiles.Trim())\9375CFF0413111d3B88A00104B2A6676\00000002"
    }

    $OutlookPathLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'

    if (Test-Path $OutlookPathLocation) {
        Get-Item -Path $OutlookProfilePath | New-ItemProperty -Name "New Signature" -Value $NombreFirma -PropertyType String -Force 
        Get-Item -Path $OutlookProfilePath | New-ItemProperty -Name "Reply-Forward Signature" -Value $NombreFirma -PropertyType String -Force 
    }
    #>
}


# Llamar a la función con los parámetros requeridos.
Set-Signatures -PathFirmaHtml ".\signature.htm" -NombreEmpresa "EMPRESA DE PRUEBAS" -OWA $true -ODA $true
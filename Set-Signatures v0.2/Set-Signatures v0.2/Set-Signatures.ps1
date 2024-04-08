# Modificar las firmas de los usuarios de una organización.
# Créditos a https://office365itpros.com/ por su inspiración y código base.

# CmdLet personalizado para conectar a Microsoft Graph usando credenciales.
# Créditos a Ondrej Sebela - https://doitpsway.com/how-to-connect-to-the-microsoft-graph-api-using-saved-user-credentials

param(
    [Parameter()]
    [string] $CorreoUsuario,

    [Parameter(Mandatory = $true)]
    [string] $NombreEmpresa,

    [Parameter(Mandatory = $true)]
    [string] $PathPlantillaHtml,

    [Parameter(Mandatory = $false)]
    [switch] $OWA,

    [Parameter(Mandatory = $false)]
    [switch] $ODA
)

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

function Get-LoggedOnUserAD {
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [ValidateScript({ Test-Connection -ComputerName $_ -Quiet -Count 1 })]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $Credentials
    )
    foreach ($comp in $ComputerName) {
        $output = @{ 'ComputerName' = $comp }
        $output.UserName = (Get-WmiObject -Class win32_computersystem -ComputerName $comp -Credential $Creds).UserName
        [PSCustomObject]$output
    }
}

#Formatear y rellenar los datos de los usuarios en la plantilla HTML del archivo que se defina
function Get-FirmasHTML {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $Credencial,

        [Parameter(Mandatory = $true)]
        [string] $NombreEmpresa,

        [Parameter(Mandatory = $true)]
        [string] $PathPlantillaHTML,

        [string] $CorreoUsuario
    )

    $FIRMA_HTML = Get-Content -Path $PathPlantillaHtml 

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
    if ($CorreoUsuario) {
        Write-Verbose "Configurando solamente la firma de un usuario"
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone | Where-Object -Property Mail -eq $CorreoUsuario
        Write-Verbose $Usuarios.Count
    }
    else {
        Write-Verbose "Configurando solamente la firma de un usuario"
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone
        Write-Verbose $Usuarios.Count
    }

    $FirmasFormateadas = @()
    # Iterar sobre usuarios para establecer firmas.
    foreach ($Usuario in $Usuarios) {
        $FirmaFormateada = ($FIRMA_HTML -f $Usuario.DisplayName, $Usuario.JobTitle, $Usuario.BusinessPhone, $Usuario.Mail, $Usuario.Mail)
        $FirmaObj = [PSCustomObject]@{
            NombreUsuario   = $Usuario.DisplayName
            Correo          = $Usuario.Mail
            FirmaFormateada = $FirmaFormateada
            NombreFirma     = "Firma-$NombreEmpresa-$($Usuario.DisplayName)"
        }

        $FirmasFormateadas = $FirmasFormateadas + $FirmaObj
    }
    return $FirmasFormateadas
}   
function Get-UsuariosEquipos {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $CredsAdminDom,
        [Parameter(Mandatory = $true)]
        [string] $filtro
    )
    
    $ErrorActionPreference = 'SilentlyContinue'
    
    $Equipos = Get-ADComputer -Filter filtro
    $UsuariosEquipos = @()
    
    foreach ($Equipo in $Equipos) {
        Write-Host "Computadora:$($Equipo.Name)"
        $Computadora = $Equipo.Name
        $Usuario = (Get-WmiObject -Class win32_computersystem -ComputerName $Equipo.Name -Credential $CredsAdminDom).UserName

        $UsuarioEquipoObj = [PSCustomObject]@{
            $CorreoUsuario    = (Get-ADUser -Filter "SamAccountName -eq '$(($Usuario -split '\\')[1])'").UserPrincipalName
            $UsuarioEquipoObj = [psobject]@{
                NombreEquipo  = $Computadora
                Usuario       = $Usuario
                CorreoUsuario = $CorreoUsuario
            }
        }
        
        $UsuariosEquipos += $UsuarioEquipoObj
    }
}

# Funcion para establecer las firmas en la App Web de Outlook (OWA)
function Set-OWASignatures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        [System.Array] $Firmas
    )

    try {
        Write-Verbose "Conectando a servicio Exchange Online..."
        Connect-ExchangeOnline
    }
    catch {
        Write-Host "Ocurrio un error al conectar al servicio de Exchange Online."
        Write-Host $_ -ErrorAction Stop
    }

    Write-Host "Servicio Exchange Online conectado."

    #Hacer la configuracion para desactivar las Roaming Signatures en Exchange Online
    try {
        Write-Verbose "Configurar organizacion para posponer el uso de las firmas en Roaming en OWA"
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
    }
    catch {
        Write-Host "Error al configurar la organizacion"
        Write-Host $_ -ErrorAction Stop
    }

    # Establecer las firmas en las cuentas
    foreach ($Firma in $Firmas) {
        Write-Host "Configurando firma de: $($Firma.Correo)"
        Set-MailboxMessageConfiguration -Identity $Firma.Correo -SignatureHTML $Firma.FirmaFormateada -AutoAddSignature $true -AutoAddSignatureOnReply $true -SignatureName $Firma.NombreFirma -SignatureHTMLBody $Firma.FirmaFormateada
        Write-Host "Firma configurada correctamente"
    }
}

# Función para establecer firmas en la app Desktop de Outlook
function Set-ODASignatures {
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential] $Credencial
    )

    #OBTENER CREDENCIALES DE USUARIO ADMINISTRADOR DE DOMINIO
    #$CredsDominio = Get-Credential

    $UsuariosEquipos = Get-UsuariosEquipos -filtro * -CredsAdminDom $CredsDominio

    foreach ($Firma in $FirmasFormateadas) {
        $CacheFirmasFormateadas = "$PATH_SCRIPT\CacheFirmas"

        #Crear la carpeta donde se guardan las firmas si no existe
        if (!(Test-Path -Path $CacheFirmasFormateadas)) {
            New-Item -ItemType Directory -Path $CacheFirmasFormateadas
        }

        #Crear las firmas .htm para cada usuario
        Add-Content -Path "$CacheFirmasFormateadas\$($firma.NombreFirma).htm" -Value $firma.FirmaFormateada

        #Conectarse a la computadora en la que se detecta el email
        $UsuariosAD = Get-ADUser -Filter * | Select-Object SamAccountName, UserPrincipalName
        $MailUsuario = $Firma.Correo    

        #Encontrar las computadoras en las que el usuario haya iniciado sesion

        <#TODO: Ver que me devuelve la expresion de arriba, necesito el nombre de la PC para conectarme a ella y ejecutar comandos para 
        1- Obtener el path de los perfiles de Outlook
        2- Copiar la firma al path adecuado
        3- Establecer la firma HTML formateada.#>
    }

    <#
    $PathFirmaOutlook = (Get-Item env:appdata).Value + '\Microsoft\Signatures'

    # Actualizar configuración de firma en Outlook.
    $PathPerfilesOutlook = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $Perfiles = (Get-ChildItem $OutlookProfilePath).PSChildName

    if (!$Perfiles -or $Perfiles.Count -ne 1) {
        Write-Host "Aplicando Firmas a todos los perfiles" 
        $PathPerfilesOutlook = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
    }
    else {
        $PathPerfilesOutlook = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$($Profiles.Trim())\9375CFF0413111d3B88A00104B2A6676\00000002"
    }

    #$OutlookPathLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'

    if (Test-Path $PathFirmaOutlook) {
        Get-Item -Path $OutlookProfilePath | New-ItemProperty -Name "New Signature" -Value $firma.NombreFirma -PropertyType String -Force 
        Get-Item -Path $OutlookProfilePath | New-ItemProperty -Name "Reply-Forward Signature" -Value $firma.NombreFirma -PropertyType String -Force 
    }
    #>
}

try{
    $PATH_SCRIPT = Split-Path $MyInvocation.MyCommand.Path -Parent
    $ErrorActionPreference = 'Stop'
    
    # Función para establecer firmas.
    # Instalar módulos necesarios si no están instalados.
    $ModulosRequeridos = @('ExchangeOnlineManagement', 'Microsoft.Graph', 'Az.Accounts')
    $ModulosInstalados = Get-InstalledModule | Select-Object -ExpandProperty Name
    $ModulosNoInstalados = $ModulosRequeridos | Where-Object { $_ -notin $ModulosInstalados }
    
    if ($ModulosNoInstalados) {
        Install-Module -Name $ModulosNoInstalados -Force
    }
    
    #Obtener las firmas HTML de los usuarios o usuario especifico
    $FirmasFormateadas = Get-FirmasHTML -Credencial $Creds -PathPlantillaHtml $PathPlantillaHtml -NombreEmpresa $NombreEmpresa -CorreoUsuario $CorreoUsuario
    
    #Si es la Outlook Web App, no se necesitan crear los archivos en Cache
    if($OWA){
        Write-Verbose "Ingresando al modulo de Configuracion OWA"
        if(-not ($FirmasFormateadas)){
            Write-Verbose "No hay datos de firmas formateadas."
        }else{
            Set-OWASignatures -Firmas $FirmasFormateadas
        }
    }
    
    if($ODA){
        Write-Verbose "Ingresando al modulo de Configuracion ODA"
        Set-ODASignatures 
    }
    
    
    Write-Host "Firmas configuradas correctamente"
}catch{
    Write-Verbose "ERROR"
    $_.Exception.Message
}
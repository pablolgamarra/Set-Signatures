# Modificar las firmas de los usuarios de una organización.
# Créditos a https://office365itpros.com/ por su inspiración y código base.

# CmdLet personalizado para conectar a Microsoft Graph usando credenciales.
# Créditos a Ondrej Sebela - https://doitpsway.com/how-to-connect-to-the-microsoft-graph-api-using-saved-user-credentials

#TODO: Ver que el script pare la ejecucion cuando no algo sale mal

$ErrorActionPreference = 'Stop'
#Utilidad para conectarse a MgGraph unicamente con credenciales
function Connect-MgGraphViaCred {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $Creds,

        [string] $tenant = $_tenantDomain
    )

    # Conectar a Azure usando credenciales.
    $param = @{
        Credential = $Creds
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

#Obtener usuario conectado a cada equipo del dominio
function Get-DeviceUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $CredsDomainAdm,
        [Parameter(Mandatory = $true)]
        [string] $filtro
    )
    
    $ErrorActionPreference = 'SilentlyContinue'
    
    $Equipos = Get-ADComputer -Filter filtro
    $UsuariosEquipos = @()
    
    foreach ($Equipo in $Equipos) {
        Write-Host "Computadora:$($Equipo.Name)"
        $Computadora = $Equipo.Name
        $Usuario = (Get-WmiObject -Class win32_computersystem -ComputerName $Equipo.Name -Credential $CredsDomainAdm).UserName

        $UsuarioEquipoObj = [PSCustomObject]@{
            $CorreoUsuario = (Get-ADUser -Filter "SamAccountName -eq '$(($Usuario -split '\\')[1])'").UserPrincipalName
            $UsuarioEquipoObj = [psobject]@{
                NombreEquipo  = $Computadora
                Usuario       = $Usuario
                CorreoUsuario = $CorreoUsuario
            }
        }
        
        $UsuariosEquipos += $UsuarioEquipoObj
    }
}

#Formatear y rellenar los datos de los usuarios en la plantilla HTML del archivo que se defina
function Get-HTMLSignatures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $Credencial,

        [Parameter(Mandatory = $true)]
        [string] $NombreEmpresa,

        [Parameter(Mandatory = $true)]
        [string] $HTMLTemplate,

        [string] $CorreoUsuario
    )

    $FIRMA_HTML = Get-Content -Path $HTMLTemplate 

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
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone | Where-Object -Property Mail -eq $CorreoUsuario
    }
    else {
        $Usuarios = Get-MgUser -All | Select-Object DisplayName, Mail, UserPrincipalName, JobTitle, BusinessPhone
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

# Funcion para establecer las firmas en la App Web de Outlook (OWA)
function Set-OWASignatures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        [System.Array] $Firmas,

        [Parameter(Mandatory = $True)]
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
        Set-MailboxMessageConfiguration $Firma.Correo -SignatureHTML $Firma.FirmaFormateada -AutoAddSignature $true -AutoAddSignatureOnReply $true
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

    throw "No implementado"
<#
    #TEMPORAL
    $UserName = "glymax\olam"
    $Pass = ConvertTo-SecureString -String "SuperGlymax80023016" -AsPlainText -Force
    $CredsDominio = New-Object System.Management.Automation.PSCredential($UserName, $Pass)

    #OBTENER CREDENCIALES DE USUARIO ADMINISTRADOR DE DOMINIO
    #$CredsDominio = Get-Credential

    $UsuariosEquipos = Get-UsuariosEquipos -filtro * -CredsDomainAdm $CredsDominio

    foreach ($Firma in $FirmasFormateadas) {
        $CacheFirmasFormateadas = "$SCRIPT_PATH\CacheFirmas"

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
        3- Establecer la firma HTML formateada.
    }

    
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

# Funcion Principal
function Set-Signatures {
    [CmdletBinding()]
    param(
        [Parameter]
        [string] $CorreoUsuario,

        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if($_.Length -eq 0){
                throw "El nombre de empresa ingresado no es valix`do."
            }
            return $true
        })]
        [string] $NombreEmpresa,

        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if((Test-Path $_) -eq $false){
                throw "La ubicacion de la plantilla HTML no es valida."
            }
            return $true
        })]
        [System.IO.FileInfo] $HTMLTemplate,

        [switch] $OWA=$false,

        [switch] $ODA =$false
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
    $FirmasFormateadas = Get-HTMLSignatures -Credencial $Creds -HTMLTemplate $HTMLTemplate -NombreEmpresa $NombreEmpresa -CorreoUsuario $CorreoUsuario
    
    
    if($OWA){
        Write-Host "Configurar firmas en Outlook Web App"
    }

    if($ODA){
        Write-Host "Configurar firmas en Outlook Desktop App"
    }

    #Si es la Outlook Web App, no se necesitan crear los archivos en Cache
    Set-OWASignatures -Firmas $FirmasFormateadas -Credencial $Creds

    Write-Host "Firmas configuradas correctamente"

}

$SCRIPT_PATH = Split-Path $MyInvocation.MyCommand.Path -Parent

# Llamar a la función con los parámetros requeridos.
try{
    Set-Signatures -HTMLTemplate ".\signature.htm" -NombreEmpresa "EMPRESA DE PRUEBAS" -OWA -ODA
}catch{
    $_.Exception.Message
}

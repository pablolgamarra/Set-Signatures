<#
.SYNOPSIS
Script para modificar las firmas de los usuarios de una organización en Outlook

.NOTES
Autor: Pablo Gamarra
Fecha de Creación: 2024-03-19
Última Modificación: 2024-03-19
Versión: 1.0
Github: https://github.com/pablolgamarra

.DESCRIPTION
Este script permite configurar las firmas de los usuarios de una organización en Outlook.
Se basa en la plantilla HTML proporcionada y los datos de los usuarios obtenidos de Microsoft Graph.

.PARAMETER CorreoUsuario
Correo electrónico del usuario para configurar su firma. Si se omite, se configurarán las firmas de todos los usuarios.

.PARAMETER NombredeNegocio
Nombre de la empresa para incluir en las firmas.

.PARAMETER PathPlantillaHtml
Ruta del archivo de plantilla HTML para las firmas.
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $NombredeNegocio,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $PathPlantillaHtml,

    [Parameter()]
    [string] $CorreoUsuario
)

begin{
    # Función para conectar a Microsoft Graph usando credenciales.
    function Connect-MgGraphViaCred {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.PSCredential] $Credential,

            [string] $Tenant = $_tenantDomain
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

    # Función para obtener las firmas HTML de los usuarios.
    function Get-FirmasHTML {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string] $NombredeNegocio,
    
            [Parameter(Mandatory = $true)]
            [string] $PathPlantillaHtml,
    
            [string] $CorreoUsuario
        )
    
        $FIRMA_HTML = Get-Content -Path $PathPlantillaHtml 
    
        # Conectar a servicios MG Graph y Exchange Online.
        try {
            Write-Verbose "Conectando a servicio MgGraph..."
            Connect-MgGraphViaCred
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
                NombreFirma     = "Firma-$NombredeNegocio-$($Usuario.DisplayName)"
            }
    
            $FirmasFormateadas = $FirmasFormateadas + $FirmaObj
        }
        return $FirmasFormateadas
    }

    # Función para configurar las firmas en la Outlook Web App (OWA).
    function Set-OWASignatures {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $True)]
            [System.Array] $Firmas
        )

        try {
            Write-Host "Conectando a servicio Exchange Online."
            Connect-ExchangeOnline
            Write-Host "Servicio Exchange Online conectado."
        }
        catch {
            Write-Host "Ocurrio un error al conectar al servicio de Exchange Online."
            Write-Host $_ -ErrorAction Stop
        }

        #Hacer la configuracion para desactivar las Roaming Signatures en la Organizacion (Necesario para configurar las firmas de esta manera)
        try {
            Write-Host "Configurar organizacion para posponer el uso de las firmas en Roaming en OWA"
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
            Write-Host "Firma configurada correctamente."
        }
    }

    # Función para configurar las firmas en la Outlook Desktop App (ODA).
    function Set-ODASignatures {
        [CmdLetBinding()]
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.PSCredential] $Credential
        )
    }
}
process{
    try {
        Write-Host "PathFirmaHTML: $PathPlantillaHtml"
        Write-Host "Nombre de la Empresa: $NombredeNegocio"
        Write-Host "Correo del Usuario: $CorreoUsuario"


        # Instalar módulos necesarios si no están instalados
        $ModulosRequeridos = @('ExchangeOnlineManagement', 'Microsoft.Graph', 'Az.Accounts')
        $ModulosInstalados = Get-InstalledModule | Select-Object -ExpandProperty Name
        $ModulosNoInstalados = $RequiredModules | Where-Object { $_ -notin $InstalledModules }

        if ($ModulosNoInstalados) {
            Install-Module -Name $ModulosNoInstalados -Force
        }

        # Obtener las firmas HTML de los usuarios
        $FirmasHTML = Get-FirmasHTML -PathPlantillaHtml $PathPlantillaHtml -NombredeNegocio $NombredeNegocio -CorreoUsuario $CorreoUsuario

        Write-Host "Cantidad Firmas Formateadas: $($FirmasHTML.Count)"
        Write-Host "Mostrando firmas formateadas:"
foreach ($firma in $FirmasHTML) {
    Write-Host "Nombre de usuario: $($firma.NombreUsuario)"
    Write-Host "Correo electrónico: $($firma.Correo)"
    Write-Host "Firma formateada: $($firma.FirmaFormateada)"
    Write-Host "Nombre de firma: $($firma.NombreFirma)"
    Write-Host "-------------------------"
}

        Set-OWASignatures -Firmas $FirmasHTML

        Write-Host "Final de configuracion de las firmas."
    }
    catch {
        Write-Error "Ocurrió un error: $($_.Exception.Message)"
    }
}

end{
    Disconnect-AzAccount
    Disconnect-MgGraph
    Disconnect-ExchangeOnline
}
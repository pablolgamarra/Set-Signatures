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

$User = "Virtual_TI@glymax.com"
$Pass = ConvertTo-SecureString -String "Virtual_TI" -AsPlainText -Force
$Cred = $CredsDominio = New-Object System.Management.Automation.PSCredential($User, $Pass)


Connect-MgGraphViaCred -credential $Cred

Get-MgDevice -All
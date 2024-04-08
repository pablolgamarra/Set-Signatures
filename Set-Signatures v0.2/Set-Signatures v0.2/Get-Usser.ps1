$UserName = "glymax\olam"
$Pass = ConvertTo-SecureString -String "SuperGlymax80023016" -AsPlainText -Force
$CredsAdminDom = New-Object System.Management.Automation.PSCredential($UserName, $Pass)   

function Get-UsuariosEquipos{
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $CredsAdminDom,
    [Parameter(Mandatory=$true)]
    [string] $filtro
)

    $ErrorActionPreference = 'SilentlyContinue'

    $Equipos = Get-ADComputer -Filter *
    $UsuariosEquipos = @()

    foreach($Equipo in $Equipos){
        Write-Host "Computadora:$($Equipo.Name)"
        $Computadora = $Equipo.Name
        $Usuario = (Get-WmiObject -Class win32_computersystem -ComputerName $Equipo.Name -Credential $CredsAdminDom).UserName

        $UsuarioEquipoObj = [PSCustomObject]@{
            NombreEquipo = $Computadora
            Usuario = $Usuario
        }

        $UsuariosEquipos += $UsuarioEquipoObj
    }
}


<#function Get-UsuariosEquipos{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $CredsAdminDom,
        [Parameter(Mandatory=$true)]
        [string] $filtro
    )
    
        $ErrorActionPreference = 'SilentlyContinue'
    
        $Equipos = Get-ADComputer -Filter filtro
        $UsuariosEquipos = @()
    
        foreach($Equipo in $Equipos){
        Write-Host "Computadora:$($Equipo.Name)"
        $Computadora = $Equipo.Name
        $Usuario = (Get-WmiObject -Class win32_computersystem -ComputerName $Equipo.Name -Credential $CredsAdminDom).UserName

        $UsuarioEquipoObj = [PSCustomObject]@{
            $CorreoUsuario = Get-ADUser -Filter {UserPrincipalName -eq ($Usuario -split '\')[1]}
            $UsuarioEquipoObj = [psobject]@{
                NombreEquipo = $Computadora
                Usuario = $Usuario
                CorreoUsuario = $CorreoUsuario
            }
    
            $UsuariosEquipos += $UsuarioEquipoObj
        }
        Write-Host $UsuariosEquipos
    }#>

    Write-Host $(Get-UsuariosEquipos -CredsAdminDom $CredsAdminDom -filtro '*')

    Get-ADUser -Filter {UserPrincipalName -eq ($Usuario -split '\\')[1]}

    Get-ADUser -Filter "SamAccountName -eq '$(($Usuario -split '\\')[1])'"
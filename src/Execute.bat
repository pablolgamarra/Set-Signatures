:: Script Batch para Facilitar la Ejecución del Comando
@echo off

title Configurar firmas en Outlook
color 9F

:inicio
echo 1- Aplicar firma a todos los usuarios
echo 2- Aplicar firma a un solo usuario
echo 3- Aplicar firma a un grupo de seguridad
set /p ELECCION=[1,2,3]?
cls

if "%ELECCION%"=="1" (
    goto ejecutarScript
)
if "%ELECCION%"=="2" (
    goto ingresarUsuario
)
if "%ELECCION%"=="3" (
    goto ejecutarScript
)

@echo on
:ingresarUsuario
set /p CORREO="Ingrese el correo del usuario cuyo correo desea colocar:"
echo correo: %CORREO%
goto ejecutarScript
exit/B 0

:ingresarGrupo
set /p CORREO="Ingrese el nombre del grupo de correos cuyas firmas desea configurar:"
echo grupo: %GRUPO%
goto ejecutarScript
exit/B 0

:ejecutarScript
set SCRIPTNAME=Set-Signatures.ps1
set SCRIPTDIR=%CD%\Scripts\%SCRIPTNAME%


if defined CORREO (
    echo "Configurar correo para un solo usuario: %CORREO%"
    PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%" -UserMail "%CORREO%" -GroupName ""
) else (
    if defined GRUPO (
        echo "Configurar correo para un grupo de seguridad: %GRUPO%"
        PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%" -UserMail "%CORREO%" -GroupName "%GRUPO%"
    )
    echo "Configurar correo para todos los usuarios"
    PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%" -UserMail "" -GroupName "%GRUPO%"
)

echo "Script Ejecutado"
pause
exit/B 0

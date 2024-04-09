:: Script Batch para Facilitar la Ejecución del Comando
@echo off

title Configurar firmas en Outlook
color 9F

:inicio
echo 1-Aplicar a todos los usuarios
echo 2-Aplicar a un solo usuario
set /p ELECCION=[1,2]?
cls

if "%ELECCION%"=="1" (
    goto ejecutarScript
) else (
    goto ingresarUsuario
)

@echo on
:ingresarUsuario
set /p CORREO="Ingrese el correo del usuario cuyo correo desea colocar:"
echo correo: %CORREO%
goto ejecutarScript
exit/B 0

:ejecutarScript
set SCRIPTNAME=Set-Signatures.ps1
set SCRIPTDIR=%CD%\%SCRIPTNAME%

if defined CORREO (
    echo "Configurar correo para un solo usuario: %CORREO%"
    PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%" -CorreoUsuario "%CORREO%"
) else (
    echo "Configurar correo para un todos los usuarios"
    PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%"
)

::

::PowerShell -ExecutionPolicy Bypass -File "%SCRIPTDIR%" -CorreoUsuario "%CORREO%"

echo "Script Ejecutado"
pause
exit/B 0

@echo off
set "script_folder=%~dp0\script"

if not exist "%script_folder%" (
    echo Error: Script folder "%script_folder%" not found.
    exit /b 1
)

echo Por favor espere, esto puede tardar un poco.
echo.

:LOOP
cd /d "%script_folder%"
cscript //nologo index.vbs
timeout /t 15 >nul
goto LOOP

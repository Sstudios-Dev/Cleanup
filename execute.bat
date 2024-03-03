@echo off
echo Welcome to the script!
echo.

set "script_folder=%~dp0\script"

if not exist "%script_folder%" (
    echo Error: Script folder "%script_folder%" not found.
    exit /b 1
)

echo Please wait, this may take a while.
echo.

:LOOP
cd /d "%script_folder%"
cscript //nologo index.vbs
timeout /t 15 >nul
goto LOOP

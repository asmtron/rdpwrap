@echo off
::
::
:: Please run this bat as an ADMIN by right click -> "Run As Administrator"
::
:: This is from https://github.com/asmtron/rdpwrap
::
::
if not exist "%~dp0RDPWInst.exe" goto :error
call "%~dp0uninstall.bat"
call "%ProgramFiles%\RDP Wrapper\helper\autoupdate__disable_autorun_on_startup.bat"
rd /s/q "%ProgramFiles%\RDP Wrapper"
echo ______________________________________________________________
echo.
echo Completed.
echo.
goto :anykey
:error
echo [-] Installer executable not found.
echo Please download from https://github.com/stascorp/rdpwrap/releases and extract all files from the downloaded package or check your anti-virus.
:anykey
pause
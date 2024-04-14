@echo off
::
::
:: Please run this bat as an ADMIN by right click -> "Run As Administrator"
::
:: This is from https://github.com/asmtron/rdpwrap
::
::
if not exist "%~dp0RDPWInst.exe" goto :error
call "%~dp0install.bat"
md "%ProgramFiles%\RDP Wrapper\helper\*.*"
copy "%~dp0\helper\*.*" "%ProgramFiles%\RDP Wrapper\helper\*.*"
copy "%~dp0autoupdate.bat" "%ProgramFiles%\RDP Wrapper\"
copy "%~dp0RDPWInst.exe" "%ProgramFiles%\RDP Wrapper\"
call "%ProgramFiles%\RDP Wrapper\autoupdate.bat"
call "%ProgramFiles%\RDP Wrapper\helper\autoupdate__enable_autorun_on_startup.bat"
: "%~dp0RDPConf.exe"
: "%~dp0RDPCheck.exe"
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
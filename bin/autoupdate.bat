<!-- : Begin of batch script
@echo off
setLocal EnableExtensions
setlocal EnableDelayedExpansion
::                                        _                   _
::              _                        | |      _          | |          _
::   ____ _   _| |_  ___  _   _ ____   _ | | ____| |_  ____  | | _   ____| |_
::  / _  | | | |  _)/ _ \| | | |  _ \ / || |/ _  |  _)/ _  ) | || \ / _  |  _)
:: ( ( | | |_| | |_| |_| | |_| | | | ( (_| ( ( | | |_( (/ / _| |_) ( ( | | |__
::  \_||_|\____|\___\___/ \____| ||_/ \____|\_||_|\___\____(_|____/ \_||_|\___)
::                             |_|
::
:: Automatic RDP Wrapper installer and updater v.1.4       asmtron (2025-10-23)
:: ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
:: Options:
::   -log        = redirect display output to the file autoupdate.log
::   -taskadd    = add autorun of autoupdate.bat on startup in schedule task
::   -taskremove = remove autorun of autoupdate.bat on startup in schedule task
::
:: Info:
::   The autoupdater first use and check the local rdpwrap.ini file.
::   If a new termsrv.dll is not supported in the local rdpwrap.ini file,
::   autoupdate uses the online rdpwrap.ini files from the community.
::
::   Autogeneration: v.1.4 or newer supports autogeneration of rdpwrap.ini entries
::
::  { Special thanks to binarymaster, sebaxakerhtc and all other contributors }
::
::
:: -----------------------------------------
:: Location of new/updated rdpwrap.ini files
:: -----------------------------------------
set rdpwrap_ini_update_github_1="https://raw.githubusercontent.com/asmtron/rdpwrap/master/res/rdpwrap.ini"
set rdpwrap_ini_update_github_2="https://raw.githubusercontent.com/sebaxakerhtc/rdpwrap.ini/master/rdpwrap.ini"
set rdpwrap_ini_update_github_3="https://raw.githubusercontent.com/affinityv/INI-RDPWRAP/master/rdpwrap.ini"
::set rdpwrap_ini_update_github_4="https://raw.githubusercontent.com/....Extra.4...."
::set rdpwrap_ini_update_github_5="https://raw.githubusercontent.com/....Extra.5...."
::
set autoupdate_version=1.4
set autoupdate_url="https://raw.githubusercontent.com/asmtron/rdpwrap/master/bin/autoupdate.bat"
set autoupdate_ver_url="https://raw.githubusercontent.com/asmtron/rdpwrap/master/bin/autoupdate.ver"
set autoupdate_bat="%~dp0autoupdate.bat"
set autoupdate_new_bat="%~dp0autoupdate_new.bat"
set autoupdate_log="%~dp0autoupdate.log"
set autoupdate_ver="%~dp0autoupdate.ver"
set RDPWInst_exe="%~dp0RDPWInst.exe"
set rdpwrap_dll="%~dp0rdpwrap.dll"
set rdpwrap_ini="%~dp0rdpwrap.ini"
set rdpwrap_ini_check=%rdpwrap_ini%
set rdpwrap_new_ini="%~dp0rdpwrap_new.ini"
set rdpwrap_ini_url=%rdpwrap_ini_update_github_1%
set github_location=1
set retry_network_check=0
set rdpwrap_installed=0
set version_check=0
set updated=0
::
set cvdump_url="https://raw.githubusercontent.com/asmtron/rdpwrap/master/bin/helper/cvdump.zip"
set cvdump_zip="%~dp0helper\cvdump.zip"
set cvdump_exe="%~dp0helper\cvdump.exe"
set pdblister_url="https://raw.githubusercontent.com/asmtron/rdpwrap/master/bin/helper/pdblister.zip"
set pdblister_zip="%~dp0helper\pdblister.zip"
set pdblister_exe="%~dp0helper\pdblister.exe"
set helper_path="%~dp0helper\"
set termsrv_dll="%systemroot%\system32\termsrv.dll"
set termsrv_pdb_dump="%~dp0helper\termsrv.pdb.dump"
set autogen=0
set autogen_ready=0
::
set arg_log=0
set arg_taskadd=0
set arg_taskremove=0
set arg_updated=0
set arg_nowaitnet=0
set arg_help=0
set arg_unknown=0
set args_count=0
set args=
::
:: --------------------------------
:: 1) check if admin
:: --------------------------------
fsutil dirty query %systemdrive% >nul
if not %errorlevel% == 0 (
  call :not_admin
  goto :eof
)
::
:: --------------------------------
:: 2) check arguments
:: --------------------------------
for %%i in (%*) do (
    set /a args_count=!args_count!+1
    set unknown=1
    if /i "%%i"=="-log"         set arg_log=1         &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-taskadd"     set arg_taskadd=1     &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-taskremove"  set arg_taskremove=1  &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-updated"     set arg_updated=1     &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-nowaitnet"   set arg_nowaitnet=1   &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-help"        set arg_help=1        &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="-?"           set arg_help=1        &  set unknown=0  &  set "args=!args! %%i"
    if /i "%%i"=="/?"           set arg_help=1        &  set unknown=0  &  set "args=!args! %%i"
    if !unknown!==1 (
        set arg_unknown=%%i
    )
)
if not %arg_unknown%==0 (
    call :header
    call :missing_arg
    goto :eof
)
if %arg_help%==1 (
    call :header
    call :usage_arg
    goto :eof
)
if %arg_taskadd%==1 (
    call :header
    echo [+] add autorun of %autoupdate_bat% on startup in the schedule task.
    schtasks /create /f /sc ONSTART /tn "RDP Wrapper Autoupdate" /tr "cmd.exe /C \"%~dp0autoupdate.bat\" -log" /ru SYSTEM /delay 0000:10
    powershell "$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries; Set-ScheduledTask -TaskName 'RDP Wrapper Autoupdate' -Settings $settings"
    goto :eof
)
if %arg_taskremove%==1 (
    call :header
    echo [-] remove autorun of %autoupdate_bat% on startup in the schedule task^^!
    schtasks /delete /f /tn "RDP Wrapper Autoupdate"
    goto :eof
)
if %arg_updated%==1 (
    set updated=1
    copy /y %autoupdate_new_bat% %autoupdate_bat%
) else (
    if exist %autoupdate_new_bat% del %autoupdate_new_bat%
)
if %arg_log%==1 (
    echo %autoupdate_bat% output from %date% at %time% > %autoupdate_log%
    if %args_count% == 1 (
        call %autoupdate_bat% >> %autoupdate_log%
    ) else (
        call %autoupdate_bat% %args:-log=% >> %autoupdate_log%
    )
    goto :eof
)
call :header
if %arg_nowaitnet%==1 (
    echo [+] Set option: do not wait up to 5 minutes for a valid internet connection
    set /a retry_network_check=29
)
::
:: --------------------------------
:: 3) check if "RDPWInst.exe" exist
:: --------------------------------
if not exist %RDPWInst_exe% (
    call :error_install
    goto :eof
)
:: ----------------------------------
:: 4) check if TermService is running
:: ----------------------------------
sc queryex "TermService"|find "STATE"|find /v "RUNNING" >nul&&(
    echo [-] TermService NOT running^^!
    call :install
)||(
    echo [+] TermService running.
)
:: ------------------------------------------
:: 5) check if listener session rdp-tcp exist
:: ------------------------------------------
set rdp_tcp_session=""
set rdp_tcp_session_id=0
if exist %systemroot%\system32\query.exe (
    for /f "tokens=1-2* usebackq" %%a in (
        `query session rdp-tcp`
    ) do (
        set rdp_tcp_session=%%a
        set /a rdp_tcp_session_id=%%b 2>nul
    )
) else (
    for /f "tokens=2* usebackq" %%a in (
        `reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v "fDenyTSConnections" 2^>nul`
    ) do (
        if "%%a"=="REG_DWORD" (
            set rdp_tcp_session=AllowTSConnection
            if "%%b"=="0x0" (set rdp_tcp_session_id=1)
        )
    )
)
if %rdp_tcp_session_id%==0 (
    echo [-] Listener session rdp-tcp NOT found^^!
    call :install
) else (
    echo [+] Found listener session: %rdp_tcp_session% ^(ID: %rdp_tcp_session_id%^).
)
:: -----------------------------------------
:: 6) check if rdpwrap.dll exist in registry
:: -----------------------------------------
reg query "HKLM\SYSTEM\CurrentControlSet\Services\TermService\Parameters" /f "rdpwrap.dll" >nul&&(
    echo [+] Found windows registry entry for "rdpwrap.dll".
)||(
    echo [-] NOT found windows registry entry for "rdpwrap.dll"^^!
    if %rdpwrap_installed%==0 (
        call :install
    )
)
:: -----------------------------------
:: 7) check if rdpwrap.dll file exists
:: -----------------------------------
if exist %rdpwrap_dll% (
    echo [+] Found file: %rdpwrap_dll%
) else (
    echo [-] File NOT found: %rdpwrap_dll%^^!
    if %rdpwrap_installed%==0 (
        call :install
    )
)
:: ------------------------------
:: 8) check if rdpwrap.ini exists
:: ------------------------------
if exist %rdpwrap_ini% (
    echo [+] Found file: %rdpwrap_ini%.
) else (
    echo [-] File NOT found: %rdpwrap_ini%^^!
    if %rdpwrap_installed%==0 (
        call :install
    )
)
:: ----------------------------------------------------
:: 9) get file version of %windir%\System32\termsrv.dll
:: ----------------------------------------------------
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:getFileVersion %termsrv_dll%`
) do (
    set termsrv_dll_ver=%%a
)
if "%termsrv_dll_ver%"=="" (
    echo [x] Error on getting the file version of %termsrv_dll%^^!
    goto :eof
) else (
    echo [+] Installed "termsrv.dll" version: %termsrv_dll_ver%.
)
:: -------------------------------------------------------------------------------------------
:: 10) check if installed file version is different to the last saved file version in registry
:: -------------------------------------------------------------------------------------------
echo [*] Read last "termsrv.dll" version from the windows registry...
for /f "tokens=2* usebackq" %%a in (
    `reg query "HKEY_LOCAL_MACHINE\SOFTWARE\RDP-Wrapper\Autoupdate" /v "termsrv.dll" 2^>nul`
) do (
    set last_termsrv_dll_ver=%%b
)
if "%last_termsrv_dll_ver%"=="%termsrv_dll_ver%" (
    echo [+] Current "termsrv.dll v.%termsrv_dll_ver%" same as last "termsrv.dll v.%last_termsrv_dll_ver%".
) else (
    echo [-] Current "termsrv.dll v.%termsrv_dll_ver%" different from last "termsrv.dll v.%last_termsrv_dll_ver%"^^!
    echo [*] Update current "termsrv.dll" version to the windows registry...
    reg add "HKEY_LOCAL_MACHINE\SOFTWARE\RDP-Wrapper\Autoupdate" /v "termsrv.dll" /t REG_SZ /d "%termsrv_dll_ver%" /f
    if %rdpwrap_installed%==0 (
        call :install
    )
)
:: ----------------------------------------------------------------
:: 11) check if installed termsrv.dll version exists in rdpwrap.ini
:: ----------------------------------------------------------------
call :check_dll_version
::
::
:: F I N I S H
::
echo.
goto :eof
::
:: ----------------------------------------------
:: Search termsrv.dll version in file rdpwrap.ini
:: ----------------------------------------------
:check_dll_version
if exist %rdpwrap_ini_check% (
    echo [*] Start searching [%termsrv_dll_ver%] version entry in file %rdpwrap_ini_check%...
    findstr /c:"[%termsrv_dll_ver%]" %rdpwrap_ini_check% >nul&&(
        echo [+] Found "termsrv.dll" version entry [%termsrv_dll_ver%] in file %rdpwrap_ini_check%.
        echo [.] ..............................................
        echo [*] RDP Wrapper seems to be up-to-date and working...
    )||(
        echo [-] NOT found "termsrv.dll" version entry [%termsrv_dll_ver%] in file %rdpwrap_ini_check%^^!
        if not "!rdpwrap_ini_update_github_%github_location%!" == "" (
            set rdpwrap_ini_url=!rdpwrap_ini_update_github_%github_location%!
            call :update
            call :check_dll_version
        ) else (
            if %autogen%==0 (
                set autogen=1
                set /a github_location=1
                set rdpwrap_ini_url=%rdpwrap_ini_update_github_1%
                echo [.] ......................................................................................................
                echo [~] Start the process of automatically generating entries for "termsrv.dll" with version [%termsrv_dll_ver%]
                call :update
                call :check_dll_version
            ) else (
                echo [*] NOT found "termsrv.dll" version entry [%termsrv_dll_ver%] in file %rdpwrap_ini_check%^^!
                echo [-] Give up - "termsrv.dll" version [%termsrv_dll_ver%] is not yet supported at the moment^^!
                echo [~] ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                echo [X] RDP Wrapper will probably not work correctly ...
            )
        )
    )
) else (
    echo [-] File NOT found: %rdpwrap_ini_check%.
    echo [*] Give up - Please check if Antivirus/Firewall blocks the file %rdpwrap_ini_check%^^!
)
goto :eof
::
:: -----------------------------------------------------
:: Install RDP Wrapper (exactly uninstall and reinstall)
:: -----------------------------------------------------
:install
echo.
echo [*] Uninstall and reinstall RDP Wrapper...
echo.
if exist %rdpwrap_dll% set rdpwrap_force_uninstall=1
if exist %rdpwrap_ini% set rdpwrap_force_uninstall=1
if "%rdpwrap_force_uninstall%"=="1" (
    echo [*] Set windows registry entry for "rdpwrap.dll" to force uninstall...
    reg.exe add "HKLM\SYSTEM\CurrentControlSet\Services\TermService\Parameters" /f /v ServiceDll /t REG_EXPAND_SZ /d %rdpwrap_dll%
)
call :update
set rdpwrap_installed=1
::call :setNLA
goto :eof
::
:: -------------------
:: Restart RDP Wrapper
:: -------------------
:restart
echo.
echo [*] Restart RDP Wrapper with new ini (uninstall and reinstall)...
echo.
%RDPWInst_exe% -u
if exist %rdpwrap_new_ini% (
    echo.
    echo [*] Use latest downloaded rdpwrap.ini from GitHub...
    echo     -^> %rdpwrap_ini_url%
    echo       -^> %rdpwrap_new_ini%
    echo         -^> %rdpwrap_ini%
    echo [+] copy %rdpwrap_new_ini% to %rdpwrap_ini%...
    copy %rdpwrap_new_ini% %rdpwrap_ini%
    echo.
) else (
    echo [x] ERROR - File %rdpwrap_new_ini% is missing ^^!
)
%RDPWInst_exe% -i
::call :setNLA
goto :eof
::
:: -----------------------------------
:: check if online access is available
:: -----------------------------------
:update
echo [*] check network connectivity...
call :netcheck
call :checkversion
goto :eof
::
:: -----------------------------------
:: check if online access is available
:: -----------------------------------
:netcheck
ping -n 1 google.com>nul
if errorlevel 1 (
    echo [.] Wait for network connection is available...
    ping 127.0.0.1 -n 11>nul
    set /a retry_network_check=retry_network_check+1
    :: wait for a maximum of 5 minutes
    if %retry_network_check% LSS 30 call netcheck
)
goto :eof
::
:: ---------------------------------------------------
:: check if new version of autoupdate.bat is available
:: ---------------------------------------------------
:checkversion
if %version_check%==1 (
  call :download
  goto :eof
)
set version_check=1
echo.
echo [*] get version info of autoupdate.bat from GitHub...
echo     -^> %autoupdate_ver_url%
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:saveWebBinary %autoupdate_ver_url% %autoupdate_ver%`
) do (
    set "download_status=%%a"
)
if not "%download_status%"=="-1" (
    echo [-] FAILED to get version info of autoupdate.bat from GitHub...
    echo [*] Please check you internet connection/firewall and try again^^!
    call :download
    goto :eof
)
for /f "usebackq tokens=*" %%a in (%autoupdate_ver%) do (set "autoupdate_online_version=%%a")
echo [+] Successfully get autoupdate 'v.%autoupdate_online_version%' info from GitHhub.
if %autoupdate_version% GEQ %autoupdate_online_version% (
    echo [*] autoupdate 'v.%autoupdate_online_version%' is up to date
    call :download
    goto :eof
)
echo [+] New version 'v.%autoupdate_online_version%' of autoupdate.bat available^^!
echo [*] Download new version of autoupdate.bat from GitHub...
echo     -^> %autoupdate_url%
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:saveWebBinary %autoupdate_url% %autoupdate_new_bat%`
) do (
    set "download_status=%%a"
)
if "%download_status%"=="-1" (
    echo [+] Restart with the new autoupdate 'v.%autoupdate_online_version%' ...
    if %updated%==1 (
        call :download
    )
    if "%args%" == "" (
        cmd /c %autoupdate_new_bat% -updated
    ) else (
        cmd /c %autoupdate_new_bat% -updated %args%
    )
    exit
) else (
    echo [-] FAILED to download from GitHub latest version to %autoupdate_bat%^^!
    echo [*] Please check you internet connection/firewall and try again^^!
    call :download
    goto :eof
)
::
:: ---------------------------------------------------------------
:: Download an newer up-to-date version of rdpwrap.ini from GitHub
:: ---------------------------------------------------------------
:download
echo.
if %autogen%==1 (
    echo _____________________________
    echo [~] Autogeneration attempt: %github_location%
) else (
    echo ______________
    echo [~] Attempt: %github_location%
)
echo [*] Download latest version of rdpwrap.ini from GitHub...
echo     -^> %rdpwrap_ini_url%
set /a github_location=github_location+1
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:saveWebBinary %rdpwrap_ini_url% %rdpwrap_new_ini%`
) do (
    set "download_status=%%a"
)
if "%download_status%"=="-1" (
    echo [+] Successfully downloaded from GitHhub latest version to %rdpwrap_new_ini%.
    set rdpwrap_ini_check=%rdpwrap_new_ini%
    if %autogen%==1 (
        echo [~] Try to autogenerate new "termsrv.dll" [%termsrv_dll_ver%] entries for %rdpwrap_new_ini%
        call :setup_autogen
    )
    call :restart
) else (
    echo [-] FAILED to download from GitHub latest version to %rdpwrap_new_ini%^^!
    echo [*] Please check you internet connection/firewall and try again^^!
)
goto :eof
::
:: -----------------------------------------------------
:: Setup autogen (download debug symbols of termsrv.dll)
:: -----------------------------------------------------
:setup_autogen
if %autogen_ready%==1 (
    call :start_autogen
) else (
    echo [*] Setup autogenerator...
    if not exist %pdblister_exe% (
        call :download_pdblister
    )
    if not exist %cvdump_exe% (
        call :download_cvdump
    )
    if exist %pdblister_exe% (
        if exist %cvdump_exe% (
            call :generate_symbol
            if exist %termsrv_pdb_dump% (
                call :start_autogen
            )
        )
    )
)
goto :eof
::
:: ------------------
:: Download pdblister
:: ------------------
:download_pdblister
echo [+] Download the "Symbol Checker tool pdblister v0.0.4" from Github...
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:saveWebBinary %pdblister_url% %pdblister_zip%`
) do (
    set "pdblister_status=%%a"
)
if "%pdblister_status%"=="-1" (
    echo [+] Successfully downloaded the "Symbol Checker tool" file:
    echo     -^> %pdblister_zip%
) else (
    echo [-] FAILED to download the "Symbol Checker tool" file %pdblister_zip%
    echo [*] Please check you internet connection/firewall and try again^^!
    goto :eof
)
echo [+] Extract pdblister.exe from zip file: %pdblister_zip%
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:unzip %pdblister_zip% %helper_path%`
) do (
    set "unzip_pdblister=%%a"
)
if "%unzip_pdblister%"=="-1" (
    echo [+] Successfully extracted the "Symbol Checker tool" file: %pdblister_exe%
) else (
    echo [-] FAILED to extract the Symbol Checker tool" file: %pdblister_exe%
)
goto :eof
::
:: ---------------
:: Download cvdump
:: ---------------
:download_cvdump
echo [+] Download "Microsoft Debugging Information Dumper Version 14.00.23611" from Github...
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:saveWebBinary %cvdump_url% %cvdump_zip%`
) do (
    set "cvdump_status=%%a"
)
if "%cvdump_status%"=="-1" (
    echo [+] Successfully downloaded the "Microsoft Debugging Information Dumper" file:
    echo     -^> %cvdump_zip%
) else (
    echo [-] FAILED to download the "Microsoft Debugging Information Dumper" file %cvdump_zip%
    echo [*] Please check you internet connection/firewall and try again^^!
    goto :eof
)
echo [+] Extract cvdum.exe from zip file: %cvdump_zip%
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:unzip %cvdump_zip% %helper_path%`
) do (
    set "unzip_vdump=%%a"
)
if "%unzip_vdump%"=="-1" (
    echo [+] Successfully extracted the "Microsoft Debugging Information Dumper" file: %cvdump_exe%
) else (
    echo [-] FAILED to extract the "Microsoft Debugging Information Dumper" file: %cvdump_exe%
)
goto :eof
::
:: -----------------------------------
:: Download and generate debug symbols
:: -----------------------------------
:generate_symbol
echo [+] Download and generate debug symbols dump for "termsrv.dll" version [%termsrv_dll_ver%]
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:pdbDownload %termsrv_dll%`
) do (
    set "download_pdb=%%a"
)
if "%download_pdb%"=="-1" (
    echo [+] Successfully downloaded and generated the debug symbol file:
    echo     -^> %termsrv_pdb_dump%
    set autogen_ready=1
) else (
    echo [-] FAILED to download and generate the the debug symbol file: %termsrv_pdb_dump%
    echo [*] Please check you internet connection/firewall and try again^^!
)
goto :eof
::
:: ------------------------------------------------
:: Autogenerate termsrv.dll entries for rdpwrap.ini
:: ------------------------------------------------
:start_autogen
echo [+] Start generating new entries for "termsrv.dll" version [%termsrv_dll_ver%]
echo [.] ........................................................................
for /f "tokens=* usebackq" %%a in (
    `cscript //nologo "%~f0?.wsf" //job:GenerateTermsrvINIEntries %termsrv_dll_ver% %rdpwrap_new_ini% %termsrv_dll% %termsrv_pdb_dump%`
) do (
    set "autogen_status=%%a"
)
if "%autogen_status%"=="-1" (
    echo [+] Successfully generated new entries for "termsrv.dll" with version [%termsrv_dll_ver%] in the file: %rdpwrap_new_ini%
    echo     -^> "termsrv.dll" [%termsrv_dll_ver%]
    echo       -^> %termsrv_pdb_dump%
    echo         -^> %rdpwrap_new_ini%
) else (
    echo [-] FAILED to generate new entries for "termsrv.dll" version [%termsrv_dll_ver%] in the file: %rdpwrap_new_ini%
    echo [*] Reasons for the failure:
    echo     - missing ".autogen" templates in %rdpwrap_new_ini%
    echo     - wrong ".autogen" templates in %rdpwrap_new_ini%
    echo     - incompatible ".autogen" templates in %rdpwrap_new_ini%
)
goto :eof
::
:: --------------------------------
:: Set Network Level Authentication
:: --------------------------------
:setNLA
echo [*] Set Network Level Authentication in the windows registry...
reg.exe add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v SecurityLayer /t reg_dword /d 0x2 /f
reg.exe add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v MinEncryptionLevel /t reg_dword /d 0x2 /f
reg.exe add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v UserAuthentication /t reg_dword /d 0x1 /f
goto :eof
::
:: ---------------
:: No admin rights
:: ---------------
:not_admin
color 0e
echo ___________________________________
echo [x] ERROR - No Administrator Rights
echo [*] This script must be run as administrator to work properly^^!
echo     ^<Please use 'right click' on this batch file and select "Run As Administrator"^>
echo.
timeout 60
goto :eof
::
:: ----------------
:: Error on Install
:: ----------------
:error_install
echo [-] RDP Wrapper installer executable (RDPWInst.exe) not found^^!
echo Please extract all files from the downloaded RDP Wrapper package or check your Antivirus.
echo.
goto :eof
::
:: ---------
:: Arguments
:: ---------
:header
echo _________________________________________________
echo Automatic RDP Wrapper installer and updater v.%autoupdate_version%
echo.
echo ^<check if the RDP Wrapper is up-to-date and working^>
echo.
goto :eof
:usage_arg
echo [*] Supported argments are:
echo     -log         =  redirect display output to the file autoupdate.log
echo     -taskadd     =  add autorun of autoupdate.bat on startup in the schedule task
echo     -taskremove  =  remove autorun of autoupdate.bat on startup in the schedule task
echo     -nowaitnet   =  do not wait up to 5 minutes for a valid internet connection
echo.
echo     -help        =  shows this help
goto :eof
:missing_arg
echo [x] Unknown argument specified: "%arg_unknown%"
call :usage_arg
goto :eof
::
:: -------
:: E X I T
:: -------
exit /b
::
--- : Begin of wsf script code --- saveWebBinary, getFileVersion, unzip, pdbDownload, AUTOGEN.GenerateTermsrvINIEntries  --->
<package>
  <job id="saveWebBinary"><script language="VBScript">
    Set args = WScript.Arguments
    WScript.Echo SaveWebBinary(args(0), args(1))
    Wscript.Quit
    Function SaveWebBinary(strUrl, strFile)
      On Error Resume Next
      Const adTypeBinary = 1
      Const adSaveCreateOverWrite = 2
      Const ForWriting = 2
      Dim result, web, varByteArray, strData, strBuffer, lngCounter, ado
      result = false
      'Download the file with any available object
      Err.Clear
      Set web = Nothing
      Set web = CreateObject("WinHttp.WinHttpRequest.5.1")
      If web Is Nothing Then Set web = CreateObject("WinHttp.WinHttpRequest")
      If web Is Nothing Then Set web = CreateObject("MSXML2.ServerXMLHTTP")
      If web Is Nothing Then Set web = CreateObject("Microsoft.XMLHTTP")
      web.Open "GET", strURL, False
      web.Send
      If Err.Number = 0 And web.Status = "200" Then
        varByteArray = web.ResponseBody
        'Now save the file with any available method
        Set ado = Nothing
        Set ado = CreateObject("ADODB.Stream")
        If ado Is Nothing Then
          Dim fs, ts, max
          Set fs = CreateObject("Scripting.FileSystemObject")
          Set ts = fs.OpenTextFile(strFile, ForWriting, True)
          strData = ""
          strBuffer = ""
          max = UBound(varByteArray)
          For lngCounter = 0 To max
            ts.Write Chr(255 And Ascb(Midb(varByteArray,lngCounter + 1, 1)))
          Next
          ts.Close
        Else
          ado.Type = adTypeBinary
          ado.Open
          ado.Write varByteArray
          ado.SaveToFile strFile, adSaveCreateOverWrite
          ado.Close
        End If
        result = True
      End If
      Set web = Nothing
      SaveWebBinary = result
    End Function
  </script></job>
  <job id="getFileVersion"><script language="VBScript">
    Set args = WScript.Arguments
    Set fso = CreateObject("Scripting.FileSystemObject")
    WScript.Echo fso.GetFileVersion(args(0))
    Wscript.Quit
  </script></job>
  <job id="unzip"><script language="VBScript">
    Set args = WScript.Arguments
    WScript.Echo unzip(args(0), args(1))
    Wscript.Quit
    Function unzip(ZipFile, ExtractTo)
      On Error Resume Next
      Dim fso, objShell, FilesInZip, item, result
      result = False
      Set fso = CreateObject("Scripting.FileSystemObject")
      If FSO.FolderExists(ExtractTo) = False Then
        'If the extraction location does not exist create it.
         fso.CreateFolder(ExtractTo)
      End If
      If fso.FileExists(ZipFile) = True Then
        'Extract the files from zip archive
        set objShell = CreateObject("Shell.Application")
        set FilesInZip=objShell.NameSpace(ZipFile).Items
        objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
        For Each item in FilesInZip
          If fso.FileExists(fso.BuildPath(ExtractTo, item)) = True Then
            result = True
          End If
        Next
      End If
      Set fso = Nothing
      Set objShell = Nothing
      unzip = result
    End Function
  </script></job>
  <job id="pdbDownload"><script language="VBScript">
    Set args = WScript.Arguments
    WScript.Echo pdbDownload(args(0))
    Wscript.Quit
    Function pdbDownload(file)
      On Error Resume Next
      Dim result, pdb, symbol, WshShell, pdblister
      result = false
      pdb = ""
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.CurrentDirectory = ".\helper"
      Set pdblister = WshShell.Exec("pdblister.exe download-single SRV*.\Symbols*https://msdl.microsoft.com/download/symbols " + file + " json")
      If IsObject (pdblister) Then
        Do While pdblister.Status = 0
            WScript.Sleep 100
        Loop
        Dim aStdOut: aStdOut = Split(pdblister.StdOut.ReadAll, ",", -1, 1)
        If UBound(aStdOut) = 2 Then
          If InStr(LCase(aStdOut(2)),"success") Then
            Dim aSymbol: aSymbol = Split(aStdOut(1), chr(34), -1, 1)
            If (UBound(aSymbol) = 4) Then
              If InStr(LCase(aSymbol(1)),"path") Then
                pdb = Replace(aSymbol(3), "\\", "\")
              End If
            End if
          End If
        End If
      End If
      If pdb <> "" Then
        Dim dump, fso
        dump = Right(pdb, Len(pdb) - InStrRev(pdb, "\")) + ".dump"
        WshShell.Run "cmd /c cvdump.exe " + pdb + " > " + dump, 1, True
        Set fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FileExists(dump)) Then
          result = true
        End If
        Set fso = Nothing
      End If
      Set pdblister = Nothing
      Set WshShell = Nothing
      pdbDownload = result
    End Function
  </script></job>
  <job id="GenerateTermsrvINIEntries"><script language="VBScript">
    Dim args: Set args = WScript.Arguments
    Dim oAuogen: Set oAutoGen = New AUTOGEN
    WScript.Echo oAutoGen.GenerateTermsrvINIEntries(args(0), args(1), args(2), args(3))
    Set oAutoGen = Nothing
    Wscript.Quit
    Class AUTOGEN
      Public Function GenerateTermsrvINIEntries(fileVersion, iniPath, termsrvPath, symbolFile)
        'On Error Resume Next
        Dim architecture: architecture = 0
        Dim result: result = False
        Dim aVer: aVer = Split(fileVersion, ".", -1, 1)
        If UBound(aVer) = 3 Then
          Dim pdbdump: Set pdbdump = (New TXTFile)(symbolFile)
          If pdbdump.isValid Then
            If (pdbdump.searchTextLine("amd64\termsrv.exp") <> "") Then
              architecture = 64
            ElseIf (pdbdump.searchTextLine("i386\termsrv.exp") <> "") Then
              architecture = 32
            End If
            If architecture <> 0 Then
              Dim rdpwrapini: Set rdpwrapini = (New INIFile)(iniPath)
              If rdpwrapini.isValid Then
                result = generateINIEntriesTermsrv(fileVersion, getSections_autogen(aVer, ".autogen"), pdbdump, architecture, rdpwrapini, termsrvPath)
                If result = True Then
                  result = generateINIEntriesTermsrvSLInit(fileVersion + "-SLInit", getSections_autogen(aVer, ".autogen-SLInit"), pdbdump, architecture, rdpwrapini)
                  rdpwrapini.Save
                End If
              End If
              Set rdpwrapini = Nothing
            End If
          End If
          Set pdbdump = Nothing
        End If
        GenerateTermsrvINIEntries = result
      End Function
      Private Function generateINIEntriesTermsrv(fileSection, aSection, pdbdump, architecture, rdpwrapini, termsrvPath)
        Dim result: result = False
        Dim termsrvdll: termsrvdll = False
        Dim aKeyPatch, aKeyOffset, aKeyCode, keyPatch, KeyOffset, KeyCode
        If architecture = 64 Then
          aKeyPatch = Array("LocalOnlyPatch.x64", "SingleUserPatch.x64", "DefPolicyPatch.x64", "SLInitHook.x64")
          aKeyOffset = Array("LocalOnlyOffset.x64", "SingleUserOffset.x64", "DefPolicyOffset.x64", "SLInitOffset.x64")
          aKeyCode = Array("LocalOnlyCode.x64", "SingleUserCode.x64", "DefPolicyCode.x64", "SLInitFunc.x64")
        ElseIf architecture = 32 then
          aKeyPatch = Array("LocalOnlyPatch.x86", "SingleUserPatch.x86", "DefPolicyPatch.x86", "SLInitHook.x86")
          aKeyOffset = Array("LocalOnlyOffset.x86", "SingleUserOffset.x86", "DefPolicyOffset.x86", "SLInitOffset.x86")
          aKeyCode = Array("LocalOnlyCode.x86", "SingleUserCode.x86", "DefPolicyCode.x86", "SLInitFunc.x86")
        Else
          generateINIEntriesTermsrv = False
          Exit Function
        End If
        Dim funcAddress, patchAddress, offsetType, offset, pos, posMax, sec, secMax
        pos = 0
        posMax = UBound(aKeyPatch)
        sec = 0
        secMax = UBound(aSection)
        For pos = 0 To posMax
          For sec = 0 To secMax
            keyPatch = rdpwrapini.GetValue(aSection(sec), aKeyPatch(pos))
            KeyOffset = rdpwrapini.GetValue(aSection(sec), aKeyOffset(pos))
            KeyCode = rdpwrapini.GetValue(aSection(sec), aKeyCode(pos))
            If (keyPatch <> "") And (KeyOffset <> "") And (KeyCode <> "") Then
              funcAddress = getFunctionAddress(pdbdump, keyPatch)
              If funcAddress <> 0 Then
                offsetType = Mid(KeyOffset, 1, 1)
                If offsetType = "*" Then
                  If Not IsObject(termsrvdll) then
                    Set termsrvdll = (New BINFile)(termsrvPath)
                    If termsrvdll.isValid = False Then
                      generateINIEntriesTermsrv = False
                      Exit function
                    End If
                  End If
                  patchAddress = searchBinaryDataLocation(funcAddress, termsrvdll, Mid(KeyOffset, 2, Len(KeyOffset)-1))
                ElseIf offsetType = "+" Then
                  On Error Resume Next
                  offset = CLng("&H" + Mid(KeyOffset, 2, Len(KeyOffset)-1))
                  On Error GoTo 0
                  patchAddress = funcAddress + offset
                ElseIf offsetType = "-" Then
                  On Error Resume Next
                  offset = CLng("&H" + Mid(KeyOffset, 2, Len(KeyOffset)-1))
                  On Error GoTo 0
                  patchAddress = funcAddress - offset
                Else
                  On Error Resume Next
                  offset = CLng("&H" + KeyOffset)
                  On Error GoTo 0
                  patchAddress = funcAddress + offset
                End If
                If patchAddress <> 0 Then
                  rdpwrapini.SetValue fileSection, aKeyPatch(pos), "1"
                  rdpwrapini.SetValue fileSection, aKeyOffset(pos), Hex(patchAddress)
                  rdpwrapini.SetValue fileSection, aKeyCode(pos), KeyCode
                  result = True
                End If
              End If
            End If
          Next
        Next
        Set termsrvdll = Nothing
        generateINIEntriesTermsrv = result
      End Function
      Private Function generateINIEntriesTermsrvSLInit(fileSection, aSection, pdbdump, architecture, rdpwrapini)
        Dim result: result = False
        Dim aKey
        If architecture = 64 Then
          aKey = Array("bInitialized.x64", "bServerSku.x64", "lMaxUserSessions.x64", "bAppServerAllowed.x64", "bRemoteConnAllowed.x64", "bMultimonAllowed.x64", "ulMaxDebugSessions.x64", "bFUSEnabled.x64")
        ElseIf architecture = 32 then
          aKey = Array("bInitialized.x86", "bServerSku.x86", "lMaxUserSessions.x86", "bAppServerAllowed.x86", "bRemoteConnAllowed.x86", "bMultimonAllowed.x86", "ulMaxDebugSessions.x86", "bFUSEnabled.x86")
        Else
          generateINIEntriesTermsrvSLInit = False
          Exit Function
        End If
        Dim func, funcAddress, pos, posMax, sec, secMax
        pos = 0
        posMax = UBound(aKey)
        sec = 0
        secMax = UBound(aSection)
        For pos = 0 To posMax
          For sec = 0 To secMax
            func = rdpwrapini.GetValue(aSection(sec), aKey(pos))
              If func <> "" Then
                funcAddress = getFunctionAddress(pdbdump, func)
                If funcAddress <> 0 Then
                  rdpwrapini.SetValue fileSection, aKey(pos), Hex(funcAddress)
                result = True
                End If
              End If
            Next
        Next
        generateINIEntriesTermsrvSLInit = result
      End Function
      Private Function getSections_autogen(aVer, addition)
        Dim sVer2, sArray2, sArraySize2, max2, sVer3, sArray3, sArraySize3, max3, aSections, pos, size, i , jj, str
        sVer2 = aVer(2)
        sArraySize2 = Len(sVer2) - 1
        ReDim sArray2(sArraySize2)
        For i = 0 to sArraySize2
          sArray2(i) = Mid(sVer2, i + 1, 1)
        Next
        sVer3 = aVer(3)
        sArraySize3 = Len(sVer3) - 1
        ReDim sArray3(sArraySize3)
        For i = 0 to sArraySize3
          sArray3(i) = Mid(sVer3, i + 1, 1)
        Next
        max2 = len(aVer(2))
        max3 = len(aVer(3))
        ReDim aSections(max2 + max3)
        size = UBound(sArray2)
        pos = 0
        For i=0 To max2
          str = ""
          For jj=size To 0 Step -1
            If jj >= pos Then
              str = "*" + str
            Else
              str = sArray2(jj) + str
            End If
          Next
          aSections(pos) = aVer(0) + "." + aVer(1) + "." + str + addition
          pos = pos + 1
        Next
        size = UBound(sArray3)
        pos = max2
        For i=0 To max3
          str = ""
          For jj=size To 0 Step -1
            If jj >= (pos-max2) Then
              str = "*" + str
            Else
              str = sArray3(jj) + str
            End If
          Next
          aSections(pos) = aVer(0) + "." + aVer(1) + "." + aVer(2) + "." + str + addition
          pos = pos + 1
        Next
        getSections_autogen = aSections
      End Function
      Private Function searchBinaryDataLocation(funcAddress, termsrvdll, dataString)
        Dim address, size, byteArray, byteArrayMax, arrData, pos, max, i, jj, b
        address = 0
        size = Len(dataString) / 2
        pos = 1
        max = size-1
        redim arrData(max)
        On Error Resume Next
        For i = 0 to max
          arrData(i) = CInt("&H" + Mid(dataString, pos, 2))
          pos = pos + 2
        Next
        On Error GoTo 0
        byteArray = termsrvdll.GetBytes()
        byteArrayMax = UBound(byteArray) - 1
        pos = 0
        For i = funcAddress To byteArrayMax
          b = AscB(MidB(byteArray, i + 1, 1))
          If b = arrData(0) Then
            For jj = 1 To max
              b = AscB(MidB(byteArray, i + 1 + jj, 1))
              If b = arrData(jj) Then
                If jj = max Then
                  address = i + jj
                  Exit For
                End If
              Else
                Exit For
              End If
            Next
          End If
          If address <> 0 Then Exit For
        Next
        searchBinaryDataLocation = address
      End Function
      Private Function getFunctionAddress(pdbdump, func)
        Dim S_PUB32, aS_PUB32, addressOffset, addressSection, virtualAddressText, aVirtualAddress, virtualAddress
        virtualAddress = 0
        S_PUB32 = pdbdump.searchTextLine(func)
        aS_PUB32 = SplitMultiDelimiters(S_PUB32, Array(":","[","]"), -1, 1, "|")
        If (UBound(aS_PUB32) > 3 And aS_PUB32(0) = "S_PUB32") Then
          On Error Resume Next
          addressOffset = CInt("&H" + aS_PUB32(2))
          addressSection = CLng("&H" + aS_PUB32(3))
          virtualAddressText = pdbdump.SearchTextBlockLastLine(Array("SECTION HEADER #" + Hex(addressOffset), "virtual address"))
          virtualAddressText = Trim(virtualAddressText)
          aVirtualAddress = Split(virtualAddressText, " ", -1, 1)
          virtualAddress = CLng("&H" + aVirtualAddress(0))
          On Error GoTo 0
        End If
        getFunctionAddress = virtualAddress + addressSection
      End function
      Private Function SplitMultiDelimiters(text, aDelimiters, count, compare, placeholder)
        If IsArray(aDelimiters) Then
          Dim size, i
          size = UBound(aDelimiters)
          For i = 0 To size
            text = Replace(text, aDelimiters(i), placeholder)
          Next
        End If
        SplitMultiDelimiters = Split(text, placeholder, count, compare)
      End Function
    End Class
    Const TypeBinary = 1, ForReading = 1, ForWriting = 2, ForAppending = 8
    Class BINFile
      Private FSO, inStream, FileName, FileBytes, Valid
      Private Function GetFileContents()
        If Not FileExists() Then
          GetFileContents = false
          Valid = false
        Else
          inStream.LoadFromFile(FileName)
          GetFileContents = inStream.Read()
          Valid = true
        End If
      End Function
      Private Sub WriteFileContents(ByVal MyContents)
        binaryStream.Write MyContents
        'Save binary data to disk
        binaryStream.SaveToFile FileName, ForWriting
      End Sub
      Public Default Function Init(MyFileName)
        Valid = False
        Set FSO  = CreateObject("Scripting.FileSystemObject")
        set inStream = WScript.CreateObject("ADODB.Stream")
        inStream.Open
        inStream.type= TypeBinary
        FileName = MyFileName
        Load
        Set Init = Me
      End Function
      Public Sub Load
        FileBytes = GetFileContents
      End Sub
      Public Sub Save
        WriteFileContents(FileBytes)
      End Sub
      Public Function isValid()
        isValid = Valid
      End Function
      Public Function GetFileName()
        GetFileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
      End Function
      Public Function GetFilePath()
        GetFilePath = Left(FileName, InStrRev(FileName, "\"))
      End Function
      Public Function FileExists()
        FileExists = FSO.FileExists(FileName)
      End Function
      Public Function GetBytes()
        GetBytes = FileBytes
      End Function
      Public Sub SetBytes(byteArray)
        FileBytes = byteArray
      End sub
    End Class
    Class TXTFile
      Private FSO, FileName, FileContents, Valid
      Private Function GetFileContents()
        If Not FileExists() Then
          GetFileContents = vbNullString
          Valid = false
        ElseIf FileIsEmpty() Then
          GetFileContents = vbNullString
          Valid = False
        Else
          GetFileContents = FSO.OpenTextFile(FileName, ForReading).ReadAll
          Valid = true
        End If
      End Function
      Private Sub WriteFileContents(ByVal MyContents)
        Dim FileStream: Set FileStream = FSO.OpenTextFile(FileName, ForWriting, True)
        FileStream.Write MyContents
        FileStream.Close()
      End Sub
      Public Default Function Init(MyFileName)
        Valid = false
        Set FSO  = CreateObject("Scripting.FileSystemObject")
        FileName = MyFileName
        Load
        Set Init = Me
      End Function
      Public Sub Load
        FileContents = GetFileContents
      End Sub
      Public Sub Save
        WriteFileContents(FileContents)
      End Sub
      Public Function isValid()
        isValid = Valid
      End Function
      Public Function GetFileName()
        GetFileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
      End Function
      Public Function GetFilePath()
        GetFilePath = Left(FileName, InStrRev(FileName, "\"))
      End Function
      Public Function FileExists()
        FileExists = FSO.FileExists(FileName)
      End Function
      Public Function FileIsEmpty()
        FileIsEmpty = FSO.OpenTextFile(FileName).AtEndOfStream
      End Function
      Public Function SearchTextLine(searchText)
        Dim result: result = ""
        Dim pos: pos = InStr(1, FileContents, searchText, vbTextCompare)
        If pos > 0 Then
          Dim lineStart: lineStart = InStrRev(FileContents, vbCrLf, pos, vbTextCompare) + 2
          Dim lineEnd: lineEnd = InStr(pos, FileContents, vbCrLf, vbTextCompare)
          result = Mid(FileContents, lineStart, lineEnd - lineStart)
        End If
        searchTextLine = result
      End Function
      Function SearchTextBlockLastLine(aSearchText)
        Dim result: result = ""
        If IsArray(aSearchText) Then
          Dim ln: ln = 0
          Dim max: max = UBound(aSearchText)
          Dim pos: pos = InStr(1, FileContents, aSearchText(ln), vbTextCompare)
          Do While pos > 0
            If ln = max Then
              Dim lineStart: lineStart = InStrRev(FileContents, vbCrLf, pos, vbTextCompare) + 2
              Dim lineEnd: lineEnd = InStr(pos, FileContents, vbCrLf, vbTextCompare)
              result = Mid(FileContents, lineStart, lineEnd - lineStart)
              Exit Do   ' Abort loop when text is found
            Else
              ln = ln + 1
              pos = InStr(pos, FileContents, aSearchText(ln), vbTextCompare)
              End if
          Loop
        End if
        SearchTextBlockLastLine = result
      End Function
    End Class
    Class INIFile
      Private FSO, FileName, FileContents, Valid
      Private Function GetFileContents()
        If Not FileExists() Then
          GetFileContents = vbNullString
          Valid = False
        ElseIf FileIsEmpty() Then
          GetFileContents = vbNullString
          Valid = False
        Else
          GetFileContents = FSO.OpenTextFile(FileName, ForReading).ReadAll
          Valid = true
        End If
      End Function
      Private Function GetSectionContents(MySection, PosSection, PosEndSection)
        Dim SectionContents: SectionContents = vbNullString
        PosSection = InStr(1, FileContents, "[" & MySection & "]", vbTextCompare)
        If PosSection > 0 Then
          PosEndSection = InStr(PosSection, FileContents, vbCrLf & "[")
          If PosEndSection = 0 Then PosEndSection = Len(FileContents) + 1
          SectionContents = Mid(FileContents, PosSection, PosEndSection - PosSection)
        End If
        GetSectionContents = SectionContents
      End Function
      Private Sub WriteFileContents(ByVal MyContents)
        Dim FileStream: Set FileStream = FSO.OpenTextFile(FileName, ForWriting, True)
        FileStream.Write MyContents
        FileStream.Close()
      End Sub
      Private Function ExtractValue(ByVal MyFrom, ByVal MyStart, ByVal MyEnd)
        Dim PosS: PosS = InStr(1, MyFrom, MyStart, 1)
        If PosS > 0 Then
          PosS = PosS + Len(MyStart)
          Dim PosE: PosE = InStr(PosS, MyFrom, MyEnd, 1)
          If PosE = 0 Then PosE = InStr(PosS, MyFrom, vbCrLf, 1)
          If PosE = 0 Then PosE = Len(MyFrom) + 1
          ExtractValue = Mid(MyFrom, PosS, PosE - PosS)
        Else
          ExtractValue = vbNullString
        End If
      End Function
      Public Default Function Init(MyFileName)
        Valid = False
        Set FSO  = CreateObject("Scripting.FileSystemObject")
        FileName = MyFileName
        Load
        Set Init = Me
      End Function
      Public Sub Load
        FileContents = GetFileContents
      End Sub
      Public Sub Save
        WriteFileContents(FileContents)
      End Sub
      Public Function isValid()
        isValid = Valid
      End Function
      Public Function GetFileName()
        GetFileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
      End Function
      Public Function GetFilePath()
        GetFilePath = Left(FileName, InStrRev(FileName, "\"))
      End Function
      Public Function FileExists()
        FileExists = FSO.FileExists(FileName)
      End Function
      Public Function FileIsEmpty()
        FileIsEmpty = FSO.OpenTextFile(FileName).AtEndOfStream
      End Function
      Public Function GetSections()
        Dim SectionsRegExp: Set SectionsRegExp = New RegExp
        ' Matches a [Section] on its own line. Could be at the very beginning of the file,
        ' in the middle of the file, or at the very end of the file (an empty [Section]).
        SectionsRegExp.Pattern = "([\r\n]\[|^\[)([^\]]*)(\][\r\n]|\]$)"
        SectionsRegExp.Global = True  ' Matches all occurrences, not just the first one.
        Dim SectionMatches: Set SectionMatches = SectionsRegExp.Execute(FileContents)
        Dim Sections: Sections = Array()
        Dim Index
        If SectionMatches.Count > 0 Then
          For Index = 0 To SectionMatches.Count - 1
            ReDim Preserve Sections(Index)
            Sections(Index) = SectionMatches.Item(Index).SubMatches(1)
          Next
        End If
        GetSections = Sections
      End Function
      Public Function GetKeys(MySection)
        Dim PosSection: PosSection = 0
        Dim PosEndSection: PosEndSection = 0
        Dim SectionContents: SectionContents = GetSectionContents(MySection, PosSection, PosEndSection)
        Dim KeysRegExp: Set KeysRegExp = New RegExp
        ' Matches a key= on its own line; captures the name of the key.
        KeysRegExp.Pattern = "[\r\n]{1,2}([^=]*)="
        KeysRegExp.Global = True
        Dim KeyMatches: Set KeyMatches = KeysRegExp.Execute(SectionContents)
        Dim Keys: Keys = Array()
        Dim Index
        If KeyMatches.Count > 0 Then
          For Index = 0 To KeyMatches.Count - 1
            ReDim Preserve Keys(Index)
            Keys(Index) = KeyMatches.Item(Index).SubMatches(0)
          Next
        End If
        GetKeys = Keys
      End Function
      Public Function GetValue(MySection, MyKeyName)
        Dim Value
        Dim PosSection: PosSection = 0
        Dim PosEndSection: PosEndSection = 0
        Dim SectionContents: SectionContents = GetSectionContents(MySection, PosSection, PosEndSection)
        Dim re : Set re = New RegExp
        re.Global = True
        re.Pattern = " +"  'remove one or more spaces
        SectionContents = re.Replace(SectionContents, "")
        Set re = Nothing
        If InStr(1, SectionContents, vbCrLf & MyKeyName & "=", vbTextCompare) > 0 Then
          Value = ExtractValue(SectionContents, vbCrLf & MyKeyName & "=", vbCrLf)
        End If
        GetValue = Value
      End Function
      Public Sub SetValue(MySection, MyKeyName, MyValue)
        Dim PosSection: PosSection = 0
        Dim PosEndSection: PosEndSection = 0
        Dim OldSectionContents: OldSectionContents = GetSectionContents(MySection, PosSection, PosEndSection)
        If OldSectionContents <> vbNullString Then
          Dim KeyName, Line, Found, NewSectionContents
          Dim re : Set re = New RegExp
          re.Global = True
          re.Pattern = " +"  'remove one or more spaces
          OldSectionContents = re.Replace(OldSectionContents, "")
          Set re = Nothing
          OldSectionContents = Split(OldSectionContents, vbCrLf)
          KeyName = LCase(MyKeyName & "=") ' Temp variable to find a key.
          ' Copy each line over; if the key matches, change its value first.
          For Each Line In OldSectionContents
            If LCase(Left(Line, Len(KeyName))) = KeyName Then
              Line = MyKeyName & "=" & MyValue
              Found = True
            End If
            If Line <> "" Then
              NewSectionContents = NewSectionContents & Line & vbCrLf
            End If
          Next
          If IsEmpty(Found) Then ' Key not found.
            ' Append it to the [Section].
            NewSectionContents = NewSectionContents & MyKeyName & "=" & MyValue & vbCrLf
          End If
          ' Combine pre-section, new section, and post-section data.
          FileContents = Left(FileContents, PosSection-1) & NewSectionContents & Mid(FileContents, PosEndSection)
        Else
          ' Section doesn't exist.
          ' If the file doesn't already end in a new line, and if the file isn't empty...
          If Right(FileContents, 2) <> vbCrLf And Len(FileContents) > 0 Then
            ' Add a new line to the end of the file
            FileContents = FileContents & vbCrLf
          End If
          ' Add section data at the end of file contents.
          FileContents = FileContents & vbCrLf & "[" & MySection & "]" & vbCrLf & MyKeyName & "=" & MyValue
        End If
      End Sub
    End Class
  </script></job>
</package>





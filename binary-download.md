Automatic RDP Wrapper installer and updater v.1.4
=================================
` asmtron (2025-10-23)`


Info:
  The autoupdater first use and check the local rdpwrap.ini file.
  If a new termsrv.dll is not supported in the local rdpwrap.ini file,
  autoupdate uses the online rdpwrap.ini files from the community.


  **Autogeneration: v.1.4 or newer supports autogeneration of rdpwrap.ini entries**


  ` { Special thanks to binarymaster, sebaxakerhtc and all other contributors }`



 RDP Wrapper with Autoupdater installation
 =========================================

 It is recommended to use the installer as it performs all necessary steps automatically.


###  a) For automatic installation

[ Download: [https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4-Installer.exe](https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4-Installer.exe) ]
       
[ Download: [https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4-Installer.zip](https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4-Installer.zip) ]

    1. Run "RDPWrap-with-Autoupdate-v1.4-Installer.exe" (or newer) to install RDP Wrapper with Autoupdater


###  b) For manual installation

[ Download: [https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4.zip](https://github.com/asmtron/rdpwrap/releases/download/v1.4/RDPWrap-with-Autoupdate-v1.4.zip) ]

    1. Create the directory: "%ProgramFiles%\RDP Wrapper"
       (usually C:\Program Files\RDP Wrapper)


    2. Set in your Antivirus or Windows Defender an exclusion on the folder "%ProgramFiles%\RDP Wrapper" to prevent the deletion of RDP Wrapper files...


    3. Copy the files from the zip-archive "RDPWrap-with-Autoupdate-v1.4.zip" (or newer) to the directory: "%ProgramFiles%\RDP Wrapper"

       DO NOT use other directories to install/extract the RDP Wrapper files.
       USE ONLY the directory: "%ProgramFiles%\RDP Wrapper" (usually C:\Program Files\RDP Wrapper)


    4. To enable autorun of autoupdate.bat on system startup, run the following helper batch file as administrator:

       "%ProgramFiles%\RDP Wrapper\helper\autoupdate__enable_autorun_on_startup.bat"


    5. Now you can use the autoupdate batch file to install and update the RDP Wrapper. Please run autoupdate.bat as administrator:

      "%ProgramFiles%\RDP Wrapper\autoupdate.bat"

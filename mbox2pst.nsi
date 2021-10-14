!define AppName "mbox2pst"

; The name of the installer
Name "${AppName} Installer"

; The file to write
OutFile "setup_${AppName}.exe"


; Request application privileges for Windows Vista
RequestExecutionLevel admin

; Build Unicode installer
Unicode True

; The default installation directory
InstallDir $PROGRAMFILES\${AppName}

;--------------------------------

; Pages

Page directory
Page instfiles

;--------------------------------

; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File ${AppName}.exe
  File /r Redemption
  CreateShortcut  "$DESKTOP\${AppName}.lnk"               "$INSTDIR\${AppName}.exe" "--gui" "$INSTDIR\${AppName}.exe"
  CreateDirectory "$SMPROGRAMS\${AppName}"
  CreateShortCut  "$SMPROGRAMS\${AppName}\${AppName}.lnk" "$INSTDIR\${AppName}.exe" "--gui" "$INSTDIR\${AppName}.exe"
  ExecShell "" "$INSTDIR\Redemption\Install.exe"

SectionEnd ; end the section


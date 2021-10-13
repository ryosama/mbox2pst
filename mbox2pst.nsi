; The name of the installer
Name "mbox2pst Installer"

; The file to write
OutFile "setup_mbox2pst.exe"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

; Build Unicode installer
Unicode True

; The default installation directory
InstallDir $PROGRAMFILES\mbox2pst

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
  File mbox2pst.exe
  File Redemption\*.*
  CreateShortcut "$DESKTOP\mbox2pst.lnk" "$INSTDIR\mbox2pst.exe" "--gui" "$INSTDIR\mbox2pst.exe"

SectionEnd ; end the section

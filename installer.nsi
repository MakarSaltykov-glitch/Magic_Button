; Magic Button Windows Installer
; Created with NSIS

!include "MUI2.nsh"

; General
Name "Magic Button"
OutFile "MagicButton-Installer.exe"
InstallDir "$PROGRAMFILES\MagicButton"
InstallDirRegKey HKCU "Software\MagicButton" ""

; Request admin rights
RequestExecutionLevel admin

; UI
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "Russian"

Section "Install"
  SetOutPath "$INSTDIR"
  
  ; Copy main executable
  File "dist\MagicButton.exe"
  
  ; Create Start Menu shortcuts
  CreateDirectory "$SMPROGRAMS\Magic Button"
  CreateShortCut "$SMPROGRAMS\Magic Button\Magic Button.lnk" "$INSTDIR\MagicButton.exe"
  CreateShortCut "$SMPROGRAMS\Magic Button\Uninstall.lnk" "$INSTDIR\uninstall.exe"
  
  ; Create Desktop shortcut
  CreateShortCut "$DESKTOP\Magic Button.lnk" "$INSTDIR\MagicButton.exe" "" "$INSTDIR\MagicButton.exe" 0
  
  ; Create uninstaller
  WriteUninstaller "$INSTDIR\uninstall.exe"
  WriteRegStr HKCU "Software\MagicButton" "" "$INSTDIR"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\MagicButton.exe"
  Delete "$INSTDIR\uninstall.exe"
  RMDir "$INSTDIR"
  
  Delete "$SMPROGRAMS\Magic Button\Magic Button.lnk"
  Delete "$SMPROGRAMS\Magic Button\Uninstall.lnk"
  RMDir "$SMPROGRAMS\Magic Button"
  
  Delete "$DESKTOP\Magic Button.lnk"
  
  DeleteRegKey HKCU "Software\MagicButton"
SectionEnd

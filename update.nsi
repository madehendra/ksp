;NSIS Modern User Interface version 1.70
;e-DeskSoftware Update Installer Script.
;Written by Made Hendra <mailto:made.hendra@gmail.com> <http://www.ibelog.tk>

;---------------------
;Include Modern UI

  !include "MUI.nsh"

;--------------------------------
;General

  ;Name and file
  Name "iUSPD"
  OutFile "update.exe"
  ShowInstDetails show
  BrandingText /TRIMLEFT "iUpdate Setup by Made Hendra"
  WindowIcon on

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_WELCOME
  !insertmacro MUI_PAGE_LICENSE "Licensedata.txt"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_INSTFILES
  !insertmacro MUI_PAGE_FINISH
 
;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "INDONESIAN"

;--------------------------------
;Installer Sections

Section "Main Update" SecMain
  ; rename old file
  ; Path=C:\Program Files\e-DeskSoftware\iPOS
  ; Rename "$PROGRAMFILES\e-DeskSoftware\iPOS\sweety.exe" "$PROGRAMFILES\e-DeskSoftware\iPOS\sweety.old.exe"

  ; Set output path to the installation directory.
  ; SetOutPath "$PROGRAMFILES\e-DeskSoftware\iPOS"
  
  ; Put file there
  ; File sweety.exe

  ; update untuk modul sysinfo
  SetOutPath $SYSDIR
  File "d:\MyOCX\MenuExtended.dll"
  
SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_SecMain ${LANG_INDONESIAN} "Main Update!!"

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} $(DESC_SecMain)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END
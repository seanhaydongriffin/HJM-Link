;!include nsDialogs.nsh
;!include LogicLib.nsh


; example1.nsi
;
; This script is perhaps one of the simplest NSIs you can make. All of the
; optional settings are left to their default settings. The installer simply 
; prompts the user asking them where to install, and drops a copy of example1.nsi
; there. 

XPStyle on

;--------------------------------

; The name of the installer
Name "HJM Link"

; The file to write
OutFile "setup.exe"

; The default installation directory
InstallDir "$PROGRAMFILES32\HJM Link"

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
  File "HJM Link.exe"
  ;File "HJM Link.chm"
  File *.ico
  File "curl.exe"

  CreateDirectory "$SMPROGRAMS\HJM Link"
  CreateShortCut "$SMPROGRAMS\HJM Link\HJM Link.lnk" "$INSTDIR\HJM Link.exe"

SectionEnd ; end the section

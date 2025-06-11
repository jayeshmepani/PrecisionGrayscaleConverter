; Precision Grayscale Converter Installer
; NSIS Installer Script

!define APPNAME "Precision Grayscale Converter"
!define COMPANYNAME "Precision Tools"
!define DESCRIPTION "Professional grayscale conversion with 9-decimal precision"
!define VERSIONMAJOR 1
!define VERSIONMINOR 0
!define VERSIONBUILD 0

!define INSTALLSIZE 15000 ; Estimate in KB

RequestExecutionLevel admin
InstallDir "$PROGRAMFILES\${APPNAME}"
LicenseData "license.txt"
Name "${APPNAME}"
Icon "icon.ico"
outFile "PrecisionGrayscaleConverter_Setup_${VERSIONMAJOR}.${VERSIONMINOR}.${VERSIONBUILD}.exe"

!include LogicLib.nsh

page license
page directory
page instfiles

!macro VerifyUserIsAdmin
UserInfo::GetAccountType
pop $0
${If} $0 != "admin"
    messageBox mb_iconstop "Administrator rights required!"
    setErrorLevel 740
    quit
${EndIf}
!macroend

function .onInit
    setShellVarContext all
    !insertmacro VerifyUserIsAdmin
functionEnd

section "install"
    ; Files to install
    setOutPath $INSTDIR
    file "PrecisionGrayscaleConverter.exe"
    file "README.md"
    
    ; Create uninstaller
    writeUninstaller "$INSTDIR\uninstall.exe"
    
    ; Start Menu
    createDirectory "$SMPROGRAMS\${APPNAME}"
    createShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\PrecisionGrayscaleConverter.exe"
    createShortCut "$SMPROGRAMS\${APPNAME}\Uninstall.lnk" "$INSTDIR\uninstall.exe"
    
    ; Desktop shortcut
    createShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\PrecisionGrayscaleConverter.exe"
    
    ; Registry information for add/remove programs
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "QuietUninstallString" "$\"$INSTDIR\uninstall.exe$\" /S"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation" "$\"$INSTDIR$\""
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$\"$INSTDIR\PrecisionGrayscaleConverter.exe$\""
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${COMPANYNAME}"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpLink" "${HELPURL}"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "URLUpdateInfo" "${UPDATEURL}"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "URLInfoAbout" "${ABOUTURL}"
    writeRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSIONMAJOR}.${VERSIONMINOR}.${VERSIONBUILD}"
    writeRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMajor" ${VERSIONMAJOR}
    writeRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMinor" ${VERSIONMINOR}
    writeRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
    writeRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1
    writeRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "EstimatedSize" ${INSTALLSIZE}
    
    ; File associations
    writeRegStr HKCR ".png\OpenWithList\PrecisionGrayscaleConverter.exe" "" ""
    writeRegStr HKCR ".jpg\OpenWithList\PrecisionGrayscaleConverter.exe" "" ""
    writeRegStr HKCR ".jpeg\OpenWithList\PrecisionGrayscaleConverter.exe" "" ""
    writeRegStr HKCR ".tiff\OpenWithList\PrecisionGrayscaleConverter.exe" "" ""
    writeRegStr HKCR ".bmp\OpenWithList\PrecisionGrayscaleConverter.exe" "" ""
    
sectionEnd

function un.onInit
    SetShellVarContext all
    MessageBox MB_OKCANCEL "Remove ${APPNAME}?" IDOK next IDCANCEL abort
    abort:
        quit
    next:
        !insertmacro VerifyUserIsAdmin
functionEnd

section "uninstall"
    ; Remove files
    delete "$INSTDIR\PrecisionGrayscaleConverter.exe"
    delete "$INSTDIR\README.md"
    delete "$INSTDIR\uninstall.exe"
    rmDir "$INSTDIR"
    
    ; Remove start menu
    delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
    delete "$SMPROGRAMS\${APPNAME}\Uninstall.lnk"
    rmDir "$SMPROGRAMS\${APPNAME}"
    
    ; Remove desktop shortcut
    delete "$DESKTOP\${APPNAME}.lnk"
    
    ; Remove registry entries
    deleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
    
    ; Remove file associations
    deleteRegKey HKCR ".png\OpenWithList\PrecisionGrayscaleConverter.exe"
    deleteRegKey HKCR ".jpg\OpenWithList\PrecisionGrayscaleConverter.exe"
    deleteRegKey HKCR ".jpeg\OpenWithList\PrecisionGrayscaleConverter.exe"
    deleteRegKey HKCR ".tiff\OpenWithList\PrecisionGrayscaleConverter.exe"
    deleteRegKey HKCR ".bmp\OpenWithList\PrecisionGrayscaleConverter.exe"
    
sectionEnd
Attribute VB_Name = "mMain"
Option Explicit
Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

Public quiteSTATE As axMemory

Private Sub Main()
Set quiteSTATE = New axMemory
quiteSTATE.OpenMemory "obav_quitestate"
If App.EXEName = "Uninstall" Then
    If MsgBox("Are You Sure to Uninstall ObAV from your System ?", vbQuestion + vbYesNo) = vbYes Then
        quiteSTATE.Poke "1234567890"
        tunggON 14
        unInstALL False
    End If
Else
    If isInstalled = False Then
        Fwizard.Show
    Else
        cek4Upgrade
    End If
End If
End Sub
Private Sub cek4Upgrade()
Dim ve As Long, vetem As Long
ve = GetFileVersi(appFullpatH)
vetem = GetFileVersi(fixP(ObAVdir) & "Uninstall.exe")
If ve > vetem Then
 If MsgBox("Upgrade ObAV With New Version ?", vbYesNo + vbQuestion) = vbYes Then
    quiteSTATE.Poke "1234567890"
    tunggON 14
    unInstALL True
    Fwizard.Show
 End If
Else
 MsgBox "ObAV Was Installed!", vbInformation + vbSystemModal
End If
End Sub
Private Sub unInstALL(diam As Boolean)
On Error Resume Next
Dim starprogDIR As String
        If isADMIN = False Then
        MsgBox "Please Run As ADMINISTRATOR", vbCritical + vbSystemModal
        Exit Sub
        End If
    ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-stop"
    tunggON 10
    RegObavExt False
    KillProcess fixP(GetSpecFolder(WINDOWS_DIR)) & "explorer.exe"
    ShellEXE fixP(ObAVdir) & "ObAVSvc.exe", "-uninstall"
    tunggON 10
    ShellEXE fixP(GetSpecFolder(WINDOWS_DIR)) & "explorer.exe"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", ChrW$(&H2126) & "bAV Guard"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayName"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "UninstallString"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "DisplayIcon"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "Publisher"
    DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ObAV", "HelpLink"
    
    KillProcess fixP(ObAVdir) & "ObAV.exe"
    KillProcess fixP(ObAVdir) & "ObAV.exe"
    KillProcess fixP(ObAVdir) & "ObavScanner.exe"
    KillProcess fixP(ObAVdir) & "ObavScanner.exe"
    KillProcess fixP(ObAVdir) & "ObavGuard.exe"
    KillProcess fixP(ObAVdir) & "ObavGuard.exe"

    MoveFileExW StrPtr(fixP(ObAVdir) & "Uninstall.exe"), StrPtr(fixP(obavTMP) & "Uninstall.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "KprocMon.sys"), StrPtr(fixP(obavTMP) & "KprocMon.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavExt.dll"), StrPtr(fixP(obavTMP) & "ObavExt.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavSpk.sys"), StrPtr(fixP(obavTMP) & "ObavSpk.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavScanner.exe"), StrPtr(fixP(obavTMP) & "ObavScanner.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavLoader.exe"), StrPtr(fixP(obavTMP) & "ObavLoader.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavSvc.exe"), StrPtr(fixP(obavTMP) & "ObavSpk.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "msvbvm60.dll"), StrPtr(fixP(obavTMP) & "msvbvm60.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObavGuard.exe"), StrPtr(fixP(obavTMP) & "ObavGuard.tmp"), MOVEFILE_REPLACE_EXISTING
    MoveFileExW StrPtr(fixP(ObAVdir) & "ObAV.exe"), StrPtr(fixP(obavTMP) & "ObAV.tmp"), MOVEFILE_REPLACE_EXISTING
    
    MoveFileExW StrPtr(fixP(obavTMP) & "Uninstall.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "KprocMon.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavExt.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavSpk.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavScanner.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavLoader.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavSpk.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "msvbvm60.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObavGuard.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    MoveFileExW StrPtr(fixP(obavTMP) & "ObAV.tmp"), 0&, MOVEFILE_DELAY_UNTIL_REBOOT
    
    Kill fixP(ObAVdir) & "ObAV.cfg"
    RmDir ObAVdir
    Kill fixP(GetSpecFolder(DEKSTOP_PATH)) & "ObAV.lnk"
    starprogDIR = fixP(GetSpecFolder(STAR_PROGRAMS)) & "ObAV AntiVirus"
    Kill fixP(starprogDIR) & "ObAV.lnk"
    Kill fixP(starprogDIR) & "Uninstall.lnk"
    RmDir starprogDIR
    '------hapus kernel driver------'
    Call RemovNotiv(True)
    Call ClirServisSPK
    Call delKproc
    '-------------------------------'
    If diam = False Then
        MsgBox "Uninstall Completed", vbInformation + vbSystemModal
    End If
End Sub

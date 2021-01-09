Attribute VB_Name = "bsHOOK"
Option Explicit
'oj4nBL4NK SHELLHOOK Receiver
'Created By oj4nBL4NK           '-Special Thank's For
    '-JAM 7.42 AM                   '-ALLAH SWT
    '-Tgl 03-09-2010                '-Ortu
    '-Pada Bulan R4M4DH4N           '-Indo-code.com
'Jangan Memodifikasi/Merubah Kode ini sedikitpun

'REVISI :
'tanggal 2 juni 2011 => penggantian DLL ke SYS

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_READ = &H10
Private Const WM_DESTROY = &H2
Dim hWndHOOK As Long
Dim old_proc As Long
Dim hookPAUSE As Boolean
Public Function cWINHOOK() As Boolean
hWndHOOK = CreateWindowEx(0, "#32770", "obav_22091993", 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
ShowWindow hWndHOOK, 0
old_proc = SetWindowLong(hWndHOOK, GWL_WNDPROC, AddressOf MyWndProc)
SetNotiv
hookPAUSE = False
cWINHOOK = True
End Function
Public Sub unHook()
RemovNotiv
SetWindowLong hWndHOOK, GWL_WNDPROC, old_proc
DestroyWindow hWndHOOK
End Sub
Public Sub pauseHOOK(pause As Boolean)
If pause = True Then
    hookPAUSE = True
    bsTray.ShellTrayBalloonTipShow fRtpSystem.hwnd, 1, verobAV, "ObAV EXE Guard NonActivated"
Else
    hookPAUSE = False
    bsTray.ShellTrayBalloonTipShow fRtpSystem.hwnd, 1, verobAV, "ObAV EXE Guard Activated"
End If
End Sub
Private Function MyWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim paT As String
Dim pjG As Long
Dim Hproc As Long
Dim Modules As Long
Dim cbNeeded As Long
Dim RET As Long
On Error GoTo exI
    Select Case message
        Case WM_DESTROY
            DestroyWindow (hwnd)
        Case WM_USER + 878&
            If hookPAUSE = True Then GoTo quiT
lopinG_1:
            Call Sleep(1)
            Hproc = PamzOpenProcess(wParam, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
            If Hproc > 1 Then
            RET = EnumProcessModules(Hproc, Modules, 1, cbNeeded)
            'MsgBox RET & " " & Modules
            PamzCloseHandle Hproc
            If RET < 1 Then GoTo lopinG_1
            End If
            PamzSuspendResumeProcessThreads wParam, False
            tundaTHREAD wParam, True
            paT = PathByPID(wParam)
            If Len(paT) < 3 Then GoTo exI
            pjG = scE.lenFileEX(paT)
            If pjG <= 2 Then GoTo exI
            If pjG > 1750000 Then GoTo exI
            CekVirus paT, pjG, wParam, , 3
            guardPATH.Poke paT
exI:
            tundaTHREAD wParam, False
            PamzSuspendResumeProcessThreads wParam, True
quiT:
    End Select
    MyWndProc = DefWindowProc(hwnd, message, wParam, lParam)
End Function

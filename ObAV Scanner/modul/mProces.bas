Attribute VB_Name = "bsProces"
Option Explicit

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const MAX_PATH As Long = 260

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_TERMINATE = &H1

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const token_All_Access As Long = 983551

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Type LUID
   lowpart As Long
   highpart As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExW Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As Long, ByVal nSize As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

 Public ProcessId(150) As Long
 Public pPath(150) As String
 Public nama(150) As String
 Public jmlProcess As Long
Public Sub ShellEXE(path As String, Optional argU As String)
ShellExecuteW 0&, StrPtr("open"), StrPtr(path), StrPtr(argU), StrPtr(vbNullString), 1
End Sub
Public Sub tunggON(tim As Long)
Dim i As Long
For i = 0 To tim
Sleep 100
DoEvents
Next i
End Sub
Public Sub List_Process()
 Dim w As Long
    jmlProcess = 1
    Dim hSnapshot As Long, uProcess As PROCESSENTRY32
        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    w = Process32First(hSnapshot, uProcess)
    Do While w
        ProcessId(jmlProcess) = uProcess.th32ProcessID
        pPath(jmlProcess) = PathByPID(ProcessId(jmlProcess))
        nama(jmlProcess) = uProcess.szexeFile
        w = Process32Next(hSnapshot, uProcess)
    jmlProcess = jmlProcess + 1
    Loop
    jmlProcess = jmlProcess - 1
    CloseHandle hSnapshot
    w = vbNull
End Sub

Public Function PathByPID(pId As Long) As String
    Dim cbNeeded As Long, Modules(1 To 200) As Long
    Dim RET As Long, ModuleName As String
    Dim nSize As Long, hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pId)
    If hProcess <> 0 Then
        RET = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
        If RET <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            RET = GetModuleFileNameExW(hProcess, Modules(1), StrPtr(ModuleName), nSize)
            PathByPID = VBA.Left$(ModuleName, RET)
        End If
    End If
    RET = CloseHandle(hProcess)
    If PathByPID = "" Then PathByPID = ""
    If VBA.Left(PathByPID, 4) = "\??\" Then PathByPID = ""
    If VBA.Left(PathByPID, 12) = "\SystemRoot\" Then PathByPID = ""
    cbNeeded = vbNull
    RET = vbNull
    ModuleName = vbNullString
    nSize = vbNull
    hProcess = vbNull
    Erase Modules
End Function
Public Function exPIDLV(lisPROC As ucListView, cmd As Long)
Dim j As Long, i As Long, pId As Long
j = lisPROC.ListItems.Count
For i = 1 To j
If lisPROC.ListItems.Item(j).Checked = True Then
pId = lisPROC.ListItems.Item(j).SubItem(3).Text
If pId = GetCurrentProcessId Then
Exit Function
End If

    Select Case cmd
    Case 1
    Killproc pId, False
    Case 2
    PamzSuspendResumeProcessThreads pId, False
    Case 3
    PamzSuspendResumeProcessThreads pId, True
    End Select
End If
j = j - 1
DoEvents
Next i
getProces lisPROC
End Function
Function superKill(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) As Boolean
Dim hToken As Long
Dim hProcess As Long
Dim tp As TOKEN_PRIVILEGES
If GetVersion() >= 0 Then
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or _
        TOKEN_QUERY, hToken) = 0 Then
        GoTo CleanUp
    End If
    If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
        GoTo CleanUp
    End If
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
    If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
        GoTo CleanUp
    End If
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        superKill = (TerminateProcess(hProcess, ExitCode) <> 0)
        CloseHandle hProcess
    End If
    
    PamzTerminateProcess hProcessID
    
    If GetVersion() >= 0 Then
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
CleanUp:
        If hToken Then CloseHandle hToken
    End If
End Function
Public Sub Bunuh(Namafile As String)
On Error Resume Next
Dim a As Long, Hproc As Long, buf As String
buf = UCase$(Namafile)
List_Process
For a = 1 To jmlProcess
    If UCase$(pPath(a)) = buf Then
    Hproc = OpenProcess(PROCESS_TERMINATE Or PROCESS_VM_READ, 1, ProcessId(a))
     If Hproc > 1 Then
     TerminateProcess Hproc, 0
     CloseHandle Hproc
     Else
     superKill ProcessId(a)
     End If
    End If
Next a
End Sub
Public Function ProsesadA(Namafile As String) As Boolean
Dim a As Long, buf As String
buf = UCase$(Namafile)
List_Process
For a = 1 To jmlProcess
    If UCase$(pPath(a)) = buf Then
    ProsesadA = True
    Exit For
    End If
Next a
End Function
Public Function PiDadA(pId As Long) As Boolean
Dim a As Long
List_Process
For a = 1 To jmlProcess
    If ProcessId(a) = pId Then
    PiDadA = True: Exit For
    End If
Next a
End Function
Public Sub Killproc(ByVal pId As Long, auto As Boolean)
'ojan superkiller process
On Error Resume Next
Dim ProcessH As Long
tundaTHREAD pId, True
PamzSuspendResumeProcessThreads pId, False
ProcessH = OpenProcess(PROCESS_ALL_ACCESS, False, pId)
    If ProcessH > 0 Then
    TerminateProcess ProcessH, 0
    CloseHandle ProcessH
    End If
superKill pId
Sleep 100
If PiDadA(pId) = True Then
    If auto = True Then GoTo KkilP
    If MsgBox("Kill This Process With ObavKernelProcessKiller?", vbQuestion + vbYesNo + vbSystemModal) = vbYes Then
KkilP:
    KeKillProcess pId, auto
    End If
End If
'Shell "taskkill.exe /F /PID " & piD & " /T", vbHide
DebugActiveProcess pId
End Sub
Public Sub tundaTHREAD(hpid As Long, Tunda As Boolean)
Dim Thread() As THREADENTRY32, hThread As Long, i As Long
Thread32Enum Thread(), hpid
For i = 0 To UBound(Thread)
    If Thread(i).th32OwnerProcessID = hpid Then
     hThread = OpenThread(THREAD_ALL_ACCESS, False, (Thread(i).th32ThreadID))
     If Tunda = True Then
         SuspendThread hThread
     Else
         ResumeThread hThread
     End If
    CloseHandle hThread
    End If
Next i
End Sub

Private Function Thread32Enum(ByRef Thread() As THREADENTRY32, ByVal lProcessID As Long) As Long
ReDim Thread(0)
Dim THREADENTRY32 As THREADENTRY32
Dim hThreadSnap As Long
Dim lThread As Long
hThreadSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)
THREADENTRY32.dwSize = Len(THREADENTRY32)
If Thread32First(hThreadSnap, THREADENTRY32) = False Then
    Thread32Enum = -1
    Exit Function
Else
    ReDim Thread(lThread)
    Thread(lThread) = THREADENTRY32
End If
Do
 If Thread32Next(hThreadSnap, THREADENTRY32) = False Then
 Exit Do
 Else
 lThread = lThread + 1
 ReDim Preserve Thread(lThread)
 Thread(lThread) = THREADENTRY32
 End If
Loop
Thread32Enum = lThread
CloseHandle hThreadSnap
End Function

Attribute VB_Name = "mProcess"
Option Explicit
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

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
    szexeFile As String * 260
End Type
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExW Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As Long, ByVal nSize As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION As Long = &H400

 Public ProcessId(150) As Long
 Public pPath(150) As String
 Public nama(150) As String
 Public jmlProcess As Long
Public Sub KillProcess(Namafile As String)
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
     End If
    End If
Next a
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
            ModuleName = Space(260)
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


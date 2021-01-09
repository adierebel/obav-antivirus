Attribute VB_Name = "mProcess"
Option Explicit
Private Const ANYSIZE_ARRAY = 20
Private Const TokenGroups = 2
Private Const TokenLinkedToken = 19
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const SE_GROUP_ENABLED = &H4
Private Const SE_GROUP_ENABLED_BY_DEFAULT = &H2
Private Const SE_GROUP_RESOURCE = &H20000000
Private Const SECURITY_NT_AUTHORITY = &H5
Private Const MAXIMUM_ALLOWED = &H2000000

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
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
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const token_All_Access As Long = 983551
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long

Private Enum TOKEN_TYPE
    TokenPrimary = 1
    TokenImpersonation
End Enum
Private Type TOKEN_LINKED_TOKEN
    hLinkedToken As Long
End Type

Private Declare Function DuplicateTokenEx Lib "advapi32" (ByVal hExistingToken As Long, ByVal dwDesiredAccess As Long, lpTokenAttributes As Any, ByVal ImpersonationLevel As Long, ByVal TokenType As TOKEN_TYPE, phNewToken As Long) As Long
Private Type SID_AND_ATTRIBUTES
        Sid As Long
        Attributes As Long
End Type
Private Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type
Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type
Private Declare Function AdjustTokenGroups Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal lSize As Long) As Long
Private Declare Function IsValidSid Lib "advapi32" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "advapi32" (pSid1 As Any, pSid2 As Any) As Long
Private Declare Sub FreeSid Lib "advapi32" (pSid As Any)
Private Declare Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As Long, ImpersonationLevel As Integer, DuplicateTokenHandle As Long) As Long

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const MAX_PATH As Integer = 260
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
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ = &H10
Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByRef lphModule As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByVal hModule As Long, _
    ByVal ModuleName As String, _
    ByVal nSize As Long) As Long
 Private ProcessID(100) As Long
 Private Pname(100) As String
 Private Path(100) As String
 Private jmlProcess As Integer
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

'Private Declare Function DestroyEnvironmentBlock Lib "userenv.dll" (ByVal lpEnvironment As Long) As Long
'Private Declare Function CreateEnvironmentBlock Lib "userenv.dll" (ByRef lpEnvironment As Long, ByVal hToken As Long, ByVal bInhet As Boolean) As Long
'Private Declare Function weee Lib "ObavRunner2" Alias "x" (ByVal hToken As Long, ByVal nama As String) As Long
'Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
'Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserW" (ByVal htoken As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, byval lpProcessAttributes As long, byval lpThreadAttributes As long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, byval lpStartupInfo As long, byval lpProcessInformation As long) As Long

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String

Private Const szTarget As String = "ProgMan"
Public Function ProsesadA(Namafile As String) As Boolean
Dim a As Integer
Call ListProc
For a = 1 To jmlProcess
    If Path(a) = Namafile Then
ProsesadA = True
Exit For
    End If
Next a
End Function
Private Function GetProcByName(Namafile As String) As Long
Dim a As Integer
Dim Namafilebuf As String
Call ListProc
Namafilebuf = UCase$(Namafile)
For a = 1 To jmlProcess
    If InStr(UCase$(Pname(a)), Namafilebuf) > 0 Then
GetProcByName = ProcessID(a)
Exit For
    End If
Next a
End Function
Public Function Medir() As String
Medir = fixP(App.Path)
End Function
Private Function fixP(ByVal paT As String) As String
On Error Resume Next
If Right$(paT, 1) <> ChrW$(92) Then
fixP = paT & ChrW$(92)
Else
fixP = paT
End If
End Function
Public Sub Call_ObAV()
Dim Environtt As Long
Dim tp As TOKEN_PRIVILEGES
Dim hp As Long, hT As Long
Dim hTret As Long
Dim Tgrup As TOKEN_GROUPS
Dim linkToken As TOKEN_LINKED_TOKEN
Dim BufferSize As Long
Dim InfoBuffer() As Long
Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
Dim psidAdmin As Long
Dim lResult As Long
Dim x As Long

Dim pid As Long
Dim cD As String
Dim sec As SECURITY_ATTRIBUTES
Dim sIn As STARTUPINFO
Dim pIn As PROCESS_INFORMATION
cD = CurDir
sec.bInheritHandle = &H1
sec.lpSecurityDescriptor = 0
sec.nLength = Len(sec)
sIn.cb = Len(sIn)
sIn.dwFlags = &H1
sIn.wShowWindow = vbNormalFocus

OpenProcessToken GetCurrentProcess, token_All_Access, hTret
    LookupPrivilegeValue "", "SeAssignPrimaryTokenPrivilege", tp.LuidUDT
    tp.PrivilegeCount = 1
    tp.Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges hTret, False, tp, 0, ByVal 0&, ByVal 0&
CloseHandle hTret

pid = GetProcByName("explorer")
If pid <= 0 Then Exit Sub
hp = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
OpenProcessToken hp, token_All_Access, hTret
    GetTokenInformation hTret, ByVal TokenLinkedToken, 0, 0, BufferSize
    If BufferSize Then
    ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
    lResult = GetTokenInformation(hTret, ByVal TokenLinkedToken, InfoBuffer(0), BufferSize, BufferSize)
        If lResult > 0 Then
        Call RtlMoveMemory(linkToken, InfoBuffer(0), Len(linkToken))
        DuplicateTokenEx linkToken.hLinkedToken, MAXIMUM_ALLOWED, ByVal 0&, 2, TokenPrimary, hT
        CloseHandle linkToken.hLinkedToken
        End If
    End If
CloseHandle hp
    
'    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
'    GetTokenInformation hT, ByVal TokenGroups, 0, 0, BufferSize
'    If BufferSize Then
'    ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
'    lResult = GetTokenInformation(hT, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
'        If lResult > 0 Then
'        Call RtlMoveMemory(Tgrup, InfoBuffer(0), Len(Tgrup))
'        lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
'            If lResult > 0 Then
'                If IsValidSid(psidAdmin) Then
'                For X = 0 To Tgrup.GroupCount
'                If IsValidSid(Tgrup.Groups(X).Sid) Then
'                    If EqualSid(ByVal Tgrup.Groups(X).Sid, ByVal psidAdmin) Then
'                    Tgrup.Groups(X).Attributes = SE_GROUP_ENABLED
'                    Exit For
'                    End If
'                End If
'                Next X
'                End If
'            AdjustTokenGroups hT, False, Tgrup, 0, ByVal 0&, ByVal 0&
'            Call FreeSid(psidAdmin)
'            End If
'        End If
'    End If

'Environtt = weee(hTret, "x")
'CreateEnvironmentBlock Environtt, hTret, True
'DestroyEnvironmentBlock Environtt

CreateProcessAsUser hT, vbNullString, Medir & "ObavGuard.exe -rtp", sec, sec, 0, &H20, vbNullString, cD, sIn, pIn
CloseHandle hTret
CloseHandle hT
End Sub

Private Function PathByPID(pid As Long) As String
    Dim cbNeeded As Long, Modules(1 To 200) As Long
    Dim Ret As Long, ModuleName As String
    Dim nSize As Long, hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
        Or PROCESS_VM_READ, 0, pid)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, Modules(1), _
            200, cbNeeded)
        If Ret <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, _
                Modules(1), ModuleName, nSize)
            PathByPID = Left(ModuleName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    If PathByPID = "" Then PathByPID = ""
    If Left(PathByPID, 4) = "\??\" Then PathByPID = ""
    If Left(PathByPID, 12) = "\SystemRoot\" Then PathByPID = ""
End Function

Private Sub ListProc()
 Dim w As Long
    jmlProcess = 1
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    w = Process32First(hSnapShot, uProcess)
    Do While w
        ProcessID(jmlProcess) = uProcess.th32ProcessID
        Pname(jmlProcess) = uProcess.szexeFile
        Path(jmlProcess) = PathByPID(ProcessID(jmlProcess))
        w = Process32Next(hSnapShot, uProcess)
    jmlProcess = jmlProcess + 1
    Loop
    jmlProcess = jmlProcess - 1
    CloseHandle hSnapShot
End Sub

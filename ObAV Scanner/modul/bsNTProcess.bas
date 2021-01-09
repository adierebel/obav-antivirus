Attribute VB_Name = "bsNTProcess"
Option Explicit
'*****************************************************************************************'
'---ini modul alternatif untuk pemecahan masalah akses proses milik Hirin sebelumnya.
'---menggunakan ntdll.dll secara eksplisit (visible-import-list), karena vb6 punya
'---kemampuan penanganan 'error' yang baik apabila fungsi dalam modul tidak ditemukan.
'---modul alternatif ini terfokus dengan menggunakan fungsi api milik system.
'---modul alternatif ini nggak bergantung pada psapi.dll.
'---modul alternatif ini untuk system dan proses 32 bit.
'---modul alternatif ini mendukung 100% nt---unicode.
'---modul alternatif ini tidak untuk mengatasi proses yg di-hook dgn kernel-level-driver.
'---untuk hasil lebih baik, aktifkan privilege "SeDebugPrivilege" terlebih dulu.
'---dan akan berhasil lebih baik,bila proses berjalan sebagai system service :)
'---kode ter-uji pada windows xp sp2, processor Intel-P3-600MHz.
'---"dengan komputer ber-spesifikasi rendah ini, ari bisa tahu program atau aplikasi mana
'--- yg benar-benar bisa berjalan atau eksekusi dgn cepat, bukan dgn bantuan spesifikasi,
'--- tapi dari kode yg lebih efektif dan efisien."
'---oleh    : ari pambudi[pamzlogic]10Nopember2009.
'---mulai   : 10Nopember2009@18:20.
'---selesai : 19Nopember2009@21:15.
'*****************************************************************************************'


Private Const MemoryBasicInformation = 0
Private Const MemorySectionName = 2

Private Const ProcessBasicInformation = 0

Private Const SystemProcessesInformations = 5
Private Const SystemThreadsInformations = 5 '---sama dengan proses, lha nama aslinya: "SystemProcessesAndThreadsInformations",kok.
Private Const SystemModuleInformation = 11
Private Const SystemHandleInformation = 16
Private Const SystemRangeStartInformation = 50

Private Const ObjectAllTypesInformation = 3

Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Const STATUS_ADDRESS_HANDLE_WAS_INVALID = &HC0000141
Private Const STATUS_INVALID_PARAMETER_IN_SERVICE = &HC000000D

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000

Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_SET_SESSIONID = &H4
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF

Private Const THREAD_TERMINATE = &H1
Private Const THREAD_SUSPEND_RESUME = &H2
Private Const THREAD_GET_CONTEXT = &H8
Private Const THREAD_SET_CONTEXT = &H10
Private Const THREAD_SET_INFORMATION = &H20
Private Const THREAD_QUERY_INFORMATION = &H40
Private Const THREAD_SET_THREAD_TOKEN = &H80
Private Const THREAD_IMPERSONATE = &H100
Private Const THREAD_DIRECT_IMPERSONATION = &H200
Private Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF

Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_FREE = &H10000
Private Const MEM_PRIVATE = &H20000
Private Const MEM_MAPPED = &H40000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const MEM_WRITE_WATCH = &H200000
Private Const MEM_PHYSICAL = &H400000


Private Const PAGE_NOACCESS = &H1
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_WRITECOPY = &H8
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_EXECUTE_WRITECOPY = &H80
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOCACHE = &H200
Private Const PAGE_WRITECOMBINE = &H400

Private Const LMEM_FIXED = &H0
Private Const LMEM_MOVEABLE = &H2
Private Const LMEM_NOCOMPACT = &H10
Private Const LMEM_NODISCARD = &H20
Private Const LMEM_ZEROINIT = &H40
Private Const LMEM_MODIFY = &H80
Private Const LMEM_DISCARDABLE = &HF00
Private Const LMEM_VALID_FLAGS = &HF72
Private Const LMEM_INVALID_HANDLE = &H8000

Private Const LHND = LMEM_MOVEABLE + LMEM_ZEROINIT
Private Const lPtr = LMEM_FIXED + LMEM_ZEROINIT

Private Const OBJ_CASE_INSENSITIVE = &H40
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_INSUFFICIENT_BUFFER = 122

Private Type WIN_VERSION_INDIRECT
    btMajorVersion              As Byte
    btMinorVersion              As Byte
    wMiscInfo                   As Integer
End Type

Public Type LARGE_INTEGER
    LoPart                      As Long
    HiPart                      As Long
End Type

Private Type UNICODE_STRING
    Length                      As Integer
    MaxLength                   As Integer
    pToBuffer                   As Long
End Type

Private Type OBJECT_ATTRIBUTES
    Length                      As Long
    RootDirectory               As Long
    ObjectName                  As Long '---pointer ke unicode string.
    Attributes                  As Long
    SecurityDescriptor          As Long
    SecurityQualityOfService    As Long
End Type

Private Type IO_STATUS_BLOCK
    lPointer                    As Long
    pInformation                As Long
End Type

Private Type CLIENT_ID
    UniqueProcess               As Long
    UniqueThread                As Long
End Type

Private Type GENERIC_MAPPING
    GenericRead                 As Long
    GenericWrite                As Long
    GenericExecute              As Long
    GenericAll                  As Long
End Type

Private Type MEMORY_BASIC_INFORMATION
    pBaseAddress                As Long 'baseaddr.
    pAllocationBase             As Long 'region.
    tAllocationProtect          As Long
    nRegionSize                 As Long 'lenof baseaddr.
    dState                      As Long
    tProtect                    As Long
    rgType                      As Long
End Type

Private Type SYSTEM_HANDLE_TABLE_ENTRY_INFO '==16 bytes.
    UniqueProcessId             As Integer '2
    CreatorBackTraceIndex       As Integer '2
    ObjectTypeIndex             As Byte '1
    HandleAttributes            As Byte '1
    HandleValue                 As Integer '2
    pObject                     As Long '4
    GrantedAccess               As Long '4
End Type

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                  As Long 'NTSTATUS
    PebBaseAddress              As Long 'PPEB
    AffinityMask                As Long 'ULONG_PTR
    BasePriority                As Long 'KPRIORITY
    UniqueProcessId             As Long 'ULONG_PTR
    InheritedFromUniqueProcId   As Long 'ULONG_PTR
End Type

Private Type PEB
    InheritedAddressSpace       As Byte
    ReadImageFileExecOptions    As Byte
    BeingDebugged               As Byte
    BitField                    As Byte
    Mutant                      As Long
    ImageBaseAddress            As Long '<---sip,lah.
    Ldr                         As Long 'ptr.
    ProcessParameters           As Long
    '---...masih banyak yg lainnya, tapi nggak penting :)
End Type

Private Type TEB
    ExceptionList               As Long
    StackBase                   As Long
    StackLimit                  As Long
    SubSystemTIB                As Long
    FiberData                   As Long
    ArbitraryUser               As Long
    Self                        As Long
    EnvironmentPointer          As Long
    ClientId                    As CLIENT_ID
    ActiveRpcHandle             As Long
    ThreadLocalStoragePointer   As Long
    ProcessEnvironmentBlock     As Long
    '---...masih banyak yg lainnya, tapi nggak penting :)
End Type

Private Type VM_COUNTERS
    PeakVirtualSize             As Long
    VirtualSize                 As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Long
    WorkingSetSize              As Long
    QuotaPeakPagedPoolUsage     As Long
    QuotaPagedPoolUsage         As Long
    QuotaPeakNonPagedPoolUsage  As Long
    QuotaNonPagedPoolUsage      As Long
    PagefileUsage               As Long
    PeakPagefileUsage           As Long
End Type

Private Type IO_COUNTERS
    ReadOperationCount          As LARGE_INTEGER
    WriteOperationCount         As LARGE_INTEGER
    OtherOperationCount         As LARGE_INTEGER
    ReadTransferCount           As LARGE_INTEGER
    WriteTransferCount          As LARGE_INTEGER
    OtherTransferCount          As LARGE_INTEGER
End Type

Private Type SYSTEM_PROCESSES_INFORMATION
    NextEntryDelta              As Long
    ThreadCount                 As Long
    Reserved1(5)                As Long
    CreateTime                  As LARGE_INTEGER
    UserTime                    As LARGE_INTEGER
    KernelTime                  As LARGE_INTEGER
    ProcessName                 As UNICODE_STRING
    BasePriority                As Long 'KPRIORITY
    ProcessId                   As Long
    InheritedFromProcessId      As Long
    HandleCount                 As Long
    Reserved2(1)                As Long
    VmCounters                  As VM_COUNTERS
    'IOCounters                  As IO_COUNTERS '---hanya ada mulai Win2000 ke atas.
    '---sebenarnya masih ada banyak setelah ini,tapi nggak penting saat ini :)
End Type

Private Type SYSTEM_THREADS
    KernelTime                  As LARGE_INTEGER
    UserTime                    As LARGE_INTEGER
    CreateTime                  As LARGE_INTEGER
    WaitTime                    As LARGE_INTEGER
    StartAddress                As Long
    ClientId                    As CLIENT_ID
    Priority                    As Long
    BasePriority                As Long
    ContextSwitchCount          As Long
    State                       As Long
    WaitReason                  As Long
End Type

Private Type SYSTEM_MODULE_INFORMATION 'Information Class 11
    Reserved(1)                 As Long
    Base                        As Long
    SIZE                        As Long
    Flags                       As Long
    Index                       As Integer
    Unknown                     As Integer
    LoadCount                   As Integer
    ModuleNameOffset            As Integer
    ImageName                   As String * 256 '---jangan dirubah,ketentuan(!).
End Type

Private Type OBJECT_TYPE_INFORMATION
    pszName                     As UNICODE_STRING
    ObjectCount                 As Long
    HandleCount                 As Long
    Reserved1(3)                As Long
    PeakObjectCount             As Long
    PeakHandleCount             As Long
    Reserved2(3)                As Long
    InvalidAttributes           As Long
    GenericMapping              As GENERIC_MAPPING
    ValidAccess                 As Long
    Unknown                     As Integer '==UCHAR
    MaintainHandleDatabase      As Integer '==BOOLEAN
    PoolType                    As Long '==typedef enum _POOL_TYPE
    PagedPoolUsage              As Long
    NonPagedPoolUsage           As Long
End Type


'---alternative types :)
Public Type ENUMERATE_PROCESSES_OUTPUT '---dalam unicode.
    '---nomor dan alamat di memori:
    nProcessID                  As Long '---nomor prosesnya.
    nParentProcessID            As Long '---nomor proses induknya.
    pBaseAddress                As Long '---lokasi executable program utama (dos header) di memori.
    pPEBBaseAddress             As Long '---lokasi struktur PEB(ProcessEnvironmentBlock) di memori.
    pLDRAddress                 As Long '---lokasi struktur LDR di memori.
    pProcessParamAddress        As Long '---lokasi parameter proses.
    nThreadCount                As Long '---jumlah thread yang bekerja di dalamnya.
    nObjectHandleCount          As Long '---jumlah object handle (file,event,thread,key,&lainnya) yang dibuka di dalamnya.
    nSizeOfExecutableOpInMemory As Long '---ukuran executable program utama di memori.(=maximum unpacked PE image size).
    '---pemberi tanda:
    bIsBeingDebugged            As Boolean '---bila sedang di-debug, bernilai > 0 (=1).
    bIsLockedProcess            As Boolean '---bila proses susah untuk diakses, bernilai > 0 (=1).
    bIsHiddenProcess            As Boolean '---bila proses tidak dapat dienumerasi secara normal, dianggap mencurigakan karena tersembunyi, dan bernilai > 0 (=1).
    '---nama dan alamat di disk (selalu update, bila file yg sedang berproses di-move lokasinya,alamat file akan ikut berubah):
    szNtExecutableNameW         As String '---nama (saja) executable program utama dengan format nt---unicode.
    szNtExecutablePathW         As String '---alamat dan nama executable program utama dengan format nt---unicode.
End Type

Public Type ENUMERATE_MODULES_OUTPUT
    '---nomor dan alamat di memori:
    pBaseAddress                As Long '---lokasi module di memori.
    nSizeOfModuleOpInMemory     As Long '---ukuran module di memori.(=bila PE, berilai maximum unpacked PE image size).
    '---pemberi tanda:
    bIsLockedModule             As Boolean '---bila module susah untuk diakses, bernilai > 0 (=1).
    bIsHiddenModule             As Boolean '---bila module tidak dapat dienumerasi secara normal, dianggap mencurigakan karena tersembunyi, dan bernilai > 0 (=1).
    '---nama dan alamat di disk (selalu update, bila file yg sedang berproses di-move lokasinya,alamat file akan ikut berubah):
    szNtModuleNameW             As String '---nama (saja) executable program utama dengan format nt---unicode.
    szNtModulePathW             As String '---alamat dan nama executable program utama dengan format nt---unicode.
End Type

Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long

Private Declare Function GetModuleHandleW Lib "kernel32.dll" (ByVal plpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal szlpProcName As String) As Long

'---seharusnya pakai VirtualAlloc,VirtualFree,tapi biarlah:
Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal lBytes As Long) As Long
Private Declare Function LocalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal pv_lpString As Long) As Long

Private Declare Function GetVersion Lib "kernel32.dll" () As Long

Private Declare Function QueryDosDeviceW Lib "kernel32.dll" (ByVal pp_lpDeviceName As Long, ByVal pp_lpTargetPath As Long, ByVal ucchMax As Long) As Long

Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, ByVal plpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, ByVal plpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualProtectEx Lib "kernel32.dll" (ByVal hProcess As Long, ByVal plpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByVal pplpflOldProtect As Long) As Long

Private Declare Function CreateRemoteThread Lib "kernel32.dll" (ByVal hProcess As Long, ByVal plpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal plpParameter As Long, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Sub RtlMoveMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.
Private Declare Sub RtlInitUnicodeString Lib "ntdll.dll" (ByVal pVarTypeUnicodeString As Long, ByVal pTargetUnicodeString As Long)

Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" (ByVal SystemInfoClass As Long, ByVal SystemInformation As Long, ByVal SystemInfoLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtClose Lib "ntdll.dll" (ByVal ObjectHandle As Long) As Long
Private Declare Function NtOpenProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal AccessMask As Long, ByVal ObjectAttributes As Long, ByVal ClientId As Long) As Long
Private Declare Function NtTerminateProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal ExitStatus As Long) As Long
Private Declare Function NtQueryInformationProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal ProcessInfoClass As Long, ByVal ProcessInformation As Long, ByVal ProcessInfoLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtOpenThread Lib "ntdll.dll" (ByVal ThreadHandle As Long, ByVal AccessMask As Long, ByVal ObjectAttributes As Long, ByVal ClientId As Long) As Long
Private Declare Function NtSuspendThread Lib "ntdll.dll" (ByVal ThreadHandle As Long, ByVal PreviousSuspendCount As Long) As Long
Private Declare Function NtResumeThread Lib "ntdll.dll" (ByVal ThreadHandle As Long, ByVal PreviousSuspendCount As Long) As Long
Private Declare Function NtTerminateThread Lib "ntdll.dll" (ByVal ThreadHandle As Long, ByVal ExitStatus As Long) As Long
Private Declare Function NtQueryVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal pBaseAddress As Long, ByVal MemoryInfoClass As Long, ByVal pMemoryInformation As Long, ByVal nMemoryInfoLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtReadVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal pBaseAddress As Long, ByVal pBuffer As Long, ByVal nBufferLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtWriteVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal pBaseAddress As Long, ByVal pBuffer As Long, ByVal nBufferLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal ObjectHandle As Long, ByVal ObjectInfoClass As Long, ByVal pObjectInformation As Long, ByVal nObjectInfoLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Function NtDuplicateObject Lib "ntdll.dll" (ByVal SourceProcessHandle As Long, ByVal SourceHandle As Long, ByVal TargetProcessHandle As Long, ByVal TargetHandle As Long, ByVal DesiredAccess As Long, ByVal Attributes As Long, ByVal Options As Long) As Long

Public Function PamzCloseHandle(ByVal nTargetHandle As Long) As Long
On Error Resume Next
Dim NtStatus                        As Long
    NtStatus = NtClose(nTargetHandle)
    If NtStatus = 0 Then
        PamzCloseHandle = 1
    End If
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzOpenProcess(ByVal nTargetPID As Long, ByVal nAccessMask As Long) As Long
On Error Resume Next '---PID=ProcessID.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim NtStatus                        As Long
Dim NtRetValue                      As Long
    OBJAT.Length = Len(OBJAT)
    CLIID.UniqueProcess = nTargetPID
    NtStatus = NtOpenProcess(VarPtr(NtRetValue), nAccessMask, VarPtr(OBJAT), VarPtr(CLIID))
    If NtStatus = 0 Then
        PamzOpenProcess = NtRetValue
    End If
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzOpenThread(ByVal nTargetTID As Long, ByVal nAccessMask As Long) As Long
On Error Resume Next '---TID=ThreadID. ThreadID unique terhadap system,jadi nggak ada nomor TID yg sama dalam 1 system, di waktu yang bersamaan.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim NtStatus                        As Long
Dim NtRetValue                      As Long
    OBJAT.Length = Len(OBJAT)
    CLIID.UniqueThread = nTargetTID
    NtStatus = NtOpenThread(VarPtr(NtRetValue), nAccessMask, VarPtr(OBJAT), VarPtr(CLIID))
    If NtStatus = 0 Then
        PamzOpenThread = NtRetValue
    End If
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzTerminateProcess(ByVal nTargetPID As Long) As Long
On Error Resume Next '---PID=ProcessID.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim NtStatus                        As Long
Dim NtRetValue                      As Long
Dim NtRetLength                     As Long
Dim NtProcessHandle                 As Long
Dim PPTEB                           As TEB

LBL_PROCESS_TERMINATOR_METHOD_1:
    '---coba hentikan dgn metode #1:
    NtProcessHandle = PamzOpenProcess(nTargetPID, PROCESS_TERMINATE)
    If NtProcessHandle <= 0 Then
        GoTo LBL_PROCESS_TERMINATOR_METHOD_2
    End If
    '---coba hentikan via nt-standar:
    NtStatus = NtTerminateProcess(NtProcessHandle, 0)
    If NtStatus <> 0 Then
        Call PamzCloseHandle(NtProcessHandle)
        NtProcessHandle = 0
        GoTo LBL_PROCESS_TERMINATOR_METHOD_2
    End If
    NtRetValue = 1 '---berhasil dihentikan dengan metode #1.
LBL_PROCESS_TERMINATOR_CLOSEHANDLE_1:
    Call PamzCloseHandle(NtProcessHandle)
    NtProcessHandle = 0
    GoTo LBL_BROADCAST_RESULT
LBL_PROCESS_TERMINATOR_METHOD_2: '---masih banyak metode alternatif lainnya... .
    
LBL_PROCESS_TERMINATOR_METHOD_3: '---masih banyak metode alternatif lainnya... .
    
LBL_PROCESS_TERMINATOR_METHOD_4: '---masih banyak metode alternatif lainnya... .
    
LBL_CLEAN_MEMORY:
    
LBL_BROADCAST_RESULT:
    PamzTerminateProcess = NtRetValue
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzSuspendResumeProcessThreads(ByVal nTargetPID As Long, ByVal bToResume As Boolean) As Long
On Error Resume Next '---PID=ProcessID.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim NtStatus                        As Long
Dim NtRetValue                      As Long
Dim NtRetLength                     As Long
Dim NtProcessHandle                 As Long
Dim NtThreadHandle                  As Long
Dim bIsSuccess                      As Long
Dim nThreadGetCount                 As Long
Dim nThreadHitCount                 As Long
Dim pPosTmp                         As Long
Dim pNextTmp                        As Long
Dim pPosDelta                       As Long
Dim pNextDelta                      As Long
Dim nPrevSuspendCount               As Long
Dim CTurn                           As Long
Dim DTurn                           As Long
Dim PPWVI                           As WIN_VERSION_INDIRECT
Dim PRMBI                           As MEMORY_BASIC_INFORMATION
Dim PRCBI                           As PROCESS_BASIC_INFORMATION
Dim SPRCI                           As SYSTEM_PROCESSES_INFORMATION
Dim PSSTH                           As SYSTEM_THREADS
Dim PPIOC                           As IO_COUNTERS
Dim PPTEB                           As TEB
Dim pAddressOfPEB                   As Long
Dim pBuffer                         As Long
Dim nBufferLength                   As Long
Dim szBuffer                        As String '---pakai string ajah buffernya, nggak ruwet kaya' *Alloc yang lainnya.

    nThreadGetCount = 0
    nThreadHitCount = 0
LBL_PROCESS_BREAK_METHOD_1: '---enumerasi memori,cari dan verifikasi TEB, coba akses thread:
    NtProcessHandle = PamzOpenProcess(nTargetPID, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    If NtProcessHandle = 0 Then
        GoTo LBL_PROCESS_BREAK_METHOD_2
    End If
    '---cari tahu alamat struktur PEB:
    NtStatus = NtQueryInformationProcess(NtProcessHandle, ProcessBasicInformation, VarPtr(PRCBI), Len(PRCBI), 0)
    If NtStatus <> 0 Then
        bIsSuccess = 0
        GoTo LBL_SUB_CLOSEHANDLE_1
    End If

    pAddressOfPEB = PRCBI.PebBaseAddress
    If pAddressOfPEB <= 0 Then
        bIsSuccess = 0
        GoTo LBL_SUB_CLOSEHANDLE_1
    End If
    '---cari region di memori yg punya angka sama pada lokasi tertentu sesuai struktur TEB:
LBL_SUB_LOOP_GET_TEB_PATTERN_1:
        '---dapatkan struktur region:
        NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI.pBaseAddress + PRMBI.nRegionSize, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
        Select Case NtStatus
            Case 0 '---berhasil mendapatkan info.
                '---cari nama untuk region yg bukan "Free" ajah:
                If PRMBI.pAllocationBase <> 0 Then
                    If PRMBI.pAllocationBase <> pPosTmp Then '---address yg sekarang tidak terletak pada region yg sama dgn sebelumnya:
                        pPosTmp = PRMBI.pAllocationBase '---ingat-ingat region yg sekarang.
                        '---cari tahu pattern:
                        If PRMBI.nRegionSize >= Len(PPTEB) Then
                            NtStatus = NtReadVirtualMemory(NtProcessHandle, PRMBI.pAllocationBase, VarPtr(PPTEB), Len(PPTEB), 0)
                            If PPTEB.ProcessEnvironmentBlock = pAddressOfPEB Then '---punya alamat TEB yg sama.
                                If PPTEB.ClientId.UniqueProcess = nTargetPID Then '---punya nomor PID yg sama:
                                    '66,66% yakin kalau nomor TID udah didapat :)
                                    'Debug.Print PPTEB.ClientId.UniqueThread
                                    '---sekarang,diutak-atik:
                                    nThreadGetCount = nThreadGetCount + 1
                                    NtThreadHandle = PamzOpenThread(PPTEB.ClientId.UniqueThread, THREAD_SUSPEND_RESUME)
                                    If NtThreadHandle > 0 Then
                                        If bToResume = True Then
                                            NtStatus = NtResumeThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                                            '---tambahan:upayakan thread benar-benar bisa berjalan lagi:
                                            If NtStatus = 0 Then
                                                nThreadHitCount = nThreadHitCount + 1
                                                If nPrevSuspendCount > 1 Then
                                                    CTurn = 0
                                                    pNextTmp = nPrevSuspendCount
                                                    For CTurn = 1 To pNextTmp
                                                        NtStatus = NtResumeThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                                                        If NtStatus = 0 Then
                                                            If nPrevSuspendCount = 0 Then
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Else
                                            NtStatus = NtSuspendThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                                            If NtStatus = 0 Then
                                                nThreadHitCount = nThreadHitCount + 1
                                            End If
                                        End If
                                        '---jangan lupa tutup lagi handle ke thread-nya:
                                        Call PamzCloseHandle(NtThreadHandle)
                                        NtThreadHandle = 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                GoTo LBL_SUB_LOOP_GET_TEB_PATTERN_1
            Case STATUS_INVALID_PARAMETER_IN_SERVICE
                GoTo LBL_SUB_CLOSEHANDLE_1
            Case Else
                GoTo LBL_SUB_CLOSEHANDLE_1
        End Select
    '---tutup handle bila selesai:
LBL_SUB_CLOSEHANDLE_1:
    Call PamzCloseHandle(NtProcessHandle)
    NtProcessHandle = 0
    If nThreadGetCount > 0 Then
        If nThreadHitCount = nThreadGetCount Then
            bIsSuccess = 1
        End If
    End If
    If bIsSuccess > 0 Then
        NtRetValue = 1 '---berhasil diistirahatkan dengan metode #1.
        GoTo LBL_BROADCAST_RESULT
    End If
LBL_PROCESS_BREAK_METHOD_2: '---masih banyak metode alternatif lainnya... .
    'Debug.Print "Break_Cara_2:"
    nThreadGetCount = 0
    nThreadHitCount = 0
    bIsSuccess = 0
    NtRetValue = 0
    NtStatus = NtQuerySystemInformation(SystemProcessesInformations, 0, 0, VarPtr(nBufferLength))
    If NtStatus <> STATUS_INFO_LENGTH_MISMATCH Then
        GoTo LBL_PROCESS_BREAK_METHOD_3
    End If
    If nBufferLength = 0 Then
        GoTo LBL_PROCESS_BREAK_METHOD_3
    End If
    szBuffer = String$(nBufferLength, 0) '---isi dengan bytenumber=0. '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
    pBuffer = StrPtr(szBuffer)
    '---yuk,coba panggil lagi untuk dapetin datanya:
    NtStatus = NtQuerySystemInformation(SystemProcessesInformations, pBuffer, nBufferLength, VarPtr(NtRetLength))
    If NtStatus <> 0 Then
        szBuffer = ""
        GoTo LBL_PROCESS_BREAK_METHOD_3
    End If
    If NtRetLength <> nBufferLength Then
        szBuffer = ""
        GoTo LBL_PROCESS_BREAK_METHOD_3
    End If
    pPosDelta = pBuffer
    pNextDelta = 0
LBL_LOOP_PROCESS_BREAK_METHOD_2:
    Call RtlMoveMemory(VarPtr(SPRCI), pPosDelta, Len(SPRCI))
    '---cek hanya proses yg bernomor PID sama:
    If SPRCI.ProcessId = nTargetPID Then
        '---cek setiap nomor thread yg didapatkan:
        If SPRCI.ThreadCount > 0 Then '---kalau ada thread-nya yg bisa dibaca,sih:
            '---cari tahu start alamat struktur info thread (Win2000,XP,Vista,W7,dgn sebelumnya bisa berbeda):
            '---I/O_Counters mulai Win2000:
            Call RtlMoveMemory(VarPtr(PPWVI), VarPtr(GetVersion), Len(PPWVI))
            If PPWVI.btMajorVersion >= 5 Then '---Win2000 atau lebih(Ntah kalau W7,semoga ajah sama...):
                pPosTmp = pPosDelta + Len(SPRCI) + Len(PPIOC)
            Else
                pPosTmp = pPosDelta + Len(SPRCI)
            End If
            '---cari tiap nomor TID:
            CTurn = 0
            For CTurn = 0 To (SPRCI.ThreadCount - 1)
                Call RtlMoveMemory(VarPtr(PSSTH), pPosTmp + (CTurn * Len(PSSTH)), Len(PSSTH))
                'Debug.Print SPRCI.ProcessId & vbTab & PSSTH.ClientId.UniqueThread
                '---untuk setiap nomor TID yg didapatkan:
                '---sekarang,diutak-atik:
                nThreadGetCount = nThreadGetCount + 1
                NtThreadHandle = PamzOpenThread(PSSTH.ClientId.UniqueThread, THREAD_SUSPEND_RESUME)
                'Debug.Print SPRCI.ProcessId & vbTab & PSSTH.ClientId.UniqueThread & vbTab & NtThreadHandle
                If NtThreadHandle > 0 Then
                    If bToResume = True Then
                        NtStatus = NtResumeThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                        '---tambahan:upayakan thread benar-benar bisa berjalan lagi:
                        If NtStatus = 0 Then
                            nThreadHitCount = nThreadHitCount + 1
                            If nPrevSuspendCount > 1 Then
                                DTurn = 0
                                pNextTmp = nPrevSuspendCount
                                For DTurn = 1 To pNextTmp
                                    NtStatus = NtResumeThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                                    If NtStatus = 0 Then
                                        If nPrevSuspendCount = 0 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Else
                        NtStatus = NtSuspendThread(NtThreadHandle, VarPtr(nPrevSuspendCount))
                        If NtStatus = 0 Then
                            nThreadHitCount = nThreadHitCount + 1
                        End If
                    End If
                    '---jangan lupa tutup lagi handle ke thread-nya:
                    Call PamzCloseHandle(NtThreadHandle)
                    NtThreadHandle = 0
                End If
                '---------------------------------------;
            Next
        End If
        '---akhiri bila udah selesai:
        GoTo LBL_END_LOOP_PROCESS_BREAK_METHOD_2 '---udah,nggak usah muter-muter lagi.
    End If
    '---cek untuk entry berikutnya:
    pNextDelta = SPRCI.NextEntryDelta
    If pNextDelta <> 0 Then
        pPosDelta = pPosDelta + pNextDelta
        GoTo LBL_LOOP_PROCESS_BREAK_METHOD_2
    End If
LBL_END_LOOP_PROCESS_BREAK_METHOD_2:
    Call RtlZeroMemory(VarPtr(SPRCI), Len(SPRCI))
    szBuffer = ""
    '---cek apakah sudah beres:
    If nThreadGetCount > 0 Then
        If nThreadHitCount = nThreadGetCount Then
            bIsSuccess = 1
        End If
    End If
    If bIsSuccess > 0 Then
        NtRetValue = 1 '---berhasil diistirahatkan dengan metode #1.
        GoTo LBL_BROADCAST_RESULT
    End If
LBL_PROCESS_BREAK_METHOD_3: '---masih banyak metode alternatif lainnya... .
    'Debug.Print "Break_Cara_3:"
    nThreadGetCount = 0
    nThreadHitCount = 0
    bIsSuccess = 0
    NtRetValue = 0
    
LBL_PROCESS_BREAK_METHOD_4: '---masih banyak metode alternatif lainnya... .
    
LBL_CLEAN_MEMORY:
    szBuffer = ""
LBL_BROADCAST_RESULT:
    PamzSuspendResumeProcessThreads = NtRetValue
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzEnumerateProcesses(ByRef OutputProcessesData() As ENUMERATE_PROCESSES_OUTPUT) As Long
On Error Resume Next '---kalau result bernilai > 0, berarti menunjukkan jumlah proses yg berjalan. kalau result bernilai <= 0, menunjukkan error value.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim UNSTI                           As UNICODE_STRING
Dim NtStatus                        As Long
Dim NtRetLength                     As Long
Dim NtRetValue                      As Long
Dim NtProcessHandle                 As Long
Dim SPRCI                           As SYSTEM_PROCESSES_INFORMATION
Dim PRCBI                           As PROCESS_BASIC_INFORMATION
Dim PEBSC                           As PEB
Dim PRMBI                           As MEMORY_BASIC_INFORMATION
Dim SHTEI                           As SYSTEM_HANDLE_TABLE_ENTRY_INFO
Dim OBJTI                           As OBJECT_TYPE_INFORMATION
Dim CTurn                           As Long
Dim DTurn                           As Long
Dim ETime                           As Long
Dim bAlreadyAdded                   As Long
Dim nHandleCount                    As Long
Dim pPosTmp                         As Long
Dim pNextTmp                        As Long
Dim pPosDelta                       As Long
Dim pNextDelta                      As Long
Dim nProcessesCount                 As Long
Dim pTmpName                        As Long
Dim nTmpNameLength                  As Long
Dim nObjectHandleTypeCount          As Long
Dim ThreadTypeObjectID              As Long
Dim hDupObject                      As Long
Dim szTmpName                       As String
Dim szTmpName2                      As String
Dim pBuffer                         As Long
Dim nBufferLength                   As Long
Dim szBuffer                        As String '---pakai string ajah buffernya, nggak ruwet kaya' *Alloc yang lainnya.
LBL_PRESET_SECTOR:
    Erase OutputProcessesData() '---hapus isi array sebelumnya (bila ada).
LBL_ENUM_PROCESSES_METHOD_1: '---via standar tanya ke system tentang daftar proses yang berjalan:
    NtStatus = NtQuerySystemInformation(SystemProcessesInformations, 0, 0, VarPtr(nBufferLength))
    If NtStatus <> STATUS_INFO_LENGTH_MISMATCH Then
        GoTo LBL_ENUM_PROCESSES_METHOD_2
    End If
    If nBufferLength = 0 Then
        GoTo LBL_ENUM_PROCESSES_METHOD_2
    End If
    szBuffer = String$(nBufferLength, 0) '---isi dengan bytenumber=0. '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
    pBuffer = StrPtr(szBuffer)
    '---yuk,coba panggil lagi untuk dapetin datanya:
    NtStatus = NtQuerySystemInformation(SystemProcessesInformations, pBuffer, nBufferLength, VarPtr(NtRetLength))
    If NtStatus <> 0 Then
        szBuffer = ""
        GoTo LBL_ENUM_PROCESSES_METHOD_2
    End If
    If NtRetLength <> nBufferLength Then
        szBuffer = ""
        GoTo LBL_ENUM_PROCESSES_METHOD_2
    End If
    pPosDelta = pBuffer
    pNextDelta = 0
    nProcessesCount = 0
LBL_SUB_LOOP_ENUM_1:
    Call RtlMoveMemory(VarPtr(SPRCI), pPosDelta, Len(SPRCI))
    szTmpName = String$(SPRCI.ProcessName.Length, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
    Call RtlMoveMemory(StrPtr(szTmpName), SPRCI.ProcessName.pToBuffer, SPRCI.ProcessName.Length)
    ReDim Preserve OutputProcessesData(nProcessesCount)
    With OutputProcessesData(nProcessesCount)
        .nProcessID = SPRCI.ProcessId
        .nParentProcessID = SPRCI.InheritedFromProcessId
        .nThreadCount = SPRCI.ThreadCount
        .nObjectHandleCount = SPRCI.HandleCount
        .szNtExecutableNameW = szTmpName
    End With
    NtProcessHandle = PamzOpenProcess(OutputProcessesData(nProcessesCount).nProcessID, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    If NtProcessHandle <> 0 Then
        '---asik,utak-atik memori:
        '---cari tahu alamat struktur PEB:
        NtStatus = NtQueryInformationProcess(NtProcessHandle, ProcessBasicInformation, VarPtr(PRCBI), Len(PRCBI), 0)
        If NtStatus <> 0 Then
            With OutputProcessesData(nProcessesCount)
                .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
            End With
            GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End If
        With OutputProcessesData(nProcessesCount)
            .pPEBBaseAddress = PRCBI.PebBaseAddress
        End With
        Call RtlZeroMemory(VarPtr(PEBSC), Len(PEBSC))
        NtStatus = NtReadVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pPEBBaseAddress, VarPtr(PEBSC), Len(PEBSC), 0)
        If NtStatus <> 0 Then
            With OutputProcessesData(nProcessesCount)
                .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
            End With
            GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End If
        With OutputProcessesData(nProcessesCount)
            .pBaseAddress = PEBSC.ImageBaseAddress
            .pLDRAddress = PEBSC.Ldr
            .pProcessParamAddress = PEBSC.ProcessParameters
            .bIsBeingDebugged = PEBSC.BeingDebugged
        End With
        '---cari tahu alamat lengkap executable:
        Call RtlZeroMemory(VarPtr(UNSTI), Len(UNSTI))
        nTmpNameLength = 4096 '---1 blok kecil memori cukupan,lah.
        szTmpName = String$(nTmpNameLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
        pTmpName = StrPtr(szTmpName) '---struktur unicode_string dan buffer data jadi satu.
        NtStatus = NtQueryVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pBaseAddress, MemorySectionName, pTmpName, nTmpNameLength, 0)
        If NtStatus <> 0 Then
            With OutputProcessesData(nProcessesCount)
                .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
            End With
            szTmpName = ""
            '---mencari alamat dilarang,tapi masih ada kemungkinan mengerjakan yang lainnya:
            GoTo LBL_SUB_QUERYLENGTH_ENUM_1
        End If
        Call RtlMoveMemory(VarPtr(UNSTI), pTmpName, Len(UNSTI))
        szTmpName = MidB$(szTmpName, (UNSTI.pToBuffer - pTmpName) + 1, UNSTI.Length)
        With OutputProcessesData(nProcessesCount)
            .szNtExecutablePathW = szTmpName
        End With
LBL_SUB_QUERYLENGTH_ENUM_1:
LBL_SUB_QUERYINMEM_LENGTH_ENUM_1:
        '---mendapatkan ukuran executable di memori:
        Call RtlZeroMemory(VarPtr(PRMBI), Len(PRMBI))
        '---dapatkan region yang pertama (biasanya berisi header DOS n NT, 4096 bytes).
        NtStatus = NtQueryVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pBaseAddress, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
        If NtStatus <> 0 Then
            With OutputProcessesData(nProcessesCount)
                .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
            End With
            GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End If
        With OutputProcessesData(nProcessesCount)
            .nSizeOfExecutableOpInMemory = .nSizeOfExecutableOpInMemory + PRMBI.nRegionSize '---sebenarnya dimulai dari 0 :)
        End With
LBL_SUB_LOOP_QUERYINMEM_LENGTH_ENUM_1: '---ulangi,cari sampai pAllocationBase bernilai beda dari BaseAddress-awal-nya.
        NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI.pBaseAddress + PRMBI.nRegionSize, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
        If NtStatus <> 0 Then
            With OutputProcessesData(nProcessesCount)
                .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
            End With
            GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End If
        If PRMBI.pAllocationBase = OutputProcessesData(nProcessesCount).pBaseAddress Then
            With OutputProcessesData(nProcessesCount)
                .nSizeOfExecutableOpInMemory = .nSizeOfExecutableOpInMemory + PRMBI.nRegionSize '---sebenarnya dimulai dari 0 :)
            End With
            GoTo LBL_SUB_LOOP_QUERYINMEM_LENGTH_ENUM_1
        End If
        '---jangan lupa tutup handle obyek prosesnya:
LBL_SUB_CLOSEHANDLE_ENUM_1:
        Call PamzCloseHandle(NtProcessHandle)
        NtProcessHandle = 0
    Else '---gagal "OpenProcess":
        With OutputProcessesData(nProcessesCount)
            .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
        End With
        '---oh,mengapa dilarang,sih?
    End If
    '---test hasil:
    'MsgBox "TEST_NORMAL_PROCESS_REACHED:" & nProcessesCount & vbTab & OutputProcessesData(nProcessesCount).nProcessID & vbTab & "[" & OutputProcessesData(nProcessesCount).nSizeOfExecutableOpInMemory & "]" & vbTab & "[" & OutputProcessesData(nProcessesCount).szNtExecutablePathW & "]"
    '---rekalkulasi hitungan data yang disimpan:
    nProcessesCount = nProcessesCount + 1
    '---cek untuk entry berikutnya:
    pNextDelta = SPRCI.NextEntryDelta
    If pNextDelta <> 0 Then
        pPosDelta = pPosDelta + pNextDelta
        GoTo LBL_SUB_LOOP_ENUM_1
    End If
    Call RtlZeroMemory(VarPtr(SPRCI), Len(SPRCI))
LBL_ENUM_PROCESSES_METHOD_2: '---via daftar proses yang mempunyai object-handle di memori:
    nBufferLength = 4096 '---1 blok kecil memori dulu.
    szBuffer = String$(nBufferLength, 0) '---isi dengan bytenumber=0. '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
    pBuffer = StrPtr(szBuffer)
LBL_SUB_LOOP_BUFFER_ENUM_2:
    NtStatus = NtQuerySystemInformation(SystemHandleInformation, pBuffer, nBufferLength, VarPtr(nBufferLength))
    Select Case NtStatus
        Case 0 '---berhasil,lanjut ke kode setelah switch.
        Case STATUS_INFO_LENGTH_MISMATCH
            szBuffer = String$(nBufferLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
            pBuffer = StrPtr(szBuffer)
            GoTo LBL_SUB_LOOP_BUFFER_ENUM_2
        Case Else
            szBuffer = ""
            nBufferLength = 0
            GoTo LBL_ENUM_PROCESSES_METHOD_3
    End Select
    Call RtlMoveMemory(VarPtr(nHandleCount), pBuffer, Len(nHandleCount)) '---4bytes awal adalah jumlah object-handle dalam struktur buffer.
    If nHandleCount <= 0 Then
        GoTo LBL_ENUM_PROCESSES_METHOD_3
    End If
    pBuffer = pBuffer + 4 '---setelah dikurangi dengan jumlah object-handle-keseluruhan.
    For CTurn = 0 To (nHandleCount - 1) '---untuk setiap object-handle(Base0).
        Call RtlMoveMemory(VarPtr(SHTEI), pBuffer + (CTurn * Len(SHTEI)), Len(SHTEI))
        bAlreadyAdded = 0
        For DTurn = 0 To (nProcessesCount - 1)
            If SHTEI.UniqueProcessId = OutputProcessesData(DTurn).nProcessID Then
                bAlreadyAdded = 1
                Exit For
            End If
        Next
        If bAlreadyAdded = 0 Then '---ditemukan proses 'non-visible',tapi punya handle :)
            '---tambahkan ke daftar proses:
            ReDim Preserve OutputProcessesData(nProcessesCount)
            '---isi data dalam daftar proses:
            '---masukkan 2 info: PID dan indikasi proses tersembunyi:
            With OutputProcessesData(nProcessesCount)
                .nProcessID = SHTEI.UniqueProcessId '---@@@baru tahu nomor PID-nya. SHTEI cuman dipakai untuk ini.
                .bIsHiddenProcess = 1 '---beri tahu kalau proses termasuk proses yang 'ngumpet'.
            End With
            '---cari tahu informasi lainnya:
            NtProcessHandle = PamzOpenProcess(OutputProcessesData(nProcessesCount).nProcessID, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
            If NtProcessHandle <> 0 Then
                NtStatus = NtQueryInformationProcess(NtProcessHandle, ProcessBasicInformation, VarPtr(PRCBI), Len(PRCBI), 0)
                If NtStatus <> 0 Then
                    With OutputProcessesData(nProcessesCount)
                        .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                    End With
                    GoTo LBL_SUB_CLOSEHANDLE_ENUM_2
                End If
                '---masukkan 2 info: parentPID dan PEBaddress-nya:
                With OutputProcessesData(nProcessesCount)
                    .nParentProcessID = PRCBI.InheritedFromUniqueProcId '---@@@.
                    .pPEBBaseAddress = PRCBI.PebBaseAddress '---@@@.
                End With
                Call RtlZeroMemory(VarPtr(PEBSC), Len(PEBSC))
                NtStatus = NtReadVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pPEBBaseAddress, VarPtr(PEBSC), Len(PEBSC), 0)
                If NtStatus <> 0 Then
                    With OutputProcessesData(nProcessesCount)
                        .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                    End With
                    GoTo LBL_SUB_CLOSEHANDLE_ENUM_2
                End If
                '---masukkan 4 info tambahan:
                With OutputProcessesData(nProcessesCount)
                    .pBaseAddress = PEBSC.ImageBaseAddress '---@@@.
                    .pLDRAddress = PEBSC.Ldr '---@@@.
                    .pProcessParamAddress = PEBSC.ProcessParameters '---@@@.
                    .bIsBeingDebugged = PEBSC.BeingDebugged '---@@@.
                End With
                '---cari tahu alamat lengkap executable:
                Call RtlZeroMemory(VarPtr(UNSTI), Len(UNSTI))
                nTmpNameLength = 4096 '---1 blok kecil memori cukupan,lah.
                szTmpName = String$(nTmpNameLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
                pTmpName = StrPtr(szTmpName) '---struktur unicode_string dan buffer data jadi satu.
                NtStatus = NtQueryVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pBaseAddress, MemorySectionName, pTmpName, nTmpNameLength, 0)
                If NtStatus <> 0 Then
                    With OutputProcessesData(nProcessesCount)
                        .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                    End With
                    szTmpName = ""
                    '---mencari alamat dilarang,tapi masih ada kemungkinan mengerjakan yang lainnya:
                    GoTo LBL_SUB_QUERY_OBJECT_AND_COUNT_2
                End If
                Call RtlMoveMemory(VarPtr(UNSTI), pTmpName, Len(UNSTI))
                szTmpName = MidB$(szTmpName, (UNSTI.pToBuffer - pTmpName) + 1, UNSTI.Length)
                '---masukkan 1 info tambahan:nama lengkap beserta alamat executable:
                With OutputProcessesData(nProcessesCount)
                    .szNtExecutablePathW = szTmpName '---@@@.
                End With
                '---cari tau nama executable-nya,tinggal parse ajah alamat lengkapnya :)
                pTmpName = InStrRev(szTmpName, Chr$(92), , vbTextCompare)
                szTmpName = Right$(szTmpName, (UNSTI.Length / 2) - pTmpName)
                '---masukkan 1 info tambahan:nama lengkap:
                With OutputProcessesData(nProcessesCount)
                    .szNtExecutableNameW = szTmpName '---@@@.
                End With
                '---yg belum: thread count,[*handle count*],[*ukuran PE di memori*]:
LBL_SUB_QUERY_OBJECT_AND_COUNT_2:
                '---kalau pakai NtQuerySystemInformation nggak bakalan dapet,harus pakai cara lain:
                '---dapatkan nomor ID untuk object bertipe 'thread'(direct:soalnya xp,vista,w7,n berikutnya berbeda-beda nomor ID-nya):
                nTmpNameLength = 4096 '---untuk xp,cukupan,lah.
                szTmpName = String$(nTmpNameLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
                pTmpName = StrPtr(szTmpName)
LBL_SUB_LOOP_BUFFER_QUERY_OBJECT_2:
                NtStatus = NtQueryObject(0, ObjectAllTypesInformation, pTmpName, nTmpNameLength, VarPtr(nTmpNameLength))
                Select Case NtStatus
                    Case 0
                    Case STATUS_INFO_LENGTH_MISMATCH
                        szTmpName = String$(nTmpNameLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
                        pTmpName = StrPtr(szTmpName)
                        GoTo LBL_SUB_LOOP_BUFFER_QUERY_OBJECT_2 '---coba lagi.
                End Select
                Call RtlMoveMemory(VarPtr(nObjectHandleTypeCount), pTmpName, Len(nObjectHandleTypeCount))
                DTurn = 0 '---netralisir dulu,lanjut:
                ThreadTypeObjectID = 0 '---netralisir dulu,lanjut:
                pNextTmp = 0 '---netralisir dulu,lanjut:
                pPosTmp = pTmpName + 4 '---4 bytes pertama udah dipakai untuk info jumlah object-handle-types.
                For DTurn = 1 To nObjectHandleTypeCount
                    Call RtlMoveMemory(VarPtr(OBJTI), pPosTmp, Len(OBJTI))
                    '---waduh,variabel string udah terpakai semua, masa' bikin lagi?
                    szTmpName2 = String$(OBJTI.pszName.Length / 2, 0)
                    Call RtlMoveMemory(StrPtr(szTmpName2), OBJTI.pszName.pToBuffer, OBJTI.pszName.Length)
                    'MsgBox DTurn & vbTab & "[" & szTmpName2 & "]"
                    If UCase$(szTmpName2) = "THREAD" Then '---object bertipe Thread.
                        ThreadTypeObjectID = DTurn
                        Exit For
                    End If
                    '---coba lagi berikutnya:(4 bytes tepian):
                    pPosTmp = (OBJTI.pszName.pToBuffer + OBJTI.pszName.MaxLength) + (((OBJTI.pszName.pToBuffer + OBJTI.pszName.MaxLength) - pPosTmp) Mod 4)
                Next
                szTmpName = ""
                szTmpName2 = ""
LBL_SUB_QUERY_OBJECT_AND_COUNT_PROC_2:
                '---start dimulai dari handle yg sekarang:
                DTurn = 0 '---netralisir dulu,lanjut:
                With OutputProcessesData(nProcessesCount)
                    For DTurn = CTurn To (nHandleCount - 1) '---untuk setiap object-handle(Base0).
                        Call RtlMoveMemory(VarPtr(SHTEI), pBuffer + (DTurn * Len(SHTEI)), Len(SHTEI))
                        If SHTEI.UniqueProcessId = .nProcessID Then '---cuman menghitung jumlah object-handle yg ada di dalam proses bersangkutan.
                            .nObjectHandleCount = .nObjectHandleCount + 1 '---hitung,tambahkan ke sini.
                            '---sekarang menghitung estimasi jumlah thread yg bekerja di dalamnya:
                            'If ThreadTypeObjectID > 0 Then '---bila acuan untuk tipe 'thread' diketahui.
                                'If SHTEI.ObjectTypeIndex = ThreadTypeObjectID Then '---handle merupakan handle untuk thread.
                                    'NtStatus = NtDuplicateObject(NtProcessHandle, SHTEI.HandleValue, GetCurrentProcess, VarPtr(hDupObject), 9, 8, 7)
                                    '---:
                                'End If
                            'End If
                        Else '---biasanya urut,jadi bila sudah bukan lagi sama PID-nya,berarti sudah selesai ngumpulin-nya.
                            Exit For
                        End If
                    Next
                    'MsgBox "TEST_OBJECTHANDLE_COUNT:" & vbTab & .nProcessID & vbTab & .nObjectHandleCount
                End With
LBL_SUB_QUERYINMEM_LENGTH_ENUM_2:
                '---mendapatkan ukuran executable di memori:
                Call RtlZeroMemory(VarPtr(PRMBI), Len(PRMBI))
                '---dapatkan region yang pertama (biasanya berisi header DOS n NT, 4096 bytes).
                NtStatus = NtQueryVirtualMemory(NtProcessHandle, OutputProcessesData(nProcessesCount).pBaseAddress, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
                If NtStatus <> 0 Then
                    With OutputProcessesData(nProcessesCount)
                        .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                    End With
                    GoTo LBL_SUB_CLOSEHANDLE_ENUM_2
                End If
                With OutputProcessesData(nProcessesCount)
                    .nSizeOfExecutableOpInMemory = .nSizeOfExecutableOpInMemory + PRMBI.nRegionSize '---sebenarnya dimulai dari 0 :)
                End With
LBL_SUB_LOOP_QUERYINMEM_LENGTH_ENUM_2: '---ulangi,cari sampai pAllocationBase bernilai beda dari BaseAddress-awal-nya.
                NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI.pBaseAddress + PRMBI.nRegionSize, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
                If NtStatus <> 0 Then
                    With OutputProcessesData(nProcessesCount)
                        .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                    End With
                    GoTo LBL_SUB_CLOSEHANDLE_ENUM_2
                End If
                If PRMBI.pAllocationBase = OutputProcessesData(nProcessesCount).pBaseAddress Then
                    With OutputProcessesData(nProcessesCount)
                        .nSizeOfExecutableOpInMemory = .nSizeOfExecutableOpInMemory + PRMBI.nRegionSize '---sebenarnya dimulai dari 0 :)
                    End With
                    GoTo LBL_SUB_LOOP_QUERYINMEM_LENGTH_ENUM_2
                End If
                '---jangan lupa tutup handle obyek prosesnya:
LBL_SUB_CLOSEHANDLE_ENUM_2:
                Call PamzCloseHandle(NtProcessHandle)
                NtProcessHandle = 0
            Else '---gagal "OpenProcess":
                With OutputProcessesData(nProcessesCount)
                    .bIsLockedProcess = 1 '---beri tanda 'terkunci' dulu.
                End With
                '---oh,mengapa dilarang,sih?
            End If
            '---test hasil:
            'MsgBox "TEST_HIDDEN_PROCESS_REACHED:" & nProcessesCount & vbTab & SHTEI.UniqueProcessId & vbTab & "[" & OutputProcessesData(nProcessesCount).nSizeOfExecutableOpInMemory & "]" & vbTab & "[" & OutputProcessesData(nProcessesCount).szNtExecutablePathW & "]"
            '---rekalkulasi hitungan data yang disimpan:
            nProcessesCount = nProcessesCount + 1
        End If
    Next
    
LBL_ENUM_PROCESSES_METHOD_3: '---via window (jendela) handle [nggak usah,nggak terpakai,dibuat fix-up yg udah didapat ajah]:
    
    NtRetValue = nProcessesCount
    
LBL_CLEAN_MEMORY:
    szBuffer = ""
    szTmpName = ""
LBL_BROADCAST_RESULT:
    PamzEnumerateProcesses = NtRetValue
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Public Function PamzEnumerateModules(ByVal nTargetPID As Long, ByRef OutputModulesData() As ENUMERATE_MODULES_OUTPUT) As Long
On Error Resume Next '---kalau result bernilai > 0, berarti menunjukkan jumlah modul yg berjalan. kalau result bernilai <= 0, menunjukkan error value.
Dim OBJAT                           As OBJECT_ATTRIBUTES
Dim CLIID                           As CLIENT_ID
Dim UNSTI                           As UNICODE_STRING
Dim NtStatus                        As Long
Dim NtRetLength                     As Long
Dim NtRetValue                      As Long
Dim NtProcessHandle                 As Long
Dim SPRCI                           As SYSTEM_PROCESSES_INFORMATION
Dim PRCBI                           As PROCESS_BASIC_INFORMATION
Dim PEBSC                           As PEB
Dim PRMBI                           As MEMORY_BASIC_INFORMATION
Dim PRMBI2                          As MEMORY_BASIC_INFORMATION '---nggak efisien,tapi nggak apa-apa,lah.
Dim SYSMI                           As SYSTEM_MODULE_INFORMATION
Dim CTurn                           As Long
Dim DTurn                           As Long
Dim bAlreadyAdded                   As Long
Dim pPosTmp                         As Long
Dim pNextTmp                        As Long
Dim nModulesCount                   As Long
Dim nMinUserAddress                 As Long
Dim nMaxUserAddress                 As Long
Dim pTmpName                        As Long
Dim nTmpNameLength                  As Long
Dim szTmpName                       As String
Dim pBuffer                         As Long
Dim nBufferLength                   As Long
Dim szBuffer                        As String '---pakai string ajah buffernya, nggak ruwet kaya' *Alloc yang lainnya.
    nModulesCount = 0
    NtProcessHandle = PamzOpenProcess(nTargetPID, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    If NtProcessHandle <> 0 Then
        '---asik,utak-atik memori:
        '---kalau pakai versi standar (flink-blink-lib-chain),ternyata windows versi spesifik.
        '---jadi antara windows-xp,vista,w7,berbeda-beda(walau sedikit)strukturnya,membuat
        '---enumerasi menjadi beresiko gagal cukup tinggi.
        '---jadi ari pakai enumerasi memori berdasarkan region aja,lebih lambat tapi lebih (stabil),InsyaAllah:
        '---sebelumnya, cek dulu apakah proses adalah "System" (PID:4 untuk XP,dsb):
        '---kalau "YA", enumerasi driver-nya ajah:
        '---kalau "TIDAK", enumerasi normal:
LBL_CHECK_ENUM_IS_SYSTEM_1:
        If nTargetPID > 0 And nTargetPID < 10 Then '---anggap saja ini.
            pBuffer = 0
            nBufferLength = 0
LBL_CHECK_BUFFER_FOR_IS_SYSTEM_1:
            NtStatus = NtQuerySystemInformation(SystemModuleInformation, pBuffer, nBufferLength, VarPtr(nBufferLength))
            Select Case NtStatus
                Case 0
                Case STATUS_INFO_LENGTH_MISMATCH
                    szBuffer = String$(nBufferLength, 0)
                    pBuffer = StrPtr(szBuffer)
                    GoTo LBL_CHECK_BUFFER_FOR_IS_SYSTEM_1
                Case Else
                    GoTo LBL_ENUM_MODULES_METHOD_1
            End Select
            '---parsing isi buffer menjadi daftar module n diver:
            Call RtlMoveMemory(VarPtr(NtRetValue), pBuffer, Len(NtRetValue))
            pBuffer = pBuffer + 4 '---udah dikurangi dgn info banyaknya modul n driver.
            For CTurn = 0 To (NtRetValue - 1)
                Call RtlMoveMemory(VarPtr(SYSMI), pBuffer + (CTurn * Len(SYSMI)), Len(SYSMI))
                ReDim Preserve OutputModulesData(nModulesCount)
                '---isi data dalam daftar modul:
                With OutputModulesData(nModulesCount)
                    .pBaseAddress = SYSMI.Base '---AWAS!BaseAddress yg lebih besar dari &H80000000 (kernel-mode-range) tidak dapat diakses n di-dump dari program yg berjalan di user-mode, tanpa bantuan driver atau mengakses langsung PhysicalMemory.
                    .nSizeOfModuleOpInMemory = SYSMI.SIZE
                    szTmpName = StrConv(SYSMI.ImageName, vbUnicode) '---jadikan format unicode dgn banyak nullchars.
                    pTmpName = InStr(1, szTmpName, Chr$(0), vbTextCompare)
                    If pTmpName > 0 Then
                        szTmpName = Left$(szTmpName, pTmpName - 1) '---nullchars nggak usah ikut.
                    End If
                    .szNtModulePathW = szTmpName '---karena data berupa Ansi,harus di-ubah dulu menjadi unicode standar.
                    .szNtModuleNameW = Right$(szTmpName, Len(szTmpName) - SYSMI.ModuleNameOffset)
                End With
                '---rekalkulasi hitungan data yang disimpan:
                nModulesCount = nModulesCount + 1
            Next
            '---bila berhasil mengumpulkan modul milik system, langsung tutup handle:
            GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End If
LBL_ENUM_MODULES_METHOD_1:
LBL_LOOP_PREPARE_ENUM_1:
        '---cari tahu nama lengkap region:
        nBufferLength = 4096 '---1 blok (4096) kecil memori cukupan,lah.
        szBuffer = String$(nBufferLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
        pBuffer = StrPtr(szBuffer) '---struktur unicode_string dan buffer data jadi satu.
        '---nomor error yg jadi patokan:
        '-1073741503,    "The address handle given to the transport was invalid." --->error:region_invalid.
        '-1073741811,    "An invalid parameter was passed to a service or function." --->error:region_bukan_user_address.
LBL_LOOP_NEXT_ENUM_1:
        '---dapatkan struktur region:
        NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI.pBaseAddress + PRMBI.nRegionSize, MemoryBasicInformation, VarPtr(PRMBI), Len(PRMBI), 0)
        Select Case NtStatus
            Case 0 '---berhasil mendapatkan info.
                '---cari nama untuk region yg bukan "Free" ajah:
                If PRMBI.pAllocationBase <> 0 Then
                    If PRMBI.pAllocationBase <> pPosTmp Then '---address yg sekarang tidak terletak pada region yg sama dgn sebelumnya:
                        pPosTmp = PRMBI.pAllocationBase '---ingat-ingat region yg sekarang.
                        '---cari tahu nama region:
                        Call RtlZeroMemory(pBuffer, nBufferLength)
                        NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI.pBaseAddress, MemorySectionName, pBuffer, nBufferLength, 0)
                        If NtStatus = 0 Then
                            Call RtlMoveMemory(VarPtr(UNSTI), pBuffer, Len(UNSTI))
                            If UNSTI.Length > 0 Then
                                szTmpName = MidB$(szBuffer, (UNSTI.pToBuffer - pBuffer) + 1, UNSTI.Length)
                                ReDim Preserve OutputModulesData(nModulesCount)
                                '---isi data dalam daftar modul:
                                With OutputModulesData(nModulesCount)
                                    .pBaseAddress = PRMBI.pAllocationBase
                                    .szNtModulePathW = szTmpName
                                End With
                                '---cari nama inisialnya,dan simpan:
                                pTmpName = InStrRev(szTmpName, Chr$(92), , vbTextCompare)
                                szTmpName = Right$(szTmpName, (UNSTI.Length / 2) - pTmpName)
                                With OutputModulesData(nModulesCount)
                                    .szNtModuleNameW = szTmpName
                                End With
                                '---cari ukuran di memori,dan simpan:
                                Call RtlZeroMemory(VarPtr(PRMBI2), Len(PRMBI2))
                                PRMBI2.pBaseAddress = PRMBI.pAllocationBase
                                PRMBI2.nRegionSize = 0
LBL_LOOP_MODULE_SIZE_SUM_1:
                                NtStatus = NtQueryVirtualMemory(NtProcessHandle, PRMBI2.pBaseAddress + PRMBI2.nRegionSize, MemoryBasicInformation, VarPtr(PRMBI2), Len(PRMBI2), 0)
                                If NtStatus = 0 Then
                                    If PRMBI2.pAllocationBase = PRMBI.pAllocationBase Then
                                        With OutputModulesData(nModulesCount)
                                            .nSizeOfModuleOpInMemory = .nSizeOfModuleOpInMemory + PRMBI2.nRegionSize
                                        End With
                                        GoTo LBL_LOOP_MODULE_SIZE_SUM_1
                                    End If
                                End If
                                '---rekalkulasi hitungan data yang disimpan:
                                nModulesCount = nModulesCount + 1
                            End If
                        End If
                    End If
                End If
                GoTo LBL_LOOP_NEXT_ENUM_1
            Case STATUS_INVALID_PARAMETER_IN_SERVICE
                GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
            Case Else
                GoTo LBL_SUB_CLOSEHANDLE_ENUM_1
        End Select
        
LBL_SUB_CLOSEHANDLE_ENUM_1:
        Call PamzCloseHandle(NtProcessHandle)
        NtProcessHandle = 0
    Else
        nModulesCount = 0
    End If
    szTmpName = ""
    szBuffer = ""
LBL_ENUM_MODULES_METHOD_2:

LBL_CLEAN_MEMORY:
    szBuffer = ""
    szTmpName = ""
LBL_BROADCAST_RESULT:
    NtRetValue = nModulesCount
    PamzEnumerateModules = NtRetValue
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

'++PLUS_PLUS++:jangan anggap serius fungsi-fungsi berikut :)
Public Function PamzForceUnLoadProcessModule32(ByVal nTargetPID As Long, ByVal nTargetModuleBaseAddress As Long) As Long
On Error Resume Next
'---awas!apabila proses masih mempunyai kode yg mengakses wilayah modul yg mau di-unload, proses akan mengalami 'crash'.
'---fungsi ini tidak berlaku bagi modul yg terkait dengan punya sistem, misal: ntdll,kernel32,user32,gdi32,dan lainnya.
Dim NtRetValue                      As Long
Dim NtProcessHandle                 As Long
Dim KeLibModulePtr                  As Long
Dim KeFreeLibCodePtr                As Long
Dim UniLongVar                      As Long
Dim hHitThread                      As Long
Dim nHitThreadID                    As Long
Dim PRelCodeAddress                 As Long
Dim PInjectAddress                  As Long
Dim PreCastValue                    As Long
Dim PreCastPointer                  As Long
Dim PreCastLength                   As Long
Dim pTmpName                        As Long
Dim nTmpNameLength                  As Long
Dim szTmpName                       As String
LBL_UNLOAD_MODULE_METHOD_1: '---pakai cara klasik, injeksikan kode ke dalam proses target, lalu jalankan kode sebagai thread baru:
DoEvents
    NtProcessHandle = PamzOpenProcess(nTargetPID, PROCESS_QUERY_INFORMATION Or PROCESS_CREATE_THREAD Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE)
    If NtProcessHandle <= 0 Then
        GoTo LBL_UNLOAD_MODULE_METHOD_2
    End If
    '---cari tahu base address "kernel32.dll" dan codeptr dari "FreeLibrary" (alamat sama untuk setiap proses, di semua proses):
    KeLibModulePtr = GetModuleHandleW(StrPtr("kernel32"))
    If KeLibModulePtr <= 0 Then
        GoTo LBL_SUB_FREEOBJECTS_1
    End If
    KeFreeLibCodePtr = GetProcAddress(KeLibModulePtr, "FreeLibrary")
    If KeFreeLibCodePtr <= 0 Then
        GoTo LBL_SUB_FREEOBJECTS_1
    End If
    '---ini adalah kode paling tidak efisien buatan saya :)
    '---alokasikan cukup memori dari target untuk disisipi kode:
    PInjectAddress = VirtualAllocEx(NtProcessHandle, 0, 4096, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If PInjectAddress = 0 Then
        GoTo LBL_SUB_FREEOBJECTS_1
    End If
    '---persiapkan struktur sementara:
    ReDim AsmCodeInjectInStc(128) As Byte '---akan diisi struktur kode pelepas modul sederhana.
    PreCastPointer = 0 '---preset,dimulai dari 0.
    '---isi struktur dengan kode:
    '------------------------(01):
    PreCastLength = 2
    PreCastValue = &HFF8B 'mov edi, edi
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(02):
    PreCastLength = 1
    PreCastValue = &H55 'push ebp
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(03):
    PreCastLength = 2
    PreCastValue = &HEC8B 'mov ebp, esp
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(04):
    PreCastLength = 1
    PreCastValue = &HB8 'mov eax, ...
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(05):
    PreCastLength = 4
    PreCastValue = nTargetModuleBaseAddress '---berisi nilai 32 bit alamat: baseaddress dll yg mau di-unload.
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(06):
    PreCastLength = 1
    PreCastValue = &H50 'push eax
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(07):
    PreCastLength = 1
    PreCastValue = &HE8 'call ...
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(08):
    PreCastLength = 4
    '---cari alamat relatifnya:
    PRelCodeAddress = KeFreeLibCodePtr - (PInjectAddress + PreCastPointer) - PreCastLength
    PreCastValue = PRelCodeAddress '---berisi nilai 32 bit alamat: alamat relatif kode fungsi "FreeLibrary".
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(88):
    PreCastLength = 1
    PreCastValue = &H5D 'pop ebp
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '------------------------(99):
    PreCastLength = 3
    PreCastValue = &H4C2 'ret 4
    Call RtlMoveMemory(VarPtr(AsmCodeInjectInStc(PreCastPointer)), VarPtr(PreCastValue), PreCastLength)
    PreCastPointer = PreCastPointer + PreCastLength
    '---coba tulis kode ke target proses:
    If NtWriteVirtualMemory(NtProcessHandle, PInjectAddress, VarPtr(AsmCodeInjectInStc(0)), UBound(AsmCodeInjectInStc) + 1, 0) <> 0 Then
        GoTo LBL_SUB_FREEOBJECTS_1
    End If
    '---coba jalankan kode:
    nTmpNameLength = 4096
    szTmpName = String$(nTmpNameLength, 0)
    pTmpName = StrPtr(szTmpName)
    '---cek apakah modul masih ada?:
    If NtQueryVirtualMemory(NtProcessHandle, nTargetModuleBaseAddress, MemorySectionName, pTmpName, nTmpNameLength, 0) <> 0 Then
        NtRetValue = 0 '----sepertinya nggak ada.
        GoTo LBL_SUB_FREEOBJECTS_1
    End If
    UniLongVar = 0
    For UniLongVar = 1 To 65536 'Base0=65535'---kata-nya,maksimum increment untuk *.dll adalah segitu, jadi decrement-nya juga segitu :) awas, bisa menimbulkan efek 'hung' (maks. 3 menit) bila *.dll di-load mulai startup program.
        hHitThread = CreateRemoteThread(NtProcessHandle, 0, 0, PInjectAddress, 1, 0, nHitThreadID)
        If hHitThread > 0 Then
            Call WaitForSingleObject(hHitThread, 1000) '---tungguin 1 detik. biasanya tidak terpakai, karena kecepatan eksekusi kode?.
            '---cek apakah modul masih ada?:
            If NtQueryVirtualMemory(NtProcessHandle, nTargetModuleBaseAddress, MemorySectionName, pTmpName, nTmpNameLength, 0) <> 0 Then
                NtRetValue = 1 '----sepertinya berhasil.
                Exit For
            End If
            '---tutup handle ke thread:
            Call NtClose(hHitThread)
            hHitThread = 0
        Else
            Exit For
        End If
    Next
    If NtRetValue > 0 Then
        NtRetValue = UniLongVar '---beritahu coba dihentikan sebanyak berapa kali.
    End If
    'MsgBox "MATI PADA HITUNGAN: " & vbTab & UniLongVar & vbTab & NtRetValue
LBL_SUB_FREEOBJECTS_1:
    If PInjectAddress > 0 Then
        UniLongVar = VirtualFreeEx(NtProcessHandle, PInjectAddress, 4096, MEM_DECOMMIT)
        'MsgBox "DECOMMIT:" & vbTab & Hex$(UniLongVar)
        UniLongVar = VirtualFreeEx(NtProcessHandle, PInjectAddress, 0, MEM_RELEASE)
        'MsgBox "RELEASE:" & vbTab & Hex$(UniLongVar)
    End If
    If NtProcessHandle > 0 Then
        Call PamzCloseHandle(NtProcessHandle)
        NtProcessHandle = 0
    End If
    szTmpName = "" '---hapus buffer.
    If NtRetValue > 0 Then
        GoTo LBL_CLEAN_MEMORY
    End If
LBL_UNLOAD_MODULE_METHOD_2:

LBL_CLEAN_MEMORY:
    Erase AsmCodeInjectInStc()
LBL_BROADCAST_RESULT:
    PamzForceUnLoadProcessModule32 = NtRetValue
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function


Public Function PamzNtPathToUserFriendlyPathW(ByRef szTargetNtPath As String) As String
On Error Resume Next '---dari NtPath(=\Device\...) menjadi DosPath(=C:\...).[UserFriendly=yang mudah dipahami oleh user].
Dim CTurn                   As Long
Dim szDOSDevice             As String
Dim szNTDevice              As String
Dim szBuffer                As String
Dim pBuffer                 As Long
Dim nBuffer                 As Long
Dim nRetLength              As Long
Dim Gotcha                  As Long
    nBuffer = 128
    szBuffer = String$(nBuffer, 0)
    pBuffer = StrPtr(szBuffer)
    For CTurn = 0 To 25 '---device standar untuk disk : A-Z ajah.
        szDOSDevice = Chr$(CByte(65 + CTurn)) & ":"
        nRetLength = QueryDosDeviceW(StrPtr(szDOSDevice), pBuffer, nBuffer)
        If nRetLength > 0 Then
            szNTDevice = Left$(szBuffer, nRetLength - 2)
            'MsgBox "[" & szDOSDevice & "]" & vbTab & "[" & szNTDevice & "]"
            If UCase$(Left$(szTargetNtPath, Len(szNTDevice))) = UCase$(szNTDevice) Then
                PamzNtPathToUserFriendlyPathW = szDOSDevice & Right$(szTargetNtPath, Len(szTargetNtPath) - Len(szNTDevice))
                Gotcha = 1
            End If
        End If
        Call RtlZeroMemory(pBuffer, nBuffer)
        If Gotcha > 0 Then
            Exit For
        End If
    Next
    If Gotcha <= 0 Then
        PamzNtPathToUserFriendlyPathW = szTargetNtPath
    End If
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function


Public Function App_FullPathW(Optional ByVal bIsInIDE As Boolean) As String
On Error Resume Next '---dari NtPath(=\Device\...) menjadi DosPath(=C:\...).[UserFriendly=yang mudah dipahami oleh user].
Dim hBaseAddress            As Long
Dim nTmpNameLength          As Long
Dim pTmpName                As Long
Dim NtStatus                As Long
Dim UNSTI                   As UNICODE_STRING
Dim szTmpName               As String
    If bIsInIDE = True Then '---saat debug di IDE vb6,memang hanya mendukung AnsiPath.
        App_FullPathW = AddSlashW(App.path) & App.EXEName '---ansi dalam format memori unicode.
        GoTo LBL_TERAKHIR
    End If
    hBaseAddress = GetModuleHandleW(0)
    nTmpNameLength = 1024 '---1/2 blok kecil memori cukupan,lah.
        szTmpName = String$(nTmpNameLength, 0) '---jadi 2 * Length, nggak-papa,lah,daripada ngitung lagi.
        pTmpName = StrPtr(szTmpName) '---struktur unicode_string dan buffer data jadi satu.
        NtStatus = NtQueryVirtualMemory(GetCurrentProcess, hBaseAddress, MemorySectionName, pTmpName, nTmpNameLength, 0)
        If NtStatus = 0 Then
            Call RtlMoveMemory(VarPtr(UNSTI), pTmpName, Len(UNSTI))
            szTmpName = MidB$(szTmpName, (UNSTI.pToBuffer - pTmpName) + 1, UNSTI.Length)
            App_FullPathW = PamzNtPathToUserFriendlyPathW(szTmpName) '---unicode.
        Else
            App_FullPathW = AddSlashW(App.path) & App.EXEName '---ansi dalam format memori unicode.
        End If
LBL_TERAKHIR:
    If erR.Number > 0 Then
        erR.Clear
    End If
End Function

Private Function AddSlashW(ByVal StrInW As String) As String 'OK
On Error Resume Next    'tambah "\" di sebelah kanan string unicode.
    If Right$(StrInW, 1) <> ChrW$(92) Then
        AddSlashW = StrInW & ChrW$(92) 'unicode string;
    Else
        AddSlashW = StrInW
    End If
    erR.Clear
End Function


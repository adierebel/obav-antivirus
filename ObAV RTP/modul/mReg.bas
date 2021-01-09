Attribute VB_Name = "bsReg"
Option Explicit
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Private Declare Function RegCreateKeyUn Lib "advapi32.dll" Alias "RegCreateKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExUn Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValueUn Lib "advapi32.dll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As Long) As Long
Private Declare Function RegOpenKeyUn Lib "advapi32.dll" Alias "RegOpenKeyW" (ByVal hKey As Long, ByVal lpSubKey As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExUn Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegOpenKeyExUn Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValueUn Lib "advapi32.dll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const BUFFER_SIZE As Long = 255
Private Const Pathkey  As String = "Software\Microsoft\Windows\CurrentVersion\Run"
Dim RET As Long, Result As Long

Private Const REG_SZ As Long = 1
Private Const REG_DWORD = 4

Public Const ERROR_SUCCESS As Long = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Private Enum Key
    a = HKEY_CURRENT_USER
    B = HKEY_LOCAL_MACHINE
End Enum

Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const SPI_SETNONCLIENTMETRICS As Long = &H2A
Private Const SMTO_ABORTIFHUNG As Long = &H2

Private Type SECURITY_ATTRIBUTES
   nLength                 As Long
   lpSecurityDescriptor    As Long
   bInheritHandle          As Long
End Type
Public Function isADMIN() As Boolean
    isADMIN = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
End Function
Public Sub SaveString(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal strdata As String)
Dim keyhand As Long
strdata = BuffUni(strdata)
RegCreateKeyUn hKey, StrPtr(strPath), keyhand
RegSetValueExUn keyhand, StrPtr(strValue), 0, REG_SZ, ByVal StrPtr(strdata), Len(strdata)
RegCloseKey keyhand
End Sub
Public Sub SaveDworD(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
lRegResult = RegCloseKey(hCurKey)
End Sub

Private Function BuffUni(STR As String) As String
    BuffUni = STR & String$(Len(STR), ";")
End Function
Public Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim RET As Long
    RegCreateKeyUn hKey, StrPtr(strPath), RET
    RegDeleteValueUn RET, StrPtr(strValue)
    RegCloseKey RET
End Sub
Public Function GetStringW(hKey As Long, strPath As String, strValue As String) As String
    Dim lngValueType As Long
    Dim strBuffer As String
    Dim lngDataBufferSize As Long
    Dim hCurKey As Long
    Dim intZeroPos As Long
    RegOpenKeyUn hKey, StrPtr(strPath), hCurKey
    RegQueryValueExUn hCurKey, StrPtr(strValue), 0&, lngValueType, ByVal 0&, lngDataBufferSize
    If lngValueType = REG_SZ Then
        strBuffer = String(lngDataBufferSize, " ")
        RegQueryValueExUn hCurKey, StrPtr(strValue), 0&, 0&, ByVal StrPtr(strBuffer), lngDataBufferSize
        intZeroPos = InStr(strBuffer, ChrW$(0))
        If intZeroPos > 0 Then
            GetStringW = VBA.Left$(strBuffer, intZeroPos - 1)
        Else
            GetStringW = strBuffer
        End If
    End If
    GetStringW = Trim$(GetStringW)
    RegCloseKey hCurKey
End Function
Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long = 0) As Long
On Error Resume Next
Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

lRegResult = RegOpenKeyUn(hKey, StrPtr(strPath), hCurKey)
lDataBufferSize = 4

lRegResult = RegQueryValueExUn(hCurKey, StrPtr(strValue), 0&, lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then
  If lValueType = REG_DWORD Then
    GetSettingLong = lBuffer
  End If
End If
lRegResult = RegCloseKey(hCurKey)
End Function

Sub ForceCacheRefresh()
Dim hKey As Long, dwKeyType As Long, dwDataType As Long, dwDataSize As Long, sKeyName As String, sValue As String, sData As String, sDataRet As String

Dim tmp As Long
Dim sNewValue As String
Dim dwNewValue As Long
Dim success As Long
   dwKeyType = HKEY_CURRENT_USER
   sKeyName = "Control Panel\Desktop\WindowMetrics"
   sValue = "Shell Icon Size"
   hKey = RegKeyOpen(HKEY_CURRENT_USER, sKeyName)
   If hKey <> 0 Then
      dwDataSize = RegGetStringSize(ByVal hKey, sValue, dwDataType)
      If dwDataSize > 0 Then
         sDataRet = RegGetStringValue(hKey, sValue, dwDataSize)
         If sDataRet > "" Then
            tmp = CLng(sDataRet)
            tmp = tmp - 1
            sNewValue = CStr(tmp) & Chr$(0)
            dwNewValue = Len(sNewValue)
            If RegWriteStringValue(hKey, sValue, dwDataType, sNewValue) = ERROR_SUCCESS Then
            
               Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, success)
               sDataRet = sDataRet & Chr$(0)
               Call RegWriteStringValue(hKey, sValue, dwDataType, sDataRet)
               Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, success)
               
            End If   'If RegWriteStringValue
         End If   'If sDataRet > ""
      End If   'If dwDataSize > 0
   End If   'If hKey > 0
   Call RegCloseKey(hKey)
End Sub

Function RegKeyOpen(dwKeyType As Long, sKeyPath As String) As Long
Dim hKey As Long
Dim dwOptions As Long
Dim SA As SECURITY_ATTRIBUTES
   SA.nLength = Len(SA)
   SA.bInheritHandle = False
   dwOptions = 0&
   If RegOpenKeyEx(dwKeyType, sKeyPath, dwOptions, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
      RegKeyOpen = hKey
   End If
End Function

Function RegGetStringSize(ByVal hKey As Long, ByVal sValue As String, dwDataType As Long) As Long
Dim success As Long
Dim dwDataSize As Long
success = RegQueryValueEx(hKey, sValue, 0&, dwDataType, ByVal 0&, dwDataSize)
   If success = ERROR_SUCCESS Then
    If dwDataType = REG_SZ Then
    RegGetStringSize = dwDataSize
    End If
   End If
End Function

Function RegGetStringValue(ByVal hKey As Long, ByVal sValue As String, dwDataSize As Long) As String
Dim sDataRet As String
Dim dwDataRet As Long
Dim success As Long
Dim POS As Long
sDataRet = Space$(dwDataSize)
dwDataRet = Len(sDataRet)
success = RegQueryValueEx(hKey, sValue, ByVal 0&, dwDataSize, ByVal sDataRet, dwDataRet)
   If success = ERROR_SUCCESS Then
    If dwDataRet > 0 Then
    POS = InStr(sDataRet, Chr$(0))
    RegGetStringValue = VBA.Left$(sDataRet, POS - 1)
    End If
   End If
End Function

Function RegWriteStringValue(ByVal hKey, ByVal sValue, ByVal dwDataType, sNewValue) As Long
Dim success As Long
Dim dwNewValue As Long
dwNewValue = Len(sNewValue)
If dwNewValue > 0 Then
    RegWriteStringValue = RegSetValueExString(hKey, sValue, 0&, dwDataType, sNewValue, dwNewValue)
End If
End Function
Public Sub scanStartUP()
scanS a
scanS B
End Sub
Private Sub scanS(Start As Key)
Dim CNT As Long, buf As String, Buf2 As String, retdata As Long, typ As Long
On Error Resume Next
Dim KeyName As String
Dim KeyPath As String
Dim pjG As Long
Dim lis As Object
buf = Space(BUFFER_SIZE)
Buf2 = Space(BUFFER_SIZE)
RET = BUFFER_SIZE
retdata = BUFFER_SIZE
CNT = 0
RegOpenKeyExUn Start, StrPtr(Pathkey), 0, KEY_ALL_ACCESS, Result
While RegEnumValueUn(Result, CNT, StrPtr(buf), RET, 0, typ, ByVal StrPtr(Buf2), retdata) <> ERROR_NO_MORE_ITEMS
    If typ = REG_DWORD Then
        KeyName = Left(buf, RET)
        If Trim$(Buf2) <> "" Then KeyPath = FixStartP(Left$(Asc(Buf2), retdata - 1))
    Else
        KeyName = Left(buf, RET)
        If Trim$(Buf2) <> "" Then KeyPath = FixStartP(Left$(Buf2, retdata - 1))
    End If
    
    LokasiD = "startup => " & KeyPath
    pjG = scE.lenFileEX(KeyPath)
    If pjG <= 2 Then GoTo lWT
    If pjG > 1750000 Then GoTo lWT
    If CekVirus(KeyPath, pjG) = True Then
        Set lis = FrmScanner.ucListView1(1).ListItems.Add(, KeyName)
        If Start = a Then
        lis.SubItem(2).Text = "a" & Pathkey
        Else
        lis.SubItem(2).Text = "b" & Pathkey
        End If
        lis.SubItem(3).Text = "virus startup"
        Set lis = Nothing
        registriBad = registriBad + 1
        FrmScanner.Label1(2).Caption = "Registry (" & registriBad & ")"
    End If
lWT:
    CNT = CNT + 1
    buf = Space$(BUFFER_SIZE)
    Buf2 = Space$(BUFFER_SIZE)
    RET = BUFFER_SIZE
    retdata = BUFFER_SIZE
DoEvents
Wend
RegCloseKey Result
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "axMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Const FILE_MAP_ALL_ACCESS = &HF001F
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_WRITECOPY = &H8

Private Const ERROR_ALREADY_EXISTS = 183&

Dim ptrShare As Long
Dim lngStore() As Byte
Dim hFile As Long
Sub Poke(strdata As String)
Dim a As Long
Dim X As Long
    a = Len(strdata)
    ReDim lngStore(a)
    For X = 0 To a - 1
        lngStore(X) = AscW(Mid$(strdata, X + 1, 1))
    Next
    CopyMemory ByVal ptrShare, a, 4
    CopyMemory ByVal (ptrShare + 4), lngStore(0), a
End Sub
Function Peek() As String
Dim a As Long
Dim X As Long
Dim strdata As String

CopyMemory a, ByVal ptrShare, 4
ReDim lngStore(a)
    CopyMemory lngStore(0), ByVal (ptrShare + 4), a
        For X = 0 To a
            strdata = strdata & ChrW$(lngStore(X))
        Next
    Peek = strdata
End Function
Function OpenMemory(strName As String) As Boolean
Dim e As Long
Dim a As Long
hFile = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0&, 65535, strName)
e = erR.LastDllError

If hFile Then
    ptrShare = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0&, 0&, 0&)
        If ptrShare <> 0 Then
        
                If e <> ERROR_ALREADY_EXISTS Then
                    CopyMemory ByVal ptrShare, a, 4
                End If
        Else
            MsgBox "Unable to map view of memory"
            OpenMemory = False
            Exit Function
        End If
Else
    MsgBox "Unable to get memory map handle."
    OpenMemory = False
    Exit Function
End If

OpenMemory = True
End Function

Sub CloseMemory()
    UnmapViewOfFile ptrShare
    CloseHandle hFile
End Sub

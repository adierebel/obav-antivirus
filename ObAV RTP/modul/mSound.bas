Attribute VB_Name = "mSound"
Option Explicit
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8

Public SoundBuffer() As Byte
Public var_ClassID As Boolean

Sub ExtractResource(nID, nType, nFileName As String)
    On Error Resume Next
    Kill nFileName
    Dim Buffer() As Byte
    Buffer() = LoadResData(nID, nType)
    Open nFileName For Binary As #1
        Put #1, , Buffer
    Close #1
End Sub


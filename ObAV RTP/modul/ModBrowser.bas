Attribute VB_Name = "ModBrowser"
'modul buka browser
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public Sub OpenURL(situs As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, situs, vbNullString, vbNullString, 1)
End Sub




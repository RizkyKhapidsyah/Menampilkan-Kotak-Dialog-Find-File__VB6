Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
Const SW_SHOW = 5

Public Sub ShowFindDialog(Optional InitialDirectory As String)

ShellExecute 0, "find", IIf(InitialDirectory = "", "", InitialDirectory), vbNullString, vbNullString, SW_SHOW
End Sub



Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Function OpenPassword(DBPath As String, _
Password As String) As Object


 'DBPath: Trayectoria de la base de datos Access
 'Password: La Password

  On Error Resume Next
  Dim db As DAO.Database

  Set db = DAO.OpenDatabase(DBPath, False, False, _
  ";pwd=" & Password)

  If Err.Number = 0 Then
'    Global Ruta As String
    
    Set OpenPassword = db
  Else
   'Set OpenPassword = Nothing
   MsgBox ("No se logro abrir la base de datos"), vbCritical
   End
  End If

End Function


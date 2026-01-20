VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' No olvides visitar mi WebPage http://www.vbexperto.com


Const REG_DWORD = 4   ' Número de 32 bits
Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long

Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegSetValueExLong Ret, strValue, 0&, REG_DWORD, CByte(strData), 4&
    RegCloseKey Ret
End Sub

Private Sub Command1_Click()
  
  Dim strString As String
  If Command1.Caption = "Deshabilitar el Administrador de Tareas" Then
   strString = "1"
   Command1.Caption = "Habilitar el Administrador de Tareas"
  Else
   strString = "0"
   Command1.Caption = "Deshabilitar el Administrador de Tareas"
  End If
  
  SaveStringLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", CByte(strString)
  
  MsgBox ("Presiona las teclas Ctrl+Alt+Supr para que notes el cambio"), vbInformation, "Visita: www.vbexperto.com"
End Sub

Private Sub Form_Load()
 Command1.Caption = "Deshabilitar el Administrador de Tareas"
End Sub

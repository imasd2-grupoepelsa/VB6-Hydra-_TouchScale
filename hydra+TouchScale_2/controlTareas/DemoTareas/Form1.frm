VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DemoTareas © Septiembre 2002 by RAMM"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Mini Panel de Control"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5775
      Begin VB.OptionButton Option6 
         Caption         =   "Propiedades del sistema"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Propiedades del teclado"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Propiedades del Mouse"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Configuracion regional y de idioma"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Propiedades de la pantalla"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Abrir"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opciones de accesibilidad"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " ... y de esta manera puedes activar muchas otras opciones del panel de control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   3735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Icono en el SystemTray "
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   5775
      Begin VB.CommandButton Command7 
         Caption         =   "Colocar"
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aqui te muestro como colocar el icono de tu proyecto en el SystemTray (al lado del reloj de la barra de tareas)"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abrir el navegador en una URL especifica"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Abrir"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "http://www.lawebdelprogramador.com"
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Abrir / Cerrar la puerta del CD"
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Abrir"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Abrir una base de datos protegida por password"
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   5775
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         Height          =   315
         Left            =   4680
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Antes de probar el codigo porque mejor no intentas abrirla desde access ..."
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":08CA
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DemoTareas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "Form1.frx":0965
      Top             =   5160
      Width           =   480
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu Acerca 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
 End Type
 Private Const NIM_ADD = &H0
 Private Const NIM_MODIFY = &H1
 Private Const NIM_DELETE = &H2
 Private Const WM_MOUSEMOVE = &H200
 Private Const NIF_MESSAGE = &H1
 Private Const NIF_ICON = &H2
 Private Const NIF_TIP = &H4
 Private Const WM_LBUTTONDBLCLK = &H203
 Private Const WM_LBUTTONDOWN = &H201
 Private Const WM_LBUTTONUP = &H202
 Private Const WM_RBUTTONDBLCLK = &H206
 Private Const WM_RBUTTONDOWN = &H204
 Private Const WM_RBUTTONUP = &H205
 Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 Dim Accion As NOTIFYICONDATA

Dim Datos As Recordset
Dim BaseDatos As Database
Dim Cadena  As String

Private Sub Acerca_Click()
 Form2.Show
End Sub

Private Sub Command1_Click()
 If Text1.Text <> "" Then
  Dim Resultado As Long
  res = ShellExecute(Me.hwnd, "Open", Text1.Text, "", "", 1)
 End If
End Sub

Private Sub Command2_Click()
 Dim Resultado As Long, returnstring As String * 127
 Resultado = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Private Sub Command3_Click()
 Dim Resultado As Long, returnstring As String * 127
 Resultado = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub

Private Sub Command4_Click()
 Select Case True
  Case Option1.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL access.cpl @0,1")
  Case Option2.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL desk.cpl @0,1")
  Case Option3.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL intl.cpl @0,1")
  Case Option4.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL main.cpl @0,1")
  Case Option5.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL main.cpl @1,1")
  Case Option6.Value
   Resultado = Shell("Rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @0,0")
  
  Case Else
   'megbox ("Seleccione una opcion"), vbi
 End Select
End Sub

Private Sub Command5_Click()
 Dim aux1 As Variant
 
 Cadena = App.Path & "\Sistema.mdb"
 aux1 = OpenPassword(Cadena, "sistema")
 
 Set BaseDatos = OpenDatabase(Cadena)
 Set Datos = BaseDatos.OpenRecordset("Tabla1")
 Datos.MoveFirst
 Label2.Caption = Datos(1)
 Label2.ForeColor = &HFF0000
 Label3.Caption = ""
End Sub


Private Sub Command6_Click()
 End
End Sub


Private Sub Command7_Click()
 If Command7.Caption = "Colocar" Then
  If App.PrevInstance = False Then
       Accion.cbSize = Len(Accion)
       Accion.hwnd = Form1.hwnd
       Accion.uId = vbNull
       Accion.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
       Accion.uCallBackMessage = WM_MOUSEMOVE
       Accion.hIcon = Form1.Icon
       Accion.szTip = "Espero te sea util, presiona el boton derecho del mouse" & vbNullChar
       Shell_NotifyIcon NIM_ADD, Accion
       Command7.Caption = "Eliminar"
       Label6.Caption = "QUE TE PARECE? Ahora mueve el mouse sobre el icono"
       Label6.ForeColor = &HFF0000
  End If
 Else
  Shell_NotifyIcon NIM_DELETE, Accion
  Command7.Caption = "Colocar"
  Label6.Caption = "Aqui te muestro como colocar el icono de tu proyecto en el SystemTray (al lado del reloj de la barra de tareas)"
  Label6.ForeColor = &H80000012
 End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lMsg As Long
   Dim sFilter As String
   lMsg = X / Screen.TwipsPerPixelX
   Select Case lMsg
   'you can play with other events as I did as per your use
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONUP
      Case WM_LBUTTONDBLCLK
      Case WM_RBUTTONDOWN
       PopupMenu Archivo
      Case WM_RBUTTONUP
      Case WM_RBUTTONDBLCLK
   End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Shell_NotifyIcon NIM_DELETE, Accion
  End

End Sub

Private Sub Salir_Click()
 Shell_NotifyIcon NIM_DELETE, Accion
 End
End Sub

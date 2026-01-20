VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1884
   ClientLeft      =   2832
   ClientTop       =   3480
   ClientWidth     =   3336
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1113.13
   ScaleMode       =   0  'User
   ScaleWidth      =   3126.69
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password"
      Height          =   276
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loggin As String
Dim buscar As String
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
Dim Resultado As Boolean
    'comprobar si la contraseña es correcta
    user = TxtPassword.Text
    'loggin = txtPassword.Text
    
    'If user = "admin" And loggin = "drake" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        'LoginSucceeded = True
    Me.Hide
    Unload Me
        'frmEpelsa.Show
      
    'Else
        
        'buscar = "usr = " & "'" & user & "'" & " and " & "psw = " & "'" & loggin & "'"
        'verificar buscar, resultado
        
        'If resultado = False Then
            'MsgBox "La contraseña, o el nombre de usuario no son válidos. Vuelva a intentarlo", , "Inicio de sesión"
            'txtUserName.SetFocus
            'SendKeys "{Home}+{End}"
        'Else
            'LoginSucceeded = True
            'Me.Hide
            'frmEpelsa.Show
            'frmEpelsa.mnuseuridad.Visible = False
        'End If
    
    'End If
    
End Sub

Public Sub verificar(Filtro As String, valido As Boolean)
Dim base As Database
Dim rssgr As Recordset


Set base = AbrirBase()
Set rssgr = base.OpenRecordset("select * from sgrd")

With rssgr

     If .EOF Then
         'MsgBox ("No hay ningun usuario activo")
         valido = False
         .Close
         CerrarBase base
         
         Exit Sub
     End If
        
     .FindFirst Filtro
     
     If .NoMatch Then
        valido = False
     Else
        valido = True
        If !seccion <> "" Then
        usersec = !seccion
        End If
     End If

.Close
CerrarBase base

End With
End Sub


Private Sub Form_Load()
user = ""
cmdOK.Caption = LoadResString(id + 287)
cmdcancel.Caption = LoadResString(id + 288)
End Sub

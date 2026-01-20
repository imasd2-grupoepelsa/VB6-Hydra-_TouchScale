VERSION 5.00
Begin VB.Form FrmNewLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1755
   ClientLeft      =   7050
   ClientTop       =   2235
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1212
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3492
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1812
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         Height          =   372
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "#"
         TabIndex        =   3
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label LblNombre 
         Caption         =   "Nombre :"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label LblPassword 
         Caption         =   "Contraseña :"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1332
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1332
   End
   Begin VB.CommandButton CmaAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1332
   End
End
Attribute VB_Name = "FrmNewLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nintentos As Integer
Public ElUsuario As String

Private Sub CmaAceptar_Click(Index As Integer)
    If TxtNombre.TexT = "Fabrica" Then
        If TxtPassword.TexT = " epfab" Then
            LoginSucceeded = True
            Unload Me
        Else
            LoginSucceeded = False
        End If
    End If
    If TxtNombre.TexT = "Euros" Then
        If TxtPassword.TexT = "moneda" Then
            LoginSucceeded = True
            Unload Me
        Else
            LoginSucceeded = False
        End If
    End If
    If TxtNombre.TexT = "RL_RLSS" Then
        If TxtPassword.TexT = "balanzasrl" Then
            LoginSucceeded = True
            Unload Me
        Else
            LoginSucceeded = False
        End If
    End If
    If LoginSucceeded = False Then MsgBox CargaCadena(769)  '"Contraseña Incorrecta"
    nintentos = nintentos + 1
    If nintentos = 3 Then
        End
    Else
        TxtPassword.TexT = ""
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    LoginSucceeded = False
    Unload Me
End Sub

Public Sub Form_Load()
   On Error Resume Next
    LblNombre.Caption = CargaCadena(552)
    LblPassword.Caption = CargaCadena(521)
    CmaAceptar(0).Caption = CargaCadena(287)
    Command1(1).Caption = CargaCadena(288)
    On Error GoTo 0
    If ElUsuario = "RL_RLSS" Then
        FrmNewLogin.Caption = "LOGIN"
        TxtNombre.Visible = False
        LblNombre.Visible = False
        TxtNombre.TexT = "RL_RLSS"
        TxtPassword.TexT = ""
        nintentos = 0
    End If
    If ElUsuario = "Euros" Then
        FrmNewLogin.Caption = "LOGIN"
        TxtNombre.Visible = False
        LblNombre.Visible = False
        TxtNombre.TexT = "Euros"
        TxtPassword.TexT = ""
        nintentos = 0
    End If
    If ElUsuario = "FABRICA" Then
        FrmNewLogin.Caption = "Admin"
        TxtNombre.Visible = False
        LblNombre.Visible = False
        TxtNombre.TexT = "Fabrica"
        TxtPassword.TexT = ""
        nintentos = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Login - Hydra MultiTienda"
   ClientHeight    =   1692
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1692
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Command1"
      Height          =   492
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Command1"
      Height          =   492
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   372
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   372
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2532
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      ForeColor       =   &H00C0FFFF&
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            
            If StrConv(Trim(Text1(0).Text), vbLowerCase) = "cmdline" Then
                UsuarioActual.Nombre = ""
            Else
                UsuarioActual.Nombre = StrConv(Trim(Text1(0).Text), vbLowerCase)
            End If
            UsuarioActual.Password = StrConv(Trim(Text1(1).Text), vbLowerCase)
            If Secure_Login(UsuarioActual) = False Then
                MsgBox CargaCadena(769), vbCritical
                Text1(0).Text = ""
                Text1(1).Text = ""
            Else
                FrmPrincipal.Show
                Unload Me
            End If
        Case 1
            Unload Me
            End
    End Select
End Sub

Private Sub Form_Load()
    CambiarIdioma
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
End Sub
Private Sub CambiarIdioma()
    Label1.Caption = CargaCadena(520)
    Label2.Caption = CargaCadena(521)
    Command1(0).Caption = CargaCadena(287)
    Command1(1).Caption = CargaCadena(288)
    Me.Caption = CargaCadena(99)
End Sub

VERSION 5.00
Begin VB.Form frmusuario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1635
   ClientLeft      =   3120
   ClientTop       =   3720
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   8
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    If StrConv(Trim(Text1.TexT), vbLowerCase) = "cmdline" Then
        UsuarioActual.Nombre = ""
    Else
        UsuarioActual.Nombre = StrConv(Trim(Text1.TexT), vbLowerCase)
    End If
    UsuarioActual.Password = StrConv(Trim(Text2.TexT), vbLowerCase)
    frmusuario.Hide
End Sub

Private Sub Command2_Click()
    UsuarioActual.Nombre = "NONE"
    frmusuario.Hide
End Sub

Private Sub Form_Load()
    UsuarioActual.Nombre = ""
    UsuarioActual.Password = ""
    Command1.Caption = CargaCadena(287)
    Command2.Caption = CargaCadena(288)
    Label1.Caption = CargaCadena(520)
    Label2.Caption = CargaCadena(521)
    frmusuario.Caption = CargaCadena(99)
End Sub

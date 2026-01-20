VERSION 5.00
Begin VB.Form FrmOrdenar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados de Artículos"
   ClientHeight    =   1905
   ClientLeft      =   5790
   ClientTop       =   2445
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Por Descriptivo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Por PLU"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Por Código"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "¿Cómo desea ordenar el Informe?"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "FrmOrdenar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim Bucle As Integer
    For Bucle = 0 To 2
        If Option1(Bucle).Value Then TipoOrden = Bucle
    Next Bucle
    Unload Me
End Sub

Private Sub Form_Load()
    Option1(0).Value = True
    Option1(0).Caption = CargaCadena(478) & " " & CargaCadena(69)
    Option1(1).Caption = CargaCadena(478) & " " & CargaCadena(68)
    Option1(2).Caption = CargaCadena(478) & " " & CargaCadena(70)
    Command1.Caption = CargaCadena(287)
    Me.Caption = CargaCadena(894)
    Label1.Caption = CargaCadena(1048)
End Sub

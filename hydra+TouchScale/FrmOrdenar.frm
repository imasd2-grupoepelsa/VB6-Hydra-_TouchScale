VERSION 5.00
Begin VB.Form FrmOrdenar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Artículos"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Descriptivo"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por PLU"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Código"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "¿Cómo desea Ordenar el Listado?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
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
        If Option1(Bucle).Value = True Then
            Rep_Orden = Bucle
        End If
    Next Bucle
    Unload Me
End Sub

Private Sub Form_Load()
    
    Option1(Rep_Orden).Value = True
End Sub

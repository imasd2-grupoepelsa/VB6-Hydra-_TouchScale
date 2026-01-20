VERSION 5.00
Begin VB.Form FormP 
   Caption         =   "Password"
   ClientHeight    =   1155
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandP 
      Caption         =   "Aceptar."
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox TextP 
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LabelPass 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FormP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandP_Click()
    sPassG = UCase(TextP.Text)
    DescargaForm = True
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
DescargaForm = True

End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   7200
   Begin VB.Label Label1 
      Caption         =   "CLICK DE RATÓN EN FORMULARIO PARA ICONIFICAR...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Click()
    Form2.cSysTray1.InTray = True
    Unload Me

End Sub



VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmWarning 
   Caption         =   "¡¡¡¡ WARNING... !!!!"
   ClientHeight    =   8580
   ClientLeft      =   90
   ClientTop       =   1935
   ClientWidth     =   15195
   LinkTopic       =   "Form2"
   ScaleHeight     =   8580
   ScaleWidth      =   15195
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   2655
      Left            =   2280
      OleObjectBlob   =   "frmWarning.frx":0000
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim nCont As Integer
    Me.BackColor = vbRed
    
    TextToSpeech1.Speak "TONTA"
    
    'For nCont = 1 To 100
    '    Beep
    '    DoEvents
    'Next nCont
    'Unload Me
End Sub

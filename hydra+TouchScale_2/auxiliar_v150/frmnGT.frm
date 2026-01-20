VERSION 5.00
Begin VB.Form frmnGT 
   Caption         =   "Itroduzca N.G.T"
   ClientHeight    =   1110
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   4020
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Propuesto:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Último Impreso:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmnGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim nArch As Integer
    nlastGT = Val(Text2.Text)
    nArch = FreeFile()
    Open sPathNGT & "ngtord" For Output As #nArch
    'Print #nArch, Format(nlastGT + FrmTotales.Lista.ListItems.Count, "0000")
    Print #nArch, Format(nlastGT, "000000")
    Close #nArch
    frmnGT.Visible = False
    Unload Me
End Sub

Private Sub Form_Load()
    'Dim nArch As Integer
    'Dim sParte As String
    
    'nArch = FreeFile()
    'Open sPathNGT & "ngtord" For Input As #nArch
    'Line Input #nArch, sParte
    'nlastGT = Val(sParte)
    'Close #nArch
    Text1.Text = Format(nlastGT, "000000")
    Text2.Text = Format(nlastGT, "000000") 'Format(nlastGT + 1, "0000")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii > 57 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

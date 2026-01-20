VERSION 5.00
Begin VB.Form FrmLogo 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim MiCad
    Dim Resp As String
    Dim bucle As Integer
    For bucle = 0 To 49
        MiCad = Chr(2) & "=00" & Chr(0) & "01" & Chr(0) & "1" & Chr(0) & "064" & Chr(0) & "008"
        MiCad = MiCad & Chr(0) & "050" & Chr(0) & Format(bucle, "000") & Chr(0)
        MiCad = MiCad & "0102030405060708090A0B0C0D0E0F10" & Chr(3)
        EnviarA MiCad, Resp
        
    Next bucle
    MsgBox "OK"
End Sub

VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   Caption         =   "Hard Drive"
   ClientHeight    =   675
   ClientLeft      =   1425
   ClientTop       =   2295
   ClientWidth     =   1680
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Estado"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Salir"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' The following code is required:
Option Explicit

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub Form_Load()
    Set SysTray = New clsSysTray
    Me.WindowState = vbMinimized
    DoEvents
    Me.Hide
    SysTray.Init Me, "Alimerka.exe Conexión TouchScale."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SysTray = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

' End required code
' /////////////////////////////////////////////

Private Sub mnuPopup_Click(Index As Integer)
Dim nF As Integer
Dim sF As String
    For nF = 1 To nServidoras
        sF = sF & sArrIP(nF) & " --> " & sArrSTATUS(nF) & vbCrLf
    Next nF
    Select Case Me.mnuPopup(Index).Caption
        Case "Estado":
            SysTray.ShowBalloonTip sF, beInformation, "Alimerka.exe"
            'MsgBox sF
        Case "Salir": End
        Case Else: MsgBox Me.mnuPopup(Index).Caption
    End Select
End Sub

Private Sub SysTray_BalloonClicked()
    
    SysTray.ShowBalloonTip "Versión:" & App.Major & "." & App.Minor, beInformation, "Alimerka.exe"

End Sub

Private Sub SysTray_DoubleClick()
Dim nF As Integer
Dim sF As String
    For nF = 1 To nServidoras
        sF = sF & sArrIP(nF) & " --> " & sArrSTATUS(nF) & vbCrLf
    Next nF
    SysTray.ShowBalloonTip sF, beInformation, "Alimerka.exe"
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mnuSysTray
End Sub


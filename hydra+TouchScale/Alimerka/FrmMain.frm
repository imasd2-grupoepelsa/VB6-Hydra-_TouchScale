VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "FrmMain"
   ClientHeight    =   585
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   1560
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   1560
   Begin VB.Menu mnuSystray 
      Caption         =   "mnuSystry"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp 
         Caption         =   "Salir"
         Index           =   0
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "Estado"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FrmMain"
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
    SysTray.Init Me, "G.Epelsa-TouchScale"
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

Private Sub mnuPopup_Click(Index As Integer)
    Select Case Me.mnuPopUp(Index).Caption
        Case "Balloon Tip": SysTray.ShowBalloonTip "Sample message of some length.", beInformation, "Balloon Tip"
        Case "Exit": Unload Me
        Case Else: MsgBox Me.mnuPopUp(Index).Caption
    End Select
End Sub

Private Sub SysTray_DoubleClick()
    SysTray.ShowBalloonTip "Double click tray icon", beInformation, "Balloon Tip"
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mnuSystray
End Sub


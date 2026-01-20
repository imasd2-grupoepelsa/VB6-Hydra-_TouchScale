VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7200
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nCont
Dim nSumKey

Private Function fCtrlAlt()
   
    If Not (GetAsyncKeyState(vbKeyControl) = 0) Then
        If nSumKey = 20 Or nSumKey = 0 Then
            nSumKey = nSumKey + 10
        End If
    End If
    If Not (GetAsyncKeyState(vbKeyMenu) = 0) Then
        If nSumKey = 0 Or nSumKey = 10 Then
            nSumKey = nSumKey + 20
        End If
    End If

End Function

Private Sub Form_Load()
nCont = 0
nSumKey = 0
End Sub

Private Sub Timer1_Timer()
Dim s As String
Dim nventanas As Long

Call fCtrlAlt

If nSumKey = 30 Then
    Timer1.Enabled = False
    'nventanas = keybd_event(VK_SHIFT, 0, 0, 0) 'send a key
    'DoEvents
    
   
    nventanas = keybd_event(vbKeyControl, 0, KEYEVENTF_KEYUP, 0) 'release a key
    DoEvents
    nventanas = keybd_event(vbKeyMenu, 0, KEYEVENTF_KEYUP, 0) 'release a key
    DoEvents
    frmWarning.Visible = True
    'Timer1.Enabled = False
    'Timer2.Enabled = True
    nCont = 200
    BlockInput True
    DoEvents
    'Unload frmWarning
    DoEvents
    
    nSumKey = 0
    Timer1.Enabled = True
    'MsgBox "ATENCION!!"
End If

End Sub

Private Sub Timer2_Timer()
Dim nventanas As Long
If nCont > 0 Then
    nCont = nCont - 1
    nventanas = keybd_event(vbKeyControl, 0, KEYEVENTF_KEYUP, 0) 'release a key
    DoEvents
    nventanas = keybd_event(vbKeyMenu, 0, KEYEVENTF_KEYUP, 0) 'release a key
    DoEvents
    nventanas = keybd_event(vbkeyesc, 0, 0, 0) 'send a key
    DoEvents
    nventanas = keybd_event(vbKeyMenu, 0, KEYEVENTF_KEYUP, 0) 'release a key
    DoEvents
    
    Text1.Text = CStr(nCont)
Else
    Timer2.Enabled = False
    nSumKey = 0
    BlockInput 0
    DoEvents
    MsgBox "Continuar..."
    'Timer1.Enabled = True
    'ncont = 0
    
End If

'Beep

End Sub

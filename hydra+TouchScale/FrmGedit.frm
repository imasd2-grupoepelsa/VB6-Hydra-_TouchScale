VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmGedit 
   Caption         =   "Form3"
   ClientHeight    =   2550
   ClientLeft      =   7050
   ClientTop       =   2235
   ClientWidth     =   3750
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock SockGedit 
      Left            =   1920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmGedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyBufSend As String
Dim EstoyEnviando As Boolean
Private Sub Form_Load()
    EstoyEnviando = False
    SockGedit.RemotePort = 0
    SockGedit.LocalPort = 4455
    SockGedit.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SockGedit.Close
    Do While SockGedit.State <> sckClosed
        Do_Events
    Loop
End Sub

Private Sub SockGedit_Close()
    SockGedit.Close
    Do While SockGedit.State <> sckClosed
        Do_Events
    Loop
    EstoyEnviando = False
    SockGedit.RemotePort = 0
    SockGedit.LocalPort = 4455
    SockGedit.Listen
End Sub

Private Sub SockGedit_ConnectionRequest(ByVal requestID As Long)
    SockGedit.Close
    Do While SockGedit.State <> sckClosed
        Do_Events
    Loop
    SockGedit.Accept requestID
End Sub

Private Sub SockGedit_DataArrival(ByVal bytesTotal As Long)
    Dim p As String
    Dim BufRec As String
    If SockGedit.State = 7 And Not EstoyEnviando Then
        SockGedit.GetData p, vbString
        MyBufSend = MyBufSend & p
        If InStr(1, MyBufSend, Chr(13)) > 0 Then
            'frmControl.Balanza.EnviaGedit MyBufSend, BufRec
            If (BufRec <> "") Then
                If SockGedit.State = 7 Then
                    EstoyEnviando = True
                    SockGedit.SendData BufRec
                Else
                    SockGedit.Close
                    Do While SockGedit.State <> sckClosed
                        Do_Events
                    Loop
                    EstoyEnviando = False
                    SockGedit.RemotePort = 0
                    SockGedit.LocalPort = 4455
                    SockGedit.Listen
                End If
            Else
                SockGedit.Close
                Do While SockGedit.State <> sckClosed
                    Do_Events
                Loop
                EstoyEnviando = False
                SockGedit.RemotePort = 0
                SockGedit.LocalPort = 4455
                SockGedit.Listen
            End If
        End If
    End If
End Sub

Private Sub SockGedit_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    SockGedit.Close
    Do While SockGedit.State <> sckClosed
        Do_Events
    Loop
    EstoyEnviando = False
    SockGedit.RemotePort = 0
    SockGedit.LocalPort = 4455
    SockGedit.Listen
End Sub

Private Sub SockGedit_SendComplete()
    MyBufSend = ""
    EstoyEnviando = False
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmCom 
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   795
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TmrComu 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////
'// Path de red                                                //
'////////////////////////////////////////////////////////////////
Public PathRed As String
Public PAthDesc As Integer
Public Indice As Integer
Private Sub Form_Load()
    TmrComu.Enabled = False
End Sub

'/////////////////////
'// TCP/IP Sockets ///
'/////////////////////
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim buf As String
    If Socket.State = 7 Then Socket.GetData buf, vbString
    'CadenadeLog "Recibidos:" & CStr(bytesTotal) & " Cadena:" & buf
    If buf <> "" Then
        If BufferRec = "" Then
            ' sólo acepta cadenas que empiecen por STX
            If InStr(1, buf, Chr(2)) > 0 Then
                buf = Mid(buf, InStr(1, buf, Chr(2)))
            Else
                buf = ""
            End If
        End If
        BufferRec = BufferRec & buf
        If InStr(1, BufferRec, Chr(13)) Then
            buf = "a"
            ' drena datos sobrantes
            Do While buf <> ""
                If Socket.State = 7 Then
                    Socket.GetData buf, vbString
                Else
                    buf = ""
                End If
            Loop
            BufferRec = Left(BufferRec, InStr(1, BufferRec, Chr(13)))
            TCP_REC = True
        End If
    End If
End Sub


Private Sub Socket_SendComplete()
    TCP_SENT = True
End Sub
Private Sub TmrComu_Timer()
    TmrComu.Enabled = False
    EndComu = True
End Sub


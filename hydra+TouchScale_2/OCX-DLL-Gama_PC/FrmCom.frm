VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmCom 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   945
   ClientTop       =   555
   ClientWidth     =   2220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   2220
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   1
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   2
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   3
      Left            =   0
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   4
      Left            =   0
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   5
      Left            =   0
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   6
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   7
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   8
      Left            =   360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   9
      Left            =   360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   10
      Left            =   360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   11
      Left            =   360
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   12
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   13
      Left            =   720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   14
      Left            =   720
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   15
      Left            =   720
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   16
      Left            =   720
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MultiSock 
      Index           =   17
      Left            =   720
      Top             =   1800
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

Private Sub MultiSock_Close(Index As Integer)
    MultiSock(Index).Close
    Do While MultiSock(Index).State <> sckClosed
        Do_Events
    Loop
End Sub

Private Sub MultiSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim buf As String
    If MultiSock(Index).State = 7 Then MultiSock(Index).GetData buf, vbString
    BufferRecMulti(Index) = BufferRecMulti(Index) & buf
    If InStr(1, BufferRecMulti(Index), Chr(13)) Then TCP_REC_MULTI(Index) = True
End Sub

Private Sub MultiSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MultiSock(Index).Close
    Do While MultiSock(Index).State <> sckClosed
        Do_Events
    Loop
   
End Sub

Private Sub MultiSock_SendComplete(Index As Integer)
    TCP_SENT_MULTI(Index) = True
End Sub

'/////////////////////
'// TCP/IP Sockets ///
'/////////////////////
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim buf As String
    If Socket.State = 7 Then Socket.GetData buf, vbString
    If buf <> "" Then
        If BufferRec = "" Then
            ' sólo acepta cadenas que empiecen por STX
            If InStr(1, buf, Chr(2)) > 0 Then
                buf = Mid(buf, InStr(1, buf, Chr(2)))
            Else
                '3.1.23 buf = ""
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


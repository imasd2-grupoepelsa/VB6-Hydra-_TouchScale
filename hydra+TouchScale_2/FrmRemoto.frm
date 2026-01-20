VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmRemoto 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2820
   ClientLeft      =   6915
   ClientTop       =   2070
   ClientWidth     =   4005
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrEspera 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   20
      Top             =   360
   End
   Begin MSWinsockLib.Winsock SockRem 
      Left            =   1680
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRemoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SeHaEnviado As Boolean
Private Sub Form_Load()
    SockRem.LocalPort = Val(PuertoPantallas)
    SockRem.Close
    SockRem.Listen
End Sub

Private Sub SockRem_Close()
    TmrEspera.Enabled = False
    SockRem.Close
    Do Until SockRem.State = sckClosed
        Do_Events
    Loop
    SockRem.Listen
    lCogeTiquet = True
End Sub

Private Sub SockRem_ConnectionRequest(ByVal requestID As Long)
    If SockRem.State <> sckConnected Then
        lCogeTiquet = False
        OrdenRemota = ""
        SockRem.Close
        Do Until SockRem.State = sckClosed
            Do_Events
        Loop
        TmrEspera.Enabled = True
        SockRem.Accept requestID
    Else
        
        
    End If
End Sub


Private Sub SockRem_DataArrival(ByVal bytesTotal As Long)
    Dim b As String
    SockRem.GetData b
    OrdenRemota = OrdenRemota & b
    If InStr(1, OrdenRemota, Chr(13)) Then
        TmrEspera.Enabled = False
        Procesa_Orden_Remota
        TmrEspera.Enabled = True
    End If
    
End Sub
Public Sub Enviar_Respuesta()
    SeHaEnviado = False
    If SockRem.State = sckConnected Then
        SockRem.SendData RespuestaRemota
        Do Until SeHaEnviado Or SockRem.State = sckClosed
            Do_Events
        Loop
    End If
End Sub



Private Sub SockRem_SendComplete()
    SeHaEnviado = True
    TmrEspera.Enabled = False
    'TmrEspera.Enabled = True
End Sub

Private Sub TmrEspera_Timer()
    TmrEspera.Enabled = False
    SockRem.Close
    Do Until SockRem.State = sckClosed
        Do_Events
    Loop
    SockRem.Listen
End Sub

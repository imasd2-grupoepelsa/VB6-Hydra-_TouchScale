VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMaster 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydra+ - Slave"
   ClientHeight    =   1365
   ClientLeft      =   7005
   ClientTop       =   2205
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3765
   Begin VB.Timer TimerOut 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer tmrsockslave 
      Enabled         =   0   'False
      Left            =   2880
      Top             =   240
   End
   Begin MSWinsockLib.Winsock sockslave 
      Left            =   720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Tmracabar 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3000
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1812
   End
   Begin VB.Timer TmrSlave 
      Enabled         =   0   'False
      Left            =   2880
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock SockMaster 
      Left            =   3360
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Hydra+ - SLAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3732
   End
End
Attribute VB_Name = "FrmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Probando As Boolean
Private Fallo_Slave As Boolean
Private Fase_Cierre As Integer
Private Envio_Completo As Boolean
Private Master_Responde As Boolean
Private Slave_Responde As Boolean
Private Buffer_Slave As String
Private HeConectado As Boolean
Private Master_Ordena_Salir As Boolean
Private Base_Copiada As Boolean
Private CuentaEsclavo As Long
Private MiCadena As String
Private MiCadena2 As String
Private SLProbando As Boolean
Public SlaveChange As Boolean
Private Sub Command1_Click()
    Dim Arch As Integer
    If Dir(App.Path & "\alarmas.exe") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\dirtmp\alarmas.off" For Output As #Arch
        Close #Arch
    End If
    End
End Sub

Private Sub Form_DblClick()
    FrmMaster.Visible = False
End Sub

Private Sub Form_Load()
    If Dir(MiruTa & "\slave.cfg") <> "" Then
        frmControl.cSysTray2.InTray = True
        modo_Slave
    End If
    If Dir(MiruTa & "\master.cfg") <> "" Then
        Modo_Master
    End If
        
End Sub
Private Sub Modo_Master()
    Dim Archivo As Integer
    Dim Buffer As String
    Dim MiCadena
    Archivo = FreeFile()
    Open MiruTa & "\master.cfg" For Input As Archivo
    Line Input #Archivo, Buffer
    MS_LPort = Buffer
    Line Input #Archivo, Buffer
    MS_RHost = Buffer
    Line Input #Archivo, Buffer
    MS_RPath = Buffer
    Close #Archivo
    MS_RPath = Trim(MS_RPath)
    If Right(MS_RPath, 1) = "\" Then
        MS_RPath = Left(MS_RPath, Len(MS_RPath) - 1)
    End If
    Probando = True
    Prueba_Master
    Probando = False
    SockMaster.Close
    SockMaster.LocalPort = MS_LPort
    SockMaster.Listen
    
End Sub
Private Sub Prueba_Master()
    Dim Bucle As Integer
    FrmMaster.Command1.Visible = False
    FrmMaster.Caption = "Starting Master System..."
    FrmMaster.Label1.Caption = "HYDRA - MASTER"
    FrmMaster.Show
    Do_Events
    Sleep (7000)
    FrmMaster.Enabled = False
    SockMaster.Close
    SockMaster.RemotePort = MS_LPort
    SockMaster.RemoteHost = Val(Left(MS_RHost, 3)) & "." & Val(Mid(MS_RHost, 5, 3)) & "." & Val(Mid(MS_RHost, 9, 3)) & "." & Val(Right(MS_RHost, 3))
    On Error GoTo fin
    
    SockMaster.Connect
    On Error GoTo 0
    '****************
    ' copiando base
    HeConectado = False
    TmrSlave.Interval = 4000
    TmrSlave.Enabled = True
    Fallo_Slave = False
    Do Until HeConectado Or Fallo_Slave
        Do_Events
    Loop
    
    If HeConectado Then
        TmrSlave.Interval = 20000
        TmrSlave.Enabled = True
        Fallo_Slave = False
        Do Until Base_Copiada Or Fallo_Slave
            Do_Events
        Loop
        For Bucle = 1 To 100
            Sleep (100)
            Do_Events
        Next Bucle
    End If
    FrmMaster.Visible = False
    Exit Sub
fin:
    FrmMaster.Visible = False
    SockMaster.Close
End Sub
Private Sub modo_Slave()
    Dim Reintentos As Integer
    Dim Arch As Integer
    Dim Archivo As Integer
    Dim Buffer As String
    Dim MiCadena
    MiCadena = "ALIVE?"
    Archivo = FreeFile()
    Open MiruTa & "\slave.cfg" For Input As Archivo
    Line Input #Archivo, Buffer
    MS_LPort = Buffer
    Line Input #Archivo, Buffer
    MS_RHost = Buffer
    Line Input #Archivo, Buffer
    MS_RPath = Buffer
    Close #Archivo
    SockMaster.RemoteHost = Val(Left(MS_RHost, 3)) & "." & Val(Mid(MS_RHost, 5, 3)) & "." & Val(Mid(MS_RHost, 9, 3)) & "." & Val(Right(MS_RHost, 3))
    SockMaster.RemotePort = MS_LPort
    CuentaEsclavo = 0
    Fallo_Slave = False
    Reintentos = 0
    Do While Fallo_Slave = False And Master_Ordena_Salir = False
        On Error Resume Next
        If Dir(App.Path & "\dirtmp\alarmas.off") = "" Then
            Arch = FreeFile()
            On Error Resume Next
            Open App.Path & "\dirtmp\alarmas.off" For Output As #Arch
            Close #Arch
            On Error GoTo 0
        End If
        On Error GoTo 0
        '******************
        ' intenta abrir el socket
        '******************
        'On Error GoTo fin
        HeConectado = False
        If SockMaster.State <> sckClosed Then
            SockMaster.Close
            Do Until SockMaster.State = sckClosed
                Do_Events
            Loop
        End If
        SockMaster.Connect
        'On Error GoTo 0
        TmrSlave.Interval = 5000
        TmrSlave.Enabled = True
        Do Until (HeConectado) Or (Fallo_Slave)
            Do_Events
        Loop
        TmrSlave.Enabled = False
        Envio_Completo = False
        If SockMaster.State = sckConnected Then
            '***********************
            ' envia y espera a que se termine de enviar
            '***********************
            TmrSlave.Interval = 5000
            TmrSlave.Enabled = True
            SockMaster.SendData MiCadena
            Do Until Envio_Completo Or Fallo_Slave
                Do_Events
            Loop
            TmrSlave.Enabled = False
            '***********************
            ' espera a que llegue la respuesta
            '***********************
            If Not Fallo_Slave Then
                TmrSlave.Interval = 5000
                TmrSlave.Enabled = True
                Master_Responde = False
                Buffer_Slave = ""
                Do Until Master_Responde Or Fallo_Slave
                    Do_Events
                Loop
                TmrSlave.Enabled = False
                If SockMaster.State <> sckClosed Then
                    SockMaster.Close
                End If
                TmrSlave.Interval = 4000
                TmrSlave.Enabled = True
                Do Until Fallo_Slave
                    Do_Events
                Loop
                Fallo_Slave = False
                TmrSlave.Enabled = False
            End If
        Else
            Fallo_Slave = True
        End If
        '**********************
        ' se hacen tres reintentos
        ' antes de dar por fracasada
        ' la conexión
        '**********************
        If Fallo_Slave Then
            Reintentos = Reintentos + 1
            If Reintentos >= 3 Then
                CadenadeLog "Master no responde, Reintentos : " & Reintentos
            Else
                CadenadeLog "Master no responde, reintentando"
                Fallo_Slave = False
            End If
        End If
    Loop
    If Master_Ordena_Salir Then
        If Dir(App.Path & "\alarmas.exe") <> "" Then
            Arch = FreeFile()
            Open App.Path & "\dirtmp\alarmas.off" For Output As #Arch
            Close #Arch
        End If
        End
    End If
    '*********************
    If sockslave.State <> sckClosed Then
        sockslave.Close
        Do Until sockslave.State <> sckClosed
            Do_Events
        Loop
    End If
    SlaveChange = False
    tmrsockslave.Interval = 5000
    tmrsockslave.Enabled = True
fin:
    CadenadeLog "HYDRA Esclavo no detecta conexión con el HYDRA maestro, pasa a funcionar como Reserva"
    frmControl.cSysTray2.InTray = False
    If SockMaster.State <> sckClosed Then
        SockMaster.Close
    End If
    tmrsockslave.Interval = 2000
    tmrsockslave.Enabled = True
    SoySlave = False
    SoyReserva = True
    SockMaster.LocalPort = MS_LPort
    SockMaster.Listen
End Sub





Private Sub Label1_DblClick()
FrmMaster.Visible = False
End Sub

Private Sub SockMaster_Close()
    If Probando Then
           Base_Copiada = True
    End If
End Sub

Private Sub SockMaster_Connect()
    HeConectado = True
End Sub

Private Sub SockMaster_ConnectionRequest(ByVal requestID As Long)
    '********************
    
    If SoyMaster Then
        SockMaster.Close
        SockMaster.Accept requestID
    End If
    If SoyReserva Then
        If DebugActivo Then CadenadeLog "Detectada conexión del sistema Master"
        frmEpelsa.LblSC10.Caption = "Master Connection Received"
        Fase_Cierre = 1
        lCogeTiquet = False
        frmControl.timetqt.Enabled = False
        SockMaster.Close
        SockMaster.Accept requestID
        If PlugActivo Then
            FrmHlink.STOPALL
        End If
        If DebugActivo Then CadenadeLog ("Enviando base maestra a MASTER")
        SockMaster.SendData "SENDING DATABASE" & Format(PlugStatus, "000")
        'MS_EnviaBaseRemota True
        Tmracabar.Interval = 2000
        Tmracabar.Enabled = True
    End If
        
End Sub

Private Sub SockMaster_DataArrival(ByVal bytesTotal As Long)
    Dim mibuffer
    Dim MiCadena
    MiCadena = "ALIVE!"
    MiCadena2 = "CLOSED"
    mibuffer = ""
    If SoySlave Then
        SockMaster.GetData mibuffer
        Buffer_Slave = Buffer_Slave & mibuffer
        If Len(Buffer_Slave) >= 6 Then
            Master_Responde = True
            If Buffer_Slave = "CLOSED" Then
                Master_Ordena_Salir = True
            End If
            If Left(Buffer_Slave, 6) = "ALIVE!" Then
                If Len(Buffer_Slave) >= 9 Then
                    PlugStatus = Val(Mid(Buffer_Slave, 7, 3))
                End If
            End If
        End If
    End If
    If SoyMaster Then
        SockMaster.GetData mibuffer
        Buffer_Slave = Buffer_Slave & mibuffer
        If Not Probando Then
            If Len(Buffer_Slave) >= 6 And Len(Buffer_Slave) < 16 Then
                
                If CI_Salir Or RE_TOTAL Then
                    SockMaster.SendData MiCadena2
                Else
                    SockMaster.SendData MiCadena & Format(PlugStatus, "000")
                End If
                If Buffer_Slave = "WAIT10" Then
                    Sleep (5000)
                End If
                Buffer_Slave = ""
            End If
        Else
            If Len(Buffer_Slave) >= 16 Then
                Slave_Responde = True
                If Len(Buffer_Slave) >= 19 Then
                    PlugStatus = Val(Mid(Buffer_Slave, 17, 3))
                End If
                Buffer_Slave = ""
            End If
        End If
    End If
End Sub

Private Sub SockMaster_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If SoySlave Then
        Fallo_Slave = True
    End If
    If SoyMaster Then
        If Not Probando Then
            SockMaster.Close
            SockMaster.LocalPort = MS_LPort
            SockMaster.Listen
        End If
    End If
End Sub

Private Sub SockMaster_SendComplete()
    If SoySlave Then
        Envio_Completo = True
    End If
    If SoyReserva Then
        If Fase_Cierre = 1 Then
            Fase_Cierre = 2
            If descAuto = False Then
                'MS_EnviaBaseRemota True
                SockMaster.Close
                salir_programa
            Else
                MS_Salir = True
            End If
        End If
    End If
    
End Sub

Private Sub sockslave_Connect()
    Dim Arch As Integer
    
    If Dir(App.Path & "\digitonuba.txt") = "" Then
    
    On Error Resume Next
    If SlaveChange Then sockslave.SendData "WAIT10"
    sockslave.Close
    If Dir(App.Path & "\alarmas.exe") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\dirtmp\alarmas.off" For Output As #Arch
        Close #Arch
    End If
    On Error GoTo 0
    If SlaveChange Then
        'MS_EnviaBaseRemota True
    End If
    Arch = FreeFile()
    Open MiruTa & "\reinit.slave" For Output As Arch
    Close #Arch
    Shell (MiruTa & "\hydra.exe")
    End
    
    End If
    
End Sub

Private Sub sockslave_DataArrival(ByVal bytesTotal As Long)

Dim s As String

sockslave.GetData s
sSpubS = sSpubS & s

End Sub

Private Sub sockslave_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim a
a = a
End Sub

Private Sub sockslave_SendComplete()
Dim a
a = a
End Sub

Private Sub TimerOut_Timer()
    lTimeOut = True
End Sub

Private Sub Tmracabar_Timer()
    salir_programa
End Sub

Private Sub TmrSlave_Timer()
    Fallo_Slave = True
End Sub

Private Sub tmrsockslave_Timer()
    Dim b As Integer
    tmrsockslave.Enabled = False
    If Not SLProbando Then
        SLProbando = True
        If sockslave.State <> sckClosed Then
            sockslave.Close
        End If
        Do While sockslave.State <> sckClosed
            Do_Events
        Loop
        sockslave.RemotePort = MS_LPort
        sockslave.RemoteHost = Val(Left(MS_RHost, 3)) & "." & Val(Mid(MS_RHost, 5, 3)) & "." & Val(Mid(MS_RHost, 9, 3)) & "." & Val(Right(MS_RHost, 3))
        sockslave.Connect
        SLProbando = False
    End If
    tmrsockslave.Interval = 5000
    tmrsockslave.Enabled = True
    
End Sub

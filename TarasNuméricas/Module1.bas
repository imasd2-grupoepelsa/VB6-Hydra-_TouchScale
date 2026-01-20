Attribute VB_Name = "Module1"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Winsock API declares
Public Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
' linger structure
Public Type LINGER_STRUCT
  l_onoff As Integer          ' Is linger on or off?
  l_linger As Integer         ' Linger timeout in seconds.
End Type
' Level number for (get/set)sockopt() to apply to socket itself.
Public Const SOL_SOCKET = 65535      ' Options for socket level.
Public Const SO_LINGER = &H80&       ' Linger on close if data present.
'****************************
' Timeout
Public EndComu As Boolean
'*****************************
' Buffer de recepción
Public BufferRec As String
'''''''''''''''''''''''''''''
'****************************
' Buffer de envío
Public buffersend As String
'****************************
Public TCP_SENT As Boolean
Public TCP_REC As Boolean
Public Success As Integer
Public Failed As Integer
Public sIp As String
Public nPortS As Integer

Public Function SuperLog(MiCad As String)
    '////////////////////////////////
    '// Esta función va generando  //
    '// el log de comunicaciones   //
    '// en un fichero llamado      //
    '// logepelcom.txt en la ruta  //
    '// donde esté situado el OCX  //
    '////////////////////////////////
    Dim Arch As Integer
    Dim MyVAl As Long
    Dim MiCad2 As String
    Dim Bucle As Integer
    Dim MiCad3 As String
    MiCad3 = ""
    For Bucle = 1 To Len(MiCad)
            MiCad3 = MiCad3 & Asc(Mid(MiCad, Bucle, 1)) & "|"
    Next Bucle

    Arch = FreeFile()
    On Error Resume Next
    Open App.Path & "\logepelcom.txt" For Append As #Arch
    If Err.Number = 0 Then
        Print #Arch, Now() & "-->" & MiCad3
        Close #Arch
    End If
    On Error GoTo 0
End Function

'////////////////////////////////
'// Esta función va generando  //
'// el log de comunicaciones   //
'// en un fichero llamado      //
'// logepelcom.txt en la ruta  //
'// donde esté situado el OCX  //
'////////////////////////////////
Public Function CadenadeLog(ByVal MiCad As String)
    Dim Arch As Integer
    Dim MyVAl As Long
    Dim MiCad2 As String
    Dim Bucle As Integer
    Dim MiCad3 As String
    Dim DoSuperLog As Boolean
    MiCad2 = ""
    MiCad3 = ""
    For Bucle = 1 To Len(MiCad)
        If Asc(Mid(MiCad, Bucle, 1)) < 32 Then
            DoSuperLog = True
            MiCad2 = MiCad2 & "·"
        Else
            MiCad2 = MiCad2 & Mid(MiCad, Bucle, 1)
        End If
    Next Bucle
    Arch = FreeFile()
    On Error Resume Next
    Open App.Path & "\logepelcom.txt" For Append As #Arch
    If Err.Number = 0 Then
        Print #Arch, Now() & "-->" & MiCad2
        Close #Arch
    End If
    If DoSuperLog Then SuperLog (Mid(MiCad, InStr(1, MiCad, Chr(2))))
    If Err.Number = 0 Then
        On Error Resume Next
        MyVAl = FileLen(App.Path & "\logepelcom.txt")
        If MyVAl > 1048576 Then
            On Error Resume Next
            Kill App.Path & "\logepelcom.txt"
        End If
    End If
    On Error GoTo 0
End Function


Public Function Do_Events()
    Sleep (1)
    DoEvents
    Sleep (1)
End Function

Public Function CalcCRC(ByVal nlen As Integer, MyCad As String) As Double
    '/////////////////////////////////////
    '// BINARIO:                        //
    '// Calcula el checksum de envío    //
    '/////////////////////////////////////
    Dim r, a, B As Double
    Dim T As Double
    r = 0
    For T = 0 To nlen - 1
        r = r + T * (Asc(Mid(MyCad, T + 1, 1))) + T
    Next
    a = r And 65280
    B = r And 255
    If a = 3328 Then a = 3584
    If B = 13 Then B = 14
    CalcCRC = a Or B
End Function
Public Function tonhex(ByVal sHex As String, ByVal nByte As Integer) As String
   '///////////////////////////////////
   '// auxiliar de IntToHex          //
   '// (INTERNA)                     //
   '///////////////////////////////////
    Dim sCeros As String
    Dim nlen, N As Integer
    Dim sNumero As String
    nlen = Len(sHex)
    sCeros = "00000000"
    If nlen < nByte Then
       sNumero = Mid(sCeros, 1, nByte - nlen) + sHex
    Else
       sNumero = sHex
    End If
    tonhex = sNumero
End Function
Public Function StrBinToInt(ByVal cadena As String) As Double
   '////////////////////////////////////
   '// convierte datos de varios      //
   '// bytes del formato binario      //
   '// entendible por Euroscale a     //
   '// formato en cristiano           //
   '// (INTERNA)                      //
   '////////////////////////////////////
    Dim Buffer As String
    Dim Bucle As Long
    Dim Contador As Integer
    Dim valor As Double
    valor = 0
    Contador = 0
    For Bucle = 1 To Len(cadena)
        valor = valor + Asc(Mid(cadena, Bucle, 1)) * (256 ^ Contador)
        Contador = Contador + 1
    Next Bucle
    StrBinToInt = valor
End Function

Public Function IntToHex(ByVal nNum As Double, ByVal nByte As Integer) As String
   '///////////////////////////////////
   '// convierte datos de varios     //
   '// bytes al formato binario      //
   '// entendible por Euroscale      //
   '// (INTERNA)                     //
   '///////////////////////////////////
   Dim nNumero As Variant
   Dim nlen, N As Integer
   Dim sCeros As String
   Dim MyCad As String
   MyCad = ""
   sCeros = "00000000"
   nNumero = tonhex(Hex(nNum), nByte * 2)
   For N = nByte To 1 Step -1
      MyCad = MyCad & Chr(CInt("&H" & Mid(nNumero, (N * 2) - 1, 2)))
   Next
   IntToHex = MyCad
End Function
Public Function Caracter13(cadena As String) As String
'////////////////////////////////////
'// BINARIO:                       //
'// Repone el ASCII chr(13) que se //
'// cambia en la comunicación      //
'// binaria                        //
'////////////////////////////////////
    Dim s As String
    Dim Bucle As Integer
    Dim Buffer As String
    If Len(cadena) < 5 Then
        Buffer = cadena
    Else
        s = Mid(cadena, 5, 1)
        For Bucle = 1 To Len(cadena)
            If Mid(cadena, Bucle, 1) = s Then
                Buffer = Buffer & Chr(13)
            Else
                Buffer = Buffer & Mid(cadena, Bucle, 1)
            End If
        Next Bucle
    End If
    Caracter13 = Buffer
End Function

Public Function Binary_CheckSum() As Boolean
    Dim buf As String
    Dim Resto As String
    Dim cadena As String
    Dim Crc As Double
    Dim MyOK As Boolean
    Dim Bucle As Integer
    MyOK = False

    For Bucle = 1 To 2
        If Bucle = 1 Then
            cadena = Caracter13(BufferRec)
        Else
            cadena = Caracter13(BufferRec)
        End If
        If Len(cadena) >= 8 Then
            buf = Left(cadena, 4) & Chr(0) & Mid(cadena, 6)
            buf = Left(buf, Len(buf) - 3)
            Resto = Right(cadena, 3)
            Resto = Left(Resto, 2)
            Crc = CalcCRC(Len(buf), buf)
            If Resto = IntToHex(Crc, 2) Then
                MyOK = True
                Exit For
            
            End If
        End If
    Next Bucle
    Binary_CheckSum = MyOK
End Function

Public Function AdjB(cadena As String) As String
    '/////////////////////////////////////
    '// BINARIO :                       //
    '// Crea el checksum binario        //
    '// para el envío                   //
    '/////////////////////////////////////
    Dim nVal As Integer
    Dim nAlta As Integer
    Dim nBaja As Integer
    Dim nlen As Integer
    Dim N As Integer
    Dim s As String
    Dim nPos As Integer
    Dim sBinSx As String
    Dim Crc As Double
    sBinSx = cadena
    nVal = Len(sBinSx) - 10
    nAlta = Int(nVal / 256)
    nBaja = nVal - (256 * nAlta)
    sBinSx = Mid(sBinSx, 1, 5) + Chr(nBaja) + Chr(nAlta) + Mid(sBinSx, 8)
    nlen = Len(sBinSx)
    Crc = CalcCRC(nlen, sBinSx)
    sBinSx = sBinSx & IntToHex(Crc, 2)
    For N = 0 To 255
        If N <> 13 Then
            If InStr(1, sBinSx, Chr(N)) = 0 Then
                 s = Chr(N)
                 N = 256
            End If
        End If
    Next N
    sBinSx = Mid(sBinSx, 1, 4) + s + Mid(sBinSx, 6)
    Do
        nPos = InStr(1, sBinSx, Chr(13))
        If nPos <> 0 Then
            Mid(sBinSx, nPos, 1) = s
        End If
    Loop Until InStr(1, sBinSx, Chr(13)) = 0
    sBinSx = sBinSx + Chr(13)
    AdjB = sBinSx
End Function

Public Function cierraSOCKETuno() As Boolean
    Dim hhand As Long
    Dim lresult2 As Long
    Dim linger As LINGER_STRUCT
    
    On Error Resume Next
    
    If FrmCom.Socket.State <> 0 Then
    
        hhand = FrmCom.Socket.SocketHandle
        linger.l_onoff = 1
        linger.l_linger = 0
        
        FrmCom.Socket.Close
        
        Do_Events
        EndComu = False
        FrmCom.TmrComu.Interval = 5000
        FrmCom.TmrComu.Enabled = True
        Do_Events
        Do While FrmCom.Socket.State <> 0 And (Not (EndComu))
            Do_Events
        Loop
        FrmCom.TmrComu.Enabled = False

    End If
    
    'Estado = 0
    
    If FrmCom.Socket.State = 0 Then
        'Estado = 0
        cierraSOCKETuno = True
    Else
        cierraSOCKETuno = False
    End If
    
    On Error GoTo 0
    
End Function

Function abreSOCKETuno() As Boolean
    Dim lresp As Boolean
    
    On Error Resume Next
    If FrmCom.Socket.State <> 0 And FrmCom.Socket.State <> 7 Then
        lresp = cierraSOCKETuno()
    End If
    On Error GoTo 0
    
    FrmCom.Socket.RemoteHost = sIp
    FrmCom.Socket.RemotePort = nPortS
    
    If FrmCom.Socket.State <> 7 Then
    
        FrmCom.TmrComu.Enabled = False
        EndComu = False
        FrmCom.Socket.LocalPort = 0
        FrmCom.TmrComu.Interval = 5000
        FrmCom.TmrComu.Enabled = True
        FrmCom.Socket.Connect
        Do While FrmCom.Socket.State <> 7 And (Not (EndComu))
            Do_Events
        Loop
        FrmCom.TmrComu.Enabled = False
        
        If FrmCom.Socket.State = 7 Then
            abreSOCKETuno = True
        Else
            abreSOCKETuno = False
        End If
    Else
        abreSOCKETuno = True
    End If
    
End Function

Private Sub SendRaw_TCP_Binary()
'////////////////////////////////////
'// Envíos por TCP/IP              //
'// para comunicaciones con        //
'// Formato Binario                //
'////////////////////////////////////
Dim Recibido As Integer
Dim NoCheck As Boolean
Dim Checksum As String
Dim Sarr() As Byte
Dim Rarr() As Byte
Dim sRarr As Variant
Dim carac As String
Dim TimeOut As Boolean
Dim lresp As Boolean

BufferRec = ""
TCP_REC = False
'TimeOut = abreSOCKETuno()
Select Case FrmCom.Socket.State
Case 7
    TCP_SENT = False
    TCP_REC = False
    
    FrmCom.Socket.SendData buffersend
    CadenadeLog "Envío TCP/IP -->" & buffersend
    
    EndComu = False
    FrmCom.TmrComu.Interval = 8000
    FrmCom.TmrComu.Enabled = True
    Do Until TCP_SENT Or EndComu
        Do_Events
    Loop
    FrmCom.TmrComu.Enabled = False

    Select Case TCP_SENT
        Case True
            EndComu = False
            FrmCom.TmrComu.Interval = 8000
            FrmCom.TmrComu.Enabled = True
            Do Until EndComu Or TCP_REC
                Do_Events
                'If Dir(App.Path & "\cancel.tqt") <> "" Then EndComu = True
            Loop
            FrmCom.TmrComu.Enabled = False
            If EndComu Then
                TCP_REC = False
                BufferRec = ""
            End If
            Select Case TCP_REC
                Case True
                    '************************
                    '* control de checksum  *
                    '************************
                    If Not Binary_CheckSum() Then
                        BufferRec = ""
                        CadenadeLog "Fallo TCP/IP --> Error de Checksum Binario"
                        NoCheck = True
                    End If
                Case False
                    ' timeout recibiendo
                    BufferRec = ""
                    CadenadeLog "Fallo TCP/IP --> TimeOut Recibiendo datos."
                    TimeOut = True
            End Select
        Case False
            ' timeout enviando
            BufferRec = ""
            CadenadeLog "Fallo TCP/IP --> TimeOut enviando Datos."
            TimeOut = True
    End Select
Case Else
    ' no pudo establecerse el socket
    BufferRec = ""
    CadenadeLog "Fallo TCP/IP --> No pudo establecerse el socket"
    TimeOut = True
End Select

If BufferRec <> "" Then
    CadenadeLog "Recibiendo TCP/IP --> " & BufferRec
    Success = 1
Else
    If TimeOut Then
        Failed = 1
    Else
        If NoCheck Then Failed = 2
    End If
End If
    
'TimeOut = cierraSOCKETuno()


End Sub


Public Function SendRaw_Binary(sData As String) As Long
    Dim Checksum
    Dim Reintentos As Integer
    Dim Resp As Long
    Dim MyLen1 As Long, MyLen2 As Long
    BufferRec = ""
    Reintentos = 0

        Success = 0
        Failed = 0
        Do While Reintentos < 3
            buffersend = sData
            buffersend = AdjB(buffersend)
            BufferRec = ""
            SendRaw_TCP_Binary
            '*************************
            ' protección Binaria
            '*************************
            If Len(BufferRec) >= 8 Then
                If (Left(BufferRec, 1) <> Chr(2)) Or (Right(BufferRec, 1) <> Chr(13)) Or (Mid(BufferRec, 2, 1) <> Chr(126)) Then
                    CadenadeLog "Mensaje no válido"
                    BufferRec = ""
                End If
            Else
                CadenadeLog "Mesaje no válido"
                BufferRec = ""
            End If
            If Len(BufferRec) >= 8 Then
                MyLen1 = Len(BufferRec)
                On Error Resume Next
                MyLen2 = StrBinToInt(Mid(BufferRec, 6, 2))
                If Err.Number <> 0 Then MyLen2 = -1
                On Error GoTo 0
                'LogTXT BufferRec
                If (MyLen1 - 13) <> MyLen2 Then
                    CadenadeLog "Mesaje no válido"
                    BufferRec = ""
                End If
            End If
            '**************************
            Select Case BufferRec
                Case ""
                    If Dir(App.Path & "\cancel.tqt") = "" Then
                        Reintentos = Reintentos + 1
                        CadenadeLog ("Fallo de Comunicaciones, reintentando")
                    Else
                        Reintentos = 3
                        CadenadeLog ("cancel.tqt encontrado")
                    End If
                    Select Case Failed
                        Case 1
                            
                            ' Timeout intentando comunicar
                            Resp = 1000
                        Case 3
                            
                            Resp = 1020
                        Case Else
                            ' Error de checksum
                            
                            Resp = 1001
                    End Select
                Case Else
                    '****************************************
                    '* de momento, si ha conseguido         *
                    '* comunicar, devuelve OK               *
                    '****************************************
                    Reintentos = 3
                    Resp = 0
            End Select
            If Dir(App.Path & "\cancel.tqt") <> "" Then
                Reintentos = 3
                Resp = 1000
            End If
        Loop
    
    SendRaw_Binary = Resp
End Function

Public Function SendTare(ByVal tipo As Long, ByVal valor As Long, ByVal nEqp As Byte) As Long
    Dim MyOK As Boolean
    Dim sBinSx As String
    Dim Recibe As Long
    Dim lDesc As Integer
    Dim Posiciones(4) As Integer
    Dim Bucle As Integer
    Dim B2 As Integer
    Dim Contador As Integer
    Dim Resp As Long
    MyOK = True

'Envío TCP/IP -->·~··············ÿ··········
'2|126|0|0|4|14|0|0|0|1|3|1|0|0|6|1|255|1|2|0|1|0|5|0|20|19|13|
'Recibiendo TCP/IP --> ·~··············ÿ·Ã··
'2|126|0|0|3|8|0|0|0|1|4|0|0|0|6|1|255|0|195|17|13|
'Tara N.:1 Valor:5   Respuesta:0
'Envío TCP/IP -->·~··············ÿ··········
'2|126|0|0|4|14|0|0|0|1|3|1|0|0|6|1|255|1|2|0|1|0|5|0|20|19|13|
'TCP/IP --> ·~··············ÿ·Ã··
'2|126|0|0|3|8|0|0|0|1|4|0|0|0|6|1|255|0|195|17|13|
'Tara N.:1 Valor:5   Respuesta:0
    'If tipo > 1 Then
        If MyOK Then
                'recuperar
                sBinSx = Chr(2) & Chr(126) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) _
                       & Chr(0) & Chr(0) & Chr(1)
                    ' Orden programar(3), método igual (1,0), máscara 0, dbid(6), key (1), extendida (255)
                sBinSx = sBinSx & Chr(3) & Chr(1) & Chr(0) & Chr(0) & Chr(6) & Chr(1) & Chr(255) _
                    & Chr(1) & Chr(2) & Chr(0) & Chr(1)
                    ' tipo
                sBinSx = sBinSx & Chr(tipo - 1)
                    ' valor
                sBinSx = sBinSx & IntToHex(valor, 2)
                Resp = SendRaw_Binary(sBinSx)
        Else
            Resp = 1002
        End If
    
        CadenadeLog "Tara N.:" & CStr(tipo) & " Valor:" & CStr(valor) & "   Respuesta:" & CStr(Resp)
    'End If
    SendTare = Resp
    
End Function



Attribute VB_Name = "tools"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Option Explicit
Public Type Tipo_TqCodBar
    Tipo As Integer
    NBalanza As Integer
    NMostrador As Integer
    NTicket As Long
    NVendedor As Long
    Importe As Double
End Type

Public General_Events As Integer
Public Type gianSec
    secmaq As Integer
    secmos As Integer
    secdes As String
End Type
Dim sGi(99) As gianSec
'1.7.17
Public chrOrgG(13) As Byte
Public chrDesG(13) As Byte
'1.7.19 (eco)
Public sIpEco As String
Public nPortEco As String
''''''''''''''''''''''''''
Public Function Calcula_Juliano() As String
Dim sAño As String
Dim nJuliano As Long
Dim sP As String
    sP = Format(Year(Now()), "0000")
    sAño = Mid(sP, 3, 2)
    nJuliano = DateDiff("d", "01/01/" & sAño, Now) + 1
    Calcula_Juliano = sP & "." & Format(nJuliano, "000")

End Function
Public Function Calcula_Juliano1() As String
Dim sAño As String
Dim nJuliano As Long
Dim sP As String
    sP = Format(Year(Now()), "0000")
    sAño = Mid(sP, 3, 2)
    nJuliano = DateDiff("d", "01/01/" & sAño, Now) + 1
    Calcula_Juliano1 = "." & Format(nJuliano, "000")

End Function

Public Function Fecha2Long(Dia As Integer, Mes As Integer, Anyo As Integer) As String
Dim Fecha As Double
Dim cadena As String
Dim i As Long
Fecha = 0
If Anyo < 100 Then
    If Anyo >= 80 Then
        Anyo = Anyo + 1900
    Else
        Anyo = Anyo + 2000
    End If
End If
For i = 1980 To Anyo - 1
    If ((i Mod 4) = 0) Or (((i Mod 100) = 0) And ((i Mod 1000) = 0)) Then
        Fecha = Fecha + 366
    Else
        Fecha = Fecha + 365
    End If
Next i
For i = 1 To Mes - 1
    If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 8 Or i = 10 Or i = 12 Then
        Fecha = Fecha + 31
    End If
    If i = 4 Or i = 6 Or i = 9 Or i = 11 Then
        Fecha = Fecha + 30
    End If
    If i = 2 Then
        If ((Anyo Mod 4) = 0) Or (((Anyo Mod 100) = 0) And ((Anyo Mod 1000) = 0)) Then
            Fecha = Fecha + 29
        Else
            Fecha = Fecha + 28
        End If
    End If
Next i
Fecha = Fecha + Dia - 1
Fecha = Fecha * 86400
Fecha2Long = Fecha
End Function
Public Sub Check_FileAttr(sFile As String)
    Dim b As Long
    If Dir(sFile) <> "" Then
        b = GetAttr(sFile)
        If (b And vbReadOnly) Then
            MsgBox "File called : " & sFile & " has read-only permissions.Incorrect configuration.STOP", vbCritical
            salir_programa
        End If
    End If
End Sub
Public Function AdaptaFecha(ByVal sfecha As String) As String
    AdaptaFecha = left(sfecha, 2) & "/" & Mid(sfecha, 4, 2) & "/" & Right(sfecha, 2)
End Function
Public Function GeneraTextoIP(sData As String, Optional ByRef nPort As Long) As Boolean
    Dim Buf As String
    Dim bucle As Long
    Dim nPuntos As Integer
    Dim esPunto As Boolean
    Dim MyOK As Boolean
    Dim nNum As Integer
    Dim Num(4) As String
    Dim sPort As String
    Dim lDosPuntos As Boolean
    
    Buf = sData
    MyOK = True
    lDosPuntos = False
    sPort = ""
    For bucle = 1 To Len(Buf)
        Select Case Mid(Buf, bucle, 1)
            Case "0" To "9"
                If lDosPuntos Then
                    sPort = sPort & Mid(Buf, bucle, 1)
                Else
                    esPunto = False
                    Num(nNum) = Num(nNum) & Mid(Buf, bucle, 1)
                End If
            Case "."
                If Not esPunto Then
                    nPuntos = nPuntos + 1
                    nNum = nNum + 1
                    esPunto = True
                Else
                    MyOK = False
                    Exit For
                End If
            Case ":"
                lDosPuntos = True
            Case Else
                MyOK = False
                Exit Function
        End Select
    Next bucle
    If nPuntos <> 3 Then MyOK = False
    If IsNumeric(Num(0)) And IsNumeric(Num(1)) And IsNumeric(Num(2)) And IsNumeric(Num(3)) Then
        If (Val(Num(0)) < 256) And (Val(Num(1)) < 256) And (Val(Num(2)) < 256) And (Val(Num(3)) < 256) Then
            sData = Format(Val(Num(0)), "000") & "." & _
                    Format(Val(Num(1)), "000") & "." & _
                    Format(Val(Num(2)), "000") & "." & _
                    Format(Val(Num(3)), "000")
            
        Else
            MyOK = False
        End If
    Else
        MyOK = False
    End If
    
    If lDosPuntos Then
        nPort = (Val(sPort))
    End If
    GeneraTextoIP = MyOK
End Function
Public Function ToDouble(sData As String) As Double
    Dim Valor As Double
    Dim Valor2 As Double
    Dim Contador As Integer
    Dim eNCONTRADO As Boolean
    Dim bucle As Integer
    For bucle = 1 To Len(sData)
        Select Case Mid(sData, bucle, 1)
            Case 0 To 9
                If Not eNCONTRADO Then
                    Valor = Valor * 10
                    Valor = Valor + Val(Mid(sData, bucle, 1))
                Else
                    Contador = Contador + 1
                    Valor2 = Valor2 * 10
                    Valor2 = Valor2 + Val(Mid(sData, bucle, 1))
                End If
            Case "."
                eNCONTRADO = True
            Case ","
                eNCONTRADO = True
        End Select
    Next bucle
    Valor2 = Valor2 / (10 ^ Contador)
    Valor = Valor + Valor2
    ToDouble = Valor
End Function
Public Function ToString(sData As Double) As String
    Dim Valor As String
    Dim Contador As Integer
    Dim eNCONTRADO As Boolean
    Dim bucle As Integer
    Dim Valido As String
    Dim InValido As String
    Valor = CStr(sData)
    If "1,1" = 1.1 Then
        InValido = "."
        Valido = ","
    Else
        InValido = ","
        Valido = "."
    End If
    Valor = Replace(Valor, InValido, Valido)
    ToString = Valor
End Function

Public Sub Compactar_Base()
    Dim MiProceso As Long
    Dim MiResultado As Long
    Dim BufLcoge As Boolean
    Dim Arch As Integer
    Dim MiNum As String
    Dim SalirMulti As Boolean
    '*************************
    ' Servidor de Pantallas remotas
    ' (No compacta para no interferir,
    ' ya lo hace el servidor)
    If lUpper Then Exit Sub
    If lFornes Then Exit Sub
    If GetProcesses("tiqahtouch.exe") Then Exit Sub
    
    If Dir(App.Path & "\service.exe") <> "" Then Exit Sub
    If Dir(App.Path & "\transpluasc.exe") <> "" Then Exit Sub
    '*************************
    ' Servidor de Pantallas remotas
    ' (No compacta para no interferir,
    ' ya lo hace el servidor)
    If Dir(App.Path & "\hydraserver.exe") <> "" Then Exit Sub
    '
    If Base_General = "" Then Exit Sub
    '
    If HayMulti Then
        If (Dir(App.Path & "\t" & Format(TiendaActual, "00"), vbDirectory) = "" And TiendaActual < 100) Or (Dir(App.Path & "\t" & Format(TiendaActual, "000"), vbDirectory) = "" And TiendaActual > 99) Then
            MiNum = "0"
        Else
            Arch = FreeFile()
            If (Dir(App.Path & "\t" & Format(TiendaActual, "00") & "\compact.txt") <> "" And TiendaActual < 100) Or (Dir(App.Path & "\t" & Format(TiendaActual, "000") & "\compact.txt") <> "" And TiendaActual > 99) Then
                If TiendaActual < 100 Then
                Open App.Path & "\t" & Format(TiendaActual, "00") & "\compact.txt" For Input As #Arch
                Else
                Open App.Path & "\t" & Format(TiendaActual, "000") & "\compact.txt" For Input As #Arch
                End If
                On Error Resume Next
                Line Input #Arch, MiNum
                On Error GoTo 0
                Close #Arch
            Else
                MiNum = "0"
            End If
        End If
        If Val(MiNum) < 10 Then
            MiNum = Val(MiNum) + 1
            SalirMulti = True
        Else
            MiNum = "0"
        End If
        Arch = FreeFile()
        If TiendaActual < 100 Then
        Open App.Path & "\t" & Format(TiendaActual, "00") & "\compact.txt" For Output As #Arch
        Else
        Open App.Path & "\t" & Format(TiendaActual, "000") & "\compact.txt" For Output As #Arch
        End If
        Print #Arch, MiNum
        Close #Arch
        If SalirMulti Then Exit Sub
        
    End If
    '
    

    BufLcoge = lCogeTiquet
    lCogeTiquet = False
    MiResultado = STILL_ACTIVE
    
    '2.0.22 se compacta además de mastercopy
    If (Dir(App.Path & "\" & Format(TiendaActual, "00") & "tmp.mdb") <> "" And TiendaActual < 100) Or (Dir(App.Path & "\" & Format(TiendaActual, "000") & "tmp.mdb") <> "" And TiendaActual > 99) Then
        On Error Resume Next
        If TiendaActual < 100 Then
        Kill App.Path & "\" & Format(TiendaActual, "00") & "tmp.mdb"
        Else
        Kill App.Path & "\" & Format(TiendaActual, "000") & "tmp.mdb"
        End If
        If Err.Number <> 0 Then
            MsgBox "ERROR CRÍTICO : NO ES POSIBLE BORRAR " & App.Path & "\" & Format(TiendaActual, "000") & "TMP.MDB", vbCritical
            'escribeLOg "ERROR CRÍTICO : NO ES POSIBLE BORRAR " & App.Path & "\TMP.MDB"
            End
        End If
        On Error GoTo 0
    End If
    If TiendaActual < 100 Then
    DBEngine.CompactDatabase Base_General, App.Path & "\" & Format(TiendaActual, "00") & "tmp.mdb"
    Else
    DBEngine.CompactDatabase Base_General, App.Path & "\" & Format(TiendaActual, "000") & "tmp.mdb"
    End If
    On Error Resume Next
    Kill Base_General
    If Err.Number <> 0 Then
        MsgBox "ERROR CRÍTICO : NO ES POSIBLE BORRAR " & Base_General, vbCritical
        End
    End If
    On Error GoTo 0
    If TiendaActual < 100 Then
    Name App.Path & "\" & Format(TiendaActual, "00") & "tmp.mdb" As Base_General
    Else
    Name App.Path & "\" & Format(TiendaActual, "000") & "tmp.mdb" As Base_General
    End If
    ''''''''
    
    'MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiruTa & "\mastercopy.exe COMPACT" & Base_General, 6)) 'c2f era visualbasicnormal focus en lugar de 6
    '
    'If MiProceso <> 0 Then
    '    Do While MiResultado = STILL_ACTIVE
    '        Do_Events
    '        GetExitCodeProcess MiProceso, MiResultado
    '    Loop
    'End If
    lCogeTiquet = BufLcoge
End Sub
Public Function PRsinIVA(precio As Double, Percent As Double) As Double
    Dim Psin As Double
    Psin = (precio * 100) / (100 + Percent)
    If UsaEuro Then
        Psin = Round(Psin, 2)
    Else
        Psin = Round(Psin, decimales)
    End If
    PRsinIVA = Psin
End Function
Public Function Genera_CodBar(MisDatos As Tipo_TqCodBar, MiCodigo As Long) As String
    '**********************
    ' Tipo 0 --> Red Venta
    ' Tipo 5 --> Local Venta
    '**********************
    Dim MiPlu As Integer
    Dim bucle As Integer
    Dim codigo_barras As String
    Dim codigoDpto As String
    Dim snumtiq As String
    Dim sbarras$
    Dim cadena As String
    Dim envio
    Dim Checksum
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegArt As dao.Recordset
    Dim Buffer1
    Dim Buffer As String
    Dim Trozo As String
    Dim Hay_I As Boolean
    Dim Tipo As Integer
    Dim nC As Integer
    
    'If MisDatos.Tipo < 5 Then
    '    Tipo = 0
    'Else
    '    MisDatos.Tipo = MisDatos.Tipo - 4
    '    Tipo = 1
    'End If
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from codbar where borrado=false")
    '-->red
    With Registro
        If Not .EOF Then
            Select Case MisDatos.Tipo
                Case 1, 4, 7, 8
                    codigo_barras = !CB_VNT
                Case 2, 5
                    codigo_barras = !CB_EV1
                Case 3, 6
                    codigo_barras = !CB_EV2
                'Case 4
                '    codigo_barras = !cb_ev2
                Case Else
                    codigo_barras = ""
            End Select
        Else
            codigo_barras = ""
        End If
        If InStr(1, codigo_barras, "P") Then
            Set RegArt = Base.OpenRecordset("select * from articulo where borrado=false and codigo=" & MiCodigo)
                If Not RegArt.EOF Then
                    MiPlu = RegArt!Plu
                Else
                    MiPlu = 0
                End If
            RegArt.Close
        End If
        If codigo_barras <> "" Then
            Buffer1 = left(codigo_barras, 1)
            codigo_barras = codigo_barras & "ñ"
            For bucle = 2 To Len(codigo_barras)
                If Asc(Buffer1) > 47 And Asc(Buffer1) < 58 Then
                    Trozo = Buffer1
                    Buffer = Buffer & Buffer1
                    Trozo = ""
                Else
                    If Buffer1 <> Mid(codigo_barras, bucle, 1) Then
                        Trozo = Trozo & Buffer1
                        Select Case left(Trozo, 1)
                            Case "I"
                                If (MisDatos.Importe < 0) Then 'And (Not Sabeco) Then
                                    Hay_I = True
                                Else
                                    'If (MisDatos.Importe < 0) And Sabeco Then
                                    '    Buffer = Buffer & Right("000000000000", Len(Trozo))
                                    'Else
                                        nC = 0
                                        If Int(MisDatos.Importe) <> (MisDatos.Importe) Then
                                            Do While (Int(MisDatos.Importe) <> (MisDatos.Importe)) And nC < 4
                                                MisDatos.Importe = 10 * MisDatos.Importe
                                                nC = nC + 1
                                            Loop
                                        End If
                                        Buffer = Buffer & Right(Format(MisDatos.Importe, "000000000000"), Len(Trozo))
                                    'End If
                                End If
                            Case "N"
                                Buffer = Buffer & Right(Format(MisDatos.NBalanza, "000000000000"), Len(Trozo))
                            Case "Y"
                                Buffer = Buffer & Right(Format(MisDatos.NMostrador, "000000000000"), Len(Trozo))
                            Case "T"
                                Buffer = Buffer & Right(Format(MisDatos.NTicket, "000000000000"), Len(Trozo))
                            Case "V"
                                Buffer = Buffer & Right(Format(MisDatos.NVendedor, "000000000000"), Len(Trozo))
                            Case "C"
                                Buffer = Buffer & Right(Format(MiCodigo, "000000000000"), Len(Trozo))
                            Case "P"
                                Buffer = Buffer & Right(Format(MiPlu, "000000000000"), Len(Trozo))
                            Case Else
                                Hay_I = True
                        End Select
                        Trozo = ""
                    Else
                        Trozo = Trozo & Buffer1
                    End If
                End If
                Buffer1 = Mid(codigo_barras, bucle, 1)
            Next bucle
            codigo_barras = Buffer
            If Hay_I = False Then
                codigo_barras = codigo_barras + Format(DigitoControl(codigo_barras), "0")
            Else
                codigo_barras = ""
            End If
        End If
   End With
   On Error Resume Next
   Registro.Close
   Base.Close
   On Error GoTo 0
   
   If Len(codigo_barras) <> 13 Then
        codigo_barras = ""
   End If
   Genera_CodBar = codigo_barras
End Function
Public Function calcheck(ByVal micade As String, check)
Dim i
Dim buffercheck(255)
Dim longitud_cade

longitud_cade = Len(micade)

check = 0

For i = 1 To longitud_cade Step 1
    buffercheck(i) = Mid(micade, i, 1)

    If buffercheck(i) = Chr$(126) Then
        buffercheck(i) = Chr$(0)
    End If
    check = check Xor Asc(buffercheck(i))
Next i

check = check Xor (longitud_cade + 2)
check = ((check And 63) Or 64)
End Function
Public Function DigitoControl(ByVal Valor As String) As Integer
   '**************************
   ' Calcula el Checksum de
   ' un código de barras
   '***************************
   Dim Longitud As Integer
   Dim sumapar As Integer
   Dim sumaimpar As Integer
   Dim suma As Integer
   Dim digito, i As Integer
   sumapar = 0
   sumaimpar = 0
   Longitud = Len(Valor)
   For i = 1 To Longitud
      If i Mod 2 = 0 Then
         sumapar = sumapar + Val(Mid(Valor, i, 1))
      Else
         sumaimpar = sumaimpar + Val(Mid(Valor, i, 1))
      End If
   Next
   If Longitud Mod 2 = 0 Then
      suma = sumapar * 3 + sumaimpar
   Else
      suma = sumapar + sumaimpar * 3
   End If

   digito = 10 - suma Mod 10
   If digito = 10 Then digito = 0
   DigitoControl = digito
End Function
Public Function Quitar_Coma(ByVal MiCadena As String, ByVal Longitud As Integer) As String
    ' *****************
    ' Esta función sustituye
    ' el caracter "," (coma) por
    ' el caracter "." (punto)
    Dim C1 As Integer
    Dim HayPunto As Boolean
    Dim bucle As Integer
    Dim Buffer As String
    If Len(MiCadena) > Longitud Then
        Quitar_Coma = Space(Longitud)
        Exit Function
    End If
    For bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, bucle, 1) = "," Or Mid(MiCadena, bucle, 1) = "." Then
            HayPunto = True
            Buffer = Buffer & "."
        Else
            Buffer = Buffer & Mid(MiCadena, bucle, 1)
        End If
    Next bucle
    C1 = 0
    For bucle = Len(Buffer) To 1 Step -1
        If Mid(Buffer, bucle, 1) = "." Then
            HayPunto = True
            Exit For
        Else
            C1 = C1 + 1
        End If
    Next bucle
    If HayPunto = False Then
        If UsaEuro Or decimales = 2 Then
            Buffer = Buffer & ".00"
        Else
            Select Case decimales
                Case 1
                    Buffer = Buffer & ".0"
                Case 3
                    Buffer = Buffer & ".000"
            End Select
        End If
    Else
        Do While (UsaEuro And C1 < 2) Or (Not UsaEuro And C1 < decimales)
            Buffer = Buffer & "0"
            C1 = C1 + 1
        Loop
    End If
    Do While Len(Buffer) < Longitud
        Buffer = "0" & Buffer
    Loop
    Quitar_Coma = Buffer
End Function
Public Function EmularCodBar(MachineID As Integer, MiCodigo As String, MiSeccion As Integer, MiBalanza As Integer) As String
    Dim bucle As Integer
    Dim Buclelite As Integer
    Dim Plantilla As String
    Dim Contador As Integer
    Dim Buffer As String
    If MachineID = 0 Then EmularCodBar = MiCodigo
    For bucle = 1 To Len(MiCodigo)
        Contador = 1
        Select Case StrConv(Mid(MiCodigo, bucle, 1), vbUpperCase)
            Case "Y"
                Do Until (bucle + Contador = Len(MiCodigo) + 1) Or _
                (StrConv(Mid(MiCodigo, bucle + Contador, 1), vbUpperCase) <> "Y")
                    Contador = Contador + 1
                Loop
                Plantilla = ""
                For Buclelite = 1 To Contador
                    Plantilla = Plantilla & "0"
                Next Buclelite
                Buffer = Buffer & Right(Format(MiSeccion, Plantilla), Contador)
                bucle = Contador + bucle - 1
            Case "N"
                If Dir(App.Path & "\digitonuba.txt") = "" Then
                    Do Until (bucle + Contador = Len(MiCodigo) + 1) Or _
                    (StrConv(Mid(MiCodigo, bucle + Contador, 1), vbUpperCase) <> "N")
                        Contador = Contador + 1
                    Loop
                    Plantilla = ""
                    For Buclelite = 1 To Contador
                        Plantilla = Plantilla & "0"
                    Next Buclelite
                    Buffer = Buffer & Right(Format(MiBalanza, Plantilla), Contador)
                    bucle = Contador + bucle - 1
                Else
                    Buffer = Buffer & Mid(MiCodigo, bucle, 1)
                End If
            Case Else
                Buffer = Buffer & Mid(MiCodigo, bucle, 1)
        End Select
    Next bucle
    If Buffer = "" And (MachineID = 2 Or MachineID = 1) And TipoC9 Then Buffer = "000000000000"
    EmularCodBar = Buffer
End Function

'1.7.2
Public Function new_AbrirBase(Optional SoloLectura As Boolean) As dao.Database

Dim MiBase As dao.Database
Dim Reintentar As Boolean
Dim MiErr As Long
Dim StrErr As String
            Set MiBase = OpenDatabase(Base_General)
            Set new_AbrirBase = MiBase
End Function
'********
'***********
'***********
'***********
Public Function AbrirBase(Optional SoloLectura As Boolean) As dao.Database
Dim MiBase As dao.Database
Dim Reintentar As Boolean
Dim MiErr As Long
Dim StrErr As String

    
    Reintentar = True
    Exclusivo_Fin = False
    frmControl.TmrExclusivo.Interval = 25000
    frmControl.TmrExclusivo.Enabled = True
    Do While Reintentar And Exclusivo_Fin = False
        Reintentar = False
        If SoloLectura Then
            On Error Resume Next
            Set MiBase = OpenDatabase(Base_General, , True)
            MiErr = Err.Number
            StrErr = Err.Description
            On Error GoTo 0
        Else
            On Error Resume Next
            Set MiBase = OpenDatabase(Base_General)
            MiErr = Err.Number
            StrErr = Err.Description
            On Error GoTo 0
        End If
        If MiErr <> 0 Then
            Reintentar = True
            Sleep (100)
            Do_Events
        End If
    Loop
    frmControl.TmrExclusivo.Enabled = False
    If MiErr = 3045 Then
        MsgBox CargaCadena(972), vbCritical
        salir_programa True
    Else
        If MiErr <> 0 Then
            MsgBox CargaCadena(973) & ": " & MiErr & " " & Err.Description & vbCrLf & CargaCadena(974), vbCritical
            salir_programa True
        End If
    End If
    Set AbrirBase = MiBase
End Function
Public Sub CerrarBase(ByVal MiBase As dao.Database)  '(MiBase As dao.Database)
    MiBase.Close
    '2.0.12 Workspaces(0).Close
    Set MiBase = Nothing
    
End Sub
Public Function BoolToCheck(MiVal As Boolean) As Integer
    If MiVal Then
        BoolToCheck = vbChecked
    Else
        BoolToCheck = vbUnchecked
    End If
End Function
Public Function NumeroSistemas() As Integer
    Dim MyNum As Integer
    MyNum = 0
    If UsaGamaAlta Then MyNum = MyNum + 1
    If usaSC10 Then MyNum = MyNum + 1
    If UsaECO1 Then MyNum = MyNum + 1
    If UsaECO2 Then MyNum = MyNum + 1
    If UsaEcoPrint Then MyNum = MyNum + 1
    NumeroSistemas = MyNum
End Function
Public Function PrecioReal(MiCadena As String) As Double
    Dim Buffer As String
    Dim HayComa As Boolean
    Dim Contador As Integer
    Dim bucle As Integer
    Dim MiValor As Double
    HayComa = False
    Contador = 0
    For bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, bucle, 1) = "," Or Mid(MiCadena, bucle, 1) = "." Then
            HayComa = True
        Else
            If IsNumeric(Mid(MiCadena, bucle, 1)) Then
                Buffer = Buffer & Mid(MiCadena, bucle, 1)
                If HayComa Then Contador = Contador + 1
            End If
        End If
        
    Next bucle
    If IsNumeric(Buffer) Then
        MiValor = Val(Buffer) / (10 ^ Contador)
    Else
        MiValor = 0
    End If
    PrecioReal = MiValor
End Function

Public Sub CadenadeLogTotales(MiCadena As String)
    Dim Fichero As Integer
    
    
    On Error GoTo fin
    
    
    Fichero = FreeFile
    Open Miruta & "\loghydratotales.txt" For Append As Fichero
    On Error Resume Next
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
fin:
    On Error GoTo 0
End Sub

Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    
    On Error GoTo fin
    '3.8.3-5 --> se borra al arrancar
    'If FileLen(App.Path & "\loghydra.txt") > 16000000 Then
    '    Kill App.Path & "\loghydra.txt"
    '    Sleep (100)
    'End If
fin:
    On Error GoTo 0
    On Error GoTo FIN1
    Fichero = FreeFile
    Open Miruta & "\loghydratt.txt" For Append As Fichero
    On Error Resume Next
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
FIN1:
    On Error GoTo 0
    
End Sub

Public Sub salir_programa(Optional Forzoso As Boolean)
   Dim Arch As Integer
   Dim logntq As Integer
   Dim bucle As Integer
   Dim Base As dao.Database
   Dim Resp
   canCelar = 1
   lCogeTiquet = False
   
   On Error Resume Next
   
   If Dir(App.Path & "\laststart") <> "" Then
    Kill App.Path & "\laststart"
   End If
   
   
   If Dir(App.Path & "\hydrarun.txt") <> "" Then
       On Error Resume Next
       'If Dir(App.Path & "\musgrave.txt") <> "" Or UCase(Trim(Command)) = "HISTORICO" Then
           Sleep (100)
       'Else
       '    Sleep (2000)
       'End If
       Kill App.Path & "\hydrarun.txt"
       On Error GoTo 0
   End If
   
   If DataPlugin.mode <> "" And DataPlugin.type <> "" Then
        If Dir(Miruta & "\dirtmptouch") <> "" Then
            FrmHlink.STOPALL
        End If
   End If
   'If Dir(Miruta & "\dirtmptouch\tmp.mdb") <> "" Then
   '  Kill Miruta & "\dirtmptouch\tmp.mdb"
   'End If
   
   
   On Error GoTo 0
   '********************************
   ' Compacta la base y crea copia
   ' de seguridad automática
   '********************************
   If Not Forzoso Then
        Set frmControl = Nothing
        If PlugActivo = False Then
            If Base_General <> "" Then Compactar_Base
        End If
   End If
   '********************************
   ' Crea log de seguridad
   '********************************
siguiente:
   logntq = FreeFile
   If Dir(Miruta & "\loghydra.txt") <> "" Then
        Open Miruta & "\loghydra.txt" For Append As logntq
   Else
        Open Miruta & "\loghydra.txt" For Output As logntq
   End If
   Print #logntq, Now() & " Program Finished ************************************"
   Close logntq
   
   If Forzoso Then
        Resp = MsgBox(CargaCadena(975) _
        & vbCrLf & CargaCadena(1014) & vbCrLf & _
        CargaCadena(1015), vbCritical + vbYesNo, "HydraTouch")
        If Resp = vbYes Then
            On Error Resume Next
            DBEngine.RepairDatabase App.Path & "\dbasetouch.mdb"
            If Err.Number <> 0 Then
                If Dir(App.Path & "\dirtmptouch", vbDirectory) <> "" Then
                    If Dir(App.Path & "\dirtmptouch\compacta.mdb") <> "" Then
                        On Error Resume Next
                        FileCopy App.Path & "\dirtmptouch\compacta.mdb", App.Path & "\dbasetouch.mdb"
                        If Err.Number <> 0 Then
                            MsgBox CargaCadena(976), vbCritical
                        Else
                            On Error Resume Next
                            Set Base = OpenDatabase(Base_General)
                            If Err.Number <> 0 Then
                                MsgBox CargaCadena(976), vbCritical
                            Else
                                Base.Close
                                MsgBox CargaCadena(977), vbExclamation
                                End
                            End If
                        End If
                    End If
                End If
                MsgBox CargaCadena(976), vbCritical
            Else
                If PlugActivo = False Then
                    Compactar_Base
                End If
                MsgBox CargaCadena(977), vbExclamation
            End If
        End If
   End If
   
   If Dir(App.Path & "\killhydra") <> "" Then
       Kill App.Path & "\killhydra"
   End If
   If lAgora Then
    Shell (Environ("COMSPEC") & " /C " & "taskkill /F /IM tiqahtouch.exe")
   
    Sleep (1500)
    
    Shell (App.Path & "\tiqahtouch.exe")
    
   End If
   
   End
End Sub
Public Sub Do_Events()
    If General_Events >= 5 Then
        Sleep (1)
        General_Events = 0
    Else
        General_Events = General_Events + 1
    End If
    Sleep (1)
    DoEvents
    'Sleep (1)
End Sub
Public Function CodigoToPlu(codigo As Long) As Long
    Dim Cod As String
    Dim Plu As Long
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select codigo,plu from articulo where borrado=false and codigo=" & codigo)
    If Registro.EOF Then
        Plu = 0
    Else
        'plu = Registro.Fields("plu")
        If Not IsNull(Registro.Fields("plu")) Then
            Plu = Registro.Fields("plu")
        Else
            Plu = 0
        End If
        
    End If
    Registro.Close
    Base.Close
    Set Base = Nothing
    Set Registro = Nothing
    '1.7.2 ojo --> Workspaces(0).Close
    CodigoToPlu = Plu
End Function
Public Sub MostrarMonitor()
    CancelHumano = False
    If frmMonitor.Visible = False And Trim(Command) = "" Then
        frmMonitor.LblError.Visible = False
        
        frmMonitor.barra1.Value = 0
        frmMonitor.List1.Clear
        'If Not PlugActivo Then
            If HayComandos Then
                frmMonitor.cmdmonitorstop.Visible = False
                If AhorraMas Then frmMonitor.Command1.Visible = False
            Else
                frmMonitor.cmdmonitorstop.Visible = True
            End If
            If Not Silencio Then
                frmMonitor.Show
                frmMonitor.SetFocus
        'End If
                If frmEpelsa.Visible = True Then frmEpelsa.Enabled = False
            End If
        Do_Events
        Sleep (100)
    End If
End Sub
Public Sub HabilitarSalir()
    frmMonitor.Command1.Enabled = True
    frmMonitor.barra1.Value = frmMonitor.barra1.Max
End Sub

Public Sub LogTXT(ByVal cadena As String)
Dim bucle As Long
Dim Archivo As Integer
On Error Resume Next
Archivo = FreeFile
Open "c:\log.txt" For Output As Archivo
Write #Archivo, cadena
For bucle = 1 To Len(cadena)
    Print #Archivo, "(" & bucle & ")" & Mid(cadena, bucle, 1) & "," & Asc(Mid(cadena, bucle, 1)) & "|"
Next bucle
Close #Archivo
On Error GoTo 0
End Sub

Public Function PTAtoEURO(ByVal pelas As Double) As Double
    '*************************
    ' Función Desactivada,
    ' se mantiene por compatibilidad
    ' versiones antiguas
    '*************************
    PTAtoEURO = 1
End Function
Public Function EUROtoPTA(ByVal pelas As Double) As Double
    '*************************
    ' Función Desactivada,
    ' se mantiene por compatibilidad
    ' versiones antiguas
    '*************************
    EUROtoPTA = 1
End Function
Public Function Checktexto(ByVal CodigoTecla As Integer, ByVal Maximo As Integer, ByVal LaCadena As String, SoloNumeros As Boolean, UsaDecimal As Boolean) As Boolean
    Dim Valor As Boolean
    Dim bucle As Integer
    Dim Longitud As Integer
    Dim Lentero As Integer
    Dim lDec As Integer
    Dim ZonaDec As Boolean
    Longitud = Len(LaCadena)
    Valor = False
    If SoloNumeros Then
        If CodigoTecla < 48 Or CodigoTecla > 57 Then
            If Not UsaDecimal Or (UsaDecimal And CodigoTecla <> 44 And CodigoTecla <> 46) Then
                Valor = True
            End If
        End If
    End If
    If UsaDecimal And (CodigoTecla = 44 Or CodigoTecla = 46) Then
        For bucle = 1 To Longitud
            If Mid(LaCadena, bucle, 1) = "." Or Mid(LaCadena, bucle, 1) = "," Then
                Valor = True
            End If
        Next bucle
    End If
    If Longitud >= Maximo Then Valor = True
       
    If CodigoTecla = vbKeyBack Or CodigoTecla = vbKeyDelete Then Valor = False
    Checktexto = Valor
End Function
Public Function strtranñÑ(ByVal sR As String)
   
   Dim tmp As String
   Dim nLen As Integer
   Dim Cont As Integer
   Dim s As String
   
   Dim sResul As String
   
   nLen = Len(sR)
   For Cont = 1 To nLen
      s = Mid(sR, Cont, 1)
      Select Case s
      Case "Á"
         s = Chr(181)
      Case "Í"
         s = Chr(214)
      Case "Ó"
         s = Chr(224)
      Case "Ú"
         s = Chr(233)
      Case "Ý"
         s = Chr(237)
      Case "À"
         s = Chr(183)
      Case "Ò"
         s = Chr(227)
      Case "Ù"
         s = Chr(235)
      Case "Ñ"
         s = Chr(165)
      Case "ñ"
         s = Chr(164)
      Case "¿"
         s = Chr(168)
      Case "¡"
         s = Chr(173)
      Case "Ç"
         s = Chr(128)
      Case "ç"
         s = Chr(135)
      Case "Ê"
         s = Chr(210)
      Case "Ó"
         s = Chr(224)
      Case "È"
         s = Chr(212)
      Case "â"
         s = Chr(131)
      Case "ê"
         s = Chr(136)
      Case "î"
         s = Chr(140)
      Case "ô"
         s = Chr(147)
      Case "û"
         s = Chr(150)
      Case "º"
         s = Chr(167)
      Case "ª"
         s = Chr(166)
      Case "!"
         s = Chr(33)
      Case "¦"
         s = Chr(179)
      Case "#"
         s = Chr(35)
      Case "~"
         s = Chr(126)
      Case "¬"
         s = Chr(170)
      Case ":"
         s = Chr(58)
      Case "ó"
         s = Chr(162)
      Case "á"
         s = Chr(160)
      Case "Ä"
         s = Chr(142)
      Case "É"
         s = Chr(144)
      Case "Í"
         s = Chr(214)
      Case "Ó"
         s = Chr(224)
      Case "Ú"
         s = Chr(233)
      Case "é"
         s = Chr(130)
      Case "í"
         s = Chr(161)
      Case "ú"
         s = Chr(163)
      Case "`"
         s = Chr(96)
      Case "ä"
         s = Chr(132)
      Case "ë"
         s = Chr(137)
      Case "ï"
         s = Chr(139)
      Case "ö"
         s = Chr(148)
      Case "ü"
         s = Chr(129)
    
      End Select
      
      sResul = sResul + s
   
   Next
   
   strtranñÑ = sResul

End Function

Public Function strTranArroba(ByVal s As String)
    Dim L As Integer
    Dim ncont As Integer
    L = Len(s)
    For ncont = 1 To L
        If Mid(s, ncont, 1) = "@" Then
            s = Mid(s, 1, ncont - 1) & "Ñ" & Mid(s, ncont + 1)
        End If
    Next ncont
    strTranArroba = s
End Function

Public Function toIpAdd(ByVal sR As String)
   Dim tmp As String
   Dim nLen As Integer
   Dim Cont As Integer
   Dim s As String
   nLen = Len(sR)
   For Cont = 1 To nLen
      s = Mid(sR, Cont, 1)
      If Cont > 1 Then
      If (Asc(s) < 49 Or Asc(s) > 57) And Mid(sR, Cont - 1, 1) = "." Then
      Else
         tmp = tmp + s
      End If
      Else
         If Asc(s) < 49 Or Asc(s) > 57 Then
         Else
            tmp = tmp + s
         End If
      End If
      
   Next
   
   toIpAdd = tmp
   
End Function

Public Function leeIp() As Boolean
    Dim Arch As Integer
    'Dim MyIp As String
    Dim MyPath As String
'***********************************************
    If Dir(App.Path & "\servidoratouch.ini") = "" Then
        Check_ServidoraIP
    End If
    Arch = FreeFile()
    Open App.Path & "\servidoratouch.ini" For Input As #Arch
    If Not EOF(Arch) Then
        Line Input #Arch, MyIP
    End If
    Close #Arch
    If MyIP = "" Then
        leeIp = False
    Else
        sAddIp = MyIP
        leeIp = True
    End If
End Function
Public Sub Check_ServidoraIP()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim MyIP As String
    Dim Arch As Integer
    Dim MiPath As String
    If Not HayMulti Then
        MiPath = App.Path
    Else
        If TiendaActual < 100 Then
        MiPath = App.Path & "\T" & Format(TiendaActual, "00")
        Else
            MiPath = App.Path & "\T" & Format(TiendaActual, "000")
        End If
    End If
    If Dir(MiPath, vbDirectory) <> "" Then
        MiPath = MiPath & "\dns.txt"
        If Dir(MiPath) = "" Then MiPath = ""
    Else
        MiPath = ""
    End If
    If MiPath <> "" Then
        MyIP = "127.000.000.001"
    Else
        Set Base = OpenDatabase(Base_General)
        Set Registro = Base.OpenRecordset("select * from equipos where modelo=1 and borrado=false")
        With Registro
            If Not .EOF Then
                If Len(!prog_eqp) = 15 Then
                    If IsNumeric(left(!prog_eqp, 3)) And IsNumeric(Mid(!prog_eqp, 5, 3)) And IsNumeric(Mid(!prog_eqp, 9, 3)) And IsNumeric(Right(!prog_eqp, 3)) Then
                        MyIP = Val(left(!prog_eqp, 3)) & "." & Val(Mid(!prog_eqp, 5, 3)) & "." & Val(Mid(!prog_eqp, 9, 3)) & "." & Val(Right(!prog_eqp, 3))
                    End If
                End If
            End If
        End With
        If MyIP = "999.999.999.999" Then MyIP = ""
        Registro.Close
        Base.Close
        Set Registro = Nothing
        Set Base = Nothing
        'Workspaces(0).Close c2f 2.0.25
        
    End If
    Arch = FreeFile()
    Open App.Path & "\servidoratouch.ini" For Output As #Arch
    Print #Arch, MyIP
    Close Arch
End Sub

Public Function comFecha(ByVal Fecha As String) As Boolean
Dim d As Integer
Dim m As Integer
Dim y As Integer
Dim forFecha As String

d = CInt(left(Fecha, 2))
m = CInt(Mid(Fecha, 4, 2))
y = CInt(Right(Fecha, 2))

forFecha = DateSerial(y, m, d)

    If Fecha <> forFecha Then
        'comFecha = False
        comFecha = True
    Else
        comFecha = True
    End If

End Function


Public Sub grventas()
Dim Base As dao.Database
Dim rstgtsecgra As dao.Recordset
Dim dato As Variant
Dim Filtro As String
Dim conmos As Integer
Dim sSQL As String
Dim sfechaNow As String
Dim nPos As Integer

'If Taquion Then Exit Sub '1.7.2
If Base_General = "" Then Exit Sub
If Dir(App.Path & "\nover1.txt") <> "" Then Exit Sub

If Dir(Base_General) <> "" Then
If HaySeguridad And UsuarioActual.Asociado <> 0 Then
    graUsr
    Exit Sub
End If
sfechaNow = CStr(CDate(Now))
nPos = InStr(1, sfechaNow, " ")
If nPos <> 0 Then
    sfechaNow = Mid(sfechaNow, 1, nPos - 1)
End If
Set Base = OpenDatabase(Base_General)
Set rstgtsecgra = Base.OpenRecordset("SELECT gtsecs.codi_ident, Sum(gtsecs.pesetas) AS SumaDepesetas From gtsecs where cstr(cdate(fecha))=" & Chr(34) & sfechaNow & Chr(34) & " GROUP BY gtsecs.codi_ident ")
'sSQL = Base.OpenRecordset("SELECT gtsecs.codi_ident, Sum(gtsecs.pesetas) AS SumaDepesetas From gtsecs where fecha=#" & Format(Now, "mm/dd/yy") & "# GROUP BY gtsecs.codi_ident ")
With rstgtsecgra
    If Not .EOF Then
        If frmEpelsa.gRafico.Visible = False Then
        frmEpelsa.gRafico.Visible = True
        End If
        dato = .Recordcount
        'frmEpelsa.gRafico.chartType = 1
        frmEpelsa.gRafico.ColumnCount = 1
        frmEpelsa.gRafico.Column = 1
        frmEpelsa.gRafico.RowCount = dato
        conmos = 1
        Do While conmos <= dato
        frmEpelsa.gRafico.Row = conmos
        frmEpelsa.gRafico.RowLabel = !codi_ident
        frmEpelsa.gRafico.data = !sumadepesetas
        conmos = conmos + 1
        .Movenext
        Loop
        'frmEpelsa.gRafico.ColumnLabel = cargacadena( 464)
    Else
        frmEpelsa.gRafico.Visible = False
    End If
.Close
End With
'1.7.2 CerrarBase Base
Set rstgtsecgra = Nothing
Base.Close
Set Base = Nothing

    If Trim(Command) = "" Then '2.0.26
    'Workspaces(0).close
    End If
End If
End Sub
Public Sub graUsr()
Dim Base As dao.Database
Dim rstgtsecgra As dao.Recordset
Dim dato, conmos As Integer
Dim sfechaNow As String
Dim nPos As Integer
    
    If Dir(Miruta & "\dbasetouch.mdb") <> "" Then
        
        sfechaNow = CStr(CDate(Now))
        nPos = InStr(1, sfechaNow, " ")
        If nPos <> 0 Then
            sfechaNow = Mid(sfechaNow, 1, nPos - 1)
        End If
        
        Set Base = OpenDatabase(Base_General)
        'Set rstgtsecgra = Base.OpenRecordset("SELECT gtsecs.codi_ident, gtsecs.codi_sub, Sum(gtsecs.pesetas) AS SumaDepesetas From gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Now)) & Chr(34) & " GROUP BY gtsecs.codi_ident, gtsecs.codi_sub HAVING (((gtsecs.codi_ident)= " & UsuarioActual.Asociado & " ))")
        Set rstgtsecgra = Base.OpenRecordset("SELECT gtsecs.codi_ident, gtsecs.codi_sub, Sum(gtsecs.pesetas) AS SumaDepesetas From gtsecs where cstr(cdate(fecha))=" & Chr(34) & sfechaNow & Chr(34) & " GROUP BY gtsecs.codi_ident, gtsecs.codi_sub HAVING (((gtsecs.codi_ident)= " & UsuarioActual.Asociado & " ))")
        With rstgtsecgra
        If Not .EOF Then
            If frmEpelsa.gRafico.Visible = False Then
            frmEpelsa.gRafico.Visible = True
            End If
            frmEpelsa.gRafico.chartType = VtChChartType2dBar
            dato = .Recordcount
            frmEpelsa.gRafico.RowCount = dato
            frmEpelsa.gRafico.ColumnCount = 1
            frmEpelsa.gRafico.Column = 1
            conmos = 1
            Do While conmos <= dato
            frmEpelsa.gRafico.Row = conmos
            frmEpelsa.gRafico.RowLabel = !codi_sub
            frmEpelsa.gRafico.data = !sumadepesetas
            conmos = conmos + 1
            .Movenext
            Loop
            'frmEpelsa.gRafico.ColumnLabel = cargacadena( 465)
        Else
            frmEpelsa.gRafico.Visible = False
        End If
        .Close
        End With
        'CerrarBase Base
        Set rstgtsecgra = Nothing
        Base.Close
        Set Base = Nothing
        'Workspaces(0).close
        
    End If
End Sub
'\\\\\\\\\\\
'///////////
Public Function version_app() As String
    version_app = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function
'\\\\\\\\\\\
'///////////
Public Function importe_bruto(ByVal aImporteNeto As Double, ByVal aTpcDescuento As Byte) As Double
Dim Res As Double
    Res = (aImporteNeto * 100) / (100 - aTpcDescuento)
    importe_bruto = Res
End Function
'\\\\\\\\\\\
'///////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Redondea la diferencia que se produce al tratar datos con decimales, con Access y VB
'////////////////////////////////////////////////////////////////////////////////////
Public Function redondea(ByVal dato As Double) As Double
    If UsaEuro Then
        dato = Round(dato, 2)
    Else
        dato = Round(dato, decimales)
    End If
    redondea = dato
End Function

Public Function separaFich(ByVal s As String) As String
   Dim Ls As Integer
   Dim nLs As Integer
   Dim bucle As Integer
   Ls = Len(s)
   For bucle = Ls To 1 Step -1
       If Mid(s, bucle, 1) = "\" Then
           s = Mid(s, bucle + 1)
           bucle = Ls + 1
       End If
   Next bucle
   separaFich = s
End Function

Public Function quita_ext(ByVal s As String) As String
   Dim Ls As Integer
   Dim nLs As Integer
   Dim bucle As Integer
   nLs = InStr(1, s, ".")
   If nLs <> 0 Then
     s = Mid(s, 1, nLs - 1)
   End If
   quita_ext = s
End Function
Public Function fVerLeclerc() As Integer

    'leclerc
    ' si existe fichero: spathmaj & "\bal_maj.lck" es que otro proceso accede a fichero.
    'Fichero : sPathMaj & "\MAJnaaaa.qqq" --> qqq=Ku,Ku,Ru,Ku,Ku ¡Juliaaaanooooo!
    '                                     --> n=número de orden, se almacenará en app.path & "\Nqqq.txt" --> "N"literal
    '                                     --> aaaa=año.
    ' El fichero se borrará tras su importación.
    ' si existe fichero LOGLEC.TXT se almacenará en un fichero MAJnaaaa.log los rechazos"
    Dim sJuliano As String
    Dim sFichMaj As String
    Dim nArch As Integer
    Dim xArch As Integer
    Dim sLosArch(10) As String
    Dim nLosArch As Integer
    Dim ncont As Integer
    Dim nCuantos As Integer
    Dim sTmp As String
    Dim nVerLeclerc As Integer
    Dim scad As String
    'para leclerc España
    'Dim Registro As RegMaj
    Dim Registro As RegMajNew
    Dim nCampo As Integer
    Dim sMiPlu As String
    Dim sMisDes(21) As String
    
    If lClR And sPathMaj <> "" Then
        
        nArch = FreeFile()
        Open App.Path & "\leclerc.log" For Output As #nArch
        'para versión en España... se anula, Francia se activa
        If Dir(sPathMaj & "\bal_maj.lck") = "" Then
    
            frmleclerc.Show
            
            sJuliano = Calcula_Juliano()
            'vemos cuantos ficheros se podrían importar...
            ' Si hubiera más de un fichero se unificarán en MAJXAAAA.QQQ
            nCuantos = 0
            
            'para versión en España... se anula
            For ncont = 0 To 9
                sFichMaj = sPathMaj & "\maj" & Chr(48 + ncont) & sJuliano
                If Dir(sFichMaj) <> "" Then
                    sLosArch(ncont) = sFichMaj
                    nCuantos = nCuantos + 1
                Else
                    If Dir(UCase(sFichMaj)) <> "" Then
                        sLosArch(ncont) = sFichMaj
                        nCuantos = nCuantos + 1
                    Else
                        sLosArch(ncont) = ""
                    End If
                End If
            Next ncont
            sFichMaj = sPathMaj & "\maj" & sJuliano
            
            'PRUEBA QUITAR...1.0.3
            'sFichMaj = "c:\leclerc\maj2008.115"
            ''''''''''''''''''''''
            
            If Dir(sFichMaj) <> "" Then
                ncont = 0
                sLosArch(ncont) = sFichMaj
                nCuantos = nCuantos + 1
            End If
            
            If nCuantos > 0 Then
                If nCuantos >= 1 Then
                    xArch = FreeFile()
                    
                    Open sPathMaj & "\majx" & sJuliano For Append As #xArch
                    
                    For ncont = 0 To 9
                        If sLosArch(ncont) <> "" Then
                            
                            Print #nArch, Now() & " --> Transfiriendo:" & sLosArch(ncont) & " a MAJX" & sJuliano
                            
                            nLosArch = FreeFile()
                            Open sLosArch(ncont) For Input As #nLosArch
                            Do While Not EOF(nLosArch)
                                Line Input #nLosArch, sTmp
                                If Len(sTmp) > 10 Then
                                    Print #xArch, sTmp
                                End If
                            Loop
                            Close #nLosArch
                            Kill sLosArch(ncont)
                        End If
                    
                    Next ncont
                    Close #xArch
                    sFichMaj = sPathMaj & "\majx" & sJuliano
                'Se anula para España
                '''''''''''''''''''''
                Else
                    For ncont = 0 To 9
                        If sLosArch(ncont) <> "" Then
                            sFichMaj = sLosArch(ncont)
                            Exit For
                        End If
                    Next ncont
                '''''''''''''''''''''
                End If
            End If
            
            'continuar proceso...
            If nCuantos > 0 Then
                xArch = FreeFile()
                Open sFichMaj For Input As #xArch
                nLosArch = FreeFile()
                Open App.Path & "\bal.dat" For Output As #nLosArch
                Do While Not EOF(xArch)
                    Line Input #xArch, sTmp
                    scad = ""
                    nCampo = 1
                    Registro.Tipo = Mid(sTmp, 1, 1)
                    For ncont = 3 To Len(sTmp)
                        If Mid(sTmp, ncont, 1) <> "#" Then
                            scad = scad & Mid(sTmp, ncont, 1)
                        Else
                            'Select Case nCampo
                            '    Case 1
                            '        Registro.seccion = sCad
                            '    Case 2
                            '        Registro.codigo = sCad
                            '    Case 3
                            '        Registro.des1 = sCad
                            '    Case 4
                            '        Registro.precio = sCad
                            '    Case 5
                            '        Registro.paridad = sCad
                            '    Case 6
                            '        Registro.WGH = sCad
                            '    Case 7
                            '        Registro.ean = sCad
                            '    Case 8
                            '        Registro.cad = sCad
                            '    Case 9
                            '        Registro.cad2 = sCad
                            '    Case 10
                            '        Registro.desresto = sCad
                            '    Case 11
                            '        Registro.tipotara = sCad
                            '    Case 12
                            '        Registro.tara = sCad
                            '    Case 13
                            '        Registro.parlinea = sCad
                            '    Case 14
                            '        Registro.pesofijo = sCad
                            '    Case 15
                            '        Registro.tecla = sCad
                            'End Select
                            
                            If scad <> "" Then
                            Select Case nCampo
                                Case 1
                                    Registro.seccion = scad
                                Case 2
                                    Registro.codigo = scad
                                Case 3
                                    Registro.des1 = scad
                                Case 4
                                    Registro.precio = scad
                                
                                Case 6
                                    Registro.WGH = scad
                                Case 7
                                    Registro.ean = scad
                                Case 8
                                    Registro.cad = scad
                                Case 9
                                    Registro.cad2 = scad
                                
                                Case 11
                                    Registro.tipotara = scad
                                Case 12
                                    Registro.tara = scad
                                Case 13
                                    Registro.cod_ingre = scad
                                Case 14
                                    Registro.pesofijo = scad
                                
                                Case 26
                                    Registro.des2 = scad
                                Case 27
                                    Registro.des3 = scad
                                Case 28
                                    Registro.des4 = scad
                                Case 34
                                    Registro.promocion = scad
                                    
                                Case 36
                                    Registro.Etiqueta = scad
                            End Select
                            End If
                            scad = ""
                            nCampo = nCampo + 1
                        End If
                    Next ncont
                    'Francia
                    '' en registro tenemos datos
                    '' preparar linea de bal.dat
                    '' Requisitos para admitir linea:
                    ''     - Código de artículo > 0
                    ''     - sección > 0 < 60
                    ''     - si el descriptivo de artículo > 25 (des1) se trunca a 25
                    ''     - resto de lineas se crean 20 de 25 caracteres (desresto) --> puede haber CR
                    ''     - si es un registro con tipo de tara se rechaza
                    ''     - los tipos de venta que se admiten son: Peso, unidades
                    ''     -

                    If Registro.codigo = "" And Registro.ean <> "" Then
                        Registro.ean = Mid(Registro.ean, 4, 4)
                    End If
                    If Registro.codigo = "" Then Registro.codigo = "0"

                    If Registro.WGH = "1" Then Registro.WGH = "W"
                    If Registro.WGH = "2" Then Registro.WGH = "U"
                    
                    If Registro.seccion = "" And Registro.ean <> "" Then
                        Registro.seccion = Mid(Registro.ean, 3, 1)
                    End If
                    If Registro.seccion = "" Then Registro.seccion = "1"
                    
                    If Registro.ean = "" Then Registro.ean = Space(12)
                    If Len(Registro.ean) > 12 Then Registro.ean = Mid(Registro.ean, 1, 12)
                    If Len(Registro.ean) < 12 Then
                        Registro.ean = Registro.ean & Space(12 - Len(Registro.ean))
                    End If
                    'If Val(Mid(Registro.ean, 3, 1)) = Val(Registro.seccion) And Val(Mid(Registro.ean, 8, 5)) = 0 Then
                    If Val(Mid(Registro.ean, 8, 5)) = 0 Then
                        'Francia
                        Registro.ean = Mid(Registro.ean, 1, 7) & "jjjjj" '1.0.7 "JJJJJ" '"IIIII"
                        'España
                        'Registro.ean = Mid(Registro.ean, 1, 7) & "IIIII"
                    End If
                    If Val(Registro.seccion) > 60 Then Registro.seccion = "60"
                    
                    'Francia
                    If Val(Registro.codigo) > 9999 Then
                        sMiPlu = "0000"
                    Else
                        sMiPlu = Registro.codigo
                    End If
                    'España
                    'sMiPlu = "0000"
                    
                    Registro.seccion = Format(Val(Registro.seccion), "00")
                    Registro.codigo = Format(Val(Registro.codigo), "000000")
                    sMiPlu = Format(Val(sMiPlu), "0000")
                    If Len(Trim(Registro.des1)) < 25 Then
                        Registro.des1 = Trim(Registro.des1) '& Space(25 - Len(Registro.des1))
                        If Len(Registro.des1) < 25 Then
                            Registro.des1 = Registro.des1 & Space(25 - Len(Registro.des1))
                        End If
                        sMisDes(1) = Space(25)
                    Else
                        sMisDes(1) = Mid(Registro.des1, 26)
                        If Len(sMisDes(1)) > 25 Then
                            sMisDes(1) = Mid(sMisDes(1), 1, 25)
                        End If
                        Registro.des1 = Mid(Registro.des1, 1, 25)
                    End If
                    sMisDes(0) = Registro.des1
                    
                    'España
                    'Registro.desresto = Trim(Registro.des2) & " " & Trim(Registro.des3) & " " & Trim(Registro.des4)
                    'Registro.desresto = Trim(Registro.desresto)
                    '''''''
                    
                    Registro.precio = Format(Val(Registro.precio), "0000000")
                    If Registro.Tipo = "S" Then
                        sMisDes(0) = Space(25)
                        Registro.precio = "0000000"
                    End If
                    If Val(Registro.codigo) > 0 Then
                        If Registro.WGH <> "W" And Registro.WGH <> "U" Then
                            Print #nArch, Now() & " --> " & sTmp & " record no procesado forma de venta de artículo no admitida."
                        Else
                        
                            'descriptivos adicionales
                            
                            If Trim(sMisDes(1)) <> "" Then
                                For ncont = 2 To 20
                                    sMisDes(ncont) = Space(25)
                                Next ncont
                            Else
                                For ncont = 1 To 20
                                    sMisDes(ncont) = Space(25)
                                Next ncont
                            
                            End If
                            If Len(Registro.desresto) > 0 Then
                                nCuantos = 0
                                ncont = InStr(1, Registro.desresto, "~M~J", 1)
                                Do While ncont <> 0 And nCuantos < 20
                                    nCuantos = nCuantos + 1
                                    If ncont <> (Len(Registro.desresto) - 3) Then
                                        Registro.desresto = Mid(Registro.desresto, 1, (ncont - 1)) & Mid(Registro.desresto, ncont + 4)
                                    Else
                                        Registro.desresto = Mid(Registro.desresto, 1, Len(Registro.desresto) - 4)
                                    End If
                                    ncont = InStr(1, Registro.desresto, "~M~J", 1)
                                
                                Loop
                                If Len(Registro.desresto) < 500 Then
                                    Registro.desresto = Registro.desresto & Space(500 - Len(Registro.desresto))
                                Else
                                    Registro.desresto = Mid(Registro.desresto, 1, 500)
                                End If
                                
                                If sMisDes(1) <> Space(25) Then
                                
                                'sMisDes(1) = Mid(Registro.desresto, 1, 25)
                                sMisDes(2) = Mid(Registro.desresto, 1, 25)
                                sMisDes(3) = Mid(Registro.desresto, 26, 25)
                                sMisDes(4) = Mid(Registro.desresto, 51, 25)
                                sMisDes(5) = Mid(Registro.desresto, 76, 25)
                                sMisDes(6) = Mid(Registro.desresto, 101, 25)
                                sMisDes(7) = Mid(Registro.desresto, 126, 25)
                                sMisDes(8) = Mid(Registro.desresto, 151, 25)
                                sMisDes(9) = Mid(Registro.desresto, 176, 25)
                                sMisDes(10) = Mid(Registro.desresto, 201, 25)
                                sMisDes(11) = Mid(Registro.desresto, 226, 25)
                                sMisDes(12) = Mid(Registro.desresto, 251, 25)
                                sMisDes(13) = Mid(Registro.desresto, 276, 25)
                                sMisDes(14) = Mid(Registro.desresto, 301, 25)
                                sMisDes(15) = Mid(Registro.desresto, 326, 25)
                                sMisDes(16) = Mid(Registro.desresto, 351, 25)
                                sMisDes(17) = Mid(Registro.desresto, 376, 25)
                                sMisDes(18) = Mid(Registro.desresto, 401, 25)
                                sMisDes(19) = Mid(Registro.desresto, 426, 25)
                                sMisDes(20) = Mid(Registro.desresto, 451, 25)
                                'sMisDes(20) = Mid(Registro.desresto, 476, 25)
                                Else
                                sMisDes(1) = Mid(Registro.desresto, 1, 25)
                                'sMisDes(2) = Mid(Registro.desresto, 1, 25)
                                sMisDes(2) = Mid(Registro.desresto, 26, 25)
                                sMisDes(3) = Mid(Registro.desresto, 51, 25)
                                sMisDes(4) = Mid(Registro.desresto, 76, 25)
                                sMisDes(5) = Mid(Registro.desresto, 101, 25)
                                sMisDes(6) = Mid(Registro.desresto, 126, 25)
                                sMisDes(7) = Mid(Registro.desresto, 151, 25)
                                sMisDes(8) = Mid(Registro.desresto, 176, 25)
                                sMisDes(9) = Mid(Registro.desresto, 201, 25)
                                sMisDes(10) = Mid(Registro.desresto, 226, 25)
                                sMisDes(11) = Mid(Registro.desresto, 251, 25)
                                sMisDes(12) = Mid(Registro.desresto, 276, 25)
                                sMisDes(13) = Mid(Registro.desresto, 301, 25)
                                sMisDes(14) = Mid(Registro.desresto, 326, 25)
                                sMisDes(15) = Mid(Registro.desresto, 351, 25)
                                sMisDes(16) = Mid(Registro.desresto, 376, 25)
                                sMisDes(17) = Mid(Registro.desresto, 401, 25)
                                sMisDes(18) = Mid(Registro.desresto, 426, 25)
                                sMisDes(19) = Mid(Registro.desresto, 451, 25)
                                sMisDes(20) = Mid(Registro.desresto, 476, 25)
                                
                                End If
                            
                            End If
                            
                            
                            scad = Registro.seccion & Registro.codigo & sMiPlu & Registro.des1 & Registro.precio & _
                                   Registro.WGH & Registro.ean & Format(Val(Registro.cad), "000")
                            'en principio cuento con 10 descriptivos despues se amplia
                            For ncont = 1 To 20
                                scad = scad & sMisDes(ncont)
                            Next
                            scad = scad & Format(Val(Registro.tara), "00000")
                            'leclerc --> en principio sin tecla directa.
                            'sCad = sCad & Format(Val(Registro.tecla), "000")
                            scad = scad & "0" '--> Juego
                            scad = scad & "000" '--> Tecla
                            scad = scad & "001" '--> Familia
                            'España
                            'If Registro.tecla = "" Then Registro.tecla = "0000"
                            'Registro.tecla = Format(Val(Registro.tecla), "0000")
                            'sCad = sCad & Format(Val(Registro.tecla), "0000")
                            
                            Print #nLosArch, scad
                            
                        End If
                
                    Else
                    
                        Print #nArch, Now() & " --> " & sTmp & " record no procesado Código de artículo inexistente en el record."
                        
                    End If
                
                Loop
                Close #xArch
                Close #nLosArch
            
                Kill sFichMaj
                
            Else
                Print #nArch, Now() & " --> No existe fichero MAJ válido para importar."
                nVerLeclerc = 1
            End If

            frmleclerc.Hide
            
        Else
        
            Print #nArch, Now() & " --> Existe fichero de bloqueo"
            nVerLeclerc = 2 'existe fichero bloqueo
            
        End If
    
        Close #nArch
        
    End If

    fVerLeclerc = nVerLeclerc
    
End Function
    
Public Function AdaptaChar(ByVal sR As String, ByVal nT As Integer) As String
    Dim s As String
    s = sR
    If Dir(App.Path & "\nocht.txt") <> "" Then
        AdaptaChar = s
    Else
        Select Case nT
            Case 0
                If Dir(App.Path & "\mipel.cht") = "" Then
                    AdaptaChar = EurChar(sR)
                Else
                    AdaptaChar = s
                End If
            Case 1 To 9
                AdaptaChar = V12Char(sR)
        End Select
    End If
End Function
Public Function EurChar(ByVal sR As String) As String
   Dim tmp As String
   Dim nLen As Integer
   Dim Cont As Integer
   Dim s As String
   Dim sResul As String
   nLen = Len(sR)
   For Cont = 1 To nLen
      s = Mid(sR, Cont, 1)
      Select Case s
      Case "Á"
         s = Chr(181)
      Case "Í"
         s = Chr(214)
      Case "Ó"
         s = Chr(224)
      Case "Ú"
         s = Chr(233)
      Case "Ý"
         s = Chr(237)
      Case "À"
         s = Chr(183)
      Case "Ò"
         s = Chr(227)
      Case "Ù"
         s = Chr(235)
      Case "Ñ"
         s = Chr(165)
      Case "ñ"
         s = Chr(164)
      Case "¿"
         s = Chr(168)
      Case "¡"
         s = Chr(173)
      Case "Ç"
         s = Chr(128)
      Case "ç"
         s = Chr(135)
      Case "Ê"
         s = Chr(210)
      Case "Ó"
         s = Chr(224)
      Case "È"
         s = Chr(212)
      Case "â"
         s = Chr(131)
      Case "ê"
         s = Chr(136)
      Case "î"
         s = Chr(140)
      Case "ô"
         s = Chr(147)
      Case "û"
         s = Chr(150)
      Case "º"
         s = Chr(167)
      Case "²"
         s = Chr(167)
      Case "ª"
         s = Chr(166)
      Case "!"
         s = Chr(33)
      Case "¦"
         s = Chr(179)
      Case "#"
         s = Chr(35)
      Case "~"
         s = Chr(126)
      Case "¬"
         s = Chr(170)
      Case ":"
         s = Chr(58)
      Case "ó"
         s = Chr(162)
      Case "á"
         s = Chr(160)
      Case "Ä"
         s = Chr(142)
      Case "É"
         s = Chr(144)
      Case "Í"
         s = Chr(214)
      Case "Ó"
         s = Chr(224)
      Case "Ú"
         s = Chr(233)
      Case "é"
         s = Chr(130)
      Case "í"
         s = Chr(161)
      Case "ú"
         s = Chr(163)
      Case "`"
         s = Chr(96)
      Case "ä"
         s = Chr(132)
      Case "ë"
         s = Chr(137)
      Case "ï"
         s = Chr(139)
      Case "ö"
         s = Chr(148)
      Case "ü"
         s = Chr(129)
      End Select
      sResul = sResul + s
   Next
   EurChar = sResul
End Function
Public Function V12Char(sR As String) As String
    Dim Buffer As String
    Dim Buffer2 As String
    Dim bucle As Integer
    Buffer2 = ""
    Buffer = StrConv(sR, vbUpperCase)
    For bucle = 1 To Len(Buffer)
        Select Case Mid(Buffer, bucle, 1)
        Case "Ñ"
            Buffer2 = Buffer2 & "¥"
        Case "@"
            Buffer2 = Buffer2 & "Ñ"
        Case "Á"
            Buffer2 = Buffer2 & "A"
        Case "É"
            Buffer2 = Buffer2 & "E"
        Case "Í"
            Buffer2 = Buffer2 & "I"
        Case "Ó"
            Buffer2 = Buffer2 & "O"
        Case "Ú"
            Buffer2 = Buffer2 & "U"
        Case "À"
            Buffer2 = Buffer2 & "A"
        Case "È"
            Buffer2 = Buffer2 & "E"
        Case "Ì"
            Buffer2 = Buffer2 & "I"
        Case "Ò"
            Buffer2 = Buffer2 & "O"
        Case "Ù"
            Buffer2 = Buffer2 & "U"
        Case "Ä"
            Buffer2 = Buffer2 & "A"
        Case "Ë"
            Buffer2 = Buffer2 & "E"
        Case "Ï"
            Buffer2 = Buffer2 & "I"
        Case "Ö"
            Buffer2 = Buffer2 & "O"
        Case "Ü"
            Buffer2 = Buffer2 & "U"
        Case "¥"
            Buffer2 = Buffer2 & "¥"
            'If Scaletype <> CLng(8) Then
            '    Buffer2 = Buffer2 & "@"
            'Else
            '    Buffer2 = Buffer2 & Chr(148)
            'End If
        'Case "Ñ"
        '    If Scaletype <> CLng(8) Then
        '        Buffer2 = Buffer2 & "@"
        '    Else
        '        Buffer2 = Buffer2 & Chr(148)
        '    End If
        Case "A" To "Z"
            Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
        Case "0" To "9"
            Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
        Case "."
            Buffer2 = Buffer2 & "."
        Case ","
            Buffer2 = Buffer2 & ","
        Case "<"
            Buffer2 = Buffer2 & "<"
        Case ">"
            Buffer2 = Buffer2 & ">"
        Case "_"
            Buffer2 = Buffer2 & "_"
        Case "-"
            Buffer2 = Buffer2 & "-"
        Case "/"
            Buffer2 = Buffer2 & "/"
        Case Else
            Buffer2 = Buffer2 & " "
        End Select
    Next bucle
    V12Char = Buffer2
End Function
    
Public Function StrBinToInt2(ByVal cadena As String) As Double
    '////////////////////////////////////
   '// convierte datos de varios      //
   '// bytes del formato binario      //
   '// entendible por Euroscale a     //
   '// formato en cristiano           //
   '// (INTERNA)                      //
   '////////////////////////////////////
    Dim Buffer As String
    Dim bucle As Long
    Dim Contador As Integer
    Dim Valor As Double
    Valor = 0
    Contador = 0
    If (Asc(Right(cadena, 1)) And 128) = 128 Then
        Valor = (-1) * StrBinToIntConSigno(cadena, True)
    Else
        For bucle = 1 To Len(cadena)
            Valor = Valor + Asc(Mid(cadena, bucle, 1)) * (256 ^ Contador)
            Contador = Contador + 1
        Next bucle
    End If
    StrBinToInt2 = Valor
End Function
    
Public Function StrBinToIntConSigno(ByVal cadena As String, EsNegativo As Boolean) As Double
    '////////////////////////////////////
   '// convierte datos de varios      //
   '// bytes del formato binario      //
   '// entendible por Euroscale a     //
   '// formato en cristiano           //
   '// (INTERNA)                      //
   '////////////////////////////////////
    Dim Buffer As String
    Dim bucle As Long
    Dim Contador As Integer
    Dim Valor As Double
    Valor = 0
    Contador = 0
    Buffer = cadena
    If EsNegativo Then
        If (128 And Asc(Right(cadena, 1))) <> 0 Then
            Buffer = ""
            For bucle = 1 To Len(cadena)
                 Buffer = Buffer & Chr(255 - Asc(Mid(cadena, bucle, 1)))
                Contador = Contador + 1
            Next bucle
        End If
    End If
    Contador = 0
    For bucle = 1 To Len(cadena)
        Valor = Valor + Asc(Mid(Buffer, bucle, 1)) * (256 ^ Contador)
        Contador = Contador + 1
    Next bucle
    If EsNegativo Then
        StrBinToIntConSigno = Valor + 1
    Else
        StrBinToIntConSigno = Valor
    End If
End Function
    
Public Sub generaGIAN()
    Dim rst As dao.Recordset
    Dim baseGi As dao.Database
    Dim sParte As String
    Dim nFich As Integer
    Dim sTipo As String
    
    
    If GrupoActual < 2 Then
        sParte = sPathGigante & "fart-flp.dat"
    Else
        sParte = sPathGigante & Format(GrupoActual, "00") & "fart-flp.dat"
    End If

    nFich = FreeFile()
    Open sParte For Output As #nFich
    Set baseGi = OpenDatabase(Base_General)
    Set rst = baseGi.OpenRecordset("select * from articulo order by codigo")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            sTipo = ""
            If rst.Fields("borrado") Then 'baja
            
                sTipo = "B"
                
            Else
                If rst.Fields("tran_plu") <> "*" Then 'Modificación
                    sTipo = "M"
                End If
                If rst.Fields("prc3") = 0 And rst.Fields("tran_plu") <> "*" Then 'Alta
                    sTipo = "A"
                End If
                
            
            End If
            If sTipo <> "" Then 'escribir registro
                sParte = Format(rst.Fields("codi_ident"), "00") & sTipo
                sParte = sParte & Format(rst.Fields("plu"), "0000")
                If rst.Fields("codi_pes") = "W" Then
                    sParte = sParte & "P"
                Else
                    sParte = sParte & "N"
                End If
                sParte = sParte & Format((rst.Fields("euros")) * 100, "00000")
                sParte = sParte & Format(rst.Fields("codigo"), "000000")
                sParte = sParte & Mid(rst.Fields("des_plu1"), 1, 18)
                sParte = sParte & Format(rst.Fields("codi_fam"), "00")
                sParte = sParte & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("codi_fam"), "00") & _
                                  Format(rst.Fields("plu"), "000") & "000"
                Print #nFich, sParte
            
            End If
        
            rst.Movenext
        Loop
    
    
    
    End If
    
    Close #nFich
    rst.Close
    Set rst = Nothing
    baseGi.Close
    Set baseGi = Nothing

End Sub

'de momento FART.SEC
'2.0.36 Gigante
Public Sub GeneraGianT(Optional d As String)  'llamar tras borrar totales (Fin de dia)
'                                                            Un./Importe
'"210001@2100@00@000001@DETERGENTE MAQUINA@P@ 1@2100001000@N@ 0@ 0....desde aquí fijo     "@    @ 0@ 0@0     @ 0@ 0@    @ 0@ 0@ 0@ 0@0     @ 0@ 0@ 0@ 0@    @ 0@ 0@0     @ 0@ 0@ 0@ 0@ 0@ 0@
Dim baseGi As dao.Database
Dim baseGiSeg As dao.Database

Dim rst As dao.Recordset
Dim rst1 As dao.Recordset
Dim baseFox As dao.Database

Dim nFich As Integer
Dim sParte As String
Dim sPath As String

Dim nCodiS As Integer
Dim nNumB As Integer
Dim nN As String
Dim nC As Integer
Dim nCC As Integer
Dim nNumeroSec As Integer
Dim F As String
Dim sBaseSeg As String
Dim lHaySeg As Boolean

If d <> "" Then
    F = Format(d, "dd/mm/yy")
Else
    F = Format(Now, "dd/mm/yy")
End If
If sUnidadGigante = "" Then
    sPath = "c:\"
Else
    sPath = sUnidadGigante
End If
sUnidadGigante = sPath
nFich = FreeFile()
On Error GoTo finerr
Open sUnidadGigante & "txt.txt" For Output As #nFich
Close #nFich
Kill sUnidadGigante & "txt.txt"

If TiendaActual < 100 Then
sBaseSeg = App.Path & "\t" & Format(TiendaActual, "00") & "\s\" & Format(Year(F), "0000") & Format(Month(F), "00") & ".mdb"
Else
sBaseSeg = App.Path & "\t" & Format(TiendaActual, "000") & "\s\" & Format(Year(F), "0000") & Format(Month(F), "00") & ".mdb"
End If

sPath = sPath & Format(Year(F), "0000") & "\" & Format(Month(F), "00") & "\" & Format(Day(F), "00") & "\T" & CStr(TiendaActual)
If Dir(sPath, vbDirectory) = "" Then
    On Error Resume Next
    MkDir sUnidadGigante & Format(Year(F), "0000")
    MkDir sUnidadGigante & Format(Year(F), "0000") & "\" & Format(Month(F), "00")
    MkDir sUnidadGigante & Format(Year(F), "0000") & "\" & Format(Month(F), "00") & "\" & Format(Day(F), "00")
    MkDir sUnidadGigante & Format(Year(F), "0000") & "\" & Format(Month(F), "00") & "\" & Format(Day(F), "00") & "\T" & CStr(TiendaActual)
    On Error GoTo 0
End If
nNumeroSec = 0

If d <> "" Then
    F = Format(d, "mm/dd/yy")
Else
    F = Format(Now, "mm/dd/yy")
End If

Set baseGi = OpenDatabase(Base_General)
If Dir(sBaseSeg) <> "" Then
    lHaySeg = True
    Set baseGiSeg = OpenDatabase(sBaseSeg)
End If

Set rst = baseGi.OpenRecordset("select * from seccion")
If Not rst.EOF Then
    nFich = FreeFile()
    Open sPath & "\fsec.sec" For Output As #nFich
    rst.MoveFirst
    Do While Not rst.EOF
        nN = rst.Fields("descripcio")
        If Len(Trim(nN)) < 25 Then
            nN = Space(25 - Len(Trim(nN))) & Trim(nN)
        End If
        sParte = Chr(34) & Format(rst.Fields("codi_ident"), "00") & "@" & nN & "@" & Chr(34)
        sGi(rst.Fields("secc_maqui")).secmaq = rst.Fields("secc_maqui")
        sGi(rst.Fields("secc_maqui")).secmos = rst.Fields("codi_ident")
        sGi(rst.Fields("secc_maqui")).secdes = nN
        nNumeroSec = nNumeroSec + 1
        Print #nFich, sParte

        rst.Movenext
    Loop
    Close #nFich
    
End If

If nNumeroSec > 0 Then 'solo si existen secciones...


Set rst = baseGi.OpenRecordset("select * from gtarti where fecha=#" & F & "# order by codigo")
If Not rst.EOF Then
    nFich = FreeFile()
    Open sPath & "\fart.sec" For Output As #nFich
    rst.MoveFirst
    Do While Not rst.EOF
        sParte = Chr(34) & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("plu"), "0000") & "@"
        sParte = sParte & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("codi_fam"), "00") & "@" & Format(rst.Fields("codi_fam"), "00") & "@"
        'duda si Código o plu
        sParte = sParte & Format(rst.Fields("plu"), "000000") & "@" & Mid(rst.Fields("des_plu1"), 1, 18) & "@"
        If rst.Fields("peso") = 0 Then
            sParte = sParte & "N" & "@"
        Else
            sParte = sParte & "P" & "@"
        End If
        sParte = sParte & " " & CStr(Int(rst.Fields("precio") * 100)) & "@"
        sParte = sParte & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("codi_fam"), "00") & Format(rst.Fields("plu"), "000") & "000" & "@" & "N" & "@"
        If rst.Fields("peso") = 0 Then
            sParte = sParte & " " & CStr(rst.Fields("unidades")) & "@"
        Else
            sParte = sParte & " " & CStr(Int(rst.Fields("peso") * 1000)) & "@"
        End If
        sParte = sParte & " " & CStr(Int(rst.Fields("pesetas") * 100)) & "@"
        sParte = sParte & "    @ 0@ 0@0     @ 0@ 0@    @ 0@ 0@ 0@ 0@0     @ 0@ 0@ 0@ 0@    @ 0@ 0@0     @ 0@ 0@ 0@ 0@ 0@ 0@" & Chr(34)
        Print #nFich, sParte
        
        
        
        
        rst.Movenext
    Loop
    Close #nFich
    
Else
    If lHaySeg Then
        Set rst = baseGiSeg.OpenRecordset("select * from gtarti where fecha=#" & F & "# order by codigo")
        If Not rst.EOF Then
            nFich = FreeFile()
            Open sPath & "\fart.sec" For Output As #nFich
            rst.MoveFirst
            Do While Not rst.EOF
                sParte = Chr(34) & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("plu"), "0000") & "@"
                sParte = sParte & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("codi_fam"), "00") & "@" & Format(rst.Fields("codi_fam"), "00") & "@"
                'duda si Código o plu
                sParte = sParte & Format(rst.Fields("plu"), "000000") & "@" & Mid(rst.Fields("des_plu1"), 1, 18) & "@"
                If rst.Fields("peso") = 0 Then
                    sParte = sParte & "N" & "@"
                Else
                    sParte = sParte & "P" & "@"
                End If
                sParte = sParte & " " & CStr(Int(rst.Fields("precio") * 100)) & "@"
                sParte = sParte & Format(rst.Fields("codi_ident"), "00") & Format(rst.Fields("codi_fam"), "00") & Format(rst.Fields("plu"), "000") & "000" & "@" & "N" & "@"
                If rst.Fields("peso") = 0 Then
                    sParte = sParte & " " & CStr(rst.Fields("unidades")) & "@"
                Else
                    sParte = sParte & " " & CStr(Int(rst.Fields("peso") * 1000)) & "@"
                End If
                sParte = sParte & " " & CStr(Int(rst.Fields("pesetas") * 100)) & "@"
                sParte = sParte & "    @ 0@ 0@0     @ 0@ 0@    @ 0@ 0@ 0@ 0@0     @ 0@ 0@ 0@ 0@    @ 0@ 0@0     @ 0@ 0@ 0@ 0@ 0@ 0@" & Chr(34)
                Print #nFich, sParte
                
                
                
                
                rst.Movenext
            Loop
            Close #nFich
            
        End If
        
    End If

End If

Set rst = baseGi.OpenRecordset("select * from tickets where d_fecha=#" & F & "# order by val(vendor),val(nume),val(numlin)")
If Not rst.EOF Then
    'CadenadeLog "Generación FTID.SEC desde dbasetouch.mdb en Uso..."
    nFich = FreeFile()
    Open sPath & "\ftid.sec" For Output As #nFich
    rst.MoveFirst
    Do While Not rst.EOF
        'Set rst1 = baseGi.OpenRecordset("select * from seccion where secc_maqui=" & CStr(Val(rst.Fields("secc"))))
        'If Not rst1.EOF Then
        '    nCodiS = rst1.Fields("codi_ident")
        'Else
        '    nCodiS = Val(rst.Fields("secc"))
        'End If
        'rst1.Close
        nCodiS = sGi(Val(rst.Fields("secc"))).secmos
        'CadenadeLog "Codigo Sección:" & CStr(nCodiS)
        'CadenadeLog "select * from equipos where numero_eqp=" & CStr(rst.Fields("numbal"))
        'CadenadeLog "Número de Tique:" & CStr(rst.Fields("nume"))
        'Set rst1 = Nothing
        Set rst1 = baseGi.OpenRecordset("select * from equipos where numero_eqp=" & CStr(rst.Fields("numbal")))
        If Not rst1.EOF Then
         If Len(Trim(rst1.Fields("descripcio"))) > 0 Then
            nNumB = Val(left(rst1.Fields("descripcio"), 1))
         Else
            nNumB = rst.Fields("numbal")
         End If
        Else
            nNumB = rst.Fields("numbal")
        End If
        'CadenadeLog "Balanza:" & CStr(nNumB) & " antes cierre recordset equipos"
        rst1.Close
        Set rst1 = Nothing
        'CadenadeLog "Balanza:" & CStr(nNumB) & " tras cierre recordset equipos"
        sParte = Chr(34) & Format(nCodiS, "00") & Format(nNumB, "0") & Format(Val(rst.Fields("nume")), "000")
        
        '04/07/06 se re-establece contador de tique por vendedor iniciado tras cada borrado de totales.
        '30/06/06 ... sParte = sParte & Mid(Trim(rst.Fields("numlote")), 1, 4) & Format(rst.Fields("numlin"), "00") & "@"
        'If rst.Fields("nume") < 500 Then
        '    sParte = sParte & Format(rst.Fields("nume"), "0000") & Format(rst.Fields("numlin"), "00") & "@"
        'Else
        '    sParte = sParte & Format((rst.Fields("nume")) - 500, "0000") & Format(rst.Fields("numlin"), "00") & "@"
        'End If
        sParte = sParte & Mid(Trim(rst.Fields("numlote")), 1, 4) & Format(rst.Fields("numlin"), "00") & "@"
        sParte = sParte & CStr(Val(rst.Fields("secc"))) & "000" & Mid(Trim(rst.Fields("vendor")), 2, 2) & "@"
        sParte = sParte & Format(Val(rst.Fields("plu")), "0000") & "@"
        sParte = sParte & " "
        If rst.Fields("weight") <> 0 Then
            nN = Format(((rst.Fields("weight") * 1000) / 1000), "#.###")
        Else
            nN = Format(rst.Fields("units"), "#")
        End If
        If Mid(nN, 1, 1) = "," Then
            nN = "." & Mid(nN, 2)
        Else
            nC = Len(nN)
            For nCC = 1 To nC
                If Mid(nN, nCC, 1) = "," Then
                    nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                End If
            Next nCC
        End If
        sParte = sParte & nN & "@ "
        nN = Format(rst.Fields("amount"), "#.###")
        If Mid(nN, 1, 1) = "," Then
            nN = "." & Mid(nN, 2)
        Else
            nC = Len(nN)
            For nCC = 1 To nC
                If Mid(nN, nCC, 1) = "," Then
                    nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                End If
            Next nCC
        End If
        sParte = sParte & nN & "@ @" & Format(nNumB, "00") & "@"
        sParte = sParte & Format(rst.Fields("d_hora"), "hhmm") & "@" & Chr(34)
        
        Print #nFich, sParte
        
        rst.Movenext
    Loop
    Close #nFich
Else

    If lHaySeg Then
            'CadenadeLog "Generación FTID.SEC desde directorio Seguridad..."
            'Set rst = baseGiSeg.OpenRecordset("select * from tickets where d_fecha=#" & F & "# order by val(nume)")
            Set rst = baseGiSeg.OpenRecordset("select * from tickets where d_fecha=#" & F & "# order by val(vendor),val(nume),val(numlin)")
            If Not rst.EOF Then
                nFich = FreeFile()
                Open sPath & "\ftid.sec" For Output As #nFich
                rst.MoveFirst
                Do While Not rst.EOF
                    'Set rst1 = baseGi.OpenRecordset("select * from seccion where secc_maqui=" & CStr(Val(rst.Fields("secc"))))
                    'If Not rst1.EOF Then
                    '    nCodiS = rst1.Fields("codi_ident")
                    'Else
                    '    nCodiS = Val(rst.Fields("secc"))
                    'End If
                    'rst1.Close
                    nCodiS = sGi(Val(rst.Fields("secc"))).secmos
                    'Set rst1 = Nothing
                    Set rst1 = baseGi.OpenRecordset("select * from equipos where numero_eqp=" & CStr(rst.Fields("numbal")))
                    If Not rst1.EOF Then
                     If Len(Trim(rst1.Fields("descripcio"))) > 0 Then
                        nNumB = Val(left(rst1.Fields("descripcio"), 1))
                     Else
                        nNumB = rst.Fields("numbal")
                     End If
                    Else
                        nNumB = rst.Fields("numbal")
                    End If
                    rst1.Close
                    Set rst1 = Nothing
                    sParte = Chr(34) & Format(nCodiS, "00") & Format(nNumB, "0") & Format(Val(rst.Fields("nume")), "000")
                    'CadenadeLog sParte
                    '04/07/06 se re-establece contador de tique por vendedor iniciado tras cada borrado de totales.
                    '30/06/06 ... sParte = sParte & Mid(Trim(rst.Fields("numlote")), 1, 4) & Format(rst.Fields("numlin"), "00") & "@"
                    'If rst.Fields("nume") < 500 Then
                    '    sParte = sParte & Format(rst.Fields("nume"), "0000") & Format(rst.Fields("numlin"), "00") & "@"
                    'Else
                    '    sParte = sParte & Format((rst.Fields("nume")) - 500, "0000") & Format(rst.Fields("numlin"), "00") & "@"
                    'End If
                    sParte = sParte & Mid(Trim(rst.Fields("numlote")), 1, 4) & Format(rst.Fields("numlin"), "00") & "@"
                    sParte = sParte & CStr(Val(rst.Fields("secc"))) & "000" & Mid(Trim(rst.Fields("vendor")), 2, 2) & "@"
                    sParte = sParte & Format(Val(rst.Fields("plu")), "0000") & "@"
                    sParte = sParte & " "
                    If rst.Fields("weight") <> 0 Then
                        nN = Format(((rst.Fields("weight") * 1000) / 1000), "#.###")
                    Else
                        nN = Format(rst.Fields("units"), "#")
                    End If
                    If Mid(nN, 1, 1) = "," Then
                        nN = "." & Mid(nN, 2)
                    Else
                        nC = Len(nN)
                        For nCC = 1 To nC
                            If Mid(nN, nCC, 1) = "," Then
                                nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                            End If
                        Next nCC
                    End If
                    sParte = sParte & nN & "@ "
                    nN = Format(rst.Fields("amount"), "#.###")
                    If Mid(nN, 1, 1) = "," Then
                        nN = "." & Mid(nN, 2)
                    Else
                        nC = Len(nN)
                        For nCC = 1 To nC
                            If Mid(nN, nCC, 1) = "," Then
                                nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                            End If
                        Next nCC
                    End If
                    sParte = sParte & nN & "@ @" & Format(nNumB, "00") & "@"
                    sParte = sParte & Format(rst.Fields("d_hora"), "hhmm") & "@" & Chr(34)
                    
                    Print #nFich, sParte
                    
                    rst.Movenext
                Loop
                Close #nFich
            
            End If

    
    End If
    
End If

Set rst = baseGi.OpenRecordset("select * from gtvend where fecha=#" & F & "# order by ident_vend")
If Not rst.EOF Then
    nFich = FreeFile()
    Open sPath & "\fven.sec" For Output As #nFich
    rst.MoveFirst
    Do While Not rst.EOF
       nN = Format(rst.Fields("secc_maqui"), "0") & "000" & Mid(Format(rst.Fields("ident_vend"), "000"), 2, 2)
       sParte = Chr(34) & nN & "@"
       sParte = sParte & Format(rst.Fields("codi_ident"), "00") & nN & "@"
       sParte = sParte & sGi(rst.Fields("secc_maqui")).secdes & "@ "
        nN = Format(rst.Fields("pesetas"), "#.###")
        If Mid(nN, 1, 1) = "," Then
            nN = "." & Mid(nN, 2)
        Else
            nC = Len(nN)
            For nCC = 1 To nC
                If Mid(nN, nCC, 1) = "," Then
                    nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                End If
            Next nCC
        End If
       sParte = sParte & nN & "@ " & CStr((rst.Fields("tiquets")))
        'nN = Format(rst.Fields("tiquets"), "#.###")
        'If Mid(nN, 1, 1) = "," Then
        '    nN = "." & Mid(nN, 2)
        'Else
        '    nC = Len(nN)
        '    For nCC = 1 To nC
        '        If Mid(nN, nCC, 1) = "," Then
        '            nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
        '        End If
        '    Next nCC
        'End If
       sParte = sParte & "@ 0@ 0@" & Chr(34)
       
       Print #nFich, sParte
       


        rst.Movenext
    Loop
    Close #nFich

Else

    If lHaySeg Then
    
        Set rst = baseGiSeg.OpenRecordset("select * from gtvend where fecha=#" & F & "# order by ident_vend")
        If Not rst.EOF Then
            nFich = FreeFile()
            Open sPath & "\fven.sec" For Output As #nFich
            rst.MoveFirst
            Do While Not rst.EOF
               nN = Format(rst.Fields("secc_maqui"), "0") & "000" & Mid(Format(rst.Fields("ident_vend"), "000"), 2, 2)
               sParte = Chr(34) & nN & "@"
               sParte = sParte & Format(rst.Fields("codi_ident"), "00") & nN & "@"
               sParte = sParte & sGi(rst.Fields("secc_maqui")).secdes & "@ "
                nN = Format(rst.Fields("pesetas"), "#.###")
                If Mid(nN, 1, 1) = "," Then
                    nN = "." & Mid(nN, 2)
                Else
                    nC = Len(nN)
                    For nCC = 1 To nC
                        If Mid(nN, nCC, 1) = "," Then
                            nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                        End If
                    Next nCC
                End If
               sParte = sParte & nN & "@ " & CStr((rst.Fields("tiquets")))
                'nN = Format(rst.Fields("tiquets"), "#.###")
                'If Mid(nN, 1, 1) = "," Then
                '    nN = "." & Mid(nN, 2)
                'Else
                '    nC = Len(nN)
                '    For nCC = 1 To nC
                '        If Mid(nN, nCC, 1) = "," Then
                '            nN = Mid(nN, 1, nCC - 1) & "." & Mid(nN, nCC + 1)
                '        End If
                '    Next nCC
                'End If
               sParte = sParte & "@ 0@ 0@" & Chr(34)
               
               Print #nFich, sParte
        
               rst.Movenext
            Loop
            Close #nFich
        
    
        End If
        
    End If
    
End If

    On Error GoTo fin
    Set baseFox = OpenDatabase(App.Path, False, False, "DBASE IV")
    Set rst1 = baseFox.OpenRecordset("select * from fichero")
    'If Not rst1.EOF Then
    '    rst1.MoveFirst
    '    rst1.Delete
    '    'rst1.Update
    'End If
    'rst1.AddNew
    rst1.Edit
    rst1.Fields("VAR_FEC") = Format(F, "mm/dd/yy")
    rst1.Fields("HOR_COP") = Format(Time, "hh:mm:ss")
    rst1.Fields("FEC_COP") = Format(Now, "dd/mm/yy")
    rst1.Fields("VAR_LOC") = Mid(Trim(Empresa), 1, 2)
    rst1.Fields("VAR_BAS") = Mid(Trim(Empresa), 1, 2) & Format(F, "ddmm")
    rst1.Update
    rst1.Close
    Set rst1 = Nothing
    baseFox.Close
    Set baseFox = Nothing
    FileCopy App.Path & "\fichero.dbf", sPath & "\fichero.dbf"
    Set baseFox = OpenDatabase(App.Path & "\fich.mdb")
    Set rst1 = baseFox.OpenRecordset("select * from Resumen")
    rst1.AddNew
    rst1.Fields("Tienda") = Format(TiendaActual, "000")
    rst1.Fields("VAR_FEC") = Format(F, "mm/dd/yy")
    rst1.Fields("VAR_FECHA") = Format(F, "dd/mm/yy")
    rst1.Fields("HOR_COP") = Format(Time, "hh:mm:ss")
    rst1.Fields("VAR_LOC") = Mid(Trim(Empresa), 1, 2)
    rst1.Update
    rst1.Close
    Set rst1 = Nothing
    baseFox.Close
    Set baseFox = Nothing
    
End If

fin:
On Error GoTo 0
rst.Close
Set rst = Nothing
baseGi.Close
Set baseGi = Nothing
If Dir(sBaseSeg) <> "" Then
   baseGiSeg.Close
   Set baseGiSeg = Nothing
End If
Exit Sub

finerr:
nFich = MsgBox("No se accede a Unidad configurada!!!", vbCritical)


End Sub
'1.7.17
Public Function AdapTaCharGR(ByVal sR As String) As String
    Dim Buffer As String
    Dim Buffer2 As String
    Dim buffer3 As String
    Dim bucle As Integer
    Dim nFich As Integer
    
    '1.7.17
    If Dir(App.Path & "\greek.cht") <> "" Then
        nFich = FreeFile()
        Open App.Path & "\greek.cht" For Input As #nFich
        For bucle = 0 To 9
            Line Input #nFich, buffer3
            chrOrgG(bucle) = Val(Mid(buffer3, 14, 5))
            chrDesG(bucle) = Val(Mid(buffer3, 28, 5))
            'CadenadeLog "origen:" & "(" & Format(Val(Mid(buffer3, 14, 5)), "000") & ")->" & Chr(chrOrgG(Bucle)) & " Detino:" & "(" & Format(Val(Mid(buffer3, 28, 5))) & ")-->" & Chr(chrDesG(Bucle))
        Next bucle
        Close #nFich
    End If
    ''''''''''
    Buffer2 = ""
    Buffer = sR
    For bucle = 1 To Len(Buffer)
        Select Case Mid(Buffer, bucle, 1)
        Case Chr(chrDesG(0))
            Buffer2 = Buffer2 & Chr(chrOrgG(0))
        Case Chr(chrDesG(1))
            Buffer2 = Buffer2 & Chr(chrOrgG(1))
        Case Chr(chrDesG(2))
            Buffer2 = Buffer2 & Chr(chrOrgG(2))
        Case Chr(chrDesG(3))
            Buffer2 = Buffer2 & Chr(chrOrgG(3))
        Case Chr(chrDesG(4))
            Buffer2 = Buffer2 & Chr(chrOrgG(4))
        Case Chr(chrDesG(5))
            Buffer2 = Buffer2 & Chr(chrOrgG(5))
        Case Chr(chrDesG(6))
            Buffer2 = Buffer2 & Chr(chrOrgG(6))
        Case Chr(chrDesG(7))
            Buffer2 = Buffer2 & Chr(chrOrgG(7))
        Case Chr(chrDesG(8))
            Buffer2 = Buffer2 & Chr(chrOrgG(8))
        Case Chr(chrDesG(9))
            Buffer2 = Buffer2 & Chr(chrOrgG(9))
        Case "Á"
            Buffer2 = Buffer2 & "A"
        Case "É"
            Buffer2 = Buffer2 & "E"
        Case "Í"
            Buffer2 = Buffer2 & "I"
        Case "Ó"
            Buffer2 = Buffer2 & "O"
        Case "Ú"
            Buffer2 = Buffer2 & "U"
        Case "À"
            Buffer2 = Buffer2 & "A"
        Case "È"
            Buffer2 = Buffer2 & "E"
        Case "Ì"
            Buffer2 = Buffer2 & "I"
        Case "Ò"
            Buffer2 = Buffer2 & "O"
        Case "Ù"
            Buffer2 = Buffer2 & "U"
        Case "Ä"
            Buffer2 = Buffer2 & "A"
        Case "Ë"
            Buffer2 = Buffer2 & "E"
        Case "Ï"
            Buffer2 = Buffer2 & "I"
        Case "Ö"
            Buffer2 = Buffer2 & "O"
        Case "Ü"
            Buffer2 = Buffer2 & "U"
        Case "¥"
            'If Scaletype <> CLng(8) Then
            '    Buffer2 = Buffer2 & "@"
            'Else
                Buffer2 = Buffer2 & Chr(148)
            'End If
        Case "Ñ"
            'If Scaletype <> CLng(8) Then
            '    Buffer2 = Buffer2 & "@"
            'Else
                Buffer2 = Buffer2 & Chr(148)
            'End If
        Case "A" To "Z"
            Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
        Case "0" To "9"
            Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
        Case "."
            Buffer2 = Buffer2 & "."
        Case ","
            Buffer2 = Buffer2 & ","
        Case "<"
            Buffer2 = Buffer2 & "<"
        Case ">"
            Buffer2 = Buffer2 & ">"
        Case "_"
            Buffer2 = Buffer2 & "_"
        Case "-"
            Buffer2 = Buffer2 & "-"
        Case "/"
            Buffer2 = Buffer2 & "/"
        Case Else
            '1.7.16
            'Buffer2 = Buffer2 & " "
            '''''''
            Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
            
        End Select
    Next bucle
    
    'CadenadeLog Buffer2
    
    AdapTaCharGR = Buffer2
End Function

Public Function AdaptaCharPolishBPC(ByVal s As String)

    Dim Buffer As String
    Dim Buffer2 As String
    Dim buffer3 As String
    Dim bucle As Integer
    Dim nFich As Integer
    


    If Dir(App.Path & "\polish.cht") <> "" Or Dir(App.Path & "\tk.cht") <> "" Then
        If Dir(App.Path & "\polish.cht") <> "" Then
            nFich = FreeFile()
            Open App.Path & "\polish.cht" For Input As #nFich
            For bucle = 0 To 12
                Line Input #nFich, buffer3
                chrOrgG(bucle) = Val(Mid(buffer3, 14, 5))
                chrDesG(bucle) = Val(Mid(buffer3, 28, 5))
                'CadenadeLog "origen:" & "(" & Format(Val(Mid(buffer3, 14, 5)), "000") & ")->" & Chr(chrOrgG(Bucle)) & " Detino:" & "(" & Format(Val(Mid(buffer3, 28, 5))) & ")-->" & Chr(chrDesG(Bucle))
            Next bucle
            Close #nFich
        End If
        If Dir(App.Path & "\tk.cht") <> "" Then
            nFich = FreeFile()
            Open App.Path & "\polish.cht" For Input As #nFich
            For bucle = 0 To 5
                Line Input #nFich, buffer3
                chrOrgG(bucle) = Val(Mid(buffer3, 14, 5))
                chrDesG(bucle) = Val(Mid(buffer3, 28, 5))
                'CadenadeLog "origen:" & "(" & Format(Val(Mid(buffer3, 14, 5)), "000") & ")->" & Chr(chrOrgG(Bucle)) & " Detino:" & "(" & Format(Val(Mid(buffer3, 28, 5))) & ")-->" & Chr(chrDesG(Bucle))
            Next bucle
            Close #nFich
        End If
    
        Buffer2 = ""
        Buffer = s
    
        For bucle = 1 To Len(Buffer)
            Select Case Mid(Buffer, bucle, 1)
            Case Chr(chrDesG(0))
                Buffer2 = Buffer2 & Chr(chrOrgG(0))
            Case Chr(chrDesG(1))
                Buffer2 = Buffer2 & Chr(chrOrgG(1))
            Case Chr(chrDesG(2))
                Buffer2 = Buffer2 & Chr(chrOrgG(2))
            Case Chr(chrDesG(3))
                Buffer2 = Buffer2 & Chr(chrOrgG(3))
            Case Chr(chrDesG(4))
                Buffer2 = Buffer2 & Chr(chrOrgG(4))
            Case Chr(chrDesG(5))
                Buffer2 = Buffer2 & Chr(chrOrgG(5))
                
            Case Else
                Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
            
        End Select
                
        Next bucle
    
        s = Buffer2
        
    End If
    AdaptaCharPolishBPC = s
End Function

Public Function PRsinIVA3(precio As Double, Percent As Double) As Double
    Dim Psin As Double
    Psin = (precio * 100) / (100 + Percent)
    If UsaEuro Then
        Psin = Round(Psin, 2 + 1)
    Else
        Psin = Round(Psin, decimales + 1)
    End If
    PRsinIVA3 = Psin
End Function

Public Sub AsignTec()
Dim rsSec As dao.Recordset
Dim rsFam As dao.Recordset
Dim rsArt As dao.Recordset
Dim rsTec As dao.Recordset
Dim atBase As dao.Database
Dim MisTeclas As DB_Teclas
Dim ncont As Integer
Dim MiRespuesta As Integer
Dim cm As String

    Load frmWAIT
    frmWAIT.Label1.Caption = "Asignación Teclas."
    frmWAIT.Label2.Caption = "Espere unos instantes..."
    frmWAIT.Visible = True
    Do_Events
    Do_Events
    
    Set atBase = AbrirBase
    cm = "update articulo set tran_plu=' ' where codigo in(select codigo from teclas)"
    atBase.Execute cm
    cm = "delete from teclas"
    atBase.Execute cm
    
    lAsigntec = True
    
    Set rsSec = atBase.OpenRecordset("select secc_maqui from seccion where borrado=false")
    If Not rsSec.EOF Then
        rsSec.MoveFirst
        Do While Not rsSec.EOF
            Set rsFam = atBase.OpenRecordset("select codi_fam,secc_maqui from familias where borrado=false and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " order by codi_fam")
            If Not rsFam.EOF Then
                rsFam.MoveFirst
                Do While Not rsFam.EOF
                
                    Set rsTec = atBase.OpenRecordset("select max(codi_tec) from teclas where borrado=false and tabla=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")))
                    If rsTec.EOF Then 'Or IsNull(rsTec.Fields(0)) Then
                        ncont = 1
                    Else
                        If IsNull(rsTec.Fields(0)) Then
                            ncont = 1
                        Else
                            ncont = rsTec.Fields(0) + 1
                        End If
                    End If
                    rsTec.Close
                    Set rsTec = Nothing
                    
                    If Dir(App.Path & "\asign1.tst") <> "" Or (Dir(App.Path & "\asign2.tst") = "" And Dir(App.Path & "\asign3.tst") = "" And Dir(App.Path & "\asign4.tst") = "") Then
                    'ordenado por código dentro de la familia
                    'Set rsArt = atBase.OpenRecordset("select codigo,secc_maqui,plu,codi_fam from articulo where borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and codigo not in(select codigo from teclas where borrado=false) order by codigo")
                    Set rsArt = atBase.OpenRecordset("select tran_plu,codigo,secc_maqui,plu,codi_fam from articulo where borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and onkey=true order by codigo")
                    If Not rsArt.EOF Then
                        rsArt.MoveFirst
                        Do While Not rsArt.EOF
                    
                    
                            MisTeclas.CodArticulo = rsArt.Fields("codigo")
                            MisTeclas.PluArticulo = rsArt.Fields("plu")
                            MisTeclas.SecArticulo = 0
                            MisTeclas.Destino_CodiTec = ncont
                            MisTeclas.Destino_NBalanza = 0
                            MisTeclas.Destino_NMostrador = rsSec.Fields("secc_maqui")
                            MisTeclas.Destino_Tabla = rsFam.Fields("codi_fam")
                            MisTeclas.Tipo = "TOUCH"
                            MisTeclas.Section = rsSec.Fields("secc_maqui")
                            MiRespuesta = Alta_Teclas(MisTeclas)
                            rsArt.Edit
                            rsArt.Fields("tran_plu") = " "
                            rsArt.Update
                    
                            ncont = ncont + 1
                            rsArt.Movenext
                        Loop
                    
                    
                    End If
                    '''''''''''''''''''''''''''''''''''
                    End If
                    If Dir(App.Path & "\asign2.tst") <> "" Then
                    'ordenando por número de PLU y tecla correlativa
                    'Set rsArt = atBase.OpenRecordset("select codigo,secc_maqui,plu,codi_fam from articulo where plu<>0 and borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and codigo not in(select codigo from teclas where borrado=false) order by plu")
                    Set rsArt = atBase.OpenRecordset("select tran_plu,codigo,secc_maqui,plu,codi_fam from articulo where plu<>0 and borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and onkey=true order by plu")
                    If Not rsArt.EOF Then
                        rsArt.MoveFirst
                        Do While Not rsArt.EOF
                    
                    
                            MisTeclas.CodArticulo = rsArt.Fields("codigo")
                            MisTeclas.PluArticulo = rsArt.Fields("plu")
                            MisTeclas.SecArticulo = 0
                            MisTeclas.Destino_CodiTec = ncont
                            MisTeclas.Destino_NBalanza = 0
                            MisTeclas.Destino_NMostrador = rsSec.Fields("secc_maqui")
                            MisTeclas.Destino_Tabla = rsFam.Fields("codi_fam")
                            MisTeclas.Tipo = "TOUCH"
                            MisTeclas.Section = rsSec.Fields("secc_maqui")
                            MiRespuesta = Alta_Teclas(MisTeclas)
                            rsArt.Edit
                            rsArt.Fields("tran_plu") = " "
                            rsArt.Update
                    
                            ncont = ncont + 1
                            rsArt.Movenext
                        Loop
                    
                    
                    End If
                    '''''''''''''''''''''''''''''''''''
                    End If
                    If Dir(App.Path & "\asign3.tst") <> "" Then
                    'ordenando por número de PLU y tecla=PLU
                    'Set rsArt = atBase.OpenRecordset("select codigo,secc_maqui,plu,codi_fam from articulo where plu<>0 and borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and codigo not in(select codigo from teclas where borrado=false) order by plu")
                    Set rsArt = atBase.OpenRecordset("select tran_plu,codigo,secc_maqui,plu,codi_fam from articulo where plu<>0 and borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and onkey=true order by plu")
                    If Not rsArt.EOF Then
                        rsArt.MoveFirst
                        Do While Not rsArt.EOF
                    
                    
                            MisTeclas.CodArticulo = rsArt.Fields("codigo")
                            MisTeclas.PluArticulo = rsArt.Fields("plu")
                            MisTeclas.SecArticulo = 0
                            MisTeclas.Destino_CodiTec = rsArt.Fields("plu")
                            MisTeclas.Destino_NBalanza = 0
                            MisTeclas.Destino_NMostrador = rsSec.Fields("secc_maqui")
                            MisTeclas.Destino_Tabla = rsFam.Fields("codi_fam")
                            MisTeclas.Tipo = "TOUCH"
                            MisTeclas.Section = rsSec.Fields("secc_maqui")
                            MiRespuesta = Alta_Teclas(MisTeclas)
                            rsArt.Edit
                            rsArt.Fields("tran_plu") = " "
                            rsArt.Update
                    
                            ncont = ncont + 1
                            rsArt.Movenext
                        Loop
                    
                    
                    End If
                    '''''''''''''''''''''''''''''''''''
                    End If
                    If Dir(App.Path & "\asign4.tst") <> "" Then
                    'ordenando por descriptivo tecla correlativa
'CadenadeLog "-->asign1:" & "FAM:" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui"))
                    'Set rsArt = atBase.OpenRecordset("select codigo,secc_maqui,plu,codi_fam from articulo where borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " and codigo not in(select codigo from teclas where borrado=false) order by des_plu1")
                    Set rsArt = atBase.OpenRecordset("select tran_plu,codigo,secc_maqui,plu,codi_fam from articulo where borrado=false and codi_fam=" & CStr(rsFam.Fields("codi_fam")) & " and secc_maqui=" & CStr(rsSec.Fields("secc_maqui")) & " order by des_plu1")
'CadenadeLog "-->asign2"
                    If Not rsArt.EOF Then
                        rsArt.MoveFirst
                        Do While Not rsArt.EOF
                    
                    
                            MisTeclas.CodArticulo = rsArt.Fields("codigo")
                            MisTeclas.PluArticulo = rsArt.Fields("plu")
                            MisTeclas.SecArticulo = 0
                            MisTeclas.Destino_CodiTec = ncont
                            MisTeclas.Destino_NBalanza = 0
                            MisTeclas.Destino_NMostrador = rsSec.Fields("secc_maqui")
                            MisTeclas.Destino_Tabla = rsFam.Fields("codi_fam")
                            MisTeclas.Tipo = "TOUCH"
                            MisTeclas.Section = rsSec.Fields("secc_maqui")
                            MiRespuesta = Alta_Teclas(MisTeclas)
                            
                            rsArt.Edit
                            rsArt.Fields("tran_plu") = " "
                            rsArt.Update
                            
                            ncont = ncont + 1
                            rsArt.Movenext
                        Loop
                    End If
                    '''''''''''''''''''''''''''''''''''
'CadenadeLog "-->asign3"
                    End If
                    
                    rsArt.Close
                    Set rsArt = Nothing
                    rsFam.Movenext
                
                
                Loop
            End If
            rsFam.Close
            Set rsFam = Nothing
            rsSec.Movenext
        Loop
        rsSec.Close
        Set rsSec = Nothing
    End If
    
    lAsigntec = False
    
    atBase.Close
    Set atBase = Nothing

    Unload frmWAIT
    frmWAIT.Visible = False

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Registration Articles: (Order "5" Format 1)
'FIELD           LENGTH  Comment.
'-----           --------        ----------
'"5"             1 (Literal) Record Type "5"
'<R-M / U>       1 digit     "0" Register-Modification, "1" Unsubscribe
'Code            6 digit         Item Code (000001 to 999999)
'Counter         2 digit     Counter to which Article (01-99) belongs
'PLU             4 digit     Correlative PLU into Section (0001-9999)
'Subsection      3 digit         Unused by Scale. (001-999)
'Family          4 digits        Family to which Article belongs (0001-9999)
'Price           5 digits    Selling price (without decimal point)
'Sales mode      1 digit     Weighted item "W", Unit item "P" or "U"
'Expiration days 3 digits        Item expiration days (000-999)
'"0"             1 (Literal)     ( "0")
'Descriptive     25 CHARAC.  Main description
'Tare            5 digits    Tare in grams to be applied in Sale
'<CR + LF>       2 CHARAC.   End of register

'Codigo#EnVenta#Operacion#PLU#Nombre#Precio#Tipo#Tara#DiasCad#Seccion#Familia#TextoG
'Campo         Descripción
'0  Código     Código del artículo
'1  En venta   Visible o no en pantalla (autoservicio)
'2  Operación  Alta/Baja/Modificación
'3  Plu        Número de PLU
'4  Nombre     Descriptivo principal del producto
'5  precio     Precio del artículo (sin punto decimal)
'6  Tipo       Tipo de venta (Pesado / No Pesado)
'7  tara       Tara del artículo
'8  DíasCad    Días de caducidad del articulo
'9  Sección    Mostrador al que pertenece el producto
'10  familia   familia
'11  TextoG    Texto largo para ingredientes.
Public Sub transforma_Fornes()
Dim sVar() As String
Dim MyFich As Integer
Dim sPath As String
Dim srg As String
Dim nFichDes As Integer
Dim nL As Integer
'If Len(Buffer) = 60 Or Len(Buffer) = 86 Or Len(Buffer) = 91 Or Len(Buffer) = 92 Then

    If Dir(App.Path & "\pluasc.dir") <> "" Then
        MyFich = FreeFile()
        Open App.Path & "\pluasc.dir" For Input As #MyFich
        Line Input #MyFich, sPath
        Close #MyFich
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    Else
        sPath = App.Path & "\"
        MyFich = FreeFile()
        Open App.Path & "\pluasc.dir" For Output As #MyFich
        Print #MyFich, sPath
        Close #MyFich
    End If
    If Dir(sPath & "Bal_Epelsa.txt") <> "" Then
        FileCopy sPath & "Bal_Epelsa.txt", App.Path & "\fornes.TXT"
        Sleep (200)
        Kill sPath & "Bal_Epelsa.txt"
        MyFich = FreeFile()
        Open App.Path & "\fornes.TXT" For Input As #MyFich
        nFichDes = FreeFile()
        Open App.Path & "\global.dat" For Output As #nFichDes
        Do While Not EOF(MyFich)
            Erase sVar
            Line Input #MyFich, srg
            srg = Replace(srg, Chr(9), "")
            '''
            srg = Replace(srg, "#" & Chr(34), "#")
            If Right(srg, 1) = Chr(34) Then
                srg = Mid(srg, 1, Len(srg) - 1)
            End If
            '''
            sVar = Split(srg, "#")
            nL = UBound(sVar)
            'creación registro "5"
            If sVar(2) <> "B" Then
                If nL = 13 Then 'nuevo fichero
                    If sVar(13) = "0" Then
                        srg = "52" 'nuevo tipo, modificar pero hacer onkey=0
                    Else
                        srg = "50"
                    End If
                Else
                    If sVar(1) = "0" Then
                        srg = "52" 'nuevo tipo, modificar pero hacer onkey=0
                    Else
                        srg = "50"
                    End If
                End If
            Else
                srg = "51"
            End If
            srg = srg & Format(Val(sVar(0)), "000000") & Format(Val(sVar(9)), "00") & Format(Val(sVar(3)), "0000") & "000" & Format(Val(sVar(10)), "0000")
            srg = srg & Format(Val(sVar(5)), "00000")
            
            If sVar(6) = "1" Then
                srg = srg & "W"
            Else
                srg = srg & "U"
            End If
            srg = srg & Format(Val(sVar(8)), "000") & "0"
            If Len(sVar(4)) > 25 Then
                srg = srg & Mid(sVar(4), 1, 25)
            Else
                If sVar(4) = "" Then
                    srg = srg & Space(25)
                Else
                    srg = srg & sVar(4) & Space(25 - Len(sVar(4)))
                End If
            End If
            If nL = 13 Then
                srg = srg & Format(Val(sVar(7)), "00000") & Format(Val(sVar(12)), "00")
            Else
                srg = srg & Format(Val(sVar(7)), "00000")
            End If
            Print #nFichDes, srg
            'creación registro "T"
'Registration Item additional Long Text : (Order "T")
'FIELD           LENGTH  Comment.
'-----           --------        ----------
'"T"         1 (Literal) Record Type "T"
'Item Code       6 digits        Item (000001 to 999999)
'Additional Text Nr. 1 digits        (1 to 4)
'Text            X CHARAC.   Additional Text
'<CR + LF>       2 CHARAC.   End of register
            If nL = 11 Then
                If Not IsNull(sVar(11)) Then
                    srg = "T0" & Format(Val(sVar(0)), "000000") & "1" & Trim(sVar(11))
                    Print #nFichDes, srg
                End If
            End If
        Loop
        Close #MyFich
        Close #nFichDes
    
    
    End If
End Sub
'''''''''''''''''''''''''''''''''
Public Sub export_BL(ByVal lTodos As Boolean)
Dim Base As Database
Dim Registro As Recordset
Dim Reg2 As Recordset
'**************************
' Cabecera
Dim B1 As String
Dim B2 As String
Dim Codigo_mos As String
Dim Codigo_bal As String
Dim Codigo_ven As String
'Dim Cod_Bar As String
Dim Lineas As String
Dim precio As String
Dim Tipo As String
Dim Fecha As String
Dim hora As String
Dim LineaAnulada As String
'********
' Linea
Dim Num_lin As String
Dim Cod_Plu As String
Dim Cod_Cod As String
Dim Precio_Uni As String
Dim Pes_Uni As String
Dim sPes As String
Dim Total As String
Dim Signo As String
Dim Tipo_Linea As String
Dim Anulado As String
'************
Dim MiEnvio, Mienvio2, Mienvio3, sParte As String
Dim nintentos As Integer
'Dim Bucle As Integer

Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim Fichero2 As Integer
Dim sLinCan As String
Dim sCodCli As String
Dim nSysToca As Integer
Dim sFich As String
Dim nLinea As Integer
Dim PathExpor As String
Dim PathExporRX As String
Dim lFirst As Boolean
Dim nLinLa As Integer
Dim nC As Integer
Dim nMess As Integer
'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
    If lFornes Then
        If Dir("c:\telecom\epelsa", vbDirectory) = "" Then
            MkDir ("c:\telecom\epelsa")
        End If
        'If Dir(App.Path & "\exportarx", vbDirectory) = "" Then
        '    MkDir (App.Path & "\exportarx")
        'End If
    End If
    If Val(TiendaActual) < 100 Then
        If TiendaActual = 0 And lFornes Then
            PathExpor = "c:\telecom\epelsa\" 'App.Path & "\exportaBL\"
            PathExporRX = "c:\telecom\epelsa\" 'App.Path & "\exportarx\"
        Else
            PathExpor = App.Path & "\t" & Format(TiendaActual, "00") & "\"
        End If
    Else
        PathExpor = App.Path & "\t" & Format(TiendaActual, "000") & "\"
    End If
    
    Set Base = AbrirBase()
    If lTodos Then 'euroscale
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    Else
        MiTabla1 = "CabeceraSC10"
        MiTabla2 = "ticketsSC10"
    End If
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
                  ") order by val(nume)"
    
    'CadenadeLog sSQL
    
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        
        
        Registro.MoveFirst
        
        
        Do While Not Registro.EOF And CI_Salir = False
        
            If lFornes Then
                Fichero2 = FreeFile()
                If Dir(PathExporRX & "rx.txt") <> "" Then
                    Open PathExporRX & "rx.txt" For Append Shared As #Fichero2
                Else
                    Open PathExporRX & "rx.txt" For Output As #Fichero2
                End If
                lFirst = True
                nLinLa = 0
                Mienvio3 = ""
                nMess = 0
            End If
            
            With Registro
            
            Codigo_mos = Format(Val(!seccion), "00")
            Codigo_bal = Format(Val(!Balanza), "00")
            Codigo_ven = Mid(Format(Val(!vended), "0000"), 3, 2)
            Lineas = Format(Val(!numlin) - Val(!lincan), "000") ' No se exportan las canceladas...
            Fecha = Format(!Fecha, "yyyymmdd")
            hora = Format(!hora, "hhmm")
            'código de Cliente se pasa a 4 de longitud según Informático.
            'sCodCli = Format(Val(!codcli), "00000")
            sCodCli = Format(Val(!codcli), "00000")
            
            MiEnvio = Codigo_mos & Format(Val(!nume), "00000")
            
            If lFornes Then
'CLAVE  "LA"
'GRUPO  3        4           2                 VARIOS/SEVERAL
'CLAVE  5        5           1                 VARIOS/SEVERAL
'   (A). Si el registro no es la última parte del ticket.
'   (B). Si este registro es la última parte del ticket. El importe corresponderá al total con descuento de todos los tramos que componen
'       el ticket.
'   (P): Es un Pedido. No se acumula en listados y no se incrementa el contador de ticket.
'   (Q):Si es la última parte de un pedido.
'   (R). Etiquetas
'descripción            Inicio Fin   Tamaño   Formato
'NUMERO DE TICKET       6      10    5        VARIOS/SEVERAL
'CODIGO DE CLIENTE      11     15    5        VARIOS/SEVERAL --> 0 SI NO HAY CLIENTE
'NUMERO DE TRAMO        16     17    2        VARIOS/SEVERAL --> '1,2,3..
'NUMERO MENSAJE         18     19    2        VARIOS/SEVERAL --> Número mensaje dentro del tramo: 0, 1, 2..
'SIGNO                  20     20    1        VARIOS/SEVERAL
'IMPORTE                21     29    9        VARIOS/SEVERAL
'VENDEDOR               30     34    5        VARIOS/SEVERAL
                Mienvio3 = Codigo_mos & "LA" & Codigo_mos & "B" & Format(Val(!nume), "00000") & sCodCli & "00"
            End If
            
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & ") order by val(numlin)")
            
            nLinea = 0
            
            If Not Reg2.EOF Then
                Reg2.MoveFirst
                                
                Do While Not Reg2.EOF
                    With Reg2
                        
                        Cod_Cod = Format(!Code, "000000")
                        
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                sPes = Format(1000 * (Pes_Uni), "00000")
                                Pes_Uni = Format(1000 * (Pes_Uni), "000000")
                            Else
                                Pes_Uni = (-1) * !Weight
                                sPes = Format(1000 * (Pes_Uni), "00000")
                                Pes_Uni = "-" & Format(1000 * (Pes_Uni), "00000")
                            End If
                            
                            Tipo_Linea = "0"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                sPes = Format((Pes_Uni), "00000")
                                Pes_Uni = Format((Pes_Uni) * 1000, "000000")
                            Else
                                Pes_Uni = (-1) * !units
                                sPes = Format((Pes_Uni), "00000")
                                Pes_Uni = "-" & Format((Pes_Uni) * 1000, "00000")
                            End If
                            Tipo_Linea = "1"
                        End If
                        
                        If !amount < 0 Then
                            Total = (-1) * !amount
                            Total = Format(100 * (Total), "000000000")
                            Signo = "1"
                        Else
                            Total = !amount
                            Total = Format(100 * (Total), "000000000")
                            Signo = "0"
                        End If
                        If !lincan = 0 Then
                            LineaAnulada = "0"
                        Else
                            LineaAnulada = "1"
                        End If
                        
                        If !lincan = 0 Then
                            
                            nLinea = nLinea + 1
                            Num_lin = Format(nLinea, "000")
                            
                            Mienvio2 = ""
                            
                            If Val(Mid(Fecha, 5, 2)) < 10 Then
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Format(Val(Mid(Fecha, 5, 2)), "0") & ".TOT"
                            Else
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Chr(55 + Val(Mid(Fecha, 5, 2))) & ".TOT"
                            End If
                            
                            Mienvio2 = MiEnvio & Num_lin & Cod_Cod & Pes_Uni & Total & Codigo_ven & Fecha & hora & sCodCli
                            
                            Fichero1 = FreeFile()
                            If Dir(PathExpor & sFich) <> "" Then
                                Open PathExpor & sFich For Append Shared As #Fichero1
                            Else
                                Open PathExpor & sFich For Output As #Fichero1
                            End If
                            Print #Fichero1, Mienvio2
                            Close #Fichero1
                            
                            If lFornes Then
'CODIGO       35      40           6                 VARIOS/SEVERAL
'PESO 2 + 3   41      45           5                 VARIOS/SEVERAL
'SIGNO        46      46           1                 VARIOS/SEVERAL
'IMPORTE      47      55           9                 VARIOS/SEVERAL
'CODIGO       56      61           6                 VARIOS/SEVERAL
'PESO 2 + 3   62      66           5                 VARIOS/SEVERAL
'SIGNO        67      67           1                 VARIOS/SEVERAL
'IMPORTE      68      76           9                 VARIOS/SEVERAL
'CODIGO       77      82           6                 VARIOS/SEVERAL
'PESO 2 + 3   83      87           5                 VARIOS/SEVERAL
'SIGNO        88      88           1                 VARIOS/SEVERAL
'IMPORTE      89      97           9                 VARIOS/SEVERAL
'CODIGO       98    103            6                 VARIOS/SEVERAL
'PESO 2 + 3   104   108            5                 VARIOS/SEVERAL
'SIGNO        109    109           1                 VARIOS/SEVERAL
'IMPORTE      110    118           9                 VARIOS
'HORA Y FECHA 119    128          10  (HHMMDD VARIOS/SEVERAL
                                If lFirst = True Then
                                    lFirst = False
                                    sParte = ""
                                End If
                                nLinLa = nLinLa + 1
                                sParte = sParte & Cod_Cod & sPes & Signo & Total
                                
                                If nLinLa = 4 Then 'escribir registro
                                    nLinLa = 0
                                    If Registro!Importe > 0 Then 'añadir número de mensaje.
                                        sPes = Mienvio3 & Format(nMess, "00") & "0" & Format(Registro!Importe * 100, "000000000")
                                    Else
                                        sPes = Mienvio3 & Format(nMess, "00") & "1" & Format(Registro!Importe * (-100), "000000000")
                                    End If
                                    'Mienvio3 = Mienvio3 & Format(Registro!vended, "00000")
                                    nMess = nMess + 1
                                    Print #Fichero2, sPes & Format(Registro!vended, "00000") & sParte & hora & Format(CDate(Registro!Fecha), "ddmmyy")
                                    lFirst = True
                                End If
                            End If
                            
                        End If
                        
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.Movenext
                        
                        
                    
                    End With
                    
                    
                    
                Loop
                
                If lFornes Then
                    If nLinLa <> 0 Then 'completar registro
                        For nC = nLinLa + 1 To 4
                            sParte = sParte & "000000000000000000000"
                        Next nC
                        If Registro!Importe > 0 Then 'añadir número de mensaje.
                            sPes = Mienvio3 & Format(nMess, "00") & "0" & Format(Registro!Importe * 100, "000000000")
                        Else
                            sPes = Mienvio3 & Format(nMess, "00") & "1" & Format(Registro!Importe * (-100), "000000000")
                        End If
                        'Mienvio3 = Mienvio3 & Format(Registro!vended, "00000")
                        Print #Fichero2, sPes & Format(Registro!vended, "00000") & sParte & hora & Format(CDate(Registro!Fecha), "ddmmyy")
                    End If
                End If
                
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Close #Fichero2
            
            Registro.Movenext
        
            End With
            
            If lFornes Then Close #Fichero2
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    'On Error Resume Next
    'Close #Fichero1
    'On Error GoTo 0
    
    'Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    'lCogeTiquet = True

End Sub
'generar
Public Sub Bal_Epelsa_txt()
Dim Mybase As dao.Database
Dim myRs As dao.Recordset
Dim nFich As Integer
Dim sS As String
'Codigo#EnVenta#Operacion#PLU#Nombre#Precio#Tipo#Tara#DiasCad#Seccion#Familia#TextoG
    Set Mybase = AbrirBase
    Set myRs = Mybase.OpenRecordset("select * from articulo order by codigo")
    If Not myRs.EOF Then
        nFich = FreeFile()
        Open App.Path & "\bal_epelsa.txt" For Output As #nFich
        myRs.MoveFirst
        Do While Not myRs.EOF
            sS = Format(myRs.Fields("codigo"), "000000") & "#1" & "#A#" & Format(myRs.Fields("plu"), "0000") & "#" & Trim(myRs.Fields("des_plu1")) & _
                 "#" & Format(myRs.Fields("euros") * 100, "0000000")
            'codi_pes
            If myRs.Fields("codi_pes") = "W" Then
                sS = sS & "#1#"
            Else
                sS = sS & "#0#"
            End If
            sS = sS & Format(myRs.Fields("tara"), "00000") & "#" & Format(myRs.Fields("caducidad"), "00000") & "#" & Format(myRs.Fields("secc_maqui"), "00") & _
            "#" & Format(myRs.Fields("codi_fam"), "000") & "#aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaakkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiippppppppppppppppppppp"
            Print #nFich, sS
            myRs.Movenext
        Loop
        Close #nFich
    
    End If
    myRs.Close
    Mybase.Close
    
    
    
End Sub

Public Function CodSecPlu(codigo As Long) As String
    Dim Cod As String
    Dim Plu As String
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Set Base = AbrirBase()
    Set Registro = Base.OpenRecordset("select codigo,plu,secc_maqui from articulo where borrado=false and codigo=" & codigo)
    If Registro.EOF Then
        Plu = "000000"
    Else
        Plu = Format(Registro.Fields("secc_maqui"), "00") & Format(Registro.Fields("plu"), "0000")
    End If
    Base.Close
    CodSecPlu = Plu
End Function


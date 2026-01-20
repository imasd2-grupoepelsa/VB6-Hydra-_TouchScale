Attribute VB_Name = "tools"
Option Explicit
Public General_Events As Integer
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
Public Function AdaptaFecha(ByVal sfecha As String) As String
    AdaptaFecha = Left(sfecha, 2) & "/" & Mid(sfecha, 4, 2) & "/" & Right(sfecha, 2)
End Function
Public Function GeneraTextoIP(sData As String) As Boolean
    Dim Buf As String
    Dim bucle As Long
    Dim nPuntos As Integer
    Dim esPunto As Boolean
    Dim MyOK As Boolean
    Dim nNum As Integer
    Dim Num(4) As String
    Buf = sData
    MyOK = True
    For bucle = 1 To Len(Buf)
        Select Case Mid(Buf, bucle, 1)
            Case "0" To "9"
                esPunto = False
                Num(nNum) = Num(nNum) & Mid(Buf, bucle, 1)
            Case "."
                If Not esPunto Then
                    nPuntos = nPuntos + 1
                    nNum = nNum + 1
                    esPunto = True
                Else
                    MyOK = False
                    Exit For
                End If
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
Public Function PRsinIVA(precio As Double, Percent As Double) As Double
    Dim Psin As Double
    Psin = (precio * 100) / (100 + Percent)
    Psin = Round(Psin, 2)
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
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Dim RegArt As DAO.Recordset
    Dim Buffer1
    Dim Buffer As String
    Dim Trozo As String
    Dim Hay_I As Boolean
    Dim Tipo As Integer
    If MisDatos.Tipo < 5 Then
        Tipo = 0
    Else
        MisDatos.Tipo = MisDatos.Tipo - 4
        Tipo = 1
    End If
    
    If Dir(App.Path & "\bdp.txt") <> "" Then
        MisDatos.Tipo = 1
    End If
    
    Set Base = AbrirBase()
    Set Registro = Base.OpenRecordset("select * from codbar where numero_eqp=" & _
    MisDatos.NBalanza & " and borrado=false and secc_eqp=" & Tipo)
    If Registro.EOF Then
        Set Registro = Base.OpenRecordset("select * from codbar where secc_maqui=" & _
        MisDatos.NMostrador & " and (isnull(numero_eqp) or numero_eqp=0) and borrado=false and secc_eqp=" & Tipo)
    End If
    '-->red
    With Registro
        
        If Not .EOF Then
            Select Case MisDatos.Tipo
                Case 1
                    codigo_barras = !CB_VNT
                Case 2
                    codigo_barras = !cb_sup
                Case 3
                    codigo_barras = !cb_ev1
                Case 4
                    codigo_barras = !cb_ev2
                Case Else
                    codigo_barras = ""
            End Select
        Else
            If Dir(App.Path & "\bdp.txt") <> "" Then
                codigo_barras = "2TTTT" & Format(MisDatos.NMostrador, "00") & "IIIII"
            Else
                codigo_barras = ""
            End If
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
            Buffer1 = Left(codigo_barras, 1)
            codigo_barras = codigo_barras & "ñ"
            For bucle = 2 To Len(codigo_barras)
                If Asc(Buffer1) > 47 And Asc(Buffer1) < 58 Then
                    Trozo = Buffer1
                    Buffer = Buffer & Buffer1
                    Trozo = ""
                Else
                    If Buffer1 <> Mid(codigo_barras, bucle, 1) Then
                        Trozo = Trozo & Buffer1
                        Select Case Left(Trozo, 1)
                            Case "I"
                                If (MisDatos.Importe < 0) And (Not Sabeco) Then
                                    Hay_I = True
                                Else
                                    If (MisDatos.Importe < 0) And Sabeco Then
                                        Buffer = Buffer & Right("000000000000", Len(Trozo))
                                    Else
                                        Buffer = Buffer & Right(Format(MisDatos.Importe, "000000000000"), Len(Trozo))
                                    End If
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
   Base.Close
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
Public Function DigitoControl(Valor As String) As Integer
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
Public Function sin_Coma(ByVal MiCadena As String) As String

    ' *****************
    ' Esta función sustituye
    ' el caracter "," (coma) por
    ' el caracter "." (punto)
    Dim bucle As Integer
    Dim Buffer As String
    Buffer = ""
    For bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, bucle, 1) = "," Then
            Buffer = Buffer & "."
        Else
            Buffer = Buffer & Mid(MiCadena, bucle, 1)
        End If
    Next bucle
    sin_Coma = Buffer
End Function
Public Function Quitar_Coma(MiCadena As String, Longitud As Integer) As String
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
            Case Else
                Buffer = Buffer & Mid(MiCodigo, bucle, 1)
        End Select
    Next bucle
    If Buffer = "" And (MachineID = 2 Or MachineID = 1) And TipoC9 Then Buffer = "000000000000"
    EmularCodBar = Buffer
End Function
'***********
'***********
'***********
Public Function BoolToCheck(MiVal As Boolean) As Integer
    If MiVal Then
        BoolToCheck = vbChecked
    Else
        BoolToCheck = vbUnchecked
    End If
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
Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    On Error GoTo fin
    Fichero = FreeFile
    Open App.Path & "\loghydratt.txt" For Append As Fichero
    On Error Resume Next
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
fin:
    On Error GoTo 0
End Sub
Public Sub CadenadeLogTotales(MiCadena As String)
    Dim Fichero As Integer
    
    
    On Error GoTo fin
    
    
    Fichero = FreeFile
    Open App.Path & "\loghydratotales.txt" For Append As Fichero
    On Error Resume Next
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
fin:
    On Error GoTo 0
End Sub
'c2f 1.3.1 ... repetido
'Public Sub Do_Events()
'    If General_Events >= 5 Then
'        Sleep (1)
'        General_Events = 0
'    Else
'        General_Events = General_Events + 1
'    End If
'    DoEvents
'End Sub
Public Function CodigoToPlu(codigo As Long) As Long
    Dim cod As String
    Dim Plu As Long
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Set Base = AbrirBase()
    Set Registro = Base.OpenRecordset("select codigo,plu from articulo where borrado=false and codigo=" & codigo)
    If Registro.EOF Then
        Plu = 0
    Else
        Plu = Registro.Fields("plu")
    End If
    Base.Close
    CodigoToPlu = Plu
End Function

Public Function CodSecPlu(codigo As Long) As String
    Dim cod As String
    Dim Plu As String
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
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
    Dim LDec As Integer
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
   Dim nlen As Integer
   Dim Cont As Integer
   Dim s As String
   
   Dim sResul As String
   
   nlen = Len(sR)
   For Cont = 1 To nlen
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


Public Function toIpAdd(ByVal sR As String)
   Dim tmp As String
   Dim nlen As Integer
   Dim Cont As Integer
   Dim s As String
   nlen = Len(sR)
   For Cont = 1 To nlen
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


Public Function comFecha(ByVal Fecha As String) As Boolean
Dim d As Integer
Dim m As Integer
Dim Y As Integer
Dim forFecha As String

d = CInt(Left(Fecha, 2))
m = CInt(Mid(Fecha, 4, 2))
Y = CInt(Right(Fecha, 2))

forFecha = DateSerial(Y, m, d)

    If Fecha <> forFecha Then
        'comFecha = False
        comFecha = True
    Else
        comFecha = True
    End If

End Function




Public Function ObtenNombre(MiNombre As Tipo_NombreTiquet) As String
    Dim Resp As String
    Dim V1 As String
    Dim Cont As Integer
    Dim Buffer As String
    Dim Buf2 As String
    Dim Formato As String
    Dim b As Integer
    Dim B2 As Integer
    Resp = ""
    Cont = 1
    Buffer = Trim(TqNombre & "/")
    V1 = Left(Buffer, 1)
    For b = 2 To Len(Buffer)
        If Mid(Buffer, b, 1) = V1 Then
            Cont = Cont + 1
        Else
            Formato = ""
            Buf2 = ""
            For B2 = 1 To Cont
                Formato = Formato & "0"
            Next B2
            Select Case V1
                Case "R"
                    Buf2 = Mid("********", 1, Len(Formato))
                Case "B"
                    Buf2 = Format(MiNombre.Balanza, Formato)
                Case "S"
                    Buf2 = Format(MiNombre.seccion, Formato)
                Case "T"
                    Buf2 = Format(MiNombre.tiquet, Formato)
                Case "X"
                    Buf2 = Format(MiNombre.modo, Formato)
                Case "C"
                    Buf2 = Format(MiNombre.cliente, Formato)
                Case "V"
                    Buf2 = Format(MiNombre.vendedor, Formato)
            End Select
            If Len(Buf2) > Cont Then Buf2 = Right(Buf2, Cont)
            Resp = Resp & Buf2
            V1 = Mid(Buffer, b, 1)
            Cont = 1
        End If
    Next b
    ObtenNombre = Resp
End Function



Public Function Redondea(ByVal dato As Double) As Double
'************************************
' redondea la diferencia que
' se produce al tratar datos
' con decimales, con Access y VB
'************************************
    If UsaEuro Then
        dato = Round(dato, 2)
    Else
        dato = Round(dato, decimales)
    End If
    Redondea = dato
End Function

Public Sub Aux_Ejecuta(Que_Ejecuta As String, Optional Parametros As String)
Dim MyProc As Long
Dim MyResul As Long
Dim Arch As Integer
Dim BufEnable As Boolean
Dim NombrePrograma As String
    NombrePrograma = "auxiliartouch.exe"
    If LCase(Que_Ejecuta) = "hydratouch" Then NombrePrograma = "OLDhydratouch.EXE"
    If LCase(Que_Ejecuta) = "mng" Then NombrePrograma = "OLDhydratouch.EXE"
    If LCase(Que_Ejecuta) = "iv4com" Then NombrePrograma = "OLDhydratouch.EXE"
    If Dir(App.Path & "\" & NombrePrograma) <> "" Then
        Arch = FreeFile()
        Open App.Path & "\auxiliartouch.tmp" For Output As #Arch
        Print #Arch, App.Path & "\dbasetouch.mdb"
        Print #Arch, App.Path & "\hydratouch.ini"
        Print #Arch, Que_Ejecuta
        Print #Arch, Parametros
        Print #Arch, "N"
        Print #Arch, ""
        Print #Arch, 0
        Close #Arch
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\" & NombrePrograma, vbNormalFocus))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
    Else
        MsgBox "auxiliartouch.exe not found. Can't execute order", vbCritical
    End If
    If Dir(App.Path & "\auxiliartouch.tmp") <> "" Then
        On Error Resume Next
        Kill App.Path & "\auxiliartouch.tmp"
        On Error GoTo 0
    End If
End Sub
'Public Sub Do_Events()
'    Sleep (1)
'    DoEvents
'    Sleep (1)
'End Sub



Public Function versionar() As String
    versionar = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
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

Public Function ElPing(ElHost As String) As Integer
    '///////////////////////////////////////////////
    '//  Lanza ping a la dirección que se indique //
    '//  llamando al OCX de ping                  //
    '///////////////////////////////////////////////
    Dim Contador As Integer
    Dim Echo As ICMP_ECHO_REPLY
    Dim Resp As Long
    If Dir(App.Path & "\noping.txt") <> "" Then
        ElPing = 1
    Else
        Resp = ICMP_Ping(ElHost, "alive", Echo)
        If Resp = 0 Then
            ElPing = 1
        Else
            ElPing = 0
        End If
    End If
End Function

Public Sub ImportTeclas(ByVal sS As String)
    Dim Arch1 As Integer
    Dim Buffer As String
    Dim Base As DAO.Database
    Dim MyCod As Long
    Dim MySec As Integer
    Dim MyTabla As Integer
    Dim myTec As Integer
    Dim Registro As DAO.Recordset
    Dim sDes As String
        
    If Dir(sS) <> "" Then
        If FileLen(sS) > 0 Then
            'FrmBalGloDat.Caption = "Leyendo Teclas..."
            'Load FrmBalGloDat
            'FrmBalGloDat.Show
            'FrmBalGloDat.dato "Leyendo Archivo ..."
        
            Set Base = DAO.OpenDatabase(App.Path & "\dbasetouch.mdb")
            Arch1 = FreeFile()
            Open sS For Input As #Arch1
            Do While Not EOF(Arch1)
            Line Input #Arch1, Buffer
                If Len(Buffer) = 13 Then
                    MySec = Mid(Buffer, 1, 2)
                    MyTabla = Mid(Buffer, 3, 2)
                    myTec = Mid(Buffer, 5, 3)
                    MyCod = Mid(Buffer, 8)
                    If MyCod <> 0 Then
                        Set Registro = Base.OpenRecordset("select codigo,des_plu1 from articulo where codigo=" & CStr(MyCod))
                        If Not Registro.EOF Then
                            sDes = Registro.Fields("des_plu1")
                        Else
                            sDes = ""
                        End If
                        Registro.Close
                        Set Registro = Nothing
                        Base.Execute "delete from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'"
                        Set Registro = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'")
                        If Registro.EOF Then
                            Registro.AddNew
                        Else
                            Registro.Edit
                        End If
                        
                        'FrmBalGloDat.dato "Sec.:" & CStr(MySec) & " Tabla:" & CStr(MyTabla) & " Tecla:" & CStr(myTec) & " Cod.:" & CStr(MyCod)
                        
                        Registro.Fields("codi_ident") = MySec
                        Registro.Fields("secc_maqui") = MySec
                        Registro.Fields("tran_tecsc10") = MySec
                        Registro.Fields("codigo") = MyCod
                        Registro.Fields("plu") = 0
                        Registro.Fields("desc_plu") = sDes
                        Registro.Fields("secc_eqp") = False
                        Registro.Fields("numero_eqp") = False
                        Registro.Fields("tabla") = MyTabla
                        Registro.Fields("codi_tec") = myTec
                        Registro.Fields("prog_tec") = "TOUCH"
                        Registro.Fields("tran_tec") = ""
                        Registro.Fields("borrado") = False
                        
                        Registro.Update
                        Registro.Close
                        Set Registro = Nothing
                    Else
                        'Base.Execute "delete from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'"
                        Set Registro = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'")
                        If Not Registro.EOF Then
                            Registro.Edit
                            Registro.Fields("borrado") = True
                            Registro.Update
                        End If
                        Registro.Close
                        Set Registro = Nothing
                        
                        'FrmBalGloDat.dato "Sec.:" & CStr(MySec) & " Tabla:" & CStr(MyTabla) & " Tecla:" & CStr(myTec) & " Baja."
                    End If
                
                End If
            Loop
            Close #Arch1
            
            Kill sS
            
            Base.Close
            Set Base = Nothing
            'Unload FrmBalGloDat
        End If
    End If

End Sub


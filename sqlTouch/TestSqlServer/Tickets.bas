Attribute VB_Name = "Tickets"
Option Explicit

'''''''''''''''
'¡ATENCIÓN A LA CONEXIÓN...
' Por defecto en la Balanza el puerto para conectar será: 1031 --> SQLSERVER Configuration Manager -->
'                                                                  Configuración de red de SQL Server -->
'                                                                  Protocolos de SQLEXPRESS -->
'                                                                  TCP/IP -->Propiedades-->Direcciones IP-->IPALL-->Puertos Dinámicos TCP.
'''''''''''''''
Public Type Tipo_Linea
    EsPesado As Boolean
    FactorPeso As Integer
    precio As Double
    cantidad As Double
    NumeroLinea As Byte
    CodigoArticulo As Long
    IVA As Integer
    TipoLinea As Byte
    DescuentoLinea As Double
    ImporteLinea As Double
    LineaCancelada As Boolean
    SehaModificado As Boolean
    ModificadoTemporal As Boolean
    FichaVacuno As Long
    descriptivo As String
End Type


Public Type tipo_cabecera
     Ntiquet As Long
     NVendedor As Long
     NMostrador As Long
     NCliente As Long
     Fecha As Variant
     hora As Variant
     NBalanza As Long
     ImporteTotal As Double
     NLineas As Long
     LineaInicial As Long
     LineaFinal As Long
     codigo As Long
     CadenaEstadoTiquet As String
     ImporteBruto As Double
     Impuestos As Double
     descuento As Double
     LineasCanceladas As Long
     tipoTiquet As Long
     EstadoTiquet As Long
     LineasMensaje As Long
     '********************
     ' este dato se introduce
     ' tras pasar por actualiza_tabla_cabecera
     Incluir_en_Base As Boolean
     tipoPago As Long
     cantidadEntre As Double
     Cambio As Double
     lote As String
     lEncargo As Boolean
     Lines() As Tipo_Linea
End Type

Public Type Tipo_TqCodBar
    Tipo As Integer
    NBalanza As Integer
    NMostrador As Integer
    NTicket As Long
    NVendedor As Long
    Importe As Double
End Type

'Formas de Pago....
'1   Efectivo
'2   Tarjeta Crédito
'3   Tarjeta Débito
'4   EBT Cash
'5   Gift Card
'6   Cheque
'7   Pendiente
'14  Ticket Restaurant
'15  Chèque -Vacances
'16  Cheque Gourmet
'17  Restaurant Pass
'19  Cupón
'20  Others
'21  EBT FoodStamp
'22  Visa
'23  MasterCard
'24  American Express
'25  Traveller 's Check
'28  Multibanco
'30  Visa Electron
'31  Diners Club
'32  Big Pass
'33  Sodexho
'34  Pin
'35  Chipknip
'36  Cestaticket
'37  EC -Maestro
'38  Quick
'39  BPI -Express
'40  Gift Check
'41  Ripley
'42  Saga
'43  Otras tarjetas
'44  Cheque Comercial
'45  Cheque de Gobiernos
'46  CLAVE

Public Sub LeeTickets()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    Dim MyRecAux As New ADODB.Recordset
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim nCont As Long
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    ''''''''''''
    'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    ''''''''''''
    'Para conectar mediante IP y Port.
    '''''''''''''''
    '¡ATENCIÓN A LA CONEXIÓN...
    ' Por defecto en la Balanza el puerto para conectar será: 1031
    '--> SQLSERVER Configuration Manager -->
    '    Configuración de red de SQL Server -->
    '    Protocolos de SQLEXPRESS -->
    '    TCP/IP -->Propiedades-->Direcciones IP-->IPALL-->Puertos Dinámicos TCP.
    '''''''''''''''
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000" 'el puerto es normalmente 1031.
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
            
    'desmarcar tiques... PRUEBA
    ''Ffecha = Now()
    ''cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:00.000" & "'"
    ''MyConnObj.Execute cm
    cm = "update dbo.documento set iscapturado=0"
    MyConnObj.Execute cm
    '''''''''''''''''''''''''''
    
    
    '(ICG) tipo documento (0=Venta, 2=Envasado, 3=autoServicio)
    'Hydra+ Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    '*************************
    
    'En primera versión se considera:
    '   ICG           Hydra+
    '  TotalNeto      Importe
    '  TotalBruto     Importe_bruto
    '  ImporteCambio  Cambio
    '                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
    'Para marcar tique en fecha.
    'MyRecSet.Open "select * from dbo.documento where ((SUBSTRING(convert(varchar(23),fecha,121),12,8)='00:00:00') and isfinalizado=1 and numero<>0) order by tipodocumento,numero", MyConnObj, adOpenKeyset, adLockOptimistic
    MyRecSet.Open "select * from dbo.documento where ((isnull(iscapturado,0)=0 or iscapturado=0) and (isfinalizado=1) and (numero<>0)) order by tipodocumento,numero", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Ffecha = MyRecSet.Fields("fecha")
            Fhora = Format(MyRecSet.Fields("hora"), "0000")
            Fhora = Left(Fhora, 2) & ":" & Right(Fhora, 2) & ":00"
            IdDoc = MyRecSet.Fields("iddocumento")
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            'recoger datos
            Cabecera.Ntiquet = MyRecSet.Fields("numero")
            
            'If Cabecera.Ntiquet = 7 Then
            '    Cabecera.NBalanza = Cabecera.NBalanza
            'End If
            
            Cabecera.NVendedor = MyRecSet.Fields("idusuario")
            Cabecera.NMostrador = MyRecSet.Fields("idseccion")
            Cabecera.NCliente = MyRecSet.Fields("idcontacto")
            Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.hora = Fhora
            Cabecera.NBalanza = MyRecSet.Fields("numeroterminal")
            Cabecera.ImporteTotal = MyRecSet.Fields("totalneto")
            Cabecera.NLineas = nNumLin
            Cabecera.LineaInicial = 1
            Cabecera.LineaFinal = nNumLin
            Cabecera.LineasMensaje = nNumLin
            
            'cabecera.codigo ???
            
            ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
            '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
            If MyRecSet.Fields("islocal") Then
                Cabecera.CadenaEstadoTiquet = "Local"
            Else
                Cabecera.CadenaEstadoTiquet = "Red"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Cabecera.ImporteBruto = MyRecSet.Fields("totalbruto")
            Cabecera.Impuestos = MyRecSet.Fields("totalimpuestos")
            Cabecera.descuento = MyRecSet.Fields("totaldescuento")
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where (Idlineaabono = -1 and iddocumento=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.LineasCanceladas = nNumLin
            
            'Adapta TipoDocumento(ICG) a TypTic Hydra+
            Select Case MyRecSet.Fields("tipodocumento")
            Case 0
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 5
                Else
                    Cabecera.tipoTiquet = 1
                End If
            Case 2
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 7
                Else
                    Cabecera.tipoTiquet = 3
                End If
            Case 3
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 8
                Else
                    Cabecera.tipoTiquet = 4
                End If
            End Select
            'también se vé afectado si es cancelado o no pero de mamento cerrado...
            Cabecera.EstadoTiquet = 16
            '********************
            ' este dato se introduce
            ' tras pasar por actualiza_tabla_cabecera
            'Incluir_en_Base As Boolean
            Cabecera.tipoPago = MyRecSet.Fields("idmoneda")
            Cabecera.cantidadEntre = MyRecSet.Fields("totalneto") + MyRecSet.Fields("importecambio")
            Cabecera.Cambio = MyRecSet.Fields("importecambio")
            'lote As String
            'lEncargo As Boolean


            'LINEAS

            nNumLin = 0
            MyRecAux.Open "select * from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF
                    nNumLin = nNumLin + 1
                    If MyRecAux.Fields("isventaporpeso") = True Then
                        Cabecera.Lines(nNumLin).EsPesado = True
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("peso")
                    Else
                        Cabecera.Lines(nNumLin).EsPesado = False
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("unidades")
                    End If
                    Cabecera.Lines(nNumLin).FactorPeso = 1
                    Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("precio")
                    
                    Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                    Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("idarticulo")
                    'ver qué se hace IVA As Integer
                    'ver qué se hace con TipoLinea As Byte
                    Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("descuento") 'porcentaje
                    Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                    If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                        Cabecera.Lines(nNumLin).LineaCancelada = True
                    Else
                        Cabecera.Lines(nNumLin).LineaCancelada = False
                    End If
                    If MyRecAux.Fields("precio") <> MyRecAux.Fields("preciodefecto") Then
                        Cabecera.Lines(nNumLin).SehaModificado = True
                    Else
                        Cabecera.Lines(nNumLin).SehaModificado = False
                    End If
                    '¿? ModificadoTemporal As Boolean
                    
                    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("idfichatrazabilidad")
                    'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                    Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    MyRecAux.MoveNext
                Loop
                MyRecAux.Close
                Set MyRecAux = Nothing
                Actualiza_Tabla_Cabecera MyBase, Cabecera
                If Cabecera.Incluir_en_Base = True Then
                    Actualiza_Tabla_Tickets MyBase, Cabecera
                    Actualiza_Tabla_Gtarti MyBase, Cabecera
                    Actualiza_Tabla_GtVend MyBase, Cabecera
                    actualiza_tabla_gtHora MyBase, Cabecera
                    Actualiza_Tabla_Gtsecs MyBase, Cabecera
                    actualiza_tabla_gtInfVen MyBase, Cabecera
                    
                End If
            End If



            '''''''''''''''
            'marcar...
            'cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
            cm = "update dbo.documento set iscapturado=1 where iddocumento=" & CStr(IdDoc)
            MyConnObj.Execute cm
            ''''''''''

            MyRecSet.Update

            MyRecSet.MoveNext
        Loop
    End If

''''''''

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0

End Sub
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

Public Function Genera_CodBar(MisDatos As Tipo_TqCodBar, MiCodigo As Long) As String
    '**********************
    ' Tipo 0 --> Red Venta
    ' Tipo 5 --> Local Venta
    '**********************
    Dim MiPlu As Integer
    Dim Bucle As Integer
    Dim codigo_barras As String
    Dim codigoDpto As String
    Dim snumtiq As String
    Dim sbarras$
    Dim cadena As String
    Dim envio
    Dim Checksum
    Dim Base As Database
    Dim Registro As Recordset
    Dim RegArt As Recordset
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
    Set Base = OpenDatabase(App.Path & "\dbase.mdb") 'PRUEBA base_general !!!
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
            codigo_barras = ""
        End If
        If InStr(1, codigo_barras, "P") Then
            Set RegArt = Base.OpenRecordset("select * from articulo where borrado=false and codigo=" & MiCodigo)
                If Not RegArt.EOF Then
                    MiPlu = RegArt!plu
                Else
                    MiPlu = 0
                End If
            RegArt.Close
        End If
        If codigo_barras <> "" Then
            Buffer1 = Left(codigo_barras, 1)
            codigo_barras = codigo_barras & "ñ"
            For Bucle = 2 To Len(codigo_barras)
                If Asc(Buffer1) > 47 And Asc(Buffer1) < 58 Then
                    Trozo = Buffer1
                    Buffer = Buffer & Buffer1
                    Trozo = ""
                Else
                    If Buffer1 <> Mid(codigo_barras, Bucle, 1) Then
                        Trozo = Trozo & Buffer1
                        Select Case Left(Trozo, 1)
                            Case "I"
                                If (MisDatos.Importe < 0) Then
                                    Hay_I = True
                                Else
                                    Buffer = Buffer & Right(Format(MisDatos.Importe, "000000000000"), Len(Trozo))
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
                Buffer1 = Mid(codigo_barras, Bucle, 1)
            Next Bucle
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
'Public Function redondea(ByVal dato As Double) As Double
'    If UsaEuro Then
'        dato = Round(dato, 2)
'    Else
'        dato = Round(dato, decimales)
'    End If
'    redondea = dato
'End Function

'En primera versión se considera:
'   ICG           Hydra+
'  -------------  -------------
'  TotalNeto      Importe
'  TotalBruto     Importe_bruto
'  ImporteCambio  Cambio
'                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
Public Sub Actualiza_Tabla_Cabecera(Base As Database, Cabecera As tipo_cabecera)
'*************************
' Tipos de Tiquet (typtic)
' 1 --> Rollo
' 2 --> Super
' 3 --> Envasado
' 4 --> Autoservicio
' 5 --> Rollo Local
' 6 --> Super Local
' 7 --> Envasado Local
' 8 --> Autoservicio Local
'*************************
Dim regVendedor As DAO.Recordset
Dim Registro As DAO.Recordset
Dim RegSeccion As DAO.Recordset
Dim RegDes As DAO.Recordset
Dim MiCadenaDebug As String
Dim tqcodbar As Tipo_TqCodBar
Dim micodbar As String
Dim Cod_Art As Long
Dim bimpoBrutoCab As Double
Dim sHoraCab As String
Dim nMyCont As Integer
Dim sCont As Variant
Dim sTmp As String
Dim Que_base As String
Dim sSQL As String
'''''''''''''''''''''''''''''''
    
    Que_base = "GA"
    '***************************
    ' calcula el código de barras
    '****************************
    tqcodbar.Importe = Cabecera.ImporteTotal
    tqcodbar.NBalanza = Cabecera.NBalanza
    tqcodbar.NMostrador = Cabecera.NMostrador
    tqcodbar.NTicket = Cabecera.Ntiquet
    tqcodbar.NVendedor = Cabecera.NVendedor
    '1.8.8 modo trainning viene en lEncargo del OCX
    'el modo trainning se va a considerar tipoPago = 12
    ''''''''''''''''''''''''''''''''''''''''
    tqcodbar.Tipo = Cabecera.tipoTiquet
    If Cabecera.NLineas >= 1 Then
        Cod_Art = Cabecera.Lines(1).CodigoArticulo
    Else
        Cod_Art = 0
    End If
    micodbar = Genera_CodBar(tqcodbar, Cod_Art)
    
    If Que_base = "GA" Then
        MiCadenaDebug = "Euroscale / "
        Set Registro = Base.OpenRecordset("select * from cabecera where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & Cabecera.NVendedor & Chr(34))
    End If
    If Que_base = "SC10" Then
        MiCadenaDebug = "SC10 / "
        Set Registro = Base.OpenRecordset("select * from cabeceraSC10 where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & _
        Cabecera.NVendedor & Chr(34))
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    With Registro
        
        If .EOF Then
        

                'If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Cabecera de Tiquet : " _
                '& Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                '& "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.Fecha & "Hora : " & Cabecera.hora)
                .AddNew
                .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
                .Fields("seccion") = Cabecera.NMostrador
                .Fields("codbar") = micodbar
                .Fields("Balanza") = Cabecera.NBalanza
                .Fields("vended") = Cabecera.NVendedor
                .Fields("numlin") = Cabecera.NLineas
                .Fields("lincan") = Cabecera.LineasCanceladas
                
                
                'If Not UsaEuro Then 'c2f 2.0.18 ver aquí.
                '    .Fields("Importe") = Cabecera.ImporteTotal / (10 ^ decimales)
                '    bimpoBrutoCab = redondea(Cabecera.ImporteBruto / (10 ^ decimales)) 'c2f invicta
                '
                '
                'Else
                    .Fields("Importe") = Cabecera.ImporteTotal '/ 100
                    bimpoBrutoCab = Cabecera.ImporteBruto 'redondea(Cabecera.ImporteBruto  / 100) 'c2f invicta
                    
                'End If
                
                'c2f invicta
                On Error Resume Next
                .Fields("importe_bruto") = bimpoBrutoCab
                ''''''''''''
                On Error GoTo 0
                ''''''''''''
                .Fields("hora") = Format(Cabecera.hora, "hh:mm:ss")
                .Fields("d_fecha") = Cabecera.Fecha
                .Fields("d_hora") = Cabecera.hora
                .Fields("onoff") = Cabecera.CadenaEstadoTiquet
                .Fields("STPETIC") = Cabecera.EstadoTiquet
                .Fields("codcli") = Cabecera.NCliente
                .Fields("Fecha") = Format(Cabecera.Fecha, "dd/mm/yy")
                .Fields("TYPTIC") = Cabecera.tipoTiquet
                .Fields("reserv2") = "/"
                .Fields("reserv3") = "NO"  'no exportado
                '//cas.v117
                
                'If Cabecera.lEncargo Then
                '    .Fields("tipoPago") = 12
                'Else
                    .Fields("tipoPago") = Cabecera.tipoPago
                'End If
                
                'If Not UsaEuro Then
                '    .Fields("cantidadEntre") = Cabecera.cantidadEntre / (10 ^ decimales)
                '    .Fields("cambio") = Cabecera.Cambio / (10 ^ decimales)
                'Else
                    .Fields("cantidadEntre") = Cabecera.cantidadEntre '/ 100
                    .Fields("cambio") = Cabecera.Cambio '/ 100
                'End If
                '//cas.v117
                On Error Resume Next
                If Cabecera.lote <> "" Then
                    .Fields("numlote") = Cabecera.lote
                Else
                    .Fields("numlote") = " "
        
                End If
                On Error GoTo 0
                .Update
                Cabecera.Incluir_en_Base = True
                
                '\\\\\\Descuentos a Tabla Descuentos...
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If (Cabecera.descuento > 0) Then
                    sSQL = "SELECT * " & _
                           "FROM descuentos " & _
                           " WHERE fecha=" & "#" & Format(Cabecera.Fecha, "mm,dd,yy") & "#" & _
                           " AND nume=" & CStr(Cabecera.Ntiquet) & _
                           " AND seccion=" & CStr(Cabecera.NMostrador) & _
                           " AND balanza=" & CStr(Cabecera.NBalanza) & _
                           " AND vended=" & CStr(Cabecera.NVendedor) & _
                           " AND typTic=" & CStr(Cabecera.tipoTiquet)
                           '" AND hora=" & Cabecera.hora
                    Set RegDes = Base.OpenRecordset(sSQL)
                
                    With RegDes
                
                        If .EOF Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        .Fields("nume") = Cabecera.Ntiquet
                        .Fields("seccion") = Cabecera.NMostrador
                        .Fields("balanza") = Cabecera.NBalanza
                        .Fields("vended") = Cabecera.NVendedor
                        .Fields("typTic") = Cabecera.tipoTiquet
                        .Fields("hora") = Format(Cabecera.hora, "hh:mm:ss")
                        .Fields("fecha") = Format(Cabecera.Fecha, "dd/mm/yyyy")
                        .Fields("importeneto") = Cabecera.ImporteTotal
                        .Fields("importebruto") = Cabecera.ImporteTotal + Cabecera.descuento
                        .Fields("descuentototal") = Cabecera.descuento
                        .Update
                        .Close
                    End With
                    Set RegDes = Nothing
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
        End If
    End With
    
    
End Sub
Public Sub Actualiza_Tabla_GtVend(Base As Database, Cabecera As tipo_cabecera)
    '////////////////////////////////
    ' INCORPORA LOS DATOS DE TOTALES
    ' POR VENDEDOR A PARTIR DE
    ' LO RECIBIDO EN UN TICKET
    '////////////////////////////////
    Dim Encontrado As Boolean
    Dim Bucle As Long
    Dim ElPrecio As Double
    Dim ElImporte As Double
    Dim Lafecha As String
    Dim LaSeccion As Integer
    Dim Registro As Recordset
    Dim rstseccion As Recordset
    Dim rstvendedor As Recordset
    Dim LineaInicial As Boolean
    
    On Error GoTo finVend

    LineaInicial = True
    
    Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
    If Not rstseccion.EOF Then
        LaSeccion = rstseccion.Fields("codi_ident")
    Else
        LaSeccion = Cabecera.NMostrador
    End If
    rstseccion.Close
    Encontrado = False
    'If Cabecera.lEncargo Then
    'Set Registro = Base.OpenRecordset _
    '("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    '" and codi_ident=" & Val(LaSeccion) & " and modo=" & CStr(12) & _
    '" and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    '
    'Else
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    ''" and fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#") checa
    'End If
    With Registro
        If .EOF Then
            .AddNew
            .Fields("Fecha") = Cabecera.Fecha
            .Fields("codi_ident") = LaSeccion
            .Fields("ident_vend") = Cabecera.NVendedor
            .Fields("secc_Maqui") = Cabecera.NMostrador
            'If Cabecera.lEncargo Then
            '    .Fields("Modo") = 12
            'Else
                .Fields("Modo") = Cabecera.tipoTiquet
            'End If
            .Fields("tiquets") = 0
            .Fields("operacion") = 0
            .Fields("cancela") = 0
            .Fields("impor_neg") = 0
            '********************
            ' Nombre del vendedor
            '********************
            Set rstvendedor = Base.OpenRecordset _
            ("select * from vendedor where borrado=false and ident_vend=" & Cabecera.NVendedor _
            & " and codi_ident=" & LaSeccion)
            If rstvendedor.EOF Or .Fields("Nombre") = "" Then
                .Fields("Nombre") = Space(25)
            Else
                .Fields("Nombre") = rstvendedor.Fields("Nombre")
            End If
            rstvendedor.Close
            ' ****
            'pesado o unidades
            ' ****
            .Fields("Peso") = 0
            .Fields("unidad") = 0
            ' ****
            .Fields("pesetas") = 0
            ' ****
            ' Cancelada o no
            ' ****
            .Fields("totcance") = 0
            .Fields("supertot") = 0
            .Fields("Total") = 0
            .Update
        End If
    End With
    'If Cabecera.lEncargo Then
    'Set Registro = Base.OpenRecordset _
    '("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    '" and codi_ident=" & Val(LaSeccion) & " and modo=" & CStr(12) & _
    '" and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    '
    'Else
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    ''" and fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#") checa
    'End If
    With Registro
        If LineaInicial Then
            'Edit_Record Registro
            .Edit
            .Fields("tiquets") = .Fields("tiquets") + 1
            .Fields("operacion") = .Fields("operacion") + Cabecera.NLineas
            .Update
        End If
        For Bucle = 1 To Cabecera.LineasMensaje
            'Edit_Record Registro
            .Edit
            'If UsaEuro Then
                ElImporte = Cabecera.Lines(Bucle).ImporteLinea '/ 100
                ElPrecio = Cabecera.Lines(Bucle).precio '/ 100
            'Else
            '    ElImporte = Linea(Bucle).ImporteLinea / (10 ^ decimales)
            '    ElPrecio = Linea(Bucle).precio / (10 ^ decimales)
            'End If
            ' ****
            'pesado o unidades
            ' ****
            'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then '1.8.4 fanjul
            '    If Linea(Bucle).LineaCancelada = False Then
            '        If Linea(Bucle).EsPesado = False Then
            '            .Fields("unidad") = .Fields("unidad") + (-1) * (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
            '            '.Fields("Peso") = 0
            '        Else
            '            .Fields("Peso") = .Fields("Peso") + (-1) * (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
            '            '.Fields("unidad") = 0
            '        End If
            '        ' ****
            '        .Fields("pesetas") = redondea(.Fields("pesetas") + (-1) * ElImporte * Linea(Bucle).FactorPeso)
            '    End If
            'Else
                If Cabecera.Lines(Bucle).LineaCancelada = False Then
                    If Cabecera.Lines(Bucle).EsPesado = False Then
                        .Fields("unidad") = .Fields("unidad") + Cabecera.Lines(Bucle).cantidad '(Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        'If Linea(Bucle).FactorPeso = -1 Then
                        '    .Fields("impor_neg") = redondea(.Fields("impor_neg") + ElImporte)
                        'End If
                        ''.Fields("Peso") = 0
                        If Cabecera.Lines(Bucle).ImporteLinea < 0 Then
                            .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                        End If
                    Else
                        .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(Bucle).cantidad '(Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        ''.Fields("unidad") = 0
                        'If Linea(Bucle).FactorPeso = -1 Then
                        '    .Fields("impor_neg") = redondea(.Fields("impor_neg") + ElImporte)
                        'End If
                        If Cabecera.Lines(Bucle).ImporteLinea < 0 Then
                            .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                        End If
                    End If
                    ' ****
                    '.Fields("pesetas") = redondea(.Fields("pesetas") + ElImporte * Linea(Bucle).FactorPeso)
                    .Fields("pesetas") = .Fields("pesetas") + ElImporte
                End If
            'End If
            ' ****
            ' Cancelada o no
            ' ****
            If Cabecera.Lines(Bucle).LineaCancelada Then
                If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                    .Fields("totcance") = .Fields("totcance") - ElImporte 'redondea(.Fields("totcance") - ElImporte)
                    .Fields("supertot") = .Fields("supertot") - ElImporte 'redondea(.Fields("supertot") - ElImporte)
                    .Fields("cancela") = .Fields("cancela") - 1 'redondea(.Fields("cancela") - 1)
                Else
                    .Fields("totcance") = .Fields("totcance") + ElImporte 'redondea(.Fields("totcance") + ElImporte)
                    .Fields("supertot") = .Fields("supertot") + ElImporte 'redondea(.Fields("supertot") + ElImporte)
                    .Fields("cancela") = .Fields("cancela") + 1 'redondea(.Fields("cancela") + 1)
                End If
            Else
                If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                    .Fields("Total") = .Fields("Total") + ElImporte 'redondea(.Fields("Total") + (-1) * ElImporte * Linea(Bucle).FactorPeso)
                    .Fields("supertot") = .Fields("Total") + ElImporte 'redondea(.Fields("Total") + (-1) * ElImporte)
                Else
                    .Fields("Total") = .Fields("Total") + ElImporte 'redondea(.Fields("Total") + ElImporte * Linea(Bucle).FactorPeso)
                    .Fields("supertot") = .Fields("Total") + ElImporte 'redondea(.Fields("Total") + ElImporte)
                End If
            End If
            .Update
        Next Bucle
    End With
    
finVend:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'CadenadeLog "Actualiza tabla vendedores:" & Err.Description
    End If
    
End Sub
'//////
'//////
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES HORARIOS A PARTIR DE LO RECIBIDO EN EL TICKET
'//////////////////////////////////////////////////////////////////////////////////
Function actualiza_tabla_gtHora(Base As Database, Cabecera As tipo_cabecera)
    Dim rst As Recordset
    Dim cnt As Long
    Dim sSQL As String
    Dim lclImporte As Double
    Dim LineaInicial As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    sSQL = "SELECT * " & _
'           "FROM gtInfVen " & _
'           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
'           " AND codi_ident=" & Val(lclSec) & _
'           " AND modo=" & Cabecera.tipoTiquet & _
'           " AND fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#"
'    lclFecha = Trim(CDate(Format(Cabecera.Fecha, "dd/mm/yyyy")))
'    lclHora = Trim(CStr(Format(Cabecera.hora, "hh")))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    sSQL = "SELECT * " & _
'           "FROM gthora " & _
'           "WHERE FECHA=" & "#" & lclFecha & "#" & _
'           " AND HORA=" & "'" & lclHora & "' " & _
'           "ORDER BY FECHA"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'" AND hora=" & Chr(35) & Format(Cabecera.hora, "hh") & Chr(35)
    '"WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _ checa
    sSQL = "SELECT * " & _
           "FROM gthora " & _
           "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
           " AND cstr(hour(hora))=" & Chr(34) & CStr(Hour(Cabecera.hora)) & Chr(34)
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        If .EOF Then
            .AddNew
            .Fields("FECHA") = Cabecera.Fecha
            .Fields("HORA") = Cabecera.hora
            .Fields("SECCION") = 0
            .Fields("EQUIPO") = 0
            .Fields("VENDEDOR") = 0
            .Fields("TIQUET") = 0
            .Fields("OPERACIONES") = 0
            .Fields("OPERACIONES_CAN") = 0
            .Fields("OPERACIONES_NEG") = 0
            .Fields("IMPORTE") = 0
            .Fields("IMPORTE_CAN") = 0
            .Fields("IMPORTE_NEG") = 0
            .Fields("PESO") = 0
            .Fields("PESO_CAN") = 0
            .Fields("PESO_NEG") = 0
            .Fields("UNIDADES") = 0
            .Fields("UNIDADES_CAN") = 0
            .Fields("UNIDADES_NEG") = 0
            .Update
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    'Edit_Record rst
    rst.Edit
    rst.Fields("TIQUET") = rst.Fields("TIQUET") + 1
    For cnt = 1 To Cabecera.LineasMensaje
        'If UsaEuro Then
            lclImporte = Cabecera.Lines(cnt).ImporteLinea  'redondea((Linea(cnt).ImporteLinea / 100))
        'Else
        '    lclImporte = redondea((Linea(cnt).ImporteLinea / (10 ^ decimales)))
        'End If
        If Cabecera.Lines(cnt).LineaCancelada Then
            rst.Fields("IMPORTE_CAN") = rst.Fields("IMPORTE_CAN") + lclImporte 'redondea(rst.Fields("IMPORTE_CAN") + lclImporte)
            rst.Fields("OPERACIONES_CAN") = rst.Fields("OPERACIONES_CAN") + 1
            If Cabecera.Lines(cnt).EsPesado Then
                rst.Fields("PESO_CAN") = rst.Fields("PESO_CAN") + Cabecera.Lines(cnt).cantidad
            Else
                rst.Fields("UNIDADES_CAN") = rst.Fields("UNIDADES_CAN") + Cabecera.Lines(cnt).cantidad
            End If
        Else
            'If Cabecera.Lines(cnt).FactorPeso = -1 Then
            '    rst.Fields("IMPORTE_NEG") = redondea(rst.Fields("IMPORTE_NEG") + lclImporte)
            '    rst.Fields("OPERACIONES_NEG") = rst.Fields("OPERACIONES_NEG") + 1
            '    If Linea(cnt).EsPesado Then
            '        rst.Fields("PESO_NEG") = rst.Fields("PESO_NEG") + Linea(cnt).cantidad
            '    Else
            '        rst.Fields("UNIDADES_NEG") = rst.Fields("UNIDADES_NEG") + Linea(cnt).cantidad
            '    End If
            'Else
                rst.Fields("IMPORTE") = rst.Fields("IMPORTE") + lclImporte 'redondea(rst.Fields("IMPORTE") + lclImporte)
                rst.Fields("OPERACIONES") = rst.Fields("OPERACIONES") + 1
                If Cabecera.Lines(cnt).EsPesado Then
                    rst.Fields("PESO") = rst.Fields("PESO") + Cabecera.Lines(cnt).cantidad
                Else
                    rst.Fields("UNIDADES") = rst.Fields("UNIDADES") + Cabecera.Lines(cnt).cantidad
                End If
            'End If
        End If
    Next cnt
    rst.Update
    rst.Close
    Set rst = Nothing
End Function
'\\\\\\\\\\\
'///////////
Public Sub Actualiza_Tabla_Gtsecs(Base As Database, Cabecera As tipo_cabecera)
'////////////////////////////////
' INCORPORA LOS DATOS DE TOTALES POR SECCIÓN A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////
Dim Bucle As Long
Dim Registro As Recordset
Dim RegFam As Recordset
Dim rstseccion As Recordset
Dim rstarticulo As Recordset
Dim LaSeccion As Integer
Dim LaSeccionMaquina As Integer
Dim LaFamilia As Integer
Dim LaSubseccion As Integer
Dim ElPlu As Integer
Dim LaTara As Long
Dim LaCaducidad As Long
Dim ElDescriptivo As String
Dim ElIVA As Long
Dim ElPrecio As Double
Dim ElImporte As Double
Dim LineaInicial As Boolean

    LineaInicial = True
    '**********************
    ' Determina Sección de gestión
    '*******************
    For Bucle = 1 To Cabecera.LineasMensaje
    '****************************
    ' Si no existe, nuevo registro
    '****************************
        If Cabecera.Lines(Bucle).LineaCancelada = False Then
            'If UsaEuro Then
                ElImporte = Cabecera.Lines(Bucle).ImporteLinea  '/ 100
                ElPrecio = Cabecera.Lines(Bucle).precio  '/ 100
            'Else
            '    ElImporte = Linea(Bucle).ImporteLinea / (10 ^ decimales)
            '    ElPrecio = Linea(Bucle).precio / (10 ^ decimales)
            'End If
            
            'c2f 1.6.12
            'If Dir(App.Path & "\aseccion.txt") <> "" Then
            '        Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
            '        If Not rstseccion.EOF Then
            '            LaSeccion = rstseccion.Fields("codi_ident")
            '        Else
            '            LaSeccion = Cabecera.NMostrador
            '        End If
            '        LaSeccionMaquina = Cabecera.NMostrador
            '        LaSubseccion = 0
            '        LaFamilia = 0
            '        ElPlu = 0
            '        LaTara = 0
            '        LaCaducidad = 0
            '        ElIVA = 0
            '        ElDescriptivo = CargaCadena(515) '"Varios..."
            '
            'Else
            '
            'Datos del artículo de la línea
            Set rstarticulo = Base.OpenRecordset _
            ("select caducidad,tara,codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tipo_iva,des_plu1 from articulo where borrado=false and codigo=" & Cabecera.Lines(Bucle).CodigoArticulo)
            With rstarticulo
                If Not .EOF Then
                    If Not IsNull(.Fields("codi_ident")) Then
                        LaSeccion = .Fields("codi_ident")
                    Else
                        LaSeccion = 0
                    End If
                    If Not IsNull(.Fields("secc_Maqui")) Then
                        LaSeccionMaquina = .Fields("secc_Maqui")
                    Else
                        LaSeccionMaquina = 0
                    End If
                    If Not IsNull(.Fields("codi_sub")) Then
                        LaSubseccion = .Fields("codi_sub")
                    Else
                        LaSubseccion = 0
                    End If
                    If Not IsNull(.Fields("codi_fam")) Then
                        LaFamilia = .Fields("codi_fam")
                    Else
                        LaFamilia = 0
                    End If
                    If Not IsNull(.Fields("plu")) Then
                        ElPlu = .Fields("plu")
                    Else
                        ElPlu = 0
                    End If
                    If Not IsNull(.Fields("tara")) Then
                        LaTara = .Fields("tara")
                    Else
                        LaTara = 0
                    End If
                    If Not IsNull(.Fields("caducidad")) Then
                        LaCaducidad = .Fields("caducidad")
                    Else
                        LaCaducidad = 0
                    End If
                    If Not IsNull(.Fields("tipo_iva")) Then
                        ElIVA = .Fields("tipo_iva")
                    Else
                        ElIVA = 0
                    End If
                    If .Fields("des_plu1") <> "" Then
                        ElDescriptivo = .Fields("des_plu1")
                    Else
                        ElDescriptivo = Space(25)
                    End If
                Else
                    Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
                    If Not rstseccion.EOF Then
                        LaSeccion = rstseccion.Fields("codi_ident")
                    Else
                        LaSeccion = Cabecera.NMostrador
                    End If
                    LaSeccionMaquina = Cabecera.NMostrador
                    LaSubseccion = 0
                    LaFamilia = 0
                    ElPlu = 0
                    LaTara = 0
                    LaCaducidad = 0
                    ElIVA = 0
                    ElDescriptivo = "Varios..." 'CargaCadena(515) '"Varios..."
                End If
            End With
            
            'End If
            
            'If Cabecera.lEncargo Then
            ''("select * from gtsecs where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") _ checa
            '    Set Registro = Base.OpenRecordset _
            '    ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
            '    " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            '    " and codi_fam=" & LaFamilia & " and modo=" & CStr(12))
            'Else
            ''("select * from gtsecs where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") _ checa
                Set Registro = Base.OpenRecordset _
                ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
                " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
            'End If
            
            With Registro
                If .EOF Then
                    .AddNew
                    .Fields("Fecha") = Format(Cabecera.Fecha, "dd/mm/yy")
                    .Fields("codi_ident") = LaSeccion
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codi_fam") = LaFamilia
                    If LaSeccion = 0 Then
                        .Fields("txt_fam") = ""
                    Else
                        Set RegFam = Base.OpenRecordset("select txt_fam from fam_code where codi_ident=" & LaSeccion & " and codi_fam=" & LaFamilia)
                        If Not RegFam.EOF Then
                            .Fields("txt_fam") = RegFam.Fields("txt_fam")
                        Else
                            .Fields("txt_fam") = ""
                        End If
                    End If
                    If LaSeccion = 0 Then
                        .Fields("txt_sub") = ""
                    Else
                        Set RegFam = Base.OpenRecordset("select txt_sub from sub_code where codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion)
                        If Not RegFam.EOF Then
                            .Fields("txt_sub") = RegFam.Fields("txt_sub")
                        Else
                            .Fields("txt_sub") = ""
                        End If
                    End If
                    .Fields("secc_Maqui") = LaSeccionMaquina
                    'If Cabecera.lEncargo Then
                    '    .Fields("Modo") = 12
                    'Else
                        .Fields("Modo") = Cabecera.tipoTiquet
                    'End If
                    .Fields("tiquets") = 0
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Update
                End If
            End With
            '("select * from gtsecs where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") _ Checa
            'If Cabecera.lEncargo Then
            '    Set Registro = Base.OpenRecordset _
            '    ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
            '    " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            '    " and codi_fam=" & LaFamilia & " and modo=" & CStr(12))
            '
            'Else
                
                Set Registro = Base.OpenRecordset _
                ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
                " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
                
                'prueba Mipel
                'Set Registro = Base.OpenRecordset _
                '("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) _
                '& " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
                '"and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
            
            'End If
            With Registro
                
                If Not Registro.EOF Then
                    'Edit_Record Registro
                    .Edit
                End If
                
                If LineaInicial Then
                    .Fields("tiquets") = .Fields("tiquets") + 1
                    LineaInicial = False
                End If
                .Fields("operacion") = .Fields("operacion") + 1
                If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(Bucle).EsPesado Then
                            .Fields("Peso") = .Fields("Peso") + (-1) * Cabecera.Lines(Bucle).cantidad  '(Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + (-1) * Cabecera.Lines(Bucle).cantidad  'Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + (-1) * ElImporte 'redondea((.Fields("pesetas") + (-1) * ElImporte * Linea(Bucle).FactorPeso))
                    End If
                Else
                    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(Bucle).EsPesado Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(Bucle).cantidad  '(Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(Bucle).cantidad  'Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte 'redondea(.Fields("pesetas") + ElImporte * Linea(Bucle).FactorPeso)
                    End If
                End If
                .Update
                .Close
            End With
        End If
    Next Bucle
End Sub
'//////
'//////
'//////
'////////////////////////////////////////////////////////////////////////////////
'INCORPORA LOS DATOS DE TOTALES POR ARTÍCULO A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////////////////////////////////////////////////////
Public Sub Actualiza_Tabla_Gtarti(Base As Database, Cabecera As tipo_cabecera)
'/////////////////////////////
'\ argTipoTqt => tipo de tiquet
'\ 1 --> 386
'\ 2 --> C9/C10
'\ 3 --> SC10
'\ 4 --> IV4TQ
'\\\\\\\\\\\\\
Dim rst As Recordset
Dim rstAux As Recordset
Dim LaSeccion As Integer
Dim LaSeccionMaquina As Integer
Dim LaFamilia As Integer
Dim LaSubseccion As Integer
Dim ElPlu As Integer
Dim LaTara As Long
Dim LaCaducidad As Long
Dim ElDescriptivo As String
Dim ElIVA As Long
Dim Bucle As Long
Dim ElPrecio As Double
Dim ElImporte As Double
Dim MiFecha As Date
Dim sSQL As String
Dim ElCodigo As String
Dim sEan As String
Dim argTipoTqt As Integer

    On Error GoTo finArti
    argTipoTqt = 1
    For Bucle = 1 To Cabecera.LineasMensaje
        If Cabecera.Lines(Bucle).LineaCancelada = False Then
            'If UsaEuro Then
                ElPrecio = Cabecera.Lines(Bucle).precio  '/ 100
                ElImporte = Cabecera.Lines(Bucle).ImporteLinea  '/ 100
            'Else
            '    ElPrecio = Linea(Bucle).precio / (10 ^ decimales)
            '    ElImporte = Linea(Bucle).ImporteLinea / (10 ^ decimales)
            'End If
            'Datos del artículo de la línea
            'If TipoC9 Then
            'If argTipoTqt = 2 Then
            '    sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1 " & _
            '           "FROM articulo " & _
            '           "WHERE borrado=false " & _
            '           " AND plu=" & Linea(Bucle).CodigoArticulo & _
            '           " AND codi_ident=" & Cabecera.NMostrador
            'Else
                sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1,art_cb " & _
                       "FROM articulo " & _
                       "WHERE borrado=false" & _
                       " AND codigo=" & Cabecera.Lines(Bucle).CodigoArticulo
            'End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                If Not .EOF Then
                    If Not IsNull(.Fields("codigo")) Then
                        ElCodigo = .Fields("codigo")
                    Else
                        ElCodigo = 0
                    End If
                    If Not IsNull(.Fields("plu")) Then
                        ElPlu = .Fields("plu")
                    Else
                        ElPlu = 0
                    End If
                    If Not IsNull(.Fields("codi_ident")) Then
                        LaSeccion = .Fields("codi_ident")
                    Else
                        LaSeccion = 0
                    End If
                    If Not IsNull(.Fields("secc_Maqui")) Then
                        LaSeccionMaquina = .Fields("secc_Maqui")
                    Else
                        LaSeccionMaquina = 0
                    End If
                    If Not IsNull(.Fields("codi_sub")) Then
                        LaSubseccion = .Fields("codi_sub")
                    Else
                        LaSubseccion = 0
                    End If
                    If Not IsNull(.Fields("codi_fam")) Then
                        LaFamilia = .Fields("codi_fam")
                    Else
                        LaFamilia = 0
                    End If
                    If Not IsNull(.Fields("tara")) Then
                        LaTara = .Fields("tara")
                    Else
                        LaTara = 0
                    End If
                    If Not IsNull(.Fields("caducidad")) Then
                        LaCaducidad = .Fields("caducidad")
                    Else
                        LaCaducidad = 0
                    End If
                    If Not IsNull(.Fields("tipo_iva")) Then
                        ElIVA = .Fields("tipo_iva")
                    Else
                        ElIVA = 0
                    End If
                    If .Fields("des_plu1") <> "" Then
                        ElDescriptivo = .Fields("des_plu1")
                    Else
                        ElDescriptivo = Space(25)
                    End If
                    If .Fields("art_cb") <> "" Then
                        sEan = .Fields("art_cb")
                           If Len(sEan) < 12 Then
                               sEan = sEan & Mid("000000000000", 1, 12 - Len(sEan))
                           End If
                           sEan = sEan & Format(DigitoControl(sEan), "0")
                        
                    Else
                        sEan = Space(13)
                    End If
                Else
                    sSQL = "SELECT * " & _
                           "FROM seccion " & _
                           "WHERE secc_maqui=" & Cabecera.NMostrador
                    Set rstAux = Base.OpenRecordset(sSQL)
                    If Not rstAux.EOF Then
                        LaSeccion = rstAux.Fields("codi_ident")
                    Else
                        LaSeccion = Cabecera.NMostrador
                    End If
                    rstAux.Close
                    Set rstAux = Nothing
                    LaSeccionMaquina = Cabecera.NMostrador
                    ElCodigo = 0
                    ElPlu = 0
                    LaSubseccion = 0
                    LaFamilia = 0
                    LaTara = 0
                    LaCaducidad = 0
                    ElIVA = 0
                    ElDescriptivo = "Varios..." 'CargaCadena(515) '"Varios..."
                    sEan = Space(13)
                End If
                .Close
            End With
            Set rst = Nothing
            '////////////////
            If Cabecera.Lines(Bucle).CodigoArticulo = 0 Then
                ' "WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _ Checa
                'If Cabecera.lEncargo Then
                '
                'sSQL = "SELECT * " & _
                '       "FROM gtarti " & _
                '       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                '       " AND codigo=" & Format(ElCodigo, "000000") & _
                '       " AND modo=" & CStr(12) & _
                '       " AND secc_maqui=" & LaSeccionMaquina
               '
                'Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & LaSeccionMaquina
                'End If
            Else
                '"WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _ checa
                'If Cabecera.lEncargo Then
                'sSQL = "SELECT * " & _
                '       "FROM gtarti " & _
                '       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                '       " AND codigo=" & Format(ElCodigo, "000000") & _
                '       " AND modo=" & CStr(12)
                '
                'Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
               ' End If
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                If .EOF Then
                    .AddNew
                    
                    .Fields("codigo") = ElCodigo
                    .Fields("plu") = ElPlu
                    .Fields("codi_ident") = LaSeccion
                    .Fields("Fecha") = Cabecera.Fecha
                    .Fields("des_plu1") = ElDescriptivo
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codi_fam") = LaFamilia
                    If ElCodigo = 0 Then
                        .Fields("txt_fam") = ""
                    Else
                        sSQL = "SELECT txt_fam " & _
                               "FROM fam_code " & _
                               "WHERE codi_ident=" & LaSeccion & _
                               " AND codi_fam=" & LaFamilia
                        Set rstAux = Base.OpenRecordset(sSQL)
                        If Not rstAux.EOF Then
                            .Fields("txt_fam") = rstAux.Fields("txt_fam")
                        Else
                            .Fields("txt_fam") = ""
                        End If
                        rstAux.Close
                        Set rstAux = Nothing
                    End If
                    If ElCodigo = 0 Then
                        'If Not lClR Then
                        .Fields("txt_sub") = ""
                        'Else
                        '.Fields("txt_sub") = sEan
                        'End If
                    Else
                        'If Not lClR Then
                        sSQL = "SELECT txt_sub " & _
                               "FROM sub_code " & _
                               "WHERE codi_ident=" & LaSeccion & _
                               " AND codi_sub=" & LaSubseccion
                        Set rstAux = Base.OpenRecordset(sSQL)
                        If Not rstAux.EOF Then
                            .Fields("txt_sub") = rstAux.Fields("txt_sub")
                        Else
                            .Fields("txt_sub") = ""
                        End If
                        rstAux.Close
                        Set rstAux = Nothing
                        'Else
                        '    .Fields("txt_sub") = sEan
                        'End If
                    End If
                    .Fields("secc_Maqui") = LaSeccionMaquina
                    .Fields("precio") = ElPrecio
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    'If Cabecera.lEncargo Then
                    '    .Fields("Modo") = 12
                    'Else
                        .Fields("Modo") = Cabecera.tipoTiquet
                    'End If
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Fields("tara") = LaTara
                    .Fields("caducidad") = LaCaducidad
                    .Fields("tipo_iva") = ElIVA
                    .Fields("Modificado") = Cabecera.Lines(Bucle).SehaModificado
                    .Update
                End If
                .Close
            End With
            Set rst = Nothing
            If ElCodigo = 0 Then
                'If Cabecera.lEncargo Then
                'sSQL = "SELECT * " & _
                '       "FROM gtarti " & _
                '       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                '       " AND codigo=" & Format(ElCodigo, "000000") & _
                '       " AND modo=" & CStr(12) & _
                '       " AND secc_maqui=" & CStr(LaSeccionMaquina)
                '
                'Else
                ''"WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _ checa
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & CStr(LaSeccionMaquina)
                'End If
            Else
                'If Cabecera.lEncargo Then
                'sSQL = "SELECT * " & _
                '       "FROM gtarti " & _
                '       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                '       " AND codigo=" & Format(ElCodigo, "000000") & _
                '       " AND modo=" & CStr(12)
               '
                'Else
                
                ''"WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _ checa
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
                'End If
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                'Edit_Record rst
                .Edit
                .Fields("operacion") = .Fields("operacion") + 1
                If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(Bucle).EsPesado = True Then
                            .Fields("Peso") = .Fields("Peso") + (-1) * ((Cabecera.Lines(Bucle).cantidad) * Cabecera.Lines(Bucle).FactorPeso)
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + (-1) * (Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso)
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + (-1) * (ElImporte * Cabecera.Lines(Bucle).FactorPeso)  'redondea(.Fields("pesetas") + (-1) * (ElImporte * Linea(Bucle).FactorPeso))
                    End If
                Else
                    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(Bucle).EsPesado = True Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(Bucle).cantidad ' * Linea(Bucle).FactorPeso
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(Bucle).cantidad  '* Linea(Bucle).FactorPeso
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte 'redondea(.Fields("pesetas") + ElImporte * Linea(Bucle).FactorPeso)
                    'Else
                        'actborrar = True
                    End If
                End If
                .Update
                'If actborrar Then
                '    .Delete
                '    actborrar = False
                'End If
                .Close
            End With
            Set rst = Nothing
        End If
    Next Bucle
    
finArti:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'CadenadeLog "Actualiza tabla artículos:" & Err.Description
    End If
End Sub
'\\\\\\
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES DE VENDEDORES A PARTIR DE LO RECIBIDO EN LOS TICKETS
'/////////////////////////////////////////////////////////////////////////////////////////
Function actualiza_tabla_gtInfVen(Base As Database, Cabecera As tipo_cabecera)
Dim Encontrado As Boolean
Dim lclSec As Integer
Dim lclPrecio As Double
Dim lclImporte As Double
Dim lclFecha As String
Dim rst As Recordset
Dim cnt As Long
Dim sSQL As String
    sSQL = "SELECT * " & _
           "FROM seccion " & _
           "WHERE secc_maqui=" & Cabecera.NMostrador
    Set rst = Base.OpenRecordset(sSQL)
    If Not rst.EOF Then
        lclSec = rst.Fields("codi_ident")
    Else
        lclSec = Cabecera.NMostrador
    End If
    rst.Close
    Set rst = Nothing
    Encontrado = False
    'If Cabecera.lEncargo Then
    'sSQL = "SELECT * " & _
    '       "FROM gtInfVen " & _
    '       "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
    '       " AND codi_ident=" & Val(lclSec) & _
    '       " AND modo=" & CStr(12) & _
    '       " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34)
    '       '" AND fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" checa
    '
    'Else
    sSQL = "SELECT * " & _
           "FROM gtInfVen " & _
           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
           " AND codi_ident=" & Val(lclSec) & _
           " AND modo=" & Cabecera.tipoTiquet & _
           " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34)
           '" AND fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" checa
    'End If
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        If .EOF Then
            .AddNew
            .Fields("ident_vend") = Cabecera.NVendedor
            .Fields("codi_ident") = lclSec
            .Fields("secc_Maqui") = Cabecera.NMostrador
            .Fields("imporCod") = 0
            .Fields("operCod") = 0
            .Fields("imporDir") = 0
            .Fields("operDir") = 0
            .Fields("imporCan") = 0
            .Fields("operCan") = 0
            .Fields("imporNeg") = 0
            .Fields("operNeg") = 0
            'If Cabecera.lEncargo Then
            '    .Fields("modo") = 12
            'Else
                .Fields("modo") = Cabecera.tipoTiquet
            'End If
            .Fields("fecha") = Cabecera.Fecha
            .Update
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        For cnt = 1 To Cabecera.LineasMensaje
            'If UsaEuro Then
                lclImporte = Cabecera.Lines(cnt).ImporteLinea  'redondea((Linea(cnt).ImporteLinea / 100))
            'Else
            '    lclImporte = redondea((Linea(cnt).ImporteLinea / (10 ^ decimales)))
            'End If
            'Edit_Record rst
            .Edit
            '//////////////
            'Cancelada o no
            '//////////////
            If Cabecera.Lines(cnt).LineaCancelada Then
                .Fields("imporCan") = .Fields("imporCan") + lclImporte 'redondea(.Fields("imporCan") + lclImporte)
                .Fields("operCan") = .Fields("operCan") + 1
            Else
                'If Linea(cnt).FactorPeso = -1 Then
                '    .Fields("imporNeg") = redondea(.Fields("imporNeg") + lclImporte)
                '    .Fields("operNeg") = .Fields("operNeg") + 1
                'Else
                    If Cabecera.Lines(cnt).CodigoArticulo = 0 Then
                        .Fields("imporDir") = .Fields("imporDir") + lclImporte 'redondea(.Fields("imporDir") + lclImporte)
                        .Fields("operDir") = .Fields("operDir") + 1
                    Else
                        .Fields("imporCod") = .Fields("imporCod") + lclImporte 'redondea(.Fields("imporCod") + lclImporte)
                        .Fields("operCod") = .Fields("operCod") + 1
                    End If
                'End If
            End If
            .Update
        Next cnt
        .Close
    End With
    Set rst = Nothing
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' 1º TABLA TICKETS (datos de cada línea recibida)
'////////////////////////////////////////////////
' Tipos de Tiquet (typtic)
' 1 --> Rollo
' 2 --> Super
' 3 --> Envasado
' 4 --> Autoservicio
' 5 --> Rollo Local
' 6 --> Super Local
' 7 --> Envasado Local
' 8 --> Autoservicio Local
'/////////////////////////
Public Sub Actualiza_Tabla_Tickets(Base As Database, Cabecera As tipo_cabecera)
Dim BucleGrande As Integer
Dim D_FECHA As Date
Dim D_HORA As Date
Dim Maximo As Integer
Dim Bucle As Long
Dim MiFactor As Double
Dim Registro As DAO.Recordset
Dim Buffer As String
Dim MiCadenaDebug As String
Dim RegArt As DAO.Recordset
Dim Que_base As String
Dim nMin As Integer

On Error GoTo finTick

nMin = 1

    Set RegArt = Base.OpenRecordset("select codigo,plu,des_plu1,codi_fam from articulo where borrado=false")
    Buffer = Cabecera.CadenaEstadoTiquet
    
    Que_base = "GA"
    
    If Que_base = "GA" Then
        MiCadenaDebug = "Euroscale / "
        Set Registro = Base.OpenRecordset("select * from tickets where cstr(cdate(date))=" _
        & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
    End If
    If Que_base = "SC10" Then
        MiCadenaDebug = "SC10 / "
        Set Registro = Base.OpenRecordset("select * from ticketsSC10 where cstr(cdate(date))=" _
        & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
    End If
    For Bucle = 1 To Cabecera.LineasMensaje
        Buffer = Cabecera.CadenaEstadoTiquet
        'If Que_base = "GA" Then
            If (Cabecera.Lines(Bucle).EsPesado) Then
                Buffer = Buffer & "peso "
            Else
                Buffer = Buffer & "unidades "
            End If
        'End If
        Maximo = 1
        nMin = 1
        
        'lYaDescontado = False
        
        For BucleGrande = nMin To Maximo
            With Registro
                .AddNew
                'If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Linea de Tiquet : " _
                '& Linea(Bucle).NumeroLinea & "Tiquet : " & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                '& "Importe : " & Linea(Bucle).ImporteLinea & Linea(Bucle).FactorPeso)
                .Fields("Date") = Format(Cabecera.Fecha, "dd/mm/yy")
                .Fields("Hour") = Format(Cabecera.hora, "hh:mm:ss")
                .Fields("d_fecha") = Cabecera.Fecha
                .Fields("d_hora") = Cabecera.hora
                .Fields("numbal") = Cabecera.NBalanza
                .Fields("Vendor") = Cabecera.NVendedor
                .Fields("IVA") = Cabecera.Lines(Bucle).IVA
                
                .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
                .Fields("numlin") = Cabecera.Lines(Bucle).NumeroLinea
                .Fields("SECC") = Cabecera.NMostrador
                .Fields("Code") = Cabecera.Lines(Bucle).CodigoArticulo
                RegArt.FindFirst "codigo=" & Cabecera.Lines(Bucle).CodigoArticulo
                If RegArt.EOF Or RegArt.NoMatch Then
                    .Fields("plu") = 0
                    .Fields("Item") = ""
                    'c2f invicta
                    .Fields("DPT") = "0000"
                Else
                    .Fields("plu") = RegArt.Fields("plu")
                    .Fields("Item") = RegArt.Fields("des_plu1")
                    'c2f invicta
                    .Fields("DPT") = Format(RegArt.Fields("CODI_FAM"), "0000")
                End If
                'If UsaEuro Then
                    .Fields("Price") = (Cabecera.Lines(Bucle).precio) ' / 100)
                'Else
                '    .Fields("Price") = (Cabecera.Lines(Bucle).precio / (10 ^ decimales))
                'End If
                If BucleGrande = 2 Then 'Or (Dir(App.Path & "\upper.txt") <> 0 And Linea(Bucle).LineaCancelada = True) Then
                    MiFactor = (-1)
                Else
                    MiFactor = 1
                End If
                If ((Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128) Then
                    'If UsaEuro Then
                        .Fields("amount") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso)) ' / 100)
                    'Else
                    '    .Fields("amount") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso) / (10 ^ decimales))
                    'End If
                    If Cabecera.Lines(Bucle).EsPesado Then
                        .Fields("Weight") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                        .Fields("units") = 0
                    Else
                        .Fields("Weight") = 0
                        .Fields("units") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                    End If
                Else
                    'If UsaEuro Then
                        .Fields("amount") = MiFactor * (Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso) ' / 100
                    'Else
                    '    .Fields("amount") = MiFactor * (Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso) / (10 ^ decimales)
                    'End If
                    If Cabecera.Lines(Bucle).EsPesado Then
                        .Fields("Weight") = MiFactor * (Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso)
                        .Fields("units") = 0
                    Else
                        .Fields("Weight") = 0
                        .Fields("units") = MiFactor * (Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso)
                    End If
                End If
                .Fields("STPETIC") = Cabecera.EstadoTiquet
                If Cabecera.Lines(Bucle).LineaCancelada = True Then
                   .Fields("lincan") = 1
                Else
                   .Fields("lincan") = 0
                End If
                
                .Fields("TYPTIC") = Cabecera.tipoTiquet
                .Fields("reserv1") = Cabecera.Lines(Bucle).FichaVacuno
                
                .Fields("importe_bruto") = .Fields("amount") + Cabecera.Lines(Bucle).DescuentoLinea

                If Cabecera.lote <> "" Then
                    .Fields("numlote") = Cabecera.lote
                Else
                    .Fields("numlote") = " "
                End If
                If Not IsNull(Cabecera.NCliente) Then
                    .Fields("cliente") = Cabecera.NCliente
                End If
                .Fields("reserv2") = "/"
                .Fields("reserv3") = "NO"
                
                
                .Update
                
                
            End With
            
        Next BucleGrande
    Next Bucle
    
finTick:
    
    'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
    '    Cabecera.Incluir_en_Base = False
    'End If
    
    'If Err.Number <> 0 Then
    '    On Error GoTo 0
    '    CadenadeLog "Actualiza tabla Tickets:" & Err.Description
    'End If
End Sub



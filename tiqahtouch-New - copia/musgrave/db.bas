Attribute VB_Name = "db"
Option Explicit
Public Type DB_Articulo
    codigo As Long
    Mostrador As Integer
    Plu As Integer
    subsec As Integer
    familia As Integer
    precio As Double
    caducidad As Integer
    tara As Double
    WGH As Boolean
    Balenv As Boolean
    GrpConserv As Integer
    Descriptivos(21) As String
    TipoLetra(10) As Integer
    Preferente As Integer
    Etiqueta As Integer
    ean13 As String
    TRM1 As Double
    TRM2 As Double
    PRC1 As Double
    PRC2 As Double
    IVA As Integer
    Merma As Integer
    Presel As Integer
    PRC100G As Boolean ' EcoLabel, precio por 100 gramos
    
    descriptivos_2040(20) As String 'cas->v171
    lMix As Boolean
    nPoid As Integer
    
    PRC3 As Double
    Imagen As String
    posicion As Integer
End Type

Public Sub Actualiza_Tabla_Cabecera(Base As dao.Database, Cabecera As tipo_cabecera)
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
Dim regVendedor As dao.Recordset
Dim Registro As dao.Recordset
Dim RegSeccion As dao.Recordset
Dim RegDes As dao.Recordset
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
Dim nNumCorre As Double
Dim sFactura As String
'''''''
Dim PrecioSinIVA As Double
Dim PrecioConIVA As Double
Dim BufferIVA As String
Dim BufferIVA2 As String
Dim MiIva(5) As Tipo_Grupo_IVA
Dim BucleIVA As Integer
Dim RegIVA As dao.Recordset
Dim RegArt As dao.Recordset
Dim bucle As Integer
''''''''''''''''

    Que_base = "GA"
    
    'Registro de articulos
    Set RegArt = Base.OpenRecordset("select codigo,tipo_iva from articulo")
    For bucle = 1 To Cabecera.LineasMensaje
        RegArt.FindFirst "codigo=" & Cabecera.Lines(bucle).CodigoArticulo
        If RegArt.EOF Or RegArt.NoMatch Then
            Cabecera.Lines(bucle).IVA = 0
        Else
            Cabecera.Lines(bucle).IVA = RegArt!tipo_iva
        End If
    Next bucle
    RegArt.Close
    Set RegArt = Nothing
    
    Set RegIVA = Base.OpenRecordset("select * from iva")
    'carga porcentajes en lineas
    For bucle = 1 To Cabecera.LineasMensaje
        RegIVA.FindFirst "tipo=" & Cabecera.Lines(bucle).IVA
        If RegIVA.EOF Or RegIVA.NoMatch Then
            Cabecera.Lines(bucle).porcentaje = 0
        Else
            Cabecera.Lines(bucle).porcentaje = RegIVA!Valor
        End If
    Next bucle
    '********************
    ' carga tipos de IVA
    '********************
    For bucle = 0 To 4
        MiIva(bucle).Base = 0
        MiIva(bucle).Importe = 0
        MiIva(bucle).porcentaje = 0
        RegIVA.FindFirst "tipo=" & (bucle)
        If RegIVA.EOF Or RegIVA.NoMatch Then
            MiIva(bucle).porcentaje = 0
        Else
            MiIva(bucle).porcentaje = RegIVA!Valor
        End If
    Next bucle
    RegIVA.Close
    Set RegIVA = Nothing
    
    'If Dir(App.Path & "\automotor.txt") = "" Then
        'Cálculos Iva
        If Cabecera.Trainning = False And ((Cabecera.EstadoTiquet And 64) <> 64 And (Cabecera.EstadoTiquet And 128) <> 128) Then
            For bucle = 1 To Cabecera.LineasMensaje
                If Cabecera.Lines(bucle).LineaCancelada <> True Then
                    If Not UsaEuro Then
                            If (Cabecera.Lines(bucle).TipoLinea > 1) Then
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe - (Cabecera.Lines(bucle).ImporteLinea)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base - (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            Else
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe + (Cabecera.Lines(bucle).ImporteLinea)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            End If
                    
                    Else
                            If (Cabecera.Lines(bucle).TipoLinea > 1) Then
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe - (Cabecera.Lines(bucle).ImporteLinea)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base - (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            Else
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe + (Cabecera.Lines(bucle).ImporteLinea)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            End If
                    End If
                End If
            Next bucle
        End If
    'End If
    
    '***************************
    ' calcula el código de barras
    '****************************
    tqcodbar.Importe = Cabecera.ImporteTotal
    tqcodbar.NBalanza = Cabecera.NBalanza
    tqcodbar.NMostrador = Cabecera.NMostrador
    tqcodbar.NTicket = Cabecera.Ntiquet
    tqcodbar.NVendedor = Cabecera.NVendedor
    'el modo trainning se va a considerar tipoPago = 12
    ''''''''''''''''''''''''''''''''''''''''
    tqcodbar.Tipo = Cabecera.tipoTiquet
    If Cabecera.nLineas >= 1 Then
        Cod_Art = Cabecera.Lines(1).CodigoArticulo
    Else
        Cod_Art = 0
    End If
    micodbar = Genera_CodBar(tqcodbar, Cod_Art)
    If Trim(Cabecera.ean13) <> "" Then
        micodbar = Trim(Cabecera.ean13)
    End If
    If Que_base = "GA" Then
        MiCadenaDebug = "K-T/Scale / "
        Set Registro = Base.OpenRecordset("select * from cabecera where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & Cabecera.NVendedor & Chr(34) & " and onoff=" & Chr(34) & Cabecera.CadenaEstadoTiquet & Chr(34))
    End If
    With Registro
        
        If .EOF Then

            If DEBUGACTIVO Then CadenadeLog (MiCadenaDebug & "Almacenando Cabecera de Tiquet : " _
            & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
            & "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.fecha & "Hora : " & Cabecera.hora)
            
            .AddNew
            .Fields("factura") = Cabecera.factura
            .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
            .Fields("seccion") = Cabecera.NMostrador
            .Fields("codbar") = micodbar
            .Fields("Balanza") = Cabecera.NBalanza
            .Fields("vended") = Cabecera.NVendedor
            .Fields("numlin") = Cabecera.nLineas
            .Fields("lincan") = Cabecera.LineasCanceladas
            .Fields("Importe") = Cabecera.ImporteTotal '/ 100
            bimpoBrutoCab = Cabecera.ImporteBruto 'redondea(Cabecera.ImporteBruto  / 100) 'c2f invicta
                
            'c2f invicta
            On Error Resume Next
            .Fields("importe_bruto") = bimpoBrutoCab
            ''''''''''''
            On Error GoTo 0
            ''''''''''''
            .Fields("hora") = Format(Cabecera.hora, "hh:mm:ss")
            .Fields("d_fecha") = Cabecera.fecha
            .Fields("d_hora") = Cabecera.hora
            .Fields("onoff") = Cabecera.CadenaEstadoTiquet
            .Fields("STPETIC") = Cabecera.EstadoTiquet
            .Fields("codcli") = Cabecera.NCliente
            .Fields("Fecha") = Format(Cabecera.fecha, "dd/mm/yy")
            .Fields("TYPTIC") = Cabecera.tipoTiquet
            .Fields("reserv2") = "/"
            .Fields("reserv3") = "NO"  'no exportado
            .Fields("tipoPago") = Cabecera.tipoPago
            .Fields("cantidadEntre") = Cabecera.cantidadEntre '/ 100
            .Fields("cambio") = Cabecera.Cambio '/ 100
            On Error Resume Next
            If Cabecera.lote <> "" Then
                .Fields("numlote") = Cabecera.lote
            Else
                .Fields("numlote") = " "
    
            End If
            
            '.Fields.Append .CreateField("abonado", dbBoolean)
            '.Fields.Append .CreateField("ticketabono", dbLong)
            '.Fields.Append .CreateField("ticketorigen", dbLong)
            '.Fields.Append .CreateField("abono", dbBoolean)
            If Cabecera.documentoabono <> 0 Then
                If Cabecera.documentoabono > 0 Then
                    .Fields("abonado") = True
                    .Fields("ticketabono") = Cabecera.documentoabono
                    .Fields("ticketorigen") = 0
                    .Fields("abono") = False
                Else
                    .Fields("abonado") = False
                    .Fields("ticketabono") = 0
                    .Fields("ticketorigen") = Abs(Cabecera.documentoabono)
                    .Fields("abono") = True
                
                End If
            Else
                    .Fields("abonado") = False
                    .Fields("ticketabono") = 0
                    .Fields("ticketorigen") = 0
                    .Fields("abono") = False
            
            End If
            .Fields("descuento") = Cabecera.descuento
            
            '''''''''''
            .Fields("base0") = MiIva(0).Base
            .Fields("por0") = MiIva(0).porcentaje
            .Fields("imp0") = Round(MiIva(0).Importe - MiIva(0).Base, 2)
            .Fields("base1") = MiIva(1).Base
            .Fields("por1") = MiIva(1).porcentaje
            .Fields("imp1") = Round(MiIva(1).Importe - MiIva(1).Base, 2)
            .Fields("base2") = MiIva(2).Base
            .Fields("por2") = MiIva(2).porcentaje
            .Fields("imp2") = Round(MiIva(2).Importe - MiIva(2).Base, 2)
            .Fields("base3") = MiIva(3).Base
            .Fields("por3") = MiIva(3).porcentaje
            .Fields("imp3") = Round(MiIva(3).Importe - MiIva(3).Base, 2)
            .Fields("base4") = MiIva(4).Base
            .Fields("por4") = MiIva(4).porcentaje
            .Fields("imp4") = Round(MiIva(4).Importe - MiIva(4).Base, 2)
            '''''''''''
            
            On Error GoTo 0
            .Update
            Cabecera.Incluir_en_Base = True
            
            '\\\\\\Descuentos a Tabla Descuentos...
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If (Cabecera.descuento > 0) Then
                sSQL = "SELECT * " & _
                       "FROM descuentos " & _
                       " WHERE fecha=" & "#" & Format(Cabecera.fecha, "mm,dd,yy") & "#" & _
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
                    .Fields("fecha") = Format(Cabecera.fecha, "dd/mm/yyyy")
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
Public Sub Actualiza_Tabla_GtVend(Base As dao.Database, Cabecera As tipo_cabecera)
    '////////////////////////////////
    ' INCORPORA LOS DATOS DE TOTALES
    ' POR VENDEDOR A PARTIR DE
    ' LO RECIBIDO EN UN TICKET
    '////////////////////////////////
    Dim eNCONTRADO As Boolean
    Dim bucle As Long
    Dim ElPrecio As Double
    Dim ElImporte As Double
    Dim Lafecha As String
    Dim LaSeccion As Integer
    Dim Registro As dao.Recordset
    Dim rstseccion As dao.Recordset
    Dim rstvendedor As dao.Recordset
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
    eNCONTRADO = False
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34))
    With Registro
        If .EOF Then
            .AddNew
            .Fields("Fecha") = Cabecera.fecha
            .Fields("codi_ident") = LaSeccion
            .Fields("ident_vend") = Cabecera.NVendedor
            .Fields("secc_Maqui") = Cabecera.NMostrador
            .Fields("Modo") = Cabecera.tipoTiquet
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
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34))
    With Registro
        If LineaInicial Then
            .Edit
            .Fields("tiquets") = .Fields("tiquets") + 1
            .Fields("operacion") = .Fields("operacion") + Cabecera.nLineas
            .Update
        End If
        For bucle = 1 To Cabecera.LineasMensaje
        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
            .Edit
            ElImporte = Cabecera.Lines(bucle).ImporteLinea '/ 100
            ElPrecio = Cabecera.Lines(bucle).precio '/ 100
            ' ****
            'pesado o unidades
            ' ****
            If Cabecera.Lines(bucle).LineaCancelada = False Then
                If Cabecera.Lines(bucle).EsPesado = False Then
                    .Fields("unidad") = .Fields("unidad") + Cabecera.Lines(bucle).cantidad
                    If Cabecera.Lines(bucle).ImporteLinea < 0 Then
                        .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                    End If
                Else
                    .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                    If Cabecera.Lines(bucle).ImporteLinea < 0 Then
                        .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                    End If
                End If
                .Fields("pesetas") = .Fields("pesetas") + ElImporte
            End If
            ' ****
            ' Cancelada o no
            ' ****
            If Cabecera.Lines(bucle).LineaCancelada Then
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    .Fields("totcance") = .Fields("totcance") - ElImporte
                '    .Fields("supertot") = .Fields("supertot") - ElImporte
                '    .Fields("cancela") = .Fields("cancela") - 1
                'Else
                    .Fields("totcance") = .Fields("totcance") + ElImporte
                    .Fields("supertot") = .Fields("supertot") + ElImporte
                    .Fields("cancela") = .Fields("cancela") + 1
                'End If
            Else
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    .Fields("Total") = .Fields("Total") + ElImporte
                '    .Fields("supertot") = .Fields("Total") + ElImporte
                'Else
                    .Fields("Total") = .Fields("Total") + ElImporte
                    .Fields("supertot") = .Fields("Total") + ElImporte
                'End If
            End If
            .Update
        End If
        Next bucle
        
        
    End With
    
finVend:
    If Err.Number <> 0 Then
        On Error GoTo 0
        CadenadeLog "Actualiza tabla vendedores:" & Err.Description
    End If
    
End Sub
'//////
'//////
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES HORARIOS A PARTIR DE LO RECIBIDO EN EL TICKET
'//////////////////////////////////////////////////////////////////////////////////
Function actualiza_tabla_gtHora(Base As dao.Database, Cabecera As tipo_cabecera)
    Dim rst As dao.Recordset
    Dim cnt As Long
    Dim sSQL As String
    Dim lclImporte As Double
    Dim LineaInicial As Boolean
    sSQL = "SELECT * " & _
           "FROM gthora " & _
           "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
           " AND cstr(hour(hora))=" & Chr(34) & CStr(Hour(Cabecera.hora)) & Chr(34)
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        If .EOF Then
            .AddNew
            .Fields("FECHA") = Cabecera.fecha
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
    rst.Edit
    rst.Fields("TIQUET") = rst.Fields("TIQUET") + 1
    For cnt = 1 To Cabecera.LineasMensaje
    If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
        lclImporte = Cabecera.Lines(cnt).ImporteLinea
        
        If Cabecera.Lines(cnt).LineaCancelada Then
            rst.Fields("IMPORTE_CAN") = rst.Fields("IMPORTE_CAN") + lclImporte
            rst.Fields("OPERACIONES_CAN") = rst.Fields("OPERACIONES_CAN") + 1
            If Cabecera.Lines(cnt).EsPesado Then
                rst.Fields("PESO_CAN") = rst.Fields("PESO_CAN") + Cabecera.Lines(cnt).cantidad
            Else
                rst.Fields("UNIDADES_CAN") = rst.Fields("UNIDADES_CAN") + Cabecera.Lines(cnt).cantidad
            End If
        Else
            rst.Fields("IMPORTE") = rst.Fields("IMPORTE") + lclImporte 'redondea(rst.Fields("IMPORTE") + lclImporte)
            rst.Fields("OPERACIONES") = rst.Fields("OPERACIONES") + 1
            If Cabecera.Lines(cnt).EsPesado Then
                rst.Fields("PESO") = rst.Fields("PESO") + Cabecera.Lines(cnt).cantidad
            Else
                rst.Fields("UNIDADES") = rst.Fields("UNIDADES") + Cabecera.Lines(cnt).cantidad
            End If
        End If
    End If
    Next cnt
    rst.Update
    rst.Close
    Set rst = Nothing
End Function
'\\\\\\\\\\\
'///////////
Public Sub Actualiza_Tabla_Gtsecs(Base As dao.Database, Cabecera As tipo_cabecera)
'////////////////////////////////
' INCORPORA LOS DATOS DE TOTALES POR SECCIÓN A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////
Dim bucle As Long
Dim Registro As dao.Recordset
Dim RegFam As dao.Recordset
Dim rstseccion As dao.Recordset
Dim rstarticulo As dao.Recordset
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
    For bucle = 1 To Cabecera.LineasMensaje
    '****************************
    ' Si no existe, nuevo registro
    '****************************
        If Cabecera.Lines(bucle).LineaCancelada = False And (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
            ElImporte = Cabecera.Lines(bucle).ImporteLinea
            ElPrecio = Cabecera.Lines(bucle).precio
            Set rstarticulo = Base.OpenRecordset _
            ("select caducidad,tara,codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tipo_iva,des_plu1 from articulo where borrado=false and codigo=" & Cabecera.Lines(bucle).CodigoArticulo)
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
                    ElDescriptivo = "Varios..."
                End If
            End With
            
            Set Registro = Base.OpenRecordset _
            ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
            " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
            
            With Registro
                If .EOF Then
                    .AddNew
                    .Fields("Fecha") = Format(Cabecera.fecha, "dd/mm/yy")
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
                    .Fields("Modo") = Cabecera.tipoTiquet
                    .Fields("tiquets") = 0
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Update
                End If
            End With
                
            Set Registro = Base.OpenRecordset _
            ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
            " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
                
            With Registro
                
                If Not Registro.EOF Then
                    .Edit
                End If
                
                If LineaInicial Then
                    .Fields("tiquets") = .Fields("tiquets") + 1
                    LineaInicial = False
                End If
                .Fields("operacion") = .Fields("operacion") + 1
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                '        If Cabecera.Lines(Bucle).EsPesado Then
                '            .Fields("Peso") = .Fields("Peso") + (-1) * Cabecera.Lines(Bucle).cantidad
                '        Else
                '            .Fields("Unidades") = .Fields("Unidades") + (-1) * Cabecera.Lines(Bucle).cantidad
                '        End If
                '        .Fields("pesetas") = .Fields("pesetas") + (-1) * ElImporte
                '    End If
                'Else
                    'If Cabecera.Lines(Bucle).LineaCancelada = False Then '12/07/11
                        If Cabecera.Lines(bucle).EsPesado Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(bucle).cantidad
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte
                    'End If
                'End If
                .Update
                .Close
            End With
        End If
    Next bucle
End Sub
'//////
'////////////////////////////////////////////////////////////////////////////////
'INCORPORA LOS DATOS DE TOTALES POR ARTÍCULO A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////////////////////////////////////////////////////
Public Sub Actualiza_Tabla_Gtarti(Base As dao.Database, Cabecera As tipo_cabecera)
'/////////////////////////////
'\ argTipoTqt => tipo de tiquet
'\ 1 --> 386
'\ 2 --> C9/C10
'\ 3 --> SC10
'\ 4 --> IV4TQ
'\\\\\\\\\\\\\
Dim rst As dao.Recordset
Dim rstAux As dao.Recordset
Dim LaSeccion As Integer
Dim LaSeccionMaquina As Integer
Dim LaFamilia As Integer
Dim LaSubseccion As Integer
Dim ElPlu As Integer
Dim LaTara As Long
Dim LaCaducidad As Long
Dim ElDescriptivo As String
Dim ElIVA As Long
Dim bucle As Long
Dim ElPrecio As Double
Dim ElImporte As Double
Dim MiFecha As Date
Dim sSQL As String
Dim ElCodigo As String
Dim sEan As String
Dim argTipoTqt As Integer

    On Error GoTo finArti
    argTipoTqt = 1
    For bucle = 1 To Cabecera.LineasMensaje
        If Cabecera.Lines(bucle).LineaCancelada = False And (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
            ElPrecio = Cabecera.Lines(bucle).precio
            ElImporte = Cabecera.Lines(bucle).ImporteLinea
            sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1,art_cb " & _
                   "FROM articulo " & _
                   "WHERE borrado=false" & _
                   " AND codigo=" & Cabecera.Lines(bucle).CodigoArticulo
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
                           If Len(sEan) <> 13 Then
                            sEan = sEan & Format(DigitoControl(sEan), "0")
                           End If
                        
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
                    ElDescriptivo = "Varios..."
                    sEan = Space(13)
                End If
                .Close
            End With
            Set rst = Nothing
            '////////////////
            If Cabecera.Lines(bucle).CodigoArticulo = 0 Then
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & LaSeccionMaquina

            Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                If .EOF Then
                    .AddNew
                    
                    .Fields("codigo") = ElCodigo
                    .Fields("plu") = ElPlu
                    .Fields("codi_ident") = LaSeccion
                    .Fields("Fecha") = Cabecera.fecha
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
                    '1.0.3 Leclerc
                    'If ElCodigo = 0 Then
                    '    .Fields("txt_sub") = ""
                    'Else
                    '    sSQL = "SELECT txt_sub " & _
                    '           "FROM sub_code " & _
                    '           "WHERE codi_ident=" & LaSeccion & _
                    '           " AND codi_sub=" & LaSubseccion
                    '    Set rstAux = Base.OpenRecordset(sSQL)
                    '    If Not rstAux.EOF Then
                    '        .Fields("txt_sub") = rstAux.Fields("txt_sub")
                    '    Else
                    '        .Fields("txt_sub") = ""
                    '    End If
                    '    rstAux.Close
                    '    Set rstAux = Nothing
                    'End If
                    '''''''''''''''''''''''''
                    If ElCodigo = 0 Then
                        'If Not lClR Then
                            .Fields("txt_sub") = ""
                        'Else
                        '    .Fields("txt_sub") = sEan
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
                    ''''''''''''''''''''''''''''''
                    .Fields("secc_Maqui") = LaSeccionMaquina
                    .Fields("precio") = ElPrecio
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    .Fields("Modo") = Cabecera.tipoTiquet
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Fields("tara") = LaTara
                    .Fields("caducidad") = LaCaducidad
                    .Fields("tipo_iva") = ElIVA
                    .Fields("Modificado") = Cabecera.Lines(bucle).SehaModificado
                    .Update
                End If
                .Close
            End With
            Set rst = Nothing
            If ElCodigo = 0 Then
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & CStr(LaSeccionMaquina)
            Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                .Edit
                .Fields("operacion") = .Fields("operacion") + 1
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                '        If Cabecera.Lines(Bucle).EsPesado = True Then
                '            .Fields("Peso") = .Fields("Peso") + (-1) * ((Cabecera.Lines(Bucle).cantidad) * Cabecera.Lines(Bucle).FactorPeso)
                '        Else
                '            .Fields("Unidades") = .Fields("Unidades") + (-1) * (Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso)
                '        End If
                '        .Fields("pesetas") = .Fields("pesetas") + (-1) * (ElImporte * Cabecera.Lines(Bucle).FactorPeso)
                '    End If
                'Else
                    'If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(bucle).EsPesado = True Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(bucle).cantidad
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte
                    'End If
                'End If
                .Update
                .Close
            End With
            Set rst = Nothing
        End If
    Next bucle
    
finArti:
    If Err.Number <> 0 Then
        On Error GoTo 0
        CadenadeLog "Actualiza tabla artículos:" & Err.Description
    End If
End Sub
'\\\\\\
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES DE VENDEDORES A PARTIR DE LO RECIBIDO EN LOS TICKETS
'/////////////////////////////////////////////////////////////////////////////////////////
'Function actualiza_tabla_gtInfVen(Base As DAO.Database, Cabecera As tipo_cabecera)
'Dim eNCONTRADO As Boolean
'Dim lclSec As Integer
'Dim lclPrecio As Double
'Dim lclImporte As Double
'Dim lclFecha As String
'Dim rst As DAO.Recordset
'Dim cnt As Long
'Dim sSQL As String
'    sSQL = "SELECT * " & _
'           "FROM seccion " & _
'           "WHERE secc_maqui=" & Cabecera.NMostrador
'    Set rst = Base.OpenRecordset(sSQL)
'    If Not rst.EOF Then
'        lclSec = rst.Fields("codi_ident")
'    Else
'        lclSec = Cabecera.NMostrador
'    End If
'    rst.Close
'    Set rst = Nothing
'    eNCONTRADO = False
'    sSQL = "SELECT * " & _
'           "FROM gtInfVen " & _
'           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
'           " AND codi_ident=" & Val(lclSec) & _
'           " AND modo=" & Cabecera.tipoTiquet & _
'           " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34)
'    Set rst = Base.OpenRecordset(sSQL)
'    With rst
'        If .EOF Then
'            .AddNew
'            .Fields("ident_vend") = Cabecera.NVendedor
'            .Fields("codi_ident") = lclSec
'            .Fields("secc_Maqui") = Cabecera.NMostrador
'            .Fields("imporCod") = 0
'            .Fields("operCod") = 0
'            .Fields("imporDir") = 0
'            .Fields("operDir") = 0
'            .Fields("imporCan") = 0
'            .Fields("operCan") = 0
'            .Fields("imporNeg") = 0
'            .Fields("operNeg") = 0
'            .Fields("modo") = Cabecera.tipoTiquet
'            .Fields("fecha") = Cabecera.Fecha
'            .Update
'        End If
'        .Close
'    End With
'    Set rst = Nothing
'    Set rst = Base.OpenRecordset(sSQL)
'    With rst
'        For cnt = 1 To Cabecera.LineasMensaje
'        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
'            lclImporte = Cabecera.Lines(cnt).ImporteLinea
'            .Edit
'            '//////////////
'            'Cancelada o no
'            '//////////////
'            If Cabecera.Lines(cnt).LineaCancelada Then
'                .Fields("imporCan") = .Fields("imporCan") + lclImporte
'                .Fields("operCan") = .Fields("operCan") + 1
'            Else
'                If Cabecera.Lines(cnt).CodigoArticulo = 0 Then
'                    .Fields("imporDir") = .Fields("imporDir") + lclImporte
'                    .Fields("operDir") = .Fields("operDir") + 1
'                Else
'                    .Fields("imporCod") = .Fields("imporCod") + lclImporte
'                    .Fields("operCod") = .Fields("operCod") + 1
'                End If
'            End If
'            .Update
'        End If
'        Next cnt
'        .Close
'    End With
'    Set rst = Nothing
'End Function
'Balanza (2): Siempre 0#
'GRUPO (2): Siempre 32#
'CLIENTES(5): número de tickets
'IMPORTE_CLIENTES(7): importe de los tickets (suma total) con las dos ultimas posiciones como decimales.
'CLIENTES_MAYORISTAS (5): Siempre 0
'IMPORTE_MAYORISTAS (7): Siempre 0
'BORRADO (1): Siempre 1#
'ORDEN(4): numero de listado (este es el dato realmente importante).
'TOT_CLIENTES(5): numero total de cliente, igual que el campo CLIENTES.
'SIGNO: Siempre 1#
'TOT_IMPORTE(7): importe de los clientes, igual que el campo IMPORTE_CLIENTE
'CLI_MAYORISTA (5): Siempre 0
'SIGNO_MAYORISTA (1): Siempre 0
'IMPORTE_MAYORISTA (7): Siempre 0
'CLI_NEGATIVOS (5): Siempre 0
'SIGNO_NEGATIVOS (1): Siempre 0
'IMPORTE_NEGATIVOS (7): Siempre 0
'PRO_PESADOS(5): número de productos vendido por peso.
'SIGNO_PESADOS (1): Siempre 0
'IMPORTE_PESADOS(7): Importe de los artículos vendidos por peso.
'PRO_NO_PESADOS(5): número de productos vendido por unidades.
'SIGNO_MAYORISTA (1): Siempre 0
'IMPORTE_NO_PESADOS(7): Importe de los artículos vendidos por unidades
'PRO_DEVUELTO (5): Siempre 0
'SIGNO_DEVUELTO (1): Siempre 0
'IMPORTE_DEVUELTO (7): Siempre 0
'HORA(5): hora en que se genera el fichero (FORMATO HH:MM)
'FECHA(10): Fecha en que se genera el fichero (FORMATO DD-MM-AAAA)
'INTERNO(1): vacio.
'
'En el  ejemplo que te mando seria:
'BALANZA(2): 00
'GRUPO(2): 32
'CLIENTES(5): 00148
'IMPORTE_CLIENTES(7): 0057310 à siendo pues 573,10€.
'CLIENTES_MAYORISTAS(5): 00000
'IMPORTE_MAYORISTAS(7): 0000000
'BORRADO(1): 1
'ORDEN(4): 247
'TOT_CLIENTES(5): 00148
'SIGNO: 1
'TOT_IMPORTE(7): 0057310
'CLI_MAYORISTA(5): 00000
'SIGNO_MAYORISTA(1): 0
'IMPORTE_MAYORISTA(7): 0000000
'CLI_NEGATIVOS(5): 00000
'SIGNO_NEGATIVOS(1): 0
'IMPORTE_NEGATIVOS(7): 0000000
'PRO_PESADOS(5): 00085
'SIGNO_PESADOS(1): 0
'IMPORTE_PESADOS(7): 0023518 à Siendo pues 235,18€.
'PRO_NO_PESADOS(5): 24300
'SIGNO_MAYORISTA(1): 0
'IMPORTE_NO_PESADOS(7): 0033792 à Siendo pues 337,92€.
'PRO_DEVUELTO(5): 00000
'SIGNO_DEVUELTO(1): 0
'IMPORTE_DEVUELTO(7): 0000000
'HORA(5): 22:02
'FECHA(10): 21-03-2019
'INTERNO (1):
Function actualiza_tabla_gtInfVen(Base As dao.Database, Cabecera As tipo_cabecera)
Dim eNCONTRADO As Boolean
Dim lclSec As Integer
Dim lclPrecio As Double
Dim lclImporte As Double
Dim lclFecha As String
Dim rst As dao.Recordset
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
    eNCONTRADO = False
    sSQL = "SELECT * " & _
           "FROM gtInfVen " & _
           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
           " AND codi_ident=" & Val(lclSec) & _
           " AND modo=" & Cabecera.tipoTiquet & _
           " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34)
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
            .Fields("modo") = Cabecera.tipoTiquet
            .Fields("fecha") = Cabecera.fecha
            .Fields("imporpeso") = 0
            .Fields("operpeso") = 0
            .Fields("imporuni") = 0
            .Fields("operuni") = 0
            
            .Fields("ntotal") = getNtotalTS(Cabecera.fecha)
            .Update
            
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        For cnt = 1 To Cabecera.LineasMensaje
        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
            lclImporte = Cabecera.Lines(cnt).ImporteLinea
            .Edit
            '//////////////
            'Cancelada o no
            '//////////////
            If Cabecera.Lines(cnt).LineaCancelada Then
                .Fields("imporCan") = .Fields("imporCan") + lclImporte
                .Fields("operCan") = .Fields("operCan") + 1
            Else
                If Cabecera.Lines(cnt).CodigoArticulo = 0 Then
                    .Fields("imporDir") = .Fields("imporDir") + lclImporte
                    .Fields("operDir") = .Fields("operDir") + 1
                Else
                    .Fields("imporCod") = .Fields("imporCod") + lclImporte
                    .Fields("operCod") = .Fields("operCod") + 1
                End If
                If Cabecera.Lines(cnt).EsPesado Then
                    .Fields("imporpeso") = .Fields("imporpeso") + lclImporte
                    .Fields("operpeso") = .Fields("operpeso") + 1
                Else
                    .Fields("imporuni") = .Fields("imporuni") + lclImporte
                    .Fields("operuni") = .Fields("operuni") + 1
                End If
                
            End If
            .Update
        End If
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
Public Sub Actualiza_Tabla_Tickets(Base As dao.Database, Cabecera As tipo_cabecera)
Dim BucleGrande As Integer
Dim D_FECHA As Date
Dim D_HORA As Date
Dim Maximo As Integer
Dim bucle As Long
Dim MiFactor As Double
Dim Registro As dao.Recordset
Dim Buffer As String
Dim MiCadenaDebug As String
Dim RegArt As dao.Recordset
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
        & Chr(34) & CStr(CDate(Cabecera.fecha)) & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
    End If
    For bucle = 1 To Cabecera.LineasMensaje
        Buffer = Cabecera.CadenaEstadoTiquet
        If (Cabecera.Lines(bucle).EsPesado) Then
            Buffer = Buffer & "peso "
        Else
            Buffer = Buffer & "unidades "
        End If
        Maximo = 1
        nMin = 1
        If lBoka Then
            If Cabecera.Lines(bucle).LineaCancelada Then
                Maximo = 2
                nMin = 1
            End If
        End If
        
        For BucleGrande = nMin To Maximo
            With Registro
                .AddNew
                If DEBUGACTIVO Then CadenadeLog (MiCadenaDebug & "Almacenando Linea de Tiquet : " _
                & Cabecera.Lines(bucle).NumeroLinea & "Tiquet : " & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                & "Importe : " & Cabecera.Lines(bucle).ImporteLinea)
                .Fields("Date") = Format(Cabecera.fecha, "dd/mm/yy")
                .Fields("Hour") = Format(Cabecera.hora, "hh:mm:ss")
                .Fields("d_fecha") = Cabecera.fecha
                .Fields("d_hora") = Cabecera.hora
                .Fields("numbal") = Cabecera.NBalanza
                .Fields("Vendor") = Cabecera.NVendedor
                .Fields("IVA") = Cabecera.Lines(bucle).IVA
                .Fields("porcentaje") = Cabecera.Lines(bucle).porcentaje
                .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
                .Fields("numlin") = Cabecera.Lines(bucle).NumeroLinea
                .Fields("SECC") = Cabecera.NMostrador
                .Fields("Code") = Cabecera.Lines(bucle).CodigoArticulo
                RegArt.FindFirst "codigo=" & Cabecera.Lines(bucle).CodigoArticulo
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
                .Fields("Price") = (Cabecera.Lines(bucle).precio) ' / 100)
                If BucleGrande = 2 Then
                    MiFactor = (-1)
                Else
                    MiFactor = 1
                End If
                'If ((Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128) Then
                '    .Fields("amount") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso)) ' / 100)
                '    If Cabecera.Lines(Bucle).EsPesado Then
                '        .Fields("Weight") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                '        .Fields("units") = 0
                '    Else
                '        .Fields("Weight") = 0
                '        .Fields("units") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                '    End If
                'Else
                    '.Fields("amount") = MiFactor * (Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso) ' / 100
                    .Fields("amount") = MiFactor * (Cabecera.Lines(bucle).ImporteLinea) ' / 100
                    If Cabecera.Lines(bucle).EsPesado Then
                        .Fields("Weight") = MiFactor * (Cabecera.Lines(bucle).cantidad * Cabecera.Lines(bucle).FactorPeso)
                        .Fields("units") = 0
                    Else
                        .Fields("Weight") = 0
                        .Fields("units") = MiFactor * (Cabecera.Lines(bucle).cantidad * Cabecera.Lines(bucle).FactorPeso)
                    End If
                'End If
                .Fields("STPETIC") = Cabecera.EstadoTiquet
                If Cabecera.Lines(bucle).LineaCancelada = True Then
                   .Fields("lincan") = 1
                Else
                   .Fields("lincan") = 0
                End If
                
                .Fields("TYPTIC") = Cabecera.tipoTiquet
                .Fields("reserv1") = Cabecera.Lines(bucle).FichaVacuno
                
                .Fields("importe_bruto") = Cabecera.Lines(bucle).ImporteBruto   '.Fields("amount") + Cabecera.Lines(Bucle).DescuentoLinea

                If Cabecera.lote <> "" Then
                    .Fields("numlote") = Cabecera.lote
                Else
                    .Fields("numlote") = " "
                End If
                '1.2.6
                If Not IsNull(Cabecera.NCliente) Then
                    .Fields("cliente") = Cabecera.NCliente
                Else
                .Fields("cliente") = 0
                End If
                ''''''
                If Cabecera.Lines(bucle).SehaModificado Then
                    .Fields("codi_pes") = "0"
                End If
                .Fields("reserv2") = "/"
                .Fields("reserv3") = "NO"
                '.Fields("tipoPago") = Cabecera.tipoPago
                
                .Update
                
                
            End With
            
        Next BucleGrande
    Next bucle
    
finTick:
    
    'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
    '    Cabecera.Incluir_en_Base = False
    'End If
    
    If Err.Number <> 0 Then
        'On Error GoTo 0
        CadenadeLog "Actualiza tabla Tickets:" & Err.Description
        On Error GoTo 0
    End If
End Sub

'\\\\\\\\\\\
'///////////
Public Sub Actualiza_Tabla_Log(Base As dao.Database, Cabecera As tipo_cabecera)
    '////////////////////////////////
    ' Genera log de cambios de precio o rechazos
    '////////////////////////////////
    Dim Registro As dao.Recordset
    Dim rstarticulo As dao.Recordset
    Dim bucle As Long
    Dim Buffer As String
    
    Dim LaSeccion As Integer
    Dim LaSubseccion As Integer
    Dim ElPlu As Integer
    Dim LaFamilia As Integer
    Dim LaCaducidad As Long
    Dim LaTara As Long
    Dim ElDescriptivo As String
    Dim MyPrec As Double
      
    For bucle = 1 To Cabecera.LineasMensaje
        
        If UsaEuro Then
            MyPrec = Cabecera.Lines(bucle).precio / 100
        Else
            MyPrec = Cabecera.Lines(bucle).precio / (10 ^ decimales)
        End If
        
        Set Registro = Base.OpenRecordset("select * from log where codigo=" & Cabecera.Lines(bucle).CodigoArticulo & " and precio=" + Str(MyPrec))
        If Registro.EOF Then
            With Registro
                If Cabecera.Lines(bucle).SehaModificado Then
                    '**********************************
                    ' Obtiene datos del Artículo
                    '**********************************
                    Set rstarticulo = Base.OpenRecordset _
                    ("select codigo,codi_ident,codi_sub,codi_fam,plu,caducidad,tara,des_plu1 from articulo where borrado=false and codigo=" & Cabecera.Lines(bucle).CodigoArticulo)
                    
                    With rstarticulo
                        If Not .EOF Then
                            LaSeccion = .Fields("codi_ident")
                            If Not IsNull(.Fields("codi_sub")) Then
                                LaSubseccion = .Fields("codi_sub")
                            Else
                                LaSubseccion = 0
                            End If
                            If Not IsNull(.Fields("plu")) Then
                                ElPlu = .Fields("plu")
                            Else
                                ElPlu = 0
                            End If
                            If Not IsNull(.Fields("codi_fam")) Then
                                LaFamilia = .Fields("codi_fam")
                            Else
                                LaFamilia = 0
                            End If
                            If Not IsNull(.Fields("caducidad")) Then
                                LaCaducidad = .Fields("caducidad")
                            Else
                                LaCaducidad = 0
                            End If
                            If Not IsNull(.Fields("tara")) Then
                                LaTara = .Fields("tara")
                            Else
                                LaTara = 0
                            End If
                            If .Fields("des_plu1") <> "" Then
                                ElDescriptivo = .Fields("des_plu1")
                            Else
                                ElDescriptivo = Space(25)
                            End If
                        Else
                            LaSeccion = 0
                            LaSubseccion = 0
                            ElPlu = 0
                            LaFamilia = 0
                            LaCaducidad = 0
                            LaTara = 0
                            ElDescriptivo = Space(25)
                        End If
                        .Close
                    End With
                    '**********************************
                    '**********************************
                    ' Crea la cadena de log
                    '**********************************
                    Buffer = "LOG "
                    Buffer = Buffer & Format(Cabecera.fecha, "ddmmyy") & ","
                    Buffer = Buffer & "3," & Format(Cabecera.hora, "hhmm") & ","
                    Buffer = Buffer & "0," & Format(Cabecera.Lines(bucle).CodigoArticulo, "000000") & ","
                    Buffer = Buffer & "000," & Format(LaSeccion, "000") & ","
                    If UsaEuro Then
                        Buffer = Buffer & Format(MyPrec, "000000") & ","
                    Else
                        Buffer = Buffer & Format(MyPrec, "000000") & ","
                    End If
                    Buffer = Buffer & Trim(ElDescriptivo) & ","
                    Buffer = Buffer & Format(Cabecera.NMostrador, "00")
                    '**********************************
                    .AddNew
                    .Fields("codi_ident") = LaSeccion
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codigo") = Cabecera.Lines(bucle).CodigoArticulo
                    .Fields("plu") = ElPlu
                    If UsaEuro Then
                        .Fields("precio") = Cabecera.Lines(bucle).precio / 100
                    Else
                        .Fields("precio") = Cabecera.Lines(bucle).precio / (10 ^ decimales)
                    End If
                    If Cabecera.Lines(bucle).EsPesado Then
                        .Fields("tipo_ven") = "W"
                    Else
                        .Fields("tipo_ven") = "U"
                    End If
                    .Fields("codi_fam") = LaFamilia
                    .Fields("caducidad") = LaCaducidad
                    .Fields("tara") = LaTara
                    .Fields("des_plu1") = ElDescriptivo
                    .Fields("Comentario") = "Desde Balanza."
                    .Fields("Texto") = Buffer
                    .Fields("Fecha") = Cabecera.fecha
                    .Update
                End If
            End With
        End If
        Registro.Close
    Next bucle
    

        
End Sub



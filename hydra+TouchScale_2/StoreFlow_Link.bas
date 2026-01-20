Attribute VB_Name = "StoreFlow_Link"
Option Explicit
Public StoreFlow_Modificaciones As Boolean
Public Sub CI_ProcesaOrden(EL_Buffer As String)
    ' debug
    Dim Bucle As Integer
    Dim espera_borrado As Integer
    Dim strData_1 As String
    Dim LtransEncurso As Boolean
    ' end debug
    Dim tipo_orden As String
    Dim orden As String
    Dim El_char As String
    Dim conectada As String
    Dim formato_fecha As String
    Dim formato_hora As String
    Dim LA_Fecha As Date
    Dim LA_Hora As Date
    Dim codi_ident As String
    Dim estado As String
    Dim cb_pormos As String
    Dim aprecio As String
    Dim pes As String
    Dim ppor_mos As String
    Dim cab_ley As String
    Dim num_lineas As String
    Dim fuen_linea As String
    Dim cadena As String
    
    Dim cadenadelcheck As String
    
    Dim Checksum As String
    Dim envio
    Dim Instring As String
    Dim i
    Dim k
    Dim resorden As String * 1024
    '''c2f
    Dim nCont, nCont1, nMarca As Integer
    
    '''''''''''''''''''''''''''''''''''parametros del tiquet'''''''''''
    
    Dim Base As Database
    Dim Registro As Recordset
    Dim MiBucle As Integer
    Dim MiVendedor As DB_Vendedor
    Dim MiMostrador As DB_mostrador
    Dim Miequipo As DB_Balanzas
    Dim MiCabecera As DB_Cabeceras
    Dim micodbar As DB_CodigoBarras
    Dim Miarticulo As DB_Articulo
    Dim MiTeclas As DB_Teclas
    Dim MiRespuesta As Integer
    Dim Archivo As Integer
    Dim BufferArchivo As String
    Dim Cadena0 As String
    '*************************
    ' para Euroscale por TCP/IP
    '*************************
    Dim ArchivoIP As Integer
    Dim ServidoraIP As Boolean
    Dim BufferIP As String
    '*************************
    BufferArchivo = EL_Buffer
    EL_Buffer = Mid(EL_Buffer, 6)
    tipo_orden = Mid(EL_Buffer, 1, 2)
    If DebugActivo Then CadenadeLog "Storeflow Link / Procesando " & EL_Buffer
   
   
        Select Case tipo_orden
               'Conexión / desconexión RED
    Case "CR" 'es a la inversa (conectada) para gama-386 y SC10 (C2F)
        conectada = Mid(EL_Buffer, 3, 1)
        If conectada = "0" Then 'Conexión
            AN_ControlRed True
        End If
        If conectada = "1" Then 'Desconexión
            AN_ControlRed False
        End If
       
    '******************************************
    'Programación formato de fecha y hora...
    ' Es ignorado por NTQ
    '*******************************************
    Case "FF" 'Aquí lo deja fijo ¿es así? (C2F)
        
                
    Case "FH"
        'Programación fecha y hora
        ' Programa también el formato de fecha y hora
        'la_fecha = Format(Date, "DD/MM/YY")
        'LA_Fecha = Format(Mid(EL_Buffer, 3, 2) & "/" & Mid(EL_Buffer, 5, 2) & "/" & Mid(EL_Buffer, 9, 2), "dd/mm/yy")
        'LA_Hora = Format(Mid(EL_Buffer, 11, 2) & ":" & Mid(EL_Buffer, 13, 2) & ":00", "HH:MM:SS")
        'AN_FechaHora LA_Fecha, LA_Hora

    Case "AS"
        'crear mostrador
        '*********************************
        ' Debido a que el StoreFlow no soporta
        ' Definir el tipo de Mostrador (SC10 o Euroscale)
        ' Aprovechamos el último caracter de la descripción
        ' del mostrador, de modo que si es el caracter "5"
        ' se considerará que el mostrador es Euroscale
        ' si es NO numérico se considerará SC10 de momento.
        ' Reservados :
        ' 0 --> SC10, 5 --> Euroscale , 1,2,3,4,6,7,8,9 --> Uso Futuro
        MiMostrador.NMostrador = Mid(EL_Buffer, 3, 2)
        MiMostrador.NSeccion = MiMostrador.NMostrador
        MiMostrador.descripcion = Mid(EL_Buffer, 5, 21)
        If Right(MiMostrador.descripcion, 1) = "5" Then
            MiMostrador.Tipo = 0
        Else
            MiMostrador.Tipo = 1
        End If
        MiMostrador.descripcion = Left(MiMostrador.descripcion, 20) & " "
        MiRespuesta = Alta_Mostrador(MiMostrador)
        If MiRespuesta < 2 Then AN_Secciones True
       
    ' Baja de mostrador(sección)... No válido para g-386... (C2F)
    'la baja de sección dejaría inaccesible los artículos asignados a la misma
    'sería conveniente revisar la mecánica ya que acaso fuera necesario establecer
    'una nueva mecánica que eliminara artículos...
    Case "BS"
        'baja mostrador
        MiMostrador.NMostrador = Mid(EL_Buffer, 3, 2)
        MiMostrador.NSeccion = MiMostrador.NMostrador
        MiRespuesta = Baja_Mostrador(MiMostrador)
        If MiRespuesta < 2 Then AN_Secciones True
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Alta de equipos ---> consultar para G-386...(C2F)
    Case "AE"
        'alta equipos
        '************************************************
        ' En función de que el mostrador sea de Euroscale o SC10
        ' Se creará una máquina SC10 o Euroscale
        ' En cuanto a las Euroscale, el protocolo no
        ' permite contemplar Dirección IP, que habrá que
        ' crear a mano
        '***************************************************
        Miequipo.NMostrador = Mid(EL_Buffer, 3, 2)
        Miequipo.NBalanza = Mid(EL_Buffer, 5, 2)
        Miequipo.descripcion = ""
        Miequipo.DireccionIP = ""
        Set Base = AbrirBase()
        Set Registro = Base.OpenRecordset _
        ("select * from seccion where secc_maqui=" & Miequipo.NMostrador & " and borrado=false")
        If Not Registro.EOF Then
            If Registro!enviardatos = "GA" Then
                Miequipo.Tipo = 0
            Else
                Miequipo.Tipo = 100
            End If
        Else
            Miequipo.Tipo = 0
        End If
        CerrarBase Base
        '**********************************
        ' TCP/IP : el protocolo Storeflow no
        ' permite indicar la dirección IP
        ' por tanto lo leemos de un fichero
        '**********************************
        If tcpip And Miequipo.Tipo = 0 Then
            If Dir(MiruTa & "\ipStflow.txt") <> "" Then
                ArchivoIP = FreeFile()
                Open MiruTa & "\ipstflow.txt" For Input As ArchivoIP
                ServidoraIP = True
                Do Until EOF(ArchivoIP)
                    Line Input #ArchivoIP, BufferIP
                    If Len(BufferIP) = 18 Then
                        If IsNumeric(Left(BufferIP, 2)) Then
                            If Val(Left(BufferIP, 2)) = Miequipo.NBalanza Then
                                If IsNumeric(Mid(BufferIP, 4, 3)) And IsNumeric(Mid(BufferIP, 8, 3)) _
                                And IsNumeric(Mid(BufferIP, 12, 3)) And IsNumeric(Mid(BufferIP, 16, 3)) Then
                                    If Mid(BufferIP, 7, 1) = "." And Mid(BufferIP, 11, 1) = "." _
                                    And Mid(BufferIP, 15, 1) = "." Then
                                        Miequipo.DireccionIP = Mid(BufferIP, 4, 15)
                                        If ServidoraIP Then Miequipo.Tipo = 1
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    End If
                    ServidoraIP = False
                Loop
            End If
        End If
        MiRespuesta = Alta_Balanza(Miequipo)
        If MiRespuesta < 2 Then AN_Equipos True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Baja de equipos ---> consultar para G-386...(c2f)
    'para dar de baja el equipo debe estar desconectado de red...
    Case "BE"
        'baja equipos
        Miequipo.NBalanza = Mid(EL_Buffer, 3, 2)
        MiRespuesta = Baja_Balanza(Miequipo)
        If MiRespuesta < 2 Then AN_Equipos True
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'El SC10 no tiene posibilidad de YY en código de Barras...
    'Se crea una función que sustituya YY por sección. En el caso
    'de que la programación sea por equipo se debe consultar a que
    'sección pertenece el equipo...
    Case "PB"
        'codigo de barras
        nMarca = 0
        If Mid(EL_Buffer, 3, 1) = "0" Then
            micodbar.equipo = 0
            micodbar.Mostrador = Mid(EL_Buffer, 4, 2)
        Else
            Set Base = AbrirBase()
            Set Registro = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & Mid(EL_Buffer, 4, 2))
            If Registro.EOF Then
                micodbar.Mostrador = 0
            Else
                micodbar.Mostrador = Registro!codi_ident
            End If
            CerrarBase Base
            micodbar.equipo = Mid(EL_Buffer, 4, 2)
        End If
        micodbar.Datos(0) = Mid(EL_Buffer, 6, 12)
        micodbar.Datos(5) = Mid(EL_Buffer, 18, 12)
        MiRespuesta = Alta_CodigoBarras(micodbar)
        If MiRespuesta < 2 Then AN_CodigosBarras True
        Case "PT"
            'Teclas rapidas
            cb_pormos = Mid(EL_Buffer, 3, 1)
            If cb_pormos = 0 Then
                MiTeclas.Destino_NMostrador = Mid(EL_Buffer, 4, 2)
                MiTeclas.SecArticulo = MiTeclas.Destino_NMostrador
            Else
                MiTeclas.Destino_NBalanza = Mid(EL_Buffer, 4, 2)
                Set Base = AbrirBase()
                Set Registro = Base.OpenRecordset("select * from equipos where numero_eqp=" & _
                MiTeclas.Destino_NBalanza & " and borrado=false")
                If Registro.EOF Then
                    MiTeclas.SecArticulo = 0
                Else
                    MiTeclas.SecArticulo = Registro!codi_ident
                End If
                CerrarBase Base
            End If
            MiTeclas.Destino_CodiTec = Mid(EL_Buffer, 8, 3)
            MiTeclas.PluArticulo = Mid(EL_Buffer, 11, 4)
            MiTeclas.Destino_Tabla = 0
            MiRespuesta = Alta_Teclas(MiTeclas)
            If MiRespuesta < 2 Then AN_TeclasRapidas True
            
                    
                
    Case "AV"
        'Alta de vendedor
        MiVendedor.seccion = 0
        MiVendedor.NVendedor = Mid(EL_Buffer, 5, 2)
        MiVendedor.descripcion = Mid(EL_Buffer, 7, 30)
        MiVendedor.estado = False
        MiVendedor.tecla = 0
        MiRespuesta = Alta_Vendedor(MiVendedor)
        If MiRespuesta < 2 Then AN_Vendedores (True)
        
    Case "BV"
        'Baja de vendedor
        MiVendedor.NVendedor = Mid(EL_Buffer, 5, 2)
        MiRespuesta = Baja_vendedor(MiVendedor)
        If MiRespuesta = 0 Then AN_Vendedores (True)
                    
    Case "AK"
        'Alta de tecla de vendedor
        MiVendedor.seccion = Mid(EL_Buffer, 3, 2)
        MiVendedor.NVendedor = Mid(EL_Buffer, 5, 2)
        MiVendedor.tecla = Mid(EL_Buffer, 8, 2)
        Set Base = AbrirBase()
        Set Registro = Base.OpenRecordset("select * from vendedor where ident_vend=" _
        & MiVendedor.NVendedor & " and borrado=false")
        If Not Registro.EOF Then
            MiRespuesta = 100
            MiVendedor.descripcion = Registro!Nombre
        End If
        CerrarBase Base
        If MiRespuesta = 100 Then
            MiRespuesta = Alta_Vendedor(MiVendedor)
            If MiRespuesta < 2 Then AN_Vendedores True
        End If
    Case "TV"
        'Baja tecla vendedor
        MiVendedor.NVendedor = Mid(EL_Buffer, 5, 2)
        Set Base = AbrirBase()
        Set Registro = Base.OpenRecordset("select * from vendedor where ident_vend=" & _
        MiVendedor.NVendedor & " and borrado=false")
        If Not Registro.EOF Then
            MiRespuesta = 100
            MiVendedor.descripcion = Registro!Nombre
        End If
        CerrarBase Base
        If MiRespuesta = 100 Then
            MiRespuesta = Alta_Vendedor(MiVendedor)
            If MiRespuesta < 2 Then AN_Vendedores True
        End If
        
                    
    Case "AA"
        'Altas de articulos
        Miarticulo.Mostrador = Mid(EL_Buffer, 3, 2)
        Miarticulo.plu = Mid(EL_Buffer, 5, 4)
        Miarticulo.codigo = Mid(EL_Buffer, 9, 14)  'c2f NO SE PUEDE
        aprecio = Mid(EL_Buffer, 23, 8)
        Miarticulo.precio = PrecioReal(aprecio)
               
        pes = Mid(EL_Buffer, 31, 1)
        If pes = 0 Then
            Miarticulo.WGH = False
        Else
            Miarticulo.WGH = True
        End If
        Miarticulo.caducidad = Mid(EL_Buffer, 32, 3)
        Miarticulo.Descriptivos(0) = Mid(EL_Buffer, 36, 25)
        'Miarticulo.tara = Mid(EL_Buffer, 61, 5)
        
        '******************
        ' Mientras no usen tramos
        '******************
        Miarticulo.TRM1 = -1
        Miarticulo.TRM2 = -1
        Miarticulo.PRC1 = -1
        Miarticulo.PRC2 = -1
        
        '******************
        ' Mientras no indiquen etiqueta
        '*******************
        Miarticulo.Etiqueta = 255
        If RechazarPlu0 And Miarticulo.precio = 0 Then
            MiRespuesta = 4
        Else
            MiRespuesta = Alta_Articulo(Miarticulo)
        End If
        If MiRespuesta < 2 Then AN_Articulos True
   Case "BA"
        'Baja de articulo
        Miarticulo.Mostrador = Mid(EL_Buffer, 3, 2)
        Miarticulo.plu = Mid(EL_Buffer, 5, 4)
        MiRespuesta = Baja_Articulo(Miarticulo)
        If MiRespuesta < 2 Then AN_Articulos True
   Case "MP"
        'Cambio de precio
        Miarticulo.Mostrador = Mid(EL_Buffer, 3, 2)
        Miarticulo.plu = Mid(EL_Buffer, 5, 4)
        'aprecio = Mid(El_Buffer, 9, 8) 'punto decimal libre
        aprecio = Mid(EL_Buffer, 9, 8)
        Miarticulo.precio = PrecioReal(aprecio)
        MiRespuesta = Modificacion_Articulo_Precio(Miarticulo)
        If MiRespuesta < 2 Then AN_Articulos True
   Case "PC"
        'cabeceras y leyendas
        ppor_mos = Mid(EL_Buffer, 3, 1)
        If ppor_mos = "0" Then
            MiCabecera.equipo = 0
            MiCabecera.Mostrador = Mid(EL_Buffer, 4, 2)
        Else
            Set Base = AbrirBase()
            Set Registro = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & Mid(EL_Buffer, 4, 2))
            If Registro.EOF Then
                MiCabecera.Mostrador = 0
            Else
                MiCabecera.Mostrador = Registro!codi_ident
            End If
            CerrarBase Base
            MiCabecera.equipo = Mid(EL_Buffer, 4, 2)
        End If
        For MiBucle = 0 To 4
            MiCabecera.Cabecera(MiBucle) = ""
            MiCabecera.leyenda(MiBucle) = ""
            MiCabecera.TLCabecera(MiBucle) = 0
            MiCabecera.TLLeyenda(MiBucle) = 0
        Next MiBucle
        cab_ley = Mid(EL_Buffer, 6, 1)
        MiCabecera.Reservado = 0
        num_lineas = Mid(EL_Buffer, 7, 1)
        '******************
        ' Rellena Cabeceras o Leyendas, según falte
        '******************
        
        Set Base = AbrirBase()
        If ppor_mos = "0" Then
            Set Registro = Base.OpenRecordset _
            ("select * from cabley where borrado=false and codi_ident=" & MiCabecera.Mostrador & " and isnull(numero_eqp)")
        Else
            Set Registro = Base.OpenRecordset _
            ("select * from cabley where borrado=false and numero_eqp=" & MiCabecera.equipo)
        End If
        If Not (Registro.EOF) Then
            With Registro
            If cab_ley = "0" Then
                MiCabecera.TLLeyenda(0) = !tletraley1
                MiCabecera.TLLeyenda(1) = !tletraley2
                MiCabecera.TLLeyenda(2) = !tletraley3
                MiCabecera.TLLeyenda(3) = !tletraley4
                MiCabecera.TLLeyenda(4) = !tletraley5
                MiCabecera.leyenda(0) = !leyenda01
                MiCabecera.leyenda(1) = !leyenda02
                MiCabecera.leyenda(2) = !leyenda03
                MiCabecera.leyenda(3) = !leyenda04
                MiCabecera.leyenda(4) = !leyenda05
            Else
                MiCabecera.TLCabecera(0) = !tletracab1
                MiCabecera.TLCabecera(1) = !tletracab2
                MiCabecera.TLCabecera(2) = !tletracab3
                MiCabecera.TLCabecera(3) = !tletracab4
                MiCabecera.TLCabecera(4) = !tletracab5
                MiCabecera.Cabecera(0) = !cabecera01
                MiCabecera.Cabecera(1) = !cabecera02
                MiCabecera.Cabecera(2) = !cabecera03
                MiCabecera.Cabecera(3) = !cabecera04
                MiCabecera.Cabecera(4) = !cabecera05
            End If
            End With
        End If
        CerrarBase Base
        For i = 0 To (num_lineas - 1)
            If cab_ley = "0" Then
                MiCabecera.TLCabecera(i) = Mid(EL_Buffer, 8 + (26 * i), 1)
                MiCabecera.Cabecera(i) = Mid(EL_Buffer, 9 + (26 * i), 25)
            Else
                MiCabecera.TLLeyenda(i) = Mid(EL_Buffer, 8 + (26 * i), 1)
                MiCabecera.leyenda(i) = Mid(EL_Buffer, 9 + (26 * i), 25)
            End If
        Next i
        MiRespuesta = Alta_Cabeceras(MiCabecera)
        If MiRespuesta < 2 Then AN_Cabeceras_Leyendas True
        
                   
    Case "DT"
       'Borrar todo
       AN_BorrarBases 0, True
    Case "DM"
       'Borrar Mostradores
       AN_BorrarBases 1, True
    Case "DV"
        'borrar vendedores
        AN_BorrarBases 3, True
    Case "DE"
        ' Borrar Equipos
        AN_BorrarBases 4, True
    Case "DA"
        'borrar articulos
        AN_BorrarBases 2, True
    Case "DK"
        'borrar teclas rapidas
        AN_BorrarBases 24, True
    Case "DP"
        'borrar cabeceras y leyendas
        AN_BorrarBases 5, True
End Select
     If suspenderGA = False And SuspenderSC10 = False Then
        If DebugActivo Then CadenadeLog "Storeflow Link / " & EL_Buffer & " Procesado"
     Else
        If DebugActivo Then CadenadeLog "Storeflow Link / Problema red de balanzas --> Modificaciones almacenadas en la base"
        'For Bucle = 1 To Len(BufferArchivo)
        '    If Asc(Mid(BufferArchivo, Bucle, 1)) = 0 Then
        '        Cadena0 = Cadena0 & "|"
        '    Else
        '        Cadena0 = Cadena0 & Mid(BufferArchivo, Bucle, 1)
        '    End If
        'Next Bucle
        'Archivo = FreeFile
        'Open MiruTa & "\stflow.buffer" For Output As #Archivo
        'Print #Archivo, Cadena0
        'Close #Archivo
     End If
     'Form1.datacliente.Bind CI_UDP
  
  
  
  
End Sub

Public Sub CI_EnviaTiquetsPEndientes(Euroscale_o_SC10 As Boolean)
    '**********************
    ' Euroscale=true
    ' SC10=false
    '**********************
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    '*****
    ' Cabecera
    Dim Codigo_mos As String
    Dim Codigo_bal As String
    Dim Codigo_ven As String
    Dim Cod_Bar As String
    Dim lineas As String
    Dim precio As String
    Dim Tipo As String
    Dim Fecha As String
    Dim hora As String
    '********
    ' Linea
    Dim Num_lin As String
    Dim Cod_Plu As String
    Dim Precio_Uni As String
    Dim Pes_Uni As String
    Dim Total As String
    Dim Signo As String
    Dim Tipo_Linea As String
    Dim Anulado As String
    Dim LineaAnulada As String
    '************
    Dim MiEnvio As String
    Dim nintentos As Integer
    Dim LosNulos As String * 1029
    Dim Bucle As Integer
    Dim MiTabla1 As String
    Dim MiTabla2 As String
    Dim tqcodbar As Tipo_TqCodBar
    lCogeTiquet = False
    If Euroscale_o_SC10 Then
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    Else
        MiTabla1 = "cabeceraSC10"
        MiTabla2 = "ticketsSC10"
    End If
    Set Base = AbrirBase()
    Set Registro = Base.OpenRecordset("select * from " & MiTabla1 & " where reserv3=" _
    & Chr(34) & "NO" & Chr(34) & " or reserv3=" & Chr(34) & "--" & Chr(34) & " order by val(nume)")
    If Not Registro.EOF Then
        Registro.MoveFirst
        ' si no está enviada la cabecera, se envía
        Do Until Registro.EOF Or CI_Salir = True
        With Registro
            If !reserv3 = "NO" Then
                Codigo_mos = Format(Val(!seccion), "00")
                Codigo_bal = Format(Val(!Balanza), "00")
                Codigo_ven = Format(Val(!vended), "00")
                If !codbar <> "" Then
                    Cod_Bar = !codbar
                Else
                    Cod_Bar = Space(13)
                End If
                lineas = Format(Val(!numlin), "000")
                If (!Importe) >= 0 Then
                    precio = !Importe
                    precio = Quitar_Coma(precio, 8)
                Else
                    precio = (-1) * !Importe
                    precio = "-" & Quitar_Coma(precio, 7)
                End If
                
                ' de momento tipo venta
                Tipo = "V"
                Fecha = Format(!Fecha, "ddmmyyyy")
                hora = Format(!hora, "hhmm")
                ' Enviar la cabecera
                MiEnvio = "10" & Chr(0) & Chr(0) & Chr(0) & "H" & _
                Codigo_mos & Codigo_bal & Codigo_ven & Cod_Bar & _
                lineas & precio & Tipo & Fecha & hora
                MiEnvio = MiEnvio & Mid(LosNulos, 1, 1029 - Len(MiEnvio))
                ' Envia la cadena
                CI_EnviaTCPTiquet MiEnvio, "Cabecera"
                Do Until CI_EstadoSistema > 13
                    Do_Events
                Loop
                If CI_EstadoSistema = 14 Then
                    .Edit
                    !reserv3 = "--"
                    .Update
                    CI_EstadoSistema = 10
                Else
                    .MoveLast
                End If
            End If
        End With
        If CI_EstadoSistema = 10 Then
        Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
        Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & _
        Format(Registro!nume, "00000") & Chr(34) & " order by val(numlin)")
        With Reg2
            If Not .EOF Then
                .MoveFirst
                 Do Until .EOF
                    With Reg2
                        Num_lin = Format(!numlin, "000")
                        Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        If !Price >= 0 Then
                            Precio_Uni = !Price
                            Precio_Uni = Quitar_Coma(Precio_Uni, 8)
                        Else
                            Precio_Uni = (-1) * !Price
                            Precio_Uni = (-1) * Quitar_Coma(Precio_Uni, 7)
                        End If
                        
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            Else
                                Pes_Uni = (-1) * !Weight
                                Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            End If
                            
                            Tipo_Linea = "P"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            Else
                                Pes_Uni = (-1) * !units
                                Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            End If
                            
                            Tipo_Linea = "U"
                        End If
                        
                        If !Amount < 0 Then
                            Total = (-1) * !Amount
                            Total = "-" & Quitar_Coma(Total, 7)
                            Signo = "-"
                        Else
                            Total = !Amount
                            Total = Quitar_Coma(Total, 8)
                            Signo = "+"
                        End If
                        If !lincan = 0 Then
                            LineaAnulada = "N"
                        Else
                            LineaAnulada = "S"
                        End If
                        MiEnvio = "10" & Chr(0) & Chr(0) & Chr(0) & "L" & Codigo_mos & _
                        Cod_Bar & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Total & Signo & _
                        Tipo_Linea & LineaAnulada
                    ' Envia la cadena
                        MiEnvio = MiEnvio & Mid(LosNulos, 1, 1029 - Len(MiEnvio))
                        CI_EnviaTCPTiquet MiEnvio, "Linea"
                        Do Until CI_EstadoSistema > 13
                            Do_Events
                        Loop
                        If CI_EstadoSistema = 14 Then
                            .Edit
                            !reserv3 = "SI"
                            CI_EstadoSistema = 10
                            .Update
                            .MoveNext
                        Else
                            .MoveLast
                            Registro.MoveLast
                        End If
                       
                    End With
                 Loop
                 Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
                 Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & " order by numlin")
                 If Reg2.EOF Then
                     Registro.Edit
                     Registro!reserv3 = "SI"
                     Registro.Update
                 End If
            Else
                Registro.Edit
                Registro!reserv3 = "SI"
                Registro.Update
            End If
        End With
        End If
        Registro.MoveNext
        Loop
    End If
    lCogeTiquet = True
    CerrarBase Base
    lCogeTiquet = True
    'Form1.datacliente.Bind CI_UDP
End Sub



Private Sub CI_EnviaTCPTiquet(MiEnvio As String, Mitexto As String)
    Dim Exito As Boolean
    Dim nintentos As Integer
    Exito = False
    nintentos = 0
    ' desconecta
    If Form1.tcpclient.State <> sckClosed Then Form1.tcpclient.Close
    CI_BufferTiquet = MiEnvio
    ' conecta
    CI_EstadoSistema = 11
    CI_Paquete = ""
    Form1.CI_Fallo.Enabled = True
    Form1.tcpclient.Connect
    'Do While (Form1.tcpclient.State <> sckConnected) And (nintentos < 10000)
    '    nintentos = nintentos + 1
    '    Do_Events
    'Loop
    ' desactiva recogida contínua
    
    'If Form1.tcpclient.State = sckConnected Then
    '
    'End If
End Sub


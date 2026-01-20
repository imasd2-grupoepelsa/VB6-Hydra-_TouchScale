Attribute VB_Name = "bck"
Option Explicit


Public Sub Actualiza_Tabla_Gtarti(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea)
    Dim Registro As Recordset
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
    Dim Bucle As Long
    Dim ElPrecio As Double
    Dim ElImporte As Double
    Dim MiFecha As Date
    '*******************
    ' Determina Sección de gestión
    '*******************
    For Bucle = 1 To Cabecera.LineasMensaje
    '****************************
    ' Si no existe, nuevo registro
    '****************************
        ElPrecio = Linea(Bucle).precio / 100
        ElImporte = Linea(Bucle).ImporteLinea / 100
        
        If Linea(Bucle).CodigoArticulo = 0 Then
            Set Registro = Base.OpenRecordset _
            ("select * from gtarti where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" _
            & " and codigo=" & Format(Linea(Bucle).CodigoArticulo, "000000") & _
            " and modo=" & Cabecera.tipoTiquet & " and secc_maqui=" & Cabecera.NMostrador)
        Else
            Set Registro = Base.OpenRecordset _
            ("select * from gtarti where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" _
            & " and codigo=" & Format(Linea(Bucle).CodigoArticulo, "000000") & _
            " and modo=" & Cabecera.tipoTiquet)
        End If
        '
        'Datos del artículo de la línea
        Set rstarticulo = Base.OpenRecordset _
        ("select * from articulo where borrado=false and codigo=" & Linea(Bucle).CodigoArticulo)
        With rstarticulo
            If Not .EOF Then
                If Not IsNull(!codi_ident) Then
                    LaSeccion = !codi_ident
                Else
                    LaSeccion = 0
                End If
                If Not IsNull(!secc_maqui) Then
                    LaSeccionMaquina = !secc_maqui
                Else
                    LaSeccionMaquina = 0
                End If
                If Not IsNull(!codi_sub) Then
                    LaSubseccion = !codi_sub
                Else
                    LaSubseccion = 0
                End If
                If Not IsNull(!codi_fam) Then
                    LaFamilia = !codi_fam
                Else
                    LaFamilia = 0
                End If
                If Not IsNull(!plu) Then
                    ElPlu = !plu
                Else
                    ElPlu = 0
                End If
                If Not IsNull(!tara) Then
                    LaTara = !tara
                Else
                    LaTara = 0
                End If
                If Not IsNull(!caducidad) Then
                    LaCaducidad = !caducidad
                Else
                    LaCaducidad = 0
                End If
                If Not IsNull(!tipo_iva) Then
                    ElIVA = !tipo_iva
                Else
                    ElIVA = 0
                End If
                If !des_plu1 <> "" Then
                    ElDescriptivo = !des_plu1
                Else
                    ElDescriptivo = Space(25)
                End If
            Else
                Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
                If Not rstseccion.EOF Then
                    LaSeccion = rstseccion!codi_ident
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
        With Registro
            If .EOF Then
                .AddNew
                !Fecha = Cabecera.Fecha
                !plu = ElPlu
                !des_plu1 = ElDescriptivo
                !codi_ident = LaSeccion
                !codi_sub = LaSubseccion
                !codi_fam = LaFamilia
                !secc_maqui = LaSeccionMaquina
                !codigo = Linea(Bucle).CodigoArticulo
                !precio = ElPrecio
                !operacion = 0
                !Peso = 0
                !Modo = Cabecera.tipoTiquet
                !Unidades = 0
                !pesetas = 0
                !tara = LaTara
                !caducidad = LaCaducidad
                !tipo_iva = ElIVA
                !Modificado = Linea(Bucle).SehaModificado
                .Update
            
            End If
        End With
        '********************+
        If Linea(Bucle).CodigoArticulo = 0 Then
            Set Registro = Base.OpenRecordset _
            ("select * from gtarti where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" _
            & " and codigo=" & Format(Linea(Bucle).CodigoArticulo, "000000") & _
            " and modo=" & Cabecera.tipoTiquet & " and secc_maqui=" & Cabecera.NMostrador)
        Else
            Set Registro = Base.OpenRecordset _
            ("select * from gtarti where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" _
            & " and codigo=" & Format(Linea(Bucle).CodigoArticulo, "000000") & _
            " and modo=" & Cabecera.tipoTiquet)
        End If
        With Registro
            .Edit
            !operacion = !operacion + 1
            If (Cabecera.EstadoTiquet And 64) = 64 Then
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = True Then
                        !Peso = Round(!Peso + (-1) * ((Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso), 3)
                    Else
                        !Unidades = !Unidades + (-1) * (Linea(Bucle).cantidad * Linea(Bucle).FactorPeso)
                    End If
                    !pesetas = Redondea(!pesetas + (-1) * (ElImporte * Linea(Bucle).FactorPeso))
                End If
            Else
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = True Then
                        !Peso = Round(!Peso + (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso, 3)
                    Else
                        !Unidades = !Unidades + Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                    End If
                    !pesetas = Redondea(!pesetas + ElImporte * Linea(Bucle).FactorPeso)
                End If
            End If
            .Update
        End With
    Next Bucle
   
End Sub
Private Function Redondea(ByVal dato As Double) As Double
    '************************************
    ' redondea la diferencia que
    ' se produce al tratar datos
    ' con decimales, con Access y VB
    '************************************
    
    dato = Round(dato, 2)
    
    Redondea = dato
    
End Function





'////////////////////////////////////////////////
' 1º TABLA TICKETS (datos de cada línea recibida)
' Tipos de Tiquet (typtic)
' 1 --> Rollo
' 2 --> Super
' 3 --> Envasado
' 4 --> Autoservicio
' 5 --> Rollo Local
' 6 --> Super Local
' 7 --> Envasado Local
' 8 --> Autoservicio Local
'\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub Actualiza_Tabla_Tickets(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea, Que_base As String)
    Dim BucleGrande As Integer
    Dim Maximo As Integer
    Dim Bucle As Long
    Dim MiFactor As Double
    Dim Registro As Recordset
    Dim Buffer As String
    Dim MiCadenaDebug As String
    Dim RegArt As Recordset
    Set RegArt = Base.OpenRecordset("select * from articulo where borrado=false")
    Buffer = Cabecera.CadenaEstadoTiquet
    If Que_base = "GA" Then
        MiCadenaDebug = "Euroscale / "
        Set Registro = Base.OpenRecordset("select * from tickets where date=" _
        & Chr(34) & Format(Cabecera.Fecha, "dd/mm/yy") & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
    End If
    If Que_base = "SC10" Then
        MiCadenaDebug = "SC10 / "
        Set Registro = Base.OpenRecordset("select * from ticketsSC10 where date=" _
        & Chr(34) & Format(Cabecera.Fecha, "dd/mm/yy") & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
    End If
    For Bucle = 1 To Cabecera.LineasMensaje
        Buffer = Cabecera.CadenaEstadoTiquet
        If Que_base = "GA" Then
            If (Linea(Bucle).TipoLinea And 1) Then
                Buffer = Buffer & "unidades "
                Linea(Bucle).EsPesado = False
            Else
                Buffer = Buffer & "peso "
                Linea(Bucle).EsPesado = True
            End If
        End If
        If Linea(Bucle).LineaCancelada = True Then
            Maximo = 2
        Else
            Maximo = 1
        End If
        For BucleGrande = 1 To Maximo
            With Registro
                .AddNew
                If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Linea de Tiquet : " _
                & Linea(Bucle).NumeroLinea & "Tiquet : " & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                & "Importe : " & Linea(Bucle).ImporteLinea & Linea(Bucle).FactorPeso)
                !Date = Format(Cabecera.Fecha, "dd/mm/yy")
                !Hour = Format(Cabecera.hora, "hh:mm:ss")
                !numbal = Cabecera.NBalanza
                !Vendor = Cabecera.NVendedor
                !IVA = Linea(Bucle).IVA
                !nume = Format(Cabecera.Ntiquet, "00000")
                !numlin = Linea(Bucle).NumeroLinea
                !SECC = Cabecera.NMostrador
                !Code = Linea(Bucle).CodigoArticulo
                RegArt.FindFirst "codigo=" & Linea(Bucle).CodigoArticulo
                If RegArt.EOF Or RegArt.NoMatch Then
                    !plu = 0
                    !Item = ""
                Else
                    !plu = RegArt!plu
                    !Item = RegArt!des_plu1
                End If
          
                !Price = (Linea(Bucle).precio / 100)
                
                If BucleGrande = 2 Then
                    MiFactor = (-1)
                Else
                    MiFactor = 1
                End If
                If (Cabecera.EstadoTiquet And 64) = 64 Then
                   
                    !Amount = MiFactor * (-1) * ((Linea(Bucle).ImporteLinea * Linea(Bucle).FactorPeso) / 100)
                    
                    '!euros = !amount / ValorEuro
                    If Linea(Bucle).EsPesado Then
                        !Weight = MiFactor * (-1) * ((Linea(Bucle).cantidad * Linea(Bucle).FactorPeso))
                        !units = 0
                    Else
                        !Weight = 0
                        !units = MiFactor * (-1) * ((Linea(Bucle).cantidad * Linea(Bucle).FactorPeso))
                    End If
    
                Else
               
                    !Amount = MiFactor * (Linea(Bucle).ImporteLinea * Linea(Bucle).FactorPeso) / 100
                    
                    '!euros = !amount / ValorEuro
                    If Linea(Bucle).EsPesado Then
                        !Weight = MiFactor * (Linea(Bucle).cantidad * Linea(Bucle).FactorPeso)
                        !units = 0
                    Else
                        !Weight = 0
                        !units = MiFactor * (Linea(Bucle).cantidad * Linea(Bucle).FactorPeso)
                    End If
                End If
                !STPETIC = Cabecera.EstadoTiquet
                If BucleGrande = 2 Then
                    !lincan = 1
                Else
                    !lincan = 0
                End If
                If Linea(Bucle).ModificadoTemporal Then
                    !codi_pes = 0
                Else
                    !codi_pes = 1
                End If
                !TYPTIC = Cabecera.tipoTiquet
                If Que_base = "GA" And UsaFichasVacuno Then
                    If Linea(Bucle).FichaVacuno <> 15000 Then
                        !reserv1 = Linea(Bucle).FichaVacuno
                    Else
                        !reserv1 = "-"
                    End If
                End If
                !reserv2 = "/"
                !reserv3 = "NO"
                .Update
            
            End With
        Next BucleGrande
    Next Bucle
End Sub
Public Sub Actualiza_Tabla_Cabecera(Base As Database, Cabecera As tipo_cabecera, Que_base As String, Linea() As Tipo_Linea)
    '****************************
    ' Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    '****************************
    Dim Registro As Recordset
    Dim RegSeccion As Recordset
    Dim MiCadenaDebug As String
    Dim tqcodbar As Tipo_TqCodBar
    Dim micodbar As String
    Dim Cod_Art As Long
    '*******************
    ' calcula el código de barras
    '*******************
    tqcodbar.Importe = Cabecera.ImporteTotal
    tqcodbar.NBalanza = Cabecera.NBalanza
    tqcodbar.NMostrador = Cabecera.NMostrador
    tqcodbar.NTicket = Cabecera.Ntiquet
    tqcodbar.NVendedor = Cabecera.NVendedor
    tqcodbar.Tipo = Cabecera.tipoTiquet
    If Cabecera.NLineas >= 1 Then
        Cod_Art = Linea(1).CodigoArticulo
    Else
        Cod_Art = 0
    End If
    micodbar = Genera_CodBar(tqcodbar, Cod_Art)
    If Que_base = "GA" Then
        MiCadenaDebug = "Euroscale / "
        Set Registro = Base.OpenRecordset("select * from cabecera where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and fecha=" & Chr(34) & Format(Cabecera.Fecha, "dd/mm/yy") & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & Cabecera.NVendedor & Chr(34))
    End If
    If Que_base = "SC10" Then
        MiCadenaDebug = "SC10 / "
        Set Registro = Base.OpenRecordset("select * from cabeceraSC10 where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and fecha=" & Chr(34) & Format(Cabecera.Fecha, "dd/mm/yy") & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & _
        Cabecera.NVendedor & Chr(34))
    End If
    With Registro
        If .EOF Then
            If (Cabecera.EstadoTiquet And 64) = 0 Then
                If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Cabecera de Tiquet : " _
                & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                & "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.Fecha & "Hora : " & Cabecera.hora)
                .AddNew
                !nume = Format(Cabecera.Ntiquet, "00000")
                ' ******
                ' Convierte sección máquina a gestión
                ' ******
                'Set RegSeccion = Base.OpenRecordset("select * from seccion where borrado=false and secc_maqui=" & Cabecera.NMostrador)
                'If Not RegSeccion.EOF Then
                '    !seccion = RegSeccion!codi_ident
                'Else
                    !seccion = Cabecera.NMostrador
                'End If
                ' ******
                !codbar = micodbar
                !Balanza = Cabecera.NBalanza
                !vended = Cabecera.NVendedor
                !numlin = Cabecera.NLineas
                !lincan = Cabecera.LineasCanceladas
                
                !Importe = Cabecera.ImporteTotal / 100

                !hora = Format(Cabecera.hora, "hh:mm:ss")
                !onoff = Cabecera.CadenaEstadoTiquet
                !STPETIC = Cabecera.EstadoTiquet
                !codcli = Cabecera.NCliente
                !Fecha = Format(Cabecera.Fecha, "dd/mm/yy")
                '!euros=cabecera.importetotal / valoreuro
                !TYPTIC = Cabecera.tipoTiquet
                !reserv2 = "/"
                !reserv3 = "NO"
                .Update
            Else
                Cabecera.Incluir_en_Base = False
            End If
        Else
            If (Cabecera.EstadoTiquet And 64) Then
                If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Cabecera de Tiquet : " _
                & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                & "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.Fecha & "Hora : " & Cabecera.hora)
                .AddNew
                !nume = Format(Cabecera.Ntiquet, "00000")
                ' ******
                ' Convierte sección máquina a gestión
                ' ******
                'Set RegSeccion = Base.OpenRecordset("select * from seccion where borrado=false and secc_maqui=" & Cabecera.NMostrador)
                'If Not RegSeccion.EOF Then
                '    !seccion = RegSeccion!codi_ident
                'Else
                    !seccion = Cabecera.NMostrador
                'End If
                ' ******
                !codbar = ""
                !Balanza = Cabecera.NBalanza
                !vended = Cabecera.NVendedor
                !numlin = Cabecera.NLineas
                !lincan = Cabecera.LineasCanceladas
                
                !Importe = (-1) * (Cabecera.ImporteTotal / 100)
                !hora = Format(Cabecera.hora, "hh:mm:ss")
                !onoff = Cabecera.CadenaEstadoTiquet
                !STPETIC = Cabecera.EstadoTiquet
                !codcli = Cabecera.NCliente
                !Fecha = Format(Cabecera.Fecha, "dd/mm/yy")
                '!euros=cabecera.importetotal / valoreuro
                !TYPTIC = Cabecera.tipoTiquet
                !reserv2 = "/"
                !reserv3 = "NO"
                .Update
            Else
                Cabecera.Incluir_en_Base = False
            End If
        End If
    End With
End Sub
Public Sub Actualiza_Tabla_GtVend(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea, LineaInicial As Boolean)
    Dim Encontrado As Boolean
    Dim Bucle As Long
    Dim Lafecha As String
    Dim LaSeccion As Integer
    Dim Registro As Recordset
    Dim rstseccion As Recordset
    Dim rstvendedor As Recordset
    Dim ElImporte As Double
    Dim ElPrecio As Double
    Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
    If Not rstseccion.EOF Then
        LaSeccion = rstseccion!codi_ident
    Else
        LaSeccion = Cabecera.NMostrador
    End If
    rstseccion.Close
    Encontrado = False
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#")
    With Registro
        If .EOF Then
            .AddNew
            !Fecha = Cabecera.Fecha
            !codi_ident = LaSeccion
            !ident_vend = Cabecera.NVendedor
            !secc_maqui = Cabecera.NMostrador
            !Modo = Cabecera.tipoTiquet
            !tiquets = 0
            !operacion = 0
            !cancela = 0
            '**********
            ' Nombre del vendedor
            '***********
            Set rstvendedor = Base.OpenRecordset _
            ("select * from vendedor where borrado=false and ident_vend=" & Cabecera.NVendedor _
            & " and codi_ident=" & LaSeccion)
            If rstvendedor.EOF Or !Nombre = "" Then
                !Nombre = Space(25)
            Else
                !Nombre = rstvendedor!Nombre
            End If
            rstvendedor.Close
            ' ****
            'pesado o unidades
            ' ****
            !Peso = 0
            !unidad = 0
            ' ****
            !pesetas = 0
            ' ****
            ' Cancelada o no
            ' ****
            !totcance = 0
            !supertot = 0
            !Total = 0
            .Update
        End If
    End With
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#")
    With Registro
        If LineaInicial Then
            
            .Edit
            !tiquets = !tiquets + 1
            !operacion = !operacion + Cabecera.NLineas
            .Update
        End If
        For Bucle = 1 To Cabecera.LineasMensaje
            .Edit
            
            ElImporte = Linea(Bucle).ImporteLinea / 100
            ElPrecio = Linea(Bucle).precio / 100
            
            ' ****
            'pesado o unidades
            ' ****
            If (Cabecera.EstadoTiquet And 64) = 64 Then
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = False Then
                        !unidad = !unidad + (-1) * (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        !Peso = 0
                    Else
                        !Peso = Round(!Peso + (-1) * (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso, 3)
                        !unidad = 0
                    End If
                    ' ****
                    !pesetas = Redondea(!pesetas + (-1) * ElImporte * Linea(Bucle).FactorPeso)
                End If
            Else
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = False Then
                        !unidad = !unidad + (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                        !Peso = 0
                    Else
                        !Peso = Round(!Peso + (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso, 3)
                        !unidad = 0
                    End If
                    ' ****
                    !pesetas = Redondea(!pesetas + ElImporte * Linea(Bucle).FactorPeso)
                End If
            End If
            '!euros = Linea.ImporteLinea / ValorEuro
            ' ****
            ' Cancelada o no
            ' ****
            If Linea(Bucle).LineaCancelada Then
                If (Cabecera.EstadoTiquet And 64) = 64 Then
                    !totcance = Redondea(!totcance - ElImporte)
                    !supertot = Redondea(!supertot - ElImporte)
                    !cancela = Redondea(!cancela - 1)
                Else
                    !totcance = Redondea(!totcance + ElImporte)
                    !supertot = Redondea(!supertot + ElImporte)
                    !cancela = Redondea(!cancela + 1)
                End If
            Else
                If (Cabecera.EstadoTiquet And 64) = 64 Then
                    !Total = Redondea(!Total + (-1) * ElImporte * Linea(Bucle).FactorPeso)
                    !supertot = Redondea(!Total + (-1) * ElImporte)
                Else
                    !Total = Redondea(!Total + ElImporte * Linea(Bucle).FactorPeso)
                    !supertot = Redondea(!Total + ElImporte)
                End If
            End If
            
            .Update
        Next Bucle
    End With
End Sub
Public Sub Actualiza_Tabla_Gtsecs(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea, LineaInicial As Boolean)
    Dim Bucle As Long
    Dim Registro As Recordset
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
    '*******************
    ' Determina Sección de gestión
    '*******************
    For Bucle = 1 To Cabecera.LineasMensaje
    '****************************
    ' Si no existe, nuevo registro
    '****************************

        ElImporte = Linea(Bucle).ImporteLinea / 100
        ElPrecio = Linea(Bucle).precio / 100
        
        '
        'Datos del artículo de la línea
        Set rstarticulo = Base.OpenRecordset _
        ("select * from articulo where borrado=false and codigo=" & Linea(Bucle).CodigoArticulo)
        With rstarticulo
            If Not .EOF Then
                If Not IsNull(!codi_ident) Then
                    LaSeccion = !codi_ident
                Else
                    LaSeccion = 0
                End If
                If Not IsNull(!secc_maqui) Then
                    LaSeccionMaquina = !secc_maqui
                Else
                    LaSeccionMaquina = 0
                End If
                If Not IsNull(!codi_sub) Then
                    LaSubseccion = !codi_sub
                Else
                    LaSubseccion = 0
                End If
                If Not IsNull(!codi_fam) Then
                    LaFamilia = !codi_fam
                Else
                    LaFamilia = 0
                End If
                If Not IsNull(!plu) Then
                    ElPlu = !plu
                Else
                    ElPlu = 0
                End If
                If Not IsNull(!tara) Then
                    LaTara = !tara
                Else
                    LaTara = 0
                End If
                If Not IsNull(!caducidad) Then
                    LaCaducidad = !caducidad
                Else
                    LaCaducidad = 0
                End If
                If Not IsNull(!tipo_iva) Then
                    ElIVA = !tipo_iva
                Else
                    ElIVA = 0
                End If
                If !des_plu1 <> "" Then
                    ElDescriptivo = !des_plu1
                Else
                    ElDescriptivo = Space(25)
                End If
            Else
                Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
                If Not rstseccion.EOF Then
                    LaSeccion = rstseccion!codi_ident
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
        ("select * from gtsecs where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") _
        & "# and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
        "and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
        With Registro
            If .EOF Then
                .AddNew
                !Fecha = Cabecera.Fecha
                !codi_ident = LaSeccion
                !codi_sub = LaSubseccion
                !codi_fam = LaFamilia
                !secc_maqui = LaSeccionMaquina
                !Modo = Cabecera.tipoTiquet
                !tiquets = 0
                !operacion = 0
                !Peso = 0
                !Unidades = 0
                !pesetas = 0
                .Update
            End If
        End With
        Set Registro = Base.OpenRecordset _
        ("select * from gtsecs where fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") _
        & "# and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
        "and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
        With Registro
            .Edit
            If LineaInicial Then
                !tiquets = !tiquets + 1
                LineaInicial = False
            End If
            !operacion = !operacion + 1
            If (Cabecera.EstadoTiquet And 64) = 64 Then
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado Then
                        !Peso = Round(!Peso + (-1) * (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso, 3)
                    Else
                        !Unidades = !Unidades + (-1) * Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                    End If
                    !pesetas = Redondea((!pesetas + (-1) * ElImporte * Linea(Bucle).FactorPeso))
                End If
            Else
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado Then
                        !Peso = Round(!Peso + (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso, 3)
                    Else
                        !Unidades = !Unidades + Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                    End If
                    !pesetas = Redondea(!pesetas + ElImporte * Linea(Bucle).FactorPeso)
                End If
            End If
            .Update
        End With
    Next Bucle
End Sub
Public Sub Actualiza_Tabla_Log(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea)
    Dim Registro As Recordset
    Dim rstarticulo As Recordset
    Dim Bucle As Long
    Dim Buffer As String
    Dim LaSeccion As Integer
    Dim LaSubseccion As Integer
    Dim ElPlu As Integer
    Dim LaFamilia As Integer
    Dim LaCaducidad As Long
    Dim LaTara As Long
    Dim ElDescriptivo As String
    Dim MyPrec As Double
    For Bucle = 1 To Cabecera.LineasMensaje
        MyPrec = Linea(Bucle).precio / 100
        Set Registro = Base.OpenRecordset("select * from log where codigo=" & Linea(Bucle).CodigoArticulo & " and precio=" + Str(MyPrec))
        If Registro.EOF Then
            With Registro
                If Linea(Bucle).SehaModificado Then
                    '**********************************
                    ' Obtiene datos del Artículo
                    '**********************************
                    Set rstarticulo = Base.OpenRecordset _
                    ("select * from articulo where borrado=false and codigo=" & Linea(Bucle).CodigoArticulo)
                    With rstarticulo
                        If Not .EOF Then
                            LaSeccion = !codi_ident
                            If Not IsNull(!codi_sub) Then
                                LaSubseccion = !codi_sub
                            Else
                                LaSubseccion = 0
                            End If
                            If Not IsNull(!plu) Then
                                ElPlu = !plu
                            Else
                                ElPlu = 0
                            End If
                            If Not IsNull(!codi_fam) Then
                                LaFamilia = !codi_fam
                            Else
                                LaFamilia = 0
                            End If
                            If Not IsNull(!caducidad) Then
                                LaCaducidad = !caducidad
                            Else
                                LaCaducidad = 0
                            End If
                            If Not IsNull(!tara) Then
                                LaTara = !tara
                            Else
                                LaTara = 0
                            End If
                            If !des_plu1 <> "" Then
                                ElDescriptivo = !des_plu1
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
                    Buffer = Buffer & Format(Cabecera.Fecha, "ddmmyy") & ","
                    Buffer = Buffer & "3," & Format(Cabecera.hora, "hhmm") & ","
                    Buffer = Buffer & "0," & Format(Linea(Bucle).CodigoArticulo, "000000") & ","
                    Buffer = Buffer & "000," & Format(LaSeccion, "000") & ","
                    Buffer = Buffer & Format(MyPrec, "000000") & ","
                    Buffer = Buffer & Trim(ElDescriptivo) & ","
                    Buffer = Buffer & Format(Cabecera.NMostrador, "00")
                    '**********************************
                    .AddNew
                    !codi_ident = LaSeccion
                    !codi_sub = LaSubseccion
                    !codigo = Linea(Bucle).CodigoArticulo
                    !plu = ElPlu
                    !precio = Linea(Bucle).precio / 100
                    If Linea(Bucle).EsPesado Then
                        !tipo_ven = 1
                    Else
                        !tipo_ven = 0
                    End If
                    !codi_fam = LaFamilia
                    !caducidad = LaCaducidad
                    !tara = LaTara
                    !des_plu1 = ElDescriptivo
                    !Comentario = "Desde Balanza."
                    !Texto = Buffer
                    !Fecha = Cabecera.Fecha
                    .Update
                End If
            End With
        End If
        Registro.Close
    Next Bucle
End Sub
'////////////////////////////////////////////////////////////////////////////////
'INCORPORA LOS DATOS DE TOTALES POR ARTÍCULO A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////////////////////////////////////////////////////
Public Sub Actualiza_Tabla_Gtarti(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea, Optional argTipoTqt As Integer)
'//////////////////////////////
' argTipoTqt => tipo de tiquet
' 1 --> 386
' 2 --> C9/C10
' 3 --> SC10
' 4 --> IV4TQ
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
Dim actborrar As Boolean
    actborrar = False
    For Bucle = 1 To Cabecera.LineasMensaje
        ElPrecio = Linea(Bucle).precio / 100
        ElImporte = Linea(Bucle).ImporteLinea / 100
        'Datos del artículo de la línea
        'If TipoC9 Then
        If argTipoTqt = 2 Then
            sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1 " & _
                   "FROM articulo " & _
                   "WHERE borrado=false " & _
                   " AND plu=" & Linea(Bucle).CodigoArticulo & _
                   " AND codi_ident=" & Cabecera.NMostrador
        Else
            sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1 " & _
                   "FROM articulo " & _
                   "WHERE borrado=false" & _
                   " AND codigo=" & Linea(Bucle).CodigoArticulo
        End If
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
            End If
            .Close
        End With
        Set rst = Nothing
        If Linea(Bucle).CodigoArticulo = 0 Then
            sSQL = "SELECT * " & _
                   "FROM gtarti " & _
                   "WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _
                   " AND codigo=" & Format(ElCodigo, "000000") & _
                   " AND modo=" & Cabecera.tipoTiquet & _
                   " AND secc_maqui=" & LaSeccionMaquina
        Else
            sSQL = "SELECT * " & _
                   "FROM gtarti " & _
                   "WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _
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
                    .Fields("txt_sub") = ""
                Else
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
                End If
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
                .Fields("Modificado") = Linea(Bucle).SehaModificado
                .Update
            End If
            .Close
        End With
        Set rst = Nothing
        If ElCodigo = 0 Then
            sSQL = "SELECT * " & _
                   "FROM gtarti " & _
                   "WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _
                   " AND codigo=" & Format(ElCodigo, "000000") & _
                   " AND modo=" & Cabecera.tipoTiquet & _
                   " AND secc_maqui=" & CStr(LaSeccionMaquina)
        Else
            sSQL = "SELECT * " & _
                   "FROM gtarti " & _
                   "WHERE fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#" & _
                   " AND codigo=" & Format(ElCodigo, "000000") & _
                   " AND modo=" & Cabecera.tipoTiquet
        End If
        Set rst = Base.OpenRecordset(sSQL)
        With rst
            .Edit
            .Fields("operacion") = .Fields("operacion") + 1
            If (Cabecera.EstadoTiquet And 64) = 64 Then
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = True Then
                        .Fields("Peso") = .Fields("Peso") + (-1) * ((Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso)
                    Else
                        .Fields("Unidades") = .Fields("Unidades") + (-1) * (Linea(Bucle).cantidad * Linea(Bucle).FactorPeso)
                    End If
                    .Fields("pesetas") = Redondea(.Fields("pesetas") + (-1) * (ElImporte * Linea(Bucle).FactorPeso))
                End If
            Else
                If Linea(Bucle).LineaCancelada = False Then
                    If Linea(Bucle).EsPesado = True Then
                        .Fields("Peso") = .Fields("Peso") + (Linea(Bucle).cantidad) * Linea(Bucle).FactorPeso
                    Else
                        .Fields("Unidades") = .Fields("Unidades") + Linea(Bucle).cantidad * Linea(Bucle).FactorPeso
                    End If
                    .Fields("pesetas") = Redondea(.Fields("pesetas") + ElImporte * Linea(Bucle).FactorPeso)
                Else
                    actborrar = True
                End If
            End If
            .Update
            If actborrar Then
                .Delete
                actborrar = False
            End If
            .Close
        End With
        Set rst = Nothing
    Next Bucle
End Sub
'//////
'\\\\\\
'////////////////////////////////////////////////////////////////////////////////////////
'INCORPORA LA INFORMACION DE TOTALES DE VENDEDORES A PARTIR DE LO RECIBIDO EN LOS TICKETS
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Function actualiza_tabla_gtInfVen(Base As Database, Cabecera As tipo_cabecera, Linea() As Tipo_Linea, LineaInicial As Boolean)
Dim Encontrado As Boolean
Dim lclSec As Integer
Dim lclPrecio As Double
Dim lclImporte As Double
Dim lclfecha As String
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
    sSQL = "SELECT * " & _
           "FROM gtInfVen " & _
           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & " AND codi_ident=" & Val(lclSec) & " AND modo=" & Cabecera.tipoTiquet & " AND fecha=#" & Format(Cabecera.Fecha, "mm/dd/yy") & "#"
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
            .Fields("modo") = Cabecera.tipoTiquet
            .Fields("fecha") = Cabecera.Fecha
            .Update
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        For cnt = 1 To Cabecera.LineasMensaje
            lclImporte = Redondea((Linea(cnt).ImporteLinea / 100))
            rst.Edit
            '//////////////
            'Cancelada o no
            '//////////////
            If Linea(cnt).LineaCancelada Then
                'If (Cabecera.EstadoTiquet And 64) = 64 Then
                '    .Fields("imporCan") = .Fields("imporCan") - lclImporte
                '    .Fields("operCan") = .Fields("operCan") - 1
                'Else
                    .Fields("imporCan") = .Fields("imporCan") + lclImporte
                    .Fields("operCan") = .Fields("operCan") + 1
                'End If
            Else
                'If (Cabecera.EstadoTiquet And 64) = 64 Then
                '    .Fields("Total") = Redondea(.Fields("Total") + (-1) * lclImporte * Linea(cnt).FactorPeso)
                'Else
                '    .Fields("Total") = Redondea(.Fields("Total") + lclImporte * Linea(cnt).FactorPeso)
                'End If
                If Linea(cnt).CodigoArticulo = 0 Then
                    .Fields("imporDir") = .Fields("imporDir") + lclImporte
                    .Fields("operDir") = .Fields("operDir") + 1
                Else
                    .Fields("imporCod") = .Fields("imporCod") + lclImporte
                    .Fields("operCod") = .Fields("operCod") + 1
                End If
            End If
            .Update
        Next cnt
        .Close
    End With
    Set rst = Nothing
End Function



Attribute VB_Name = "Comunicacion_General"
Option Explicit
'\\\\\\
'//////
Private Type TipoBarras
    Tipo As String
    Numero As Integer
    estado As String
End Type
'\\\\\\
'//////
Public Sub CG_PideTicket(MachineID As Integer, aModTrab As Integer, Optional aTipoTqt As Integer, Optional ByVal nVnd As Integer, Optional ByVal nEqpAlfa As Integer)
'//////////////////////////
'// PETICIÓN DE TIQUETS  //
'// PARA EUROSCALE,      //
'// SC10 (NO PATH RED)   //
'// Y C9/C10             //
'//////////////////////////
' aModTrab => modo de trabajo
' 0 --> Venta
' 1 --> Super
' 2 --> Envasado
' 3 --> AutoServicio
' 4 --> Local
' 9 --> anulados... Fanjul 1.8.4
'///////////////////////////
' aTipoTqt => tipo de tiquet
' 1 --> 386
' 2 --> C9/C10
' 3 --> SC10
' 4 --> IV4TQ
'\\\\\\\\\\\\
Dim MiContador As Integer
Dim Bucle As Long
Dim MiEnvio
Dim MiRecibe As Long
Dim PedirTiquet As String
Dim Cabecera As tipo_cabecera
Dim Linea() As Tipo_Linea
Dim Repetir As Boolean
Dim CuentaLineas As Integer
Dim NroLineas As Integer
Dim LineaInicial As Integer
Dim LineaComun As String
Dim ToTalLineas As Integer
Dim ErrorLineas As Boolean
Dim Buffer As String
Dim MiTipo As String
Dim Base As Database
Dim RegArt As Recordset
Dim BalTick As Epel_CabTick
Dim baseID As String
Dim RegTramo As Recordset
Dim HayTramos As Boolean
Dim MyArt As DB_Articulo
Dim ScaleArt As DB_Articulo
Dim OKCambiar As Boolean
Dim CSiempre As Boolean
Dim ContadorTQ As Long
Dim SeccionTQ As Integer
Dim MaquinaTQ As Integer
Dim RegTQ As Recordset
Dim tp As Variant
Dim ce As Variant
'*********************cas.v1.1.6
Dim lclImporte As Long
Dim cnt As Long
'*********************cas.v1.1.8
Dim sSQL As String
Dim codArt As Long
'c2f/caspiunza
Dim nNumLinA As Long
Dim nContLinA As Long
Dim nMyBu As Integer

nNumLinA = 0
nContLinA = 0
'/////////////////
    ContadorTQ = 1
    If Dir(App.Path & "\csiempre.txt") <> "" Then
        CSiempre = True
    Else
        CSiempre = False
    End If
    If Not UsaModem Then
        frmControl.Switch_mode MachineID
    End If
    If MachineID >= 700 And MachineID < 800 Then MachineID = 0
    Select Case MachineID
        Case 0
            baseID = "GA"
            MiTipo = aModTrab
        Case 1 To 2
            baseID = "SC10"
            If TipoC9 Then
                frmControl.Balanza.C9_TickRewind
                'c2f/caspiunza
                nNumLinA = frmControl.Balanza.C9_NUMLIN
                ''''''''''''''
            End If
            Select Case aModTrab
                Case 0
                    MiTipo = 0
                Case 1
                    MiTipo = 1
                Case 2
                    MiTipo = 4
                Case 3
                    MiTipo = 2
                Case 4
                    MiTipo = 3
            End Select
        Case 501 To 599
            baseID = "TQ"
            MiTipo = 0
    End Select
    
    If lMarteAlfa Then
        baseID = "TQ"
        MiTipo = 0
        aModTrab = 0
    End If
    
    Cabecera.Ntiquet = 0
    Repetir = True
    
    Select Case aModTrab
        Case 0
            If lMarteAlfa Then
                frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(332) & " " & CargaCadena(336) & " VND: " & CStr(nVnd)
            Else
                frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(332) & " " & CargaCadena(336)
            End If
        Case 1
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(332) & " " & CargaCadena(741)
        Case 2
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(332) & " " & CargaCadena(195)
        Case 3
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(332) & " " & CargaCadena(196)
        Case 4
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(333) & " " & CargaCadena(336)
        Case 5
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(333) & " " & CargaCadena(741)
        Case 6
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(333) & " " & CargaCadena(195)
        Case 7
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(333) & " " & CargaCadena(196)
        Case 9
            frmMonitor.MostrarInfo CargaCadena(827) & " " & CargaCadena(333) & " " & "Null Ticket..."
    End Select
    Set Base = OpenDatabase(Base_General)
    If (MachineID > 500 And MachineID < 600) Or lMarteAlfa Then 'cas.n->si es gama baja
        If lMarteAlfa Then
            MaquinaTQ = nEqpAlfa
        Else
            MaquinaTQ = MachineID - 500
        End If
            
        Set RegTQ = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & MaquinaTQ)
        SeccionTQ = RegTQ.Fields("codi_ident")
    End If
    Do Until (Repetir = False) Or (ErrorLineas = True)
        frmEpelsa.CmdComunicaciones(3).Enabled = False
        frmEpelsa.CmdComunicaciones(4).Enabled = False
        If MachineID > 500 And MachineID < 600 Then ' IV4TQ
            MiRecibe = frmControl.Balanza.Query_One_Ticket(Val(MiTipo), 0, 0, ContadorTQ, BalTick)
            ContadorTQ = ContadorTQ + 1
        Else
            ' Resto de balanzas
            If lMarteAlfa Then
                MiRecibe = frmControl.Balanza.Query_One_Ticket(Val(MiTipo), 0, 0, ContadorTQ, BalTick, nVnd)
                ContadorTQ = ContadorTQ + 1
            Else
                MiRecibe = frmControl.Balanza.Query_Ticket(Val(MiTipo), BalTick)
            End If
        End If
        frmEpelsa.CmdComunicaciones(3).Enabled = True
        frmEpelsa.CmdComunicaciones(4).Enabled = True
        If MiRecibe >= 1000 Then
            frmMonitor.MostrarDato CargaCadena(544)
            Exit Do
        End If
        If MiRecibe = 4 Then
            frmMonitor.MostrarDato CargaCadena(540)
            Repetir = False
        Else
            '***
            ' si STS=0, lo procesa (tiquet cerrado)
            '***
            If MiRecibe = 0 Then
                Repetir = True
                '********************
                ' Toma los datos de Cabecera
                '********************
                frmMonitor.MostrarDato CargaCadena(865) & " " & BalTick.NTicket
                Cabecera.NLineas = BalTick.nLines
                ReDim Linea(Cabecera.NLineas + 1)
                Cabecera.LineasMensaje = BalTick.nLines
                Cabecera.LineaFinal = BalTick.nLines
                Cabecera.LineaInicial = 1
                '2.0.9 ya se suman más abajo las canceladas
                'If lMarteAlfa Then
                '    Cabecera.LineasCanceladas = BalTick.nLinCanceled
                'Else
                    Cabecera.LineasCanceladas = 0
                'End If
                Cabecera.Ntiquet = BalTick.NTicket
                Cabecera.NMostrador = BalTick.Section
                Cabecera.NVendedor = BalTick.Vendor
                Cabecera.NBalanza = BalTick.Machine
                Cabecera.NCliente = BalTick.Client
                '1.8.8
                Cabecera.lEncargo = BalTick.lEncargo
                ''''''
                '////////////////////////////////////////
                ' IV4TQ, recompone el mostrador y balanza
                If (MachineID > 500 And MachineID < 600) Or lMarteAlfa Then
                    Cabecera.NBalanza = MaquinaTQ
                    Cabecera.NMostrador = SeccionTQ
                End If
                '////////////////////////////////////////
                Buffer = BalTick.Hour
                Cabecera.hora = TimeSerial(Val(Left(Buffer, 2)), Val(Mid(Buffer, 3, 2)), Val(Right(Buffer, 2)))
                Buffer = BalTick.Date
                Cabecera.Fecha = DateSerial(Val(Right(Buffer, 2)), Val(Mid(Buffer, 3, 2)), Val(Left(Buffer, 2)))
                If BalTick.positive = 1 Then
                    Cabecera.ImporteTotal = BalTick.amount
                Else
                    Cabecera.ImporteTotal = (-1) * BalTick.amount
                End If
                If MachineID <> 0 Then
                    Cabecera.ImporteBruto = Cabecera.ImporteTotal
                    Cabecera.descuento = 0
                Else
                    Cabecera.ImporteBruto = BalTick.P1
                    Cabecera.descuento = BalTick.P2
                End If
                
                'c2f 2.0.18 ver aquí
                If Cabecera.descuento <> 0 Then
                    Dim bTpC As Byte
                    'Dim bImpoNetoCab As Double
                    
                    bTpC = Cabecera.descuento - CByte(128)
                    
                    'c2f para 2.0.18
                    'If UsaEuro Then
                    '    bImpoNetoCab = redondea(Linea(Bucle).ImporteLinea / 100)
                    'Else
                    '    bImpoNetoCab = redondea(Linea(Bucle).ImporteLinea / (10 ^ decimales))
                    'End If
                    
                    ''''''''''''''''
                    'If (Linea(Bucle).DescuentoLinea > 0) Then
                    '            bTpC = Linea(Bucle).DescuentoLinea - CByte(128)
                    '            bimpoBrutoLin = redondea(tools.importe_bruto(bimpoNetoLin, bTpC))
                    '            .Fields("importe_bruto") = bimpoBrutoLin
                    'End If
                    ''''''''''''
                
                
                    Cabecera.ImporteTotal = Cabecera.ImporteBruto - ((Cabecera.ImporteBruto * bTpC) / 100)
                    'If BalTick.positive = 1 Then
                    'Else
                    '    Cabecera.ImporteTotal = (-1) * Cabecera.ImporteTotal
                    '    Cabecera.ImporteBruto = (-1) * Cabecera.ImporteBruto
                    'End If
                
                End If
                '''''''''''
                
                
                '///cas.v117
                Cabecera.tipoPago = BalTick.tipoPago
                Cabecera.cantidadEntre = BalTick.cantidadEntre
                Cabecera.Cambio = Cabecera.cantidadEntre - Cabecera.ImporteBruto
                If (Cabecera.tipoPago = 255) Or (Cabecera.cantidadEntre = 0) Then Cabecera.Cambio = 0
                '///cas.v117
                If Not (IsNumeric(MiTipo)) Then MiTipo = 0
                If aModTrab > 3 Then
                    Cabecera.EstadoTiquet = 17
                    Cabecera.CadenaEstadoTiquet = "Local"
                    Cabecera.tipoTiquet = MiTipo + 1
                Else
                    Cabecera.EstadoTiquet = 16
                    Cabecera.CadenaEstadoTiquet = "Red"
                    Cabecera.tipoTiquet = MiTipo + 1
                End If
                If BalTick.Cancelled = 1 Then
                    '1.8.4 fanjul
                    'qué hacer?
                    '''''''''''''
                    Cabecera.EstadoTiquet = Cabecera.EstadoTiquet + 64
                    Cabecera.CadenaEstadoTiquet = Cabecera.CadenaEstadoTiquet & " Cancelado"
                End If
                
                Cabecera.lote = BalTick.lote
                
                '*******************
                ' Pide las lineas
                '*******************
                'c2f 1.7.8
                If TipoC9 Then
                    On Error GoTo error_C9
                End If
                ''''''''''
                CuentaLineas = 1
                LineaInicial = 1
                ToTalLineas = Cabecera.NLineas
                For Bucle = 1 To Cabecera.NLineas
                    
                    nContLinA = nContLinA + 1
                    
                    'STV (positivo o negativo)
                    Linea(Bucle).ModificadoTemporal = False
                    If BalTick.Lines(Bucle - 1).positive = 1 Then
                         Linea(Bucle).FactorPeso = 1
                         Linea(Bucle).TipoLinea = 0
                    Else
                         Linea(Bucle).FactorPeso = -1
                         Linea(Bucle).TipoLinea = 2
                    End If
                    'STP (pesado o no pesado)
                    If BalTick.Lines(Bucle - 1).Weighted = 1 Then
                         Linea(Bucle).EsPesado = True
                    Else
                         Linea(Bucle).EsPesado = False
                         Linea(Bucle).TipoLinea = Linea(Bucle).TipoLinea + 1
                    End If
                    ' STC (cancelado o no cancelado)
                    If BalTick.Lines(Bucle - 1).Cancelled = 0 Then
                         Linea(Bucle).LineaCancelada = False
                    Else
                         Cabecera.LineasCanceladas = Cabecera.LineasCanceladas + 1
                         Linea(Bucle).LineaCancelada = True
                         Linea(Bucle).TipoLinea = Linea(Bucle).TipoLinea + 16
                    End If
                    Linea(Bucle).FichaVacuno = BalTick.Lines(Bucle - 1).BovineCard
                    Linea(Bucle).NumeroLinea = BalTick.Lines(Bucle - 1).Number
                    Linea(Bucle).DescuentoLinea = BalTick.Lines(Bucle - 1).Discount
                    Linea(Bucle).CodigoArticulo = BalTick.Lines(Bucle - 1).Item_Code
                    '**************************************
                    ' IV4TQ recompone el código de artículo
                    '**************************************
                    Linea(Bucle).precio = BalTick.Lines(Bucle - 1).Item_Price
                    If BalTick.Lines(Bucle - 1).positive = 1 Then
                        Linea(Bucle).ImporteLinea = BalTick.Lines(Bucle - 1).amount
                        Linea(Bucle).cantidad = BalTick.Lines(Bucle - 1).Item_Pes
                    Else
                        Linea(Bucle).ImporteLinea = BalTick.Lines(Bucle - 1).amount
                        Linea(Bucle).cantidad = BalTick.Lines(Bucle - 1).Item_Pes
                    End If
                    '*****************************
                    If Linea(Bucle).EsPesado = True Then Linea(Bucle).cantidad = (Linea(Bucle).cantidad / 1000)
                    '**************************
                    ' ML100 --> Si el peso está
                    ' entre 15 y 30 kg, sólo
                    ' proporciona dos decimales
                    ' de peso, y hay que corregirlo
                    '******************************
                    '******************************cas.v1.1.6
                    If ML100 And aTipoTqt <> 2 And (MachineID = 0 Or MachineID = 1) And Linea(Bucle).EsPesado Then
                        If Abs(Linea(Bucle).ImporteLinea) >= 999 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                            Linea(Bucle).cantidad = Linea(Bucle).cantidad * 1000
                        Else
                           If Abs(Linea(Bucle).ImporteLinea) >= 99 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                               Linea(Bucle).cantidad = Linea(Bucle).cantidad * 100
                           Else
                               If Abs(Linea(Bucle).ImporteLinea) >= 9 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                                   Linea(Bucle).cantidad = Linea(Bucle).cantidad * 10
                               End If
                           End If
                        End If
                    End If
                    '***************************************cas.v1.1.6
                    ' verifica si se ha modificado el precio
                    '***************************************
                    If Linea(Bucle).CodigoArticulo <> 0 Then
                        If aTipoTqt = 2 Then
                            sSQL = "SELECT codigo " & _
                                   "FROM articulo " & _
                                   "WHERE plu=" & Linea(Bucle).CodigoArticulo & _
                                   " AND codi_ident=" & Cabecera.NMostrador & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If Not RegArt.EOF Then
                                codArt = RegArt.Fields("codigo")
                                Linea(Bucle).CodigoArticulo = codArt
                            End If
                            RegArt.Close
                            Set RegArt = Nothing
                        Else
                            codArt = Linea(Bucle).CodigoArticulo
                        End If
                        If Not (TipoC9) And Not (MachineID > 500 And MachineID < 600) Then
                            sSQL = "SELECT * " & _
                                   "FROM articulo " & _
                                   "WHERE codigo=" & CStr(codArt) & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If RegArt.EOF Then
                                 Linea(Bucle).SehaModificado = False
                                 Linea(Bucle).IVA = 0
                            Else
                                ScaleArt.codigo = 0
                                Linea(Bucle).IVA = RegArt!tipo_iva
                                sSQL = "SELECT * " & _
                                       "FROM tramos " & _
                                       "WHERE borrado=false and codigo=" & CStr(codArt)
                                Set RegTramo = Base.OpenRecordset(sSQL)
                                If RegTramo.EOF Then
                                   HayTramos = False
                                Else
                                   HayTramos = True
                                End If
                                RegTramo.Close
                                If HayTramos = False Then
                                   If UsaEuro Then
                                       If (Linea(Bucle).precio / 100) <> RegArt!Euros Then
                                           Linea(Bucle).SehaModificado = True
                                       Else
                                           Linea(Bucle).SehaModificado = False
                                       End If
                                   Else
                                       If (Linea(Bucle).precio / (10 ^ decimales)) <> RegArt!precio Then
                                           Linea(Bucle).SehaModificado = True
                                       Else
                                           Linea(Bucle).SehaModificado = False
                                       End If
                                   End If
                                   'If Linea(Bucle).SehaModificado And (MachineID = 0 Or MachineID = 1) Then
                                   If Linea(Bucle).SehaModificado And Not (MachineID > 500 And MachineID < 600) And Not (TipoC9) Then
                                       '*****************************************
                                       '* consulta el precio real del artículo  *
                                       '*****************************************
                                       If MachineID = 0 Then
                                           ScaleArt = Binario_PideArticulo(RegArt!codigo, False)
                                       Else
                                       End If
                                       If UsaEuro Then
                                           ScaleArt.precio = ScaleArt.precio / 100
                                       Else
                                           ScaleArt.precio = ScaleArt.precio / (10 ^ decimales)
                                       End If
                                       If (UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!Euros)) Or _
                                       (Not UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!precio)) Then
                                           If AbsorverCambios Then
                                               MyArt.codigo = Linea(Bucle).CodigoArticulo
                                               MyArt.plu = 0
                                               MyArt.Mostrador = 0
                                               MyArt.precio = ScaleArt.precio
                                               If CSiempre Then
                                                   OKCambiar = True
                                               Else
                                                   OKCambiar = False
                                                   If IsNull(RegArt!PRC3) Then
                                                       OKCambiar = True
                                                   Else
                                                       If RegArt!PRC3 < (Val(Format(Cabecera.Fecha, "yymmdd")) * 10000 + Val(Format(Cabecera.hora, "hhmm"))) Then OKCambiar = True
                                                   End If
                                               End If
                                               If OKCambiar Then
                                                   Modificacion_Articulo_Precio MyArt, "Balanzas", True
                                               End If
                                           End If
                                       Else
                                           Linea(Bucle).ModificadoTemporal = True
                                           Linea(Bucle).SehaModificado = False
                                       End If
                                   End If
                                End If
                            End If
                            On Error Resume Next
                            RegArt.Close

                            Set RegArt = Nothing
                            On Error GoTo 0
                            
                        End If
                    End If
                    MiContador = MiContador + 1
                Next Bucle
                '********************************
                ' Actualiza las tablas de totales
                '********************************
                Cabecera.Incluir_en_Base = True
                Actualiza_Tabla_Cabecera Base, Cabecera, baseID, Linea
                If Cabecera.Incluir_en_Base = True Then
                    If Dir(App.Path & "\gigante.txt") <> "" And lFinDia Then
                        If nContDiasGi = 0 Then
                            nContDiasGi = 1
                            dArrGi(nContDiasGi) = Cabecera.Fecha
                        Else
                            For nMyBu = 1 To nContDiasGi
                                If dArrGi(nMyBu) = Cabecera.Fecha Then
                                    nMyBu = 9999
                                    Exit For
                                End If
                                
                            Next nMyBu
                            If nMyBu < 9999 Then
                                nContDiasGi = nContDiasGi + 1
                                dArrGi(nContDiasGi) = Cabecera.Fecha
                            End If
                        End If
                    End If
                    
                    Actualiza_Tabla_Tickets Base, Cabecera, Linea, baseID
                    '1.7.190
                    If Cabecera.Incluir_en_Base = True Then
                    
                        If Not (lMarteAlfa) Then
                            On Error Resume Next
                            Actualiza_Tabla_Gtarti Base, Cabecera, Linea, aTipoTqt
                            If Err.Number <> 0 Then
                                CadenadeLog "se ha producido al actualizar..."
                            End If
                            On Error GoTo 0
                            
                            Actualiza_Tabla_Gtsecs Base, Cabecera, Linea, True
                            Actualiza_Tabla_GtVend Base, Cabecera, Linea, True
                        End If
                        actualiza_tabla_gtInfVen Base, Cabecera, Linea, True '\\\cas.v118
                        actualiza_tabla_gtHora Base, Cabecera, Linea, True '\\\cas.v170
                        Actualiza_Tabla_Log Base, Cabecera, Linea
                    
                    End If
                End If
            End If
        End If
        
        If canCelar = 1 Then Repetir = False
    
    Loop
    
    If nNumLinA <> 0 Then
        If nNumLinA <> nContLinA Then
            canCelar = 1
            CadenadeLog "Lineas de tique C9 distintas a recibidas"
            CadenadeLog "   Recibidas:" & Format(nContLinA, "00000")
            CadenadeLog "Lineas en C9:" & Format(nNumLinA, "00000")
        Else
            CadenadeLog "Lineas de tique C9 iguales a recibidas"
            CadenadeLog "   Recibidas:" & Format(nContLinA, "00000")
            CadenadeLog "Lineas en C9:" & Format(nNumLinA, "00000")
        
        End If
    End If
    On Error Resume Next
    Base.Close
    On Error GoTo 0
    Exit Sub 'c2f 1.7.8
error_C9: 'c2f 1.7.8
    Base.Close
    On Error GoTo 0
    canCelar = 1
End Sub
'\\\\\\
'//////
Public Sub CG_BorraTotales(MachineID As Integer)
    Dim MyFlag As Long
    Dim resp, nC As Long
    Dim BufEnable As Boolean
    BufEnable = frmMonitor.cmdmonitorstop.Enabled
    frmMonitor.cmdmonitorstop.Enabled = False
    MyFlag = 0
    If MachineID >= 700 And GA_Reinit Then MyFlag = 1
    If Not UsaModem Then
        frmControl.Switch_mode (MachineID)
    End If
    'frmMonitor.MostrarDato CargaCadena(828)
    Do_Events
    '1.7.20
    frmControl.Balanza.Set_Timeout (15)
    'If usaSC10 And Not (UsaGamaAlta) Then
    '    frmControl.Balanza.Set_Timeout (60)
    'End If
    '''''''
    resp = frmControl.Balanza.Erase_Totals(MyFlag)
    
    If usaSC10 And (Not (UsaGamaAlta)) And (Not (TipoC9)) Then
        resp = frmControl.Balanza.Set_NetStatus(1)
        If resp <> 0 Then
            For nC = 1 To 100
                
                Sleep (1000)
                Do_Events
                
                frmControl.Switch_mode 1
                resp = frmControl.Balanza.Set_NetStatus(1)
                
                If resp = 0 Then Exit For
                
            Next nC
        End If
    End If
    
    frmControl.Balanza.Set_Timeout (5)
    
    If resp = 0 Then
        frmMonitor.MostrarDato CargaCadena(282) & " --> OK"
    Else
        frmMonitor.MostrarDato CargaCadena(282) & " --> NOK"
    End If
    frmMonitor.cmdmonitorstop.Enabled = BufEnable
End Sub
Public Sub CG_EnviaFamilia(Scaletype As Integer, Modificaciones As Boolean)
    '*******************************
    ' sólo SC10 - C9/C10
    '*******************************
    Dim RegArticulo As Recordset
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Bucle As Long
    Dim Contador As Long
    Dim OrdenDat As Integer
    Dim NumOrden As Long
    Dim SeEnvia() As Boolean
    Dim MiRecibe As Long
    Dim SC10NoMarcar As Boolean
    If (Scaletype = 1 Or Scaletype = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    frmControl.Switch_mode Scaletype
    Set Base = OpenDatabase(Base_General)
    If Modificaciones Then
        Set Registro = Base.OpenRecordset _
        ("select * from familia where tran_fam<>'*' or tran_fam is null")
    Else
        Set Registro = Base.OpenRecordset("select * from familia")
    End If
    '*****************************************************************
    ' Cuenta el número de Subsecciones
    '*****************************************************************
    Contador = 0
    If Registro.EOF Then Exit Sub
    With Registro
        .MoveLast
        Contador = .Recordcount
    End With
    ' Abre el fichero orden.dat
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(678)   '"ENVÍO DE FAMILIAS"
    With Registro
        .MoveFirst
        NumOrden = 1
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = Contador + 2
        For Bucle = 1 To Contador
            ' ******************************
            ' Barra Progreso
            ' *******************************
            If canCelar = 1 Then Exit For
            MonitorBarra
            Do_Events
            '*****************************
            ' Crea
            ReDim Preserve SeEnvia(Bucle + 1)
            Set Reg2 = Base.OpenRecordset _
            ("select * from seccion where borrado=false and secc_maqui=" & !secc_Maqui)
            If Reg2!enviardatos = "TODOS" Or Reg2!enviardatos = "SC10" Or _
            Reg2!multiple = 1 Then
                SeEnvia(Bucle) = True
                If Reg2!multiple = 1 Then
                    Set RegArticulo = Base.OpenRecordset _
                    ("select * from articulo where borrado=false and balenv=true " & _
                    " and secc_maqui=" & !secc_Maqui & " and codi_sub=" & !codi_sub _
                    & " and codi_fam=" & !codi_fam)
                    If RegArticulo.EOF Then
                        SeEnvia(Bucle) = False
                    End If
                End If
                If SeEnvia(Bucle) Then
                    MiRecibe = frmControl.Balanza.Send_SubsecFamily(.Fields("secc_maqui"), .Fields("codi_sub"), .Fields("codi_fam"))
                    If MiRecibe >= 1000 Then
                        frmMonitor.MostrarDato CargaCadena(544)
                        Exit For
                    End If
                    If MiRecibe <> 0 Then
                        frmMonitor.MostrarDato "Fam. " & Format(!codi_fam, "0000") & " - Sub. " & Format(!codi_sub, "0000") & "-->NOK"
                    Else
                        frmMonitor.MostrarDato "Fam. " & Format(!codi_fam, "0000") & " - Sub. " & Format(!codi_sub, "0000") & "-->OK"
                        If Not SC10NoMarcar Then
                            Edit_Record Registro
                            !tran_fam = "*"
                            .Update
                        End If
                    End If
                End If
            End If
            .Movenext
        Next Bucle
    End With
    CerrarBase Base
    
End Sub
Public Sub CG_EnviaSubSec(Scaletype As Integer, Modificaciones As Boolean)
    '/////////////////////////////////////
    '// sólo para SC10                  //
    '/////////////////////////////////////
    Dim Base As Database
    Dim Registro As Recordset
    Dim Bucle As Long
    Dim Contador As Long
    Dim OrdenDat As Integer
    Dim NumOrden As Long
    Dim RegSeccion As Recordset
    Dim RegArticulo As Recordset
    Dim SeMarca() As Boolean
    Dim MiRecibe As Long
    Dim SC10NoMarcar As Boolean
    If (Scaletype = 1 Or Scaletype = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    Set Base = OpenDatabase(Base_General)
    frmControl.Switch_mode Scaletype
    If Modificaciones Then
        Set Registro = Base.OpenRecordset _
        ("select * from subsec where tran_sub<>'*' or tran_sub is null order by borrado")
    Else
        Set Registro = Base.OpenRecordset("select * from subsec order by borrado")
    End If
    '*****************************************************************
    ' Cuenta el número de Subsecciones
    '*****************************************************************
    Contador = 0
    If Registro.EOF Then Exit Sub
    With Registro
        .MoveLast
        Contador = .Recordcount
    End With
    ' Abre el fichero orden.dat
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(677)   '"ENVÍO DE SUBSECCIONES"
    With Registro
        .MoveFirst
        NumOrden = 1
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = Contador + 2
        For Bucle = 1 To Contador
            ' ******************************
            ' Barra Progreso
            ' *******************************
            If canCelar = 1 Then Exit For
            MonitorBarra
            Do_Events
            '*****************************
            ' Crea
            ' busca correspondencia sección máquina
            ReDim Preserve SeMarca(Bucle + 1)
            Set RegSeccion = Base.OpenRecordset _
            ("select * from seccion where codi_ident=" & !codi_ident)
            ' Envía
            If RegSeccion!enviardatos = "TODOS" Or RegSeccion!enviardatos = "SC10" _
            Or RegSeccion!multiple = 1 Then
                If RegSeccion!multiple = 1 And RegSeccion!enviardatos = "GA" Then
                    Set RegArticulo = Base.OpenRecordset _
                    ("select * from articulo where balenv=true and borrado=false and codi_sub=" _
                    & !codi_sub & " and codi_ident=" & !codi_ident)
                    If Not RegArticulo.EOF Then
                        SeMarca(Bucle) = True
                    Else
                        SeMarca(Bucle) = False
                    End If
                Else
                    SeMarca(Bucle) = True
                End If
                If SeMarca(Bucle) Then
                    MiRecibe = frmControl.Balanza.Send_SubsecFamily(RegSeccion.Fields("secc_maqui"), .Fields("codi_sub"), 0)
                    If MiRecibe >= 1000 Then
                        frmMonitor.MostrarDato CargaCadena(544)
                        Exit For
                    Else
                        If MiRecibe <> 0 Then
                            frmMonitor.MostrarDato "Sub. " & Format(!codi_sub, "0000") & " - Sec. " & Format(RegSeccion!secc_Maqui, "00") & "-->NOK"
                        Else
                            frmMonitor.MostrarDato "Sub. " & Format(!codi_sub, "0000") & " - Sec. " & Format(RegSeccion!secc_Maqui, "00") & "-->OK"
                            If Not SC10NoMarcar Then
                                Edit_Record Registro
                                !tran_sub = "*"
                                .Update
                            End If
                        End If
                    End If
                End If
            End If
            RegSeccion.Close
            .Movenext
        Next Bucle
    End With
    CerrarBase Base
End Sub
Public Sub CG_EnviaPublicidad(MachineID As Integer, SoloModificaciones As Boolean)
    Dim CadenaModificacion As String
    Dim envio As String
    Dim Cadenas() As Epel_Publi
    Dim Bucle As Long
    Dim BucleGrande As Integer
    Dim Buclelite As Integer
    Dim Condicion As String
    Dim Contador As String
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim La_Respuesta As Integer
    Dim MyStr As String
    Dim MyB As Integer
    Dim BSecEqp As Integer
    Dim MyCond As Boolean
    Dim MyAux As String
    Dim Borrados() As Boolean
    Dim MyTran(3) As String
    Dim SC10NoMarcar As Boolean
    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    frmControl.Switch_mode MachineID
    
  
    MyStr = " (modelo=100) "
    MyTran(0) = "tran_pubsc10"
    MyTran(1) = "tran_ofesc10"
    MyTran(2) = "tran_actsc10"

    Contador = 0
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(14)   '"ENVÍO DE PUBLICIDAD"
    If SoloModificaciones Then
        CadenaModificacion = " and (tran_pubsc10<>'*' or tran_ofesc10<> '*' or tran_actsc10<>'*' ) and borrado="
    Else
        CadenaModificacion = " and borrado="
    End If
    Set Base = OpenDatabase(Base_General)
    Condicion = "true"
    For BucleGrande = 0 To 1
        For BSecEqp = 0 To 1
            If MachineID = 2 And BSecEqp = 1 Then Exit For
            If BSecEqp = 0 Then
                Set Registro = Base.OpenRecordset("select distinct secc_maqui from publi")
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from publi where ((not isnull(numero_eqp)) and numero_eqp <> 0)" & CadenaModificacion & Condicion)
            End If
            If Not Registro.EOF Then
                Registro.MoveFirst
                Do Until Registro.EOF
                    If BSecEqp = 0 Then
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and " & MyStr & " and secc_maqui=" & Registro!secc_Maqui)
                    Else
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and numero_eqp=" & Registro!numero_eqp & " and " & MyStr)
                    End If
                    If Not Reg2.EOF Then
                        MyCond = True
                        If BSecEqp = 0 Then
                            Set Reg3 = Base.OpenRecordset _
                            ("select * from publi where secc_maqui=" & Registro!secc_Maqui & " and (isnull(numero_eqp) or numero_eqp=0) " & CadenaModificacion & Condicion)
                            If Reg3.EOF Then MyCond = False
                        Else
                            Set Reg3 = Base.OpenRecordset _
                            ("select * from publi where numero_eqp=" & Registro.Fields("numero_eqp") & CadenaModificacion & Condicion)
                            If Reg3.EOF Then MyCond = False
                        End If
                        If MyCond Then
                            ' Activación / Desactivación
                            If Not SoloModificaciones Or Reg3.Fields(MyTran(0)) <> "*" Then
                                Contador = Contador + 1
                                ReDim Preserve Cadenas(Contador + 1)
                                ReDim Preserve Borrados(Contador + 1)
                                Select Case Reg3.Fields("prog_pub")
                                Case 0
                                    Cadenas(Contador).Status = 0
                                Case 10
                                    Cadenas(Contador).Status = 1
                                Case 1
                                    Cadenas(Contador).Status = 2
                                End Select
                                Cadenas(Contador).Mode = 0
                                Cadenas(Contador).Number = Reg3.Fields("secc_maqui")
                                If BucleGrande = 1 Then
                                    Cadenas(Contador).Banner = Reg3.Fields("publicidad")
                                    For MyB = 1 To 10
                                        Cadenas(Contador).Offers(MyB - 1) = Reg3.Fields("oferta" & Format(MyB, "00"))
                                    Next MyB
                                    Borrados(Contador) = False
                                Else
                                    Borrados(Contador) = True
                                    Cadenas(Contador).Banner = ""
                                    For MyB = 1 To 10
                                        Cadenas(Contador).Offers(MyB - 1) = ""
                                    Next MyB
                                End If
                            End If
                        End If
                End If
                Reg2.Close
                Registro.Movenext
            Loop
        End If
        Next BSecEqp
        Condicion = "false"
    Next BucleGrande
    '**********************
    ' Envía los datos
    '**********************
    For Bucle = 1 To Contador
        If canCelar = 1 Then Exit For
        La_Respuesta = frmControl.Balanza.Send_Banners(Cadenas(Bucle))
        If La_Respuesta >= 1000 Then
            frmMonitor.MostrarDato CargaCadena(544)
            Exit For
        End If
        If Not ((La_Respuesta = 0) Or (La_Respuesta = 4 And Borrados(Bucle) = True)) Then
            frmMonitor.MostrarDato CargaCadena(231) & " : " & Cadenas(Bucle).Number & " --> NOK"
            
        Else
            frmMonitor.MostrarDato CargaCadena(231) & " : " & Cadenas(Bucle).Number & " --> OK"
            If Borrados(Bucle) = True Then
                MyAux = "true"
            Else
                MyAux = "false"
            End If
            If Cadenas(Bucle).Mode = 0 Then
                Set Registro = Base.OpenRecordset _
                ("select * from publi where (isnull(numero_eqp) or numero_eqp=0) and secc_maqui=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from publi where numero_eqp=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            End If
            With Registro
                If Not .EOF Then .MoveFirst
                Do Until .EOF
                    If Not SC10NoMarcar Then
                        Edit_Record Registro
                        .Fields(MyTran(0)) = "*"
                        .Fields(MyTran(1)) = "*"
                        .Fields(MyTran(2)) = "*"
                        
                        If Not UsaGamaAlta And Borrados(Bucle) Then
                            !tran_pub = "*"
                            !tran_act = "*"
                            !tran_ofe = "*"
                        End If
                   
                        .Update
                    End If
                    If Not .EOF Then .Movenext
                Loop
            End With
        End If
    Next Bucle
    '****************
    ' Marca los borrados que no corresponden al tipo de balanza indicado
    '****************
    Set Registro = Base.OpenRecordset("select * from publi where borrado=true")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If IsNull(!numero_eqp) Then
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and  " & MyStr & " and secc_maqui=" & !secc_Maqui)
                Else
                    Set Reg2 = Base.OpenRecordset _
                   ("select * from equipos where borrado=false and  " & MyStr & " and numero_eqp=" & !numero_eqp)
                End If
                If Reg2.EOF Then
                    Edit_Record Registro
                    .Fields(MyTran(0)) = "*"
                    .Fields(MyTran(1)) = "*"
                    .Fields(MyTran(2)) = "*"
                    
                        
                    If Not UsaGamaAlta Then
                        !tran_pub = "*"
                        !tran_act = "*"
                        !tran_ofe = "*"
                    End If
               
                    
                    .Update
                End If

                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
    
End Sub
Public Sub CG_EnviaCabecerasLeyendas(MachineID As Integer, SoloModificaciones As Boolean, Optional TL As Boolean)
    Dim CadenaModificacion As String
    Dim Cadenas() As Epel_HeadingLegend
    Dim Borrados() As Boolean
    Dim Bucle As Long
    Dim BucleGrande As Integer
    Dim Condicion As String
    Dim Contador As Integer
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim La_Respuesta As Integer
    Dim MiTabla As String
    Dim MyStr As String
    Dim BSecMaq As Integer
    Dim MyAux As String
    Dim SC10NoMarcar As Boolean
    If (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    
    MyStr = " (modelo=100) "

    Contador = 0

    MiTabla = "cabley"
    
    If SoloModificaciones Then
        CadenaModificacion = " and (tran_cabsc10<>'*') and borrado="
    Else
        CadenaModificacion = " and borrado="
    End If
    If Not TL Then
        frmMonitor.MostrarInfo CargaCadena(232)  '"ENVÍO DE CABECERAS Y LEYENDAS"
    Else
        frmMonitor.MostrarInfo "Textos Libres"
    End If
    Set Base = OpenDatabase(Base_General)
    ' Cabeceras/Leyendas por sección
    ' comprueba que la sección tiene máquinas de su tipo asociadas
    Condicion = "true"
    For BucleGrande = 0 To 1
        For BSecMaq = 0 To 1
        If BSecMaq = 0 Then
            Set Registro = Base.OpenRecordset _
            ("select distinct secc_maqui from " & MiTabla)
        Else
             Set Registro = Base.OpenRecordset _
            ("select * from " & MiTabla & " where ((not isnull(numero_eqp)) and numero_eqp<>0)" & CadenaModificacion & Condicion)
        End If
        If Not Registro.EOF Then
            Registro.MoveFirst
            Do Until Registro.EOF
                If BSecMaq = 0 Then
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and " & MyStr & " and secc_maqui=" & Registro!secc_Maqui)
                Else
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and numero_eqp=" & Registro!numero_eqp & " and " & MyStr)
                End If
                If Not Reg2.EOF Then
                        If BSecMaq = 0 Then
                            Set Reg3 = Base.OpenRecordset _
                            ("select * from " & MiTabla & " where secc_maqui=" & Registro!secc_Maqui & " and (isnull(numero_eqp) or numero_eqp=0) " & CadenaModificacion & Condicion)
                        Else
                            Set Reg3 = Base.OpenRecordset _
                            ("select * from " & MiTabla & " where numero_eqp=" & Registro.Fields("numero_eqp") & CadenaModificacion & Condicion)
                        End If
                        With Reg3
                            If Not .EOF Then
                                Contador = Contador + 1
                                ReDim Preserve Cadenas(Contador + 1)
                                ReDim Preserve Borrados(Contador + 1)
                                If BucleGrande = 0 Then Borrados(Contador) = True
                                If Not TL Then
                                    Cadenas(Contador).Heading1 = .Fields("cabecera01")
                                    Cadenas(Contador).Heading2 = .Fields("cabecera02")
                                    Cadenas(Contador).Heading3 = .Fields("cabecera03")
                                    Cadenas(Contador).Heading4 = .Fields("cabecera04")
                                    Cadenas(Contador).Heading5 = .Fields("cabecera05")
                                    Cadenas(Contador).Legend1 = .Fields("leyenda01")
                                    Cadenas(Contador).Legend2 = .Fields("leyenda02")
                                    Cadenas(Contador).Legend3 = .Fields("leyenda03")
                                    Cadenas(Contador).Legend4 = .Fields("leyenda04")
                                    Cadenas(Contador).Legend5 = .Fields("leyenda05")
                                    Cadenas(Contador).FHeading1 = Asc(.Fields("tletracab1")) - 48
                                    Cadenas(Contador).FHeading2 = Asc(.Fields("tletracab2")) - 48
                                    Cadenas(Contador).FHeading3 = Asc(.Fields("tletracab3")) - 48
                                    Cadenas(Contador).FHeading4 = Asc(.Fields("tletracab4")) - 48
                                    Cadenas(Contador).FHeading5 = Asc(.Fields("tletracab5")) - 48
                                    Cadenas(Contador).FLegend1 = Asc(.Fields("tletraley1")) - 48
                                    Cadenas(Contador).FLegend2 = Asc(.Fields("tletraley2")) - 48
                                    Cadenas(Contador).FLegend3 = Asc(.Fields("tletraley3")) - 48
                                    Cadenas(Contador).FLegend4 = Asc(.Fields("tletraley4")) - 48
                                    Cadenas(Contador).FLegend5 = Asc(.Fields("tletraley5")) - 48
                                Else
                                    Cadenas(Contador).Heading1 = .Fields("texto00")
                                    Cadenas(Contador).Heading2 = .Fields("texto01")
                                    Cadenas(Contador).Heading3 = .Fields("texto02")
                                    Cadenas(Contador).Heading4 = .Fields("texto03")
                                    Cadenas(Contador).Heading5 = .Fields("texto04")
                                    Cadenas(Contador).Legend1 = .Fields("texto05")
                                    Cadenas(Contador).Legend2 = .Fields("texto06")
                                    Cadenas(Contador).Legend3 = .Fields("texto07")
                                    Cadenas(Contador).Legend4 = .Fields("texto08")
                                    Cadenas(Contador).Legend5 = .Fields("texto09")
                                    Cadenas(Contador).FHeading1 = Asc(.Fields("tletra00")) - 48
                                    Cadenas(Contador).FHeading2 = Asc(.Fields("tletra01")) - 48
                                    Cadenas(Contador).FHeading3 = Asc(.Fields("tletra02")) - 48
                                    Cadenas(Contador).FHeading4 = Asc(.Fields("tletra03")) - 48
                                    Cadenas(Contador).FHeading5 = Asc(.Fields("tletra04")) - 48
                                    Cadenas(Contador).FLegend1 = Asc(.Fields("tletra05")) - 48
                                    Cadenas(Contador).FLegend2 = Asc(.Fields("tletra06")) - 48
                                    Cadenas(Contador).FLegend3 = Asc(.Fields("tletra07")) - 48
                                    Cadenas(Contador).FLegend4 = Asc(.Fields("tletra08")) - 48
                                    Cadenas(Contador).FLegend5 = Asc(.Fields("tletra09")) - 48
                                    
                                End If
                                If BSecMaq = 0 Then
                                    Cadenas(Contador).Mode = 0
                                    Cadenas(Contador).Number = !secc_Maqui
                                Else
                                    Cadenas(Contador).Mode = 1
                                    Cadenas(Contador).Number = !numero_eqp
                                End If
                                If TL Then Cadenas(Contador).Mode = Cadenas(Contador).Mode + 2
                            End If
                    .Close
                    End With
                End If
                Reg2.Close
                Registro.Movenext
            Loop
        End If
        Next BSecMaq
        Condicion = "false"
    Next BucleGrande
    '**********************
    ' Envía los datos
    '**********************
    If Contador >= 1 Then frmControl.Switch_mode MachineID
    For Bucle = 1 To Contador
        If canCelar = 1 Then Exit For
        On Error Resume Next
        La_Respuesta = frmControl.Balanza.Send_HeadingLegend(Cadenas(Bucle))
        On Error GoTo 0
        If La_Respuesta >= 1000 Then
            frmMonitor.MostrarDato CargaCadena(544)
            Exit For
        End If
        If Not (La_Respuesta = 0 Or (La_Respuesta = 4 And Borrados(Bucle) = True)) Then
            frmMonitor.MostrarDato Cadenas(Bucle).Number & " --> NOK"
            
        Else
            frmMonitor.MostrarDato Cadenas(Bucle).Number & " --> OK"
            If Borrados(Bucle) = True Then
                MyAux = "true"
            Else
                MyAux = "false"
            End If
            If Cadenas(Bucle).Mode = 0 Then
                Set Registro = Base.OpenRecordset _
                ("select * from " & MiTabla & " where (isnull(numero_eqp) or numero_eqp=0) and secc_maqui=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from " & MiTabla & " where numero_eqp=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            End If
            With Registro
                If Not .EOF Then .MoveFirst
                Do Until .EOF
                    If Not SC10NoMarcar Then
                        Edit_Record Registro
                        If Not TL Then
                            If (Not UsaECO1) And Borrados(Bucle) Then .Fields("tran_en") = "*"
                            If (Not UsaECO2) And Borrados(Bucle) Then .Fields("tran_el") = "*"
                            If (Not UsaEcoPrint) And Borrados(Bucle) Then .Fields("tran_ep") = "*"
                            
                            !tran_cabSC10 = "*"
                            If (Not UsaGamaAlta) And Borrados(Bucle) Then
                                !tran_cab = "*"
                            End If
                           
                        Else
                            !tran_txt = "*"
                            !tran_txtsc10 = "*"
                        End If
                        .Update
                    End If
                    If Not .EOF Then .Movenext
                Loop
            End With
        End If
    Next Bucle
    '****************
    ' Marca los borrados que no corresponden al tipo de balanza
    '****************
    Set Registro = Base.OpenRecordset("select * from " & MiTabla & " where borrado=true")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If IsNull(!numero_eqp) Then
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and " & MyStr & " and secc_maqui=" & !secc_Maqui)
                Else
                    Set Reg2 = Base.OpenRecordset _
                   ("select * from equipos where borrado=false and " & MyStr & " and numero_eqp=" & !numero_eqp)
                End If
                If Reg2.EOF Then
                    Edit_Record Registro
                    If Not TL Then
                        
                        !tran_cabSC10 = "*"
                       
                    Else
                        !tran_txt = "*"
                        !tran_txtsc10 = "*"
                    End If
                    .Update
                End If
                
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
    
End Sub
Public Sub CG_EnviaCodigosBarras(MachineID As Integer, SoloModificaciones As Boolean)
    Dim CadenaModificacion As String
    Dim envio As String
    Dim Cadenas() As Epel_BarCode
    Dim Borrados() As Boolean
    Dim Checksum As Integer
    Dim Bucle As Long
    Dim BucleGrande As Integer
    Dim BucleLocalRed As Integer
    Dim Condicion As String
    Dim Contador As String
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim La_Respuesta As Integer
    Dim MyStr As String
    Dim MyAux As String
    Dim YaDim As Boolean
    Dim BSecMaq As Integer
    Dim MySec As Integer
    Dim MyMaq As Integer
    Dim SC10NoMarcar As Boolean
    If (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    
    MyStr = " (modelo=100) "
    Contador = 0
    If SoloModificaciones Then
        CadenaModificacion = " and tran_barsc10<>'*' and borrado="
    Else
        CadenaModificacion = " and borrado="
    End If
    ' 1º Prepara los datos
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(23)   '"ENVÍO DE CÓDIGOS DE BARRAS"
    Set Base = OpenDatabase(Base_General)
    FrmSC10_MostrarMensajes = False
    Condicion = "true"
    For BucleGrande = 0 To 1
        For BSecMaq = 0 To 1
            If BSecMaq = 0 Then
                Set Registro = Base.OpenRecordset("select distinct secc_maqui from codbar where secc_maqui<1000")
            Else
                Set Registro = Base.OpenRecordset("select distinct numero_eqp from codbar where ((not isnull(numero_eqp)) and numero_eqp<>0)")
            End If
            If Not Registro.EOF Then
                Registro.MoveFirst
                Do Until Registro.EOF
                    If BSecMaq = 0 Then
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and " & MyStr & " and secc_maqui=" & Registro!secc_Maqui)
                    Else
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and " & MyStr & " and numero_eqp=" & Registro!numero_eqp)
                    End If
                    If Not Reg2.EOF Then
                        YaDim = False
                        For BucleLocalRed = 0 To 1
                            If BSecMaq = 0 Then
                                Set Reg3 = Base.OpenRecordset _
                                ("select * from codbar where secc_maqui=" & Registro!secc_Maqui & _
                                " and (isnull(numero_eqp) or numero_eqp=0) and secc_eqp=" & BucleLocalRed & CadenaModificacion & Condicion)
                            Else
                                Set Reg3 = Base.OpenRecordset _
                                ("select * from codbar where numero_eqp=" & Registro!numero_eqp & _
                                " and secc_eqp=" & BucleLocalRed & CadenaModificacion & Condicion)
                            End If
                            With Reg3
                                If Not .EOF Then
                                    If BSecMaq = 0 Then
                                        MySec = !secc_Maqui
                                        MyMaq = 0
                                    Else
                                        MySec = !secc_Maqui
                                        MyMaq = !numero_eqp
                                    End If
                                    If BucleLocalRed = 0 Then
                                        If Not YaDim Then
                                            Contador = Contador + 1
                                            ReDim Preserve Cadenas(Contador + 1)
                                            ReDim Borrados(Contador + 1)
                                            If BucleGrande = 0 Then Borrados(Contador) = True
                                            YaDim = True
                                        End If
                                        Select Case BSecMaq
                                            Case 0
                                                Cadenas(Contador).Mode = 0
                                                Cadenas(Contador).Number = !secc_Maqui
                                            Case 1
                                                Cadenas(Contador).Mode = 1
                                                Cadenas(Contador).Number = !numero_eqp
                                        End Select
                                        If Not Borrados(Contador) Then
                                         
                                            Cadenas(Contador).Net_Sale = EmularCodBar(MachineID, Trim(!CB_VNT), MySec, MyMaq)
                                            Cadenas(Contador).Net_Super = EmularCodBar(MachineID, Trim(!cb_sup), MySec, MyMaq)
                                            Cadenas(Contador).Net_Pack2 = EmularCodBar(MachineID, Trim(!cb_ev2), MySec, MyMaq)
                                            Cadenas(Contador).Net_Pack1 = EmularCodBar(MachineID, Trim(!cb_ev1), MySec, MyMaq)
                                            Cadenas(Contador).Net_Mix = EmularCodBar(MachineID, Trim(!cb_mix), MySec, MyMaq)
                                        End If
                                    Else
                                        If Not YaDim Then
                                            Contador = Contador + 1
                                            ReDim Preserve Cadenas(Contador + 1)
                                            ReDim Borrados(Contador + 1)
                                            If BucleGrande = 0 Then Borrados(Contador) = True
                                            YaDim = True
                                        End If
                                        Select Case BSecMaq
                                            Case 0
                                                Cadenas(Contador).Mode = 0
                                                Cadenas(Contador).Number = !secc_Maqui
                                            Case 1
                                                Cadenas(Contador).Mode = 1
                                                Cadenas(Contador).Number = !numero_eqp
                                        End Select
                                        If Not Borrados(Contador) Then
                                            Cadenas(Contador).Local_Sale = EmularCodBar(MachineID, Trim(!CB_VNT), MySec, MyMaq)
                                            Cadenas(Contador).Local_Super = EmularCodBar(MachineID, Trim(!cb_sup), MySec, MyMaq)
                                            Cadenas(Contador).Local_Pack1 = EmularCodBar(MachineID, Trim(!cb_ev1), MySec, MyMaq)
                                            Cadenas(Contador).Local_Pack2 = EmularCodBar(MachineID, Trim(!cb_ev2), MySec, MyMaq)
                                            Cadenas(Contador).Local_Mix = EmularCodBar(MachineID, Trim(!cb_mix), MySec, MyMaq)
                                        End If
                                    End If
                                End If
                                .Close
                            End With
                        Next BucleLocalRed
                    End If
                    Reg2.Close
                    Registro.Movenext
                Loop
            End If
        Next BSecMaq
        Condicion = "false"
    Next BucleGrande
    '**********************
    ' Envía los datos
    '**********************
    If Contador >= 1 Then frmControl.Switch_mode MachineID
    For Bucle = 1 To Contador
        If canCelar = 1 Then Exit For
        On Error Resume Next
        La_Respuesta = frmControl.Balanza.Send_BarCode(Cadenas(Bucle))
        On Error GoTo 0
        If La_Respuesta >= 1000 Then
            frmMonitor.MostrarDato CargaCadena(544)
            Exit For
        End If
        If Not (La_Respuesta = 0 Or (La_Respuesta = 4 And Borrados(Bucle) = True)) Then
            frmMonitor.MostrarDato CargaCadena(23) & " " & Cadenas(Bucle).Number & " --> NOK"
          
        Else
            frmMonitor.MostrarDato CargaCadena(23) & " " & Cadenas(Bucle).Number & " --> OK"
            If Borrados(Bucle) = True Then
                MyAux = "True"
            Else
                MyAux = "False"
            End If
            If Cadenas(Bucle).Mode = 0 Then
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where (isnull(numero_eqp) or numero_eqp=0) and secc_maqui=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where numero_eqp=" & Cadenas(Bucle).Number & CadenaModificacion & MyAux)
            End If
            With Registro
                If Not .EOF Then .MoveFirst
                Do Until .EOF
                    If Not SC10NoMarcar Then
                        Edit_Record Registro
                        If Borrados(Bucle) And (Not UsaECO1) Then !tran_en = "*"
                        If Borrados(Bucle) And (Not UsaECO2) Then !tran_el = "*"
                        If Borrados(Bucle) And (Not UsaECO2) Then !tran_ep = "*"
                        
                        !tran_barSC10 = "*"
                        If Borrados(Bucle) And (Not UsaGamaAlta) Then !tran_bar = "*"
                        .Update
                    End If
                    If Not .EOF Then .Movenext
                Loop
            End With
        End If
    Next Bucle
    '****************
    ' Marca los borrados que no corresponden a Euroscale o SC10
    '****************
    Set Registro = Base.OpenRecordset("select * from codbar where borrado=true")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If IsNull(!numero_eqp) Then
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and modelo<100 and secc_maqui=" & !secc_Maqui)
                Else
                    Set Reg2 = Base.OpenRecordset _
                   ("select * from equipos where borrado=false and modelo<100 and numero_eqp=" & !numero_eqp)
                End If
                If Reg2.EOF Then
                    Edit_Record Registro
                    !tran_barSC10 = "*"
                    .Update
                End If
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
Public Sub CG_Envia_Seccion(MachineID As Integer, Modificaciones As Boolean)
    Dim Bucle As Long
    Dim Contador As Long
    Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim MiRecibe As Integer
    Dim txt As String
    Dim SeccionAnterior As Integer
    Dim SeEnvia() As Boolean
    Dim MyStr As String
    Dim MyAux As String
    Dim SC10NoMarcar As Boolean
    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    canCelar = 0
    frmControl.Switch_mode MachineID
    Select Case MachineID
        Case 0
            MyStr = " (modelo=0 or modelo=1) "
            MyAux = "GA"
        Case 1 To 2
            MyStr = " (modelo=100) "
            MyAux = "SC10"
    End Select
    Set Base = OpenDatabase(Base_General)
    If Modificaciones Then
        Select Case MachineID
            Case 0
                Set Registro = Base.OpenRecordset("select * from seccion where (tran_sec<>'*' or tran_sec is null) ")
            Case 1 To 2
                Set Registro = Base.OpenRecordset("select * from seccion where (tran_secSC10<>'*' or tran_secSC10 is null) ")
        End Select
    Else
        Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
    End If
    '*****************************************************************
    ' Cuenta el número de líneas a tratar
    '*****************************************************************
    If Registro.EOF Then Exit Sub
    With Registro
        .MoveLast
        Contador = .Recordcount
    End With
    '**********************
    'Envía datos
    '**********************
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(9)   '"ENVÍO DE SECCIONES"
    With Registro
        .MoveFirst
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = Contador + 2
        SeccionAnterior = 0
        For Bucle = 1 To Contador
            ' ******************************
            ' Barra Progreso
            ' *******************************
            MonitorBarra
            If canCelar = 1 Then Exit For
            '*****************************
            ReDim Preserve SeEnvia(Bucle + 1)
            Set Reg2 = Base.OpenRecordset _
            ("select * from equipos where borrado=false and " & MyStr & " and secc_maqui=" & !secc_Maqui)
            If (Not Reg2.EOF) Or (!multiple = 1 Or !enviardatos = "TODOS" Or !enviardatos = MyAux) Then
                If SeccionAnterior <> !secc_Maqui Then
                    SeEnvia(Bucle) = True
                    SeccionAnterior = !secc_Maqui
                End If
            Else
                SeEnvia(Bucle) = False
                Edit_Record Registro
                Select Case MachineID
                    Case 0
                        !tran_sec = "*"
                        If Not usaSC10 And !borrado = True Then
                            !tran_secSC10 = "*"
                        End If
                    Case 1 To 2
                        !tran_secSC10 = "*"
                        If Not UsaGamaAlta And !borrado = True Then
                            !tran_sec = "*"
                        End If
                End Select
                .Update
            End If
            If IsNull(!descripcio) Then
                txt = ""
            Else
                txt = !descripcio
            End If
            If SeEnvia(Bucle) Then
                If canCelar = 1 Then Exit For
                MiRecibe = frmControl.Balanza.Send_Section(!secc_Maqui, txt)
                If MiRecibe >= 1000 Then
                    frmMonitor.MostrarDato CargaCadena(544)
                    Exit For
               End If
               If MiRecibe <> 0 Then
                    frmMonitor.MostrarDato CargaCadena(47) & " " & !secc_Maqui & "-->NOK"
               Else
                    frmMonitor.MostrarDato CargaCadena(47) & " " & !secc_Maqui & "-->OK"
                    If Not SC10NoMarcar Then
                         Edit_Record Registro
                         Select Case MachineID
                             Case 0
                                 !tran_sec = "*"
                                 If Not usaSC10 And !borrado = True Then
                                     !tran_secSC10 = "*"
                                 End If
                             Case 1 To 2
                                 !tran_secSC10 = "*"
                                 If Not UsaGamaAlta And !borrado = True Then
                                     !tran_sec = "*"
                                 End If
                         End Select
                        .Update
                   End If
                End If
            End If
            .Movenext
        Next Bucle
    End With
    CerrarBase Base
End Sub
Public Sub CG_envia_equipos(MachineID As Integer, Modificaciones As Boolean)
    Dim Base As Database
    Dim Registro As Recordset
    Dim MyStr As String
    Dim MyAux As String
    Dim MiRecibe As Long
    Dim SC10NoMarcar As Boolean
    Dim NoEnvia As Boolean
    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    frmControl.Switch_mode MachineID
    Select Case MachineID
        Case 0
            MyStr = " (modelo=0 or modelo=1) "
            MyAux = "GA"
        Case 1 To 2
            MyStr = " (modelo=100) "
            MyAux = "SC10"
    End Select
    Set Base = OpenDatabase(Base_General)
    If Modificaciones = True Then
        Set Registro = Base.OpenRecordset("select * from equipos where (tran_eqp<>'*' or tran_eqp is null) and " & MyStr)
    Else
        Set Registro = Base.OpenRecordset("select * from equipos where " & MyStr)
    End If
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(11)   '"ENVÍO DE EQUIPOS"
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If canCelar = 1 Then Exit Do
                If !borrado = True Then
                    '***************************
                    ' borrar equipo
                    '**************************
                    NoEnvia = False
                    If (MachineID = 1 Or MachineID = 2) And TipoC9 Then
                        Select Case ModoC9
                            Case 0
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 3
                                    Case 6
                                    Case Else
                                        NoEnvia = True
                                End Select
                            Case 1
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 6
                                    Case Else
                                        NoEnvia = True
                                End Select
                            Case 2
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 8
                                    Case Else
                                        NoEnvia = True
                                End Select
                        End Select
                    End If
                    If Not NoEnvia Then
                        MiRecibe = frmControl.Balanza.Send_Machine(Registro.Fields("numero_eqp"), 0, 0, 1)
                    Else
                        MiRecibe = 4
                    End If
                    If Not (MiRecibe = 4 Or MiRecibe = 0) Then
                        frmMonitor.MostrarDato CargaCadena(873) & " " & !numero_eqp & "-->NOK"
                        If MiRecibe >= 1000 Then Exit Do
                    Else
                        frmMonitor.MostrarDato CargaCadena(873) & " " & !numero_eqp & "-->OK"
                        If Not .EOF And (Not SC10NoMarcar) Then Delete_Record Registro
                    End If
                Else
                    '***************************
                    ' programar equipo
                    '**************************
                    If (MachineID = 1 Or MachineID = 2) And TipoC9 Then
                        Select Case ModoC9
                            Case 0
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 3
                                    Case 6
                                    Case Else
                                        NoEnvia = True
                                End Select
                            Case 1
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 6
                                    Case Else
                                        NoEnvia = True
                                End Select
                            Case 2
                                Select Case .Fields("secc_maqui")
                                    Case 1 To 8
                                    Case Else
                                        NoEnvia = True
                                End Select
                        End Select
                    End If
                    If Not NoEnvia Then
                        MiRecibe = frmControl.Balanza.Send_Machine(.Fields("numero_eqp"), .Fields("secc_maqui"), 0, 0)
                    Else
                        MiRecibe = 4
                    End If
                    If MiRecibe <> 0 Then
                        frmMonitor.MostrarDato CargaCadena(303) & " " & !numero_eqp & "-->NOK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(303) & " " & !numero_eqp & "-->OK"
                        If Not SC10NoMarcar Then
                            Edit_Record Registro
                            !tran_eqp = "*"
                            .Update
                        End If
                    End If
                End If
                If Not .EOF Then .Movenext
            Loop
            .MoveFirst
        End If
    End With
    CerrarBase Base
End Sub
Public Function CG_EnviaArticulo(MachineID As Integer, Modificaciones As Boolean, Optional MiCodigo As Long) As Boolean
    ' Programación básica
    Dim MiEnvio As Epel_FullItem
    Dim MiResp As Boolean
    Dim MiRecibe As Long
    Dim enviarBloque As Boolean
    Dim Registro As Recordset
    Dim RSTgrupo As Recordset
    Dim Contador As Long
    Dim Bucle As Long
    '****************************
    ' 2.0.9
    '****************************
    Dim ElMostrador() As Integer
    Dim RegMost As Recordset
    Dim BucleMost As Long
    Dim ContMost As Long
    Dim SETrata As Boolean
    Dim FactorPrecio As Long
    Dim Base As Database
    Dim BuclePlu As Integer
    Dim BufBorr As Integer
    Dim SC10NoMarcar As Boolean
    Dim StrSQL As String
    Dim Enviar4 As Boolean
    '''''''''''''''''''''
    Dim rstDesc2040 As Recordset
    Dim sFields As String
    Dim sSQLdesc2040 As String
    Dim okDesc2040 As Boolean
    ''''''''''''''''''''''''''
    Dim s As String
    '2.0.33 aho
    Dim RSTTec25 As Recordset
    '2.0.36
    'Dim nSself As Integer
    '''''''''''
    
    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    frmControl.Switch_mode MachineID
    MiResp = False
    '****************************
    If UsaEuro Then
        FactorPrecio = 100
    Else
        FactorPrecio = (10 ^ decimales)
    End If
    Set Base = OpenDatabase(Base_General)
    '*******************
    ' 2.0.9
    '*******************
    ContMost = 0
    Select Case MachineID
        Case 0
            Set RegMost = Base.OpenRecordset _
            ("select * from seccion where borrado=false and (enviardatos=" & Chr(34) & "GA" _
            & Chr(34) & " or enviardatos=" & Chr(34) & "TODOS" & Chr(34) & ")")
        Case 1 To 2
            Set RegMost = Base.OpenRecordset _
            ("select * from seccion where borrado=false and (enviardatos=" & Chr(34) & "SC10" _
            & Chr(34) & " or enviardatos=" & Chr(34) & "TODOS" & Chr(34) & ")")
    End Select
    With RegMost
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                ContMost = ContMost + 1
                ReDim Preserve ElMostrador(ContMost + 1)
                ElMostrador(ContMost) = !codi_ident
                .Movenext
            Loop
        End If
    End With
    If MiCodigo <> 0 Then
        StrSQL = "select * from articulo where borrado=false and codigo=" & MiCodigo
    Else
        If Modificaciones Then
            Select Case MachineID
                Case 0
                    If AhorraMas Then
                        StrSQL = _
                        "select * from articulo where " & _
                        "((tran_plu is null or tran_plu<>'*') or (borrado=true) or (tran_texto is null or tran_texto <> '*') or (tran_tx1 is null or tran_tx1 <> '*') or (tran_cbsc10 is null or tran_cb <> '*')) order by borrado,codigo"
                    
                    Else
                        StrSQL = _
                        "select * from articulo where " & _
                        "((tran_plu is null or tran_plu<>'*') or (tran_texto is null or tran_texto <> '*') or (tran_tx1 is null or tran_tx1 <> '*') or (tran_cbsc10 is null or tran_cb <> '*')) order by borrado,codigo"
                    End If
                Case 1 To 2
                    If AhorraMas Then
                        StrSQL = _
                        "select * from articulo where " & _
                        "((tran_pluSC10 is null or tran_pluSC10<>'*') or (borrado=true) or (tran_textoSC10 is null or tran_textoSC10 <> '*') or (tran_tx1sc10 is null or tran_tx1SC10 <> '*') or (tran_cbsc10 is null or tran_cbSC10 <> '*')) order by borrado,codigo"
                    
                    Else
                        StrSQL = _
                        "select * from articulo where " & _
                        "((tran_pluSC10 is null or tran_pluSC10<>'*') or (tran_textoSC10 is null or tran_textoSC10 <> '*') or (tran_tx1sc10 is null or tran_tx1SC10 <> '*') or (tran_cbsc10 is null or tran_cbSC10 <> '*')) order by borrado,codigo"
                    End If
            End Select
        Else
            StrSQL = "select * from articulo order by borrado,codigo"
        End If
    End If
    Set Registro = Base.OpenRecordset(StrSQL)
    '*****************************************************************
    ' Cuenta el número de Artículos
    '*****************************************************************
    Contador = 0
    If Registro.EOF Then Exit Function
    With Registro
        .MoveLast
        Contador = .Recordcount
    End With
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(12)   '"ENVÍO DE ARTÍCULOS"
    With Registro
        .MoveFirst
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = Contador + 2
        For Bucle = 1 To Contador
            If canCelar = 1 Then
                CerrarBase Base
                Exit Function
            End If
            ' ******************************
            ' Barra Progreso
            ' *******************************
            MonitorBarra
            '*****************************
            SETrata = False
            If !Balenv = True Then
                SETrata = True
            Else
                For BucleMost = 1 To ContMost
                    If ElMostrador(BucleMost) = !codi_ident Then
                        SETrata = True
                        Exit For
                    End If
                Next BucleMost
            End If
            '*******************************
            ' comprobación de envío a C9/C10
            '*******************************
            'If (MachineID = 1 Or MachineID = 2) And TipoC9 Then
            If (MachineID = 1 Or MachineID = 2) And TipoC9 And (var.C9Cod = 0) Then
                Select Case ModoC9
                    Case 0
                        If !secc_Maqui = 1 Or !secc_Maqui = 2 Or !secc_Maqui = 3 Then
                            If !plu > 999 Then SETrata = False
                        Else
                            If !secc_Maqui = 6 Then
                                If !plu > 499 Then SETrata = False
                            Else
                                SETrata = False
                            End If
                        End If
                    Case 1
                        If !secc_Maqui = 1 Or !secc_Maqui = 2 Or !secc_Maqui = 3 Or !secc_Maqui = 4 Or !secc_Maqui = 5 Then
                            If !plu > 599 Then SETrata = False
                        Else
                            If !secc_Maqui = 6 Then
                                If !plu > 499 Then SETrata = False
                            Else
                                SETrata = False
                            End If
                        End If
                    Case 2
                        If !secc_Maqui = 1 Or !secc_Maqui = 2 Or !secc_Maqui = 3 Or !secc_Maqui = 4 Or !secc_Maqui = 5 Or !secc_Maqui = 6 Or !secc_Maqui = 7 Then
                            If !plu > 399 Then SETrata = False
                        Else
                            If !secc_Maqui = 8 Then
                                If !plu > 499 Then SETrata = False
                            Else
                                SETrata = False
                            End If
                        End If
                End Select
            Else
                If TipoC9 And (var.C9Cod <> 0) Then
                    If !secc_Maqui = 1 Then
                        If !plu > 6999 Then
                            SETrata = False
                        End If
                    Else
                        SETrata = False
                    End If
                End If
            End If
            '************************
            ' fin comprobación C9/C10
            '************************
            If SETrata = False Then
                Edit_Record Registro
                Select Case MachineID
                    Case 0
                        !tran_plu = "*"
                        !tran_texto = "*"
                        !tran_tx1 = "*"
                        !tran_cb = "*"
                        If !borrado = True And Not usaSC10 Then
                           !tran_pluSC10 = "*"
                           !tran_textosc10 = "*"
                           !tran_tx1SC10 = "*"
                           !tran_cbsc10 = "*"
                        End If
                    Case 1 To 2
                        !tran_pluSC10 = "*"
                        !tran_textosc10 = "*"
                        !tran_tx1SC10 = "*"
                        !tran_cbsc10 = "*"
                        If !borrado = True And Not UsaGamaAlta Then
                           !tran_plu = "*"
                           !tran_texto = "*"
                           !tran_tx1 = "*"
                           !tran_cb = "*"
                        End If
                End Select
                .Update
            Else
                '*******************************
                ' 1º limpieza de datos anteriores
                '*******************************
                MiEnvio.Caducity = 0
                MiEnvio.Code = 0
                For BuclePlu = 0 To 19
                    MiEnvio.Desc(BuclePlu) = ""
                Next BuclePlu
                MiEnvio.EAN13 = ""
                MiEnvio.Etq = 0
                MiEnvio.Family = 0
                For BuclePlu = 0 To 8
                    MiEnvio.FDesc(BuclePlu) = 0
                Next BuclePlu
                MiEnvio.Font = 0
                MiEnvio.plu = 0
                MiEnvio.Pref = 0
                MiEnvio.Price = 0
                MiEnvio.sec = 0
                MiEnvio.subsec = 0
                MiEnvio.Tare = 0
                MiEnvio.TexT = ""
                MiEnvio.vat = 0
                MiEnvio.Weight = 0
                '********************************
                MiEnvio.Code = .Fields("codigo")
                MiEnvio.sec = .Fields("secc_maqui")
                MiEnvio.plu = .Fields("plu")
                '2.0.33 aho se hace plu=tecla para sección 25 balanzas self-service.
                If AhorraMas Then '2.0.36 And MiEnvio.sec = 25 Then
                    Set RSTTec25 = Base.OpenRecordset("select * from seccion where secc_maqui=" & CStr(MiEnvio.sec))
                    If Not RSTTec25.EOF Then
                        If RSTTec25.Fields("enviardatos") = "SC10" And RSTTec25.Fields("com") = 255 Then
                            RSTTec25.Close
                            Set RSTTec25 = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MiEnvio.sec) & " and codigo=" & CStr(MiEnvio.Code))
                            If Not RSTTec25.EOF Then
                                If RSTTec25.Fields("borrado") = False Then
                                    MiEnvio.plu = RSTTec25.Fields("codi_tec")
                                End If
                            End If
                        
                        End If
                    
                    End If
                    RSTTec25.Close
                    Set RSTTec25 = Nothing
                End If
                ''''''''
                MiEnvio.Presel = .Fields("ning")
                MiEnvio.Loss = .Fields("fcb")
                If (Not TipoC9) And (MachineID = 1) Then
                    MiEnvio.subsec = .Fields("codi_sub")
                Else
                    MiEnvio.subsec = 0
                End If
                MiEnvio.Family = .Fields("codi_fam")
                If UsaEcoPrint And TipoEcoPrint = 3 Then
                    If MiEnvio.Family = 0 Or MiEnvio.Family > 15 Then
                        MiEnvio.Family = 1
                    End If
                End If
                '******************
                If MachineID = 0 Then
                    If MiEnvio.Family > 255 Then MiEnvio.Family = 255
                End If
                '******************
                If UsaEuro Then
                    MiEnvio.Price = .Fields("euros") * FactorPrecio
                Else
                    MiEnvio.Price = .Fields("precio") * FactorPrecio
                End If
                If .Fields("codi_pes") = "W" Then
                    MiEnvio.Weight = 1
                Else
                    MiEnvio.Weight = 0
                End If
                MiEnvio.Caducity = .Fields("caducidad")
                MiEnvio.Font = .Fields("tip_let1")
                MiEnvio.TexT = .Fields("des_plu1")
                MiEnvio.Tare = .Fields("tara")
                If Not IsNull(.Fields("art_cb")) Then
                    MiEnvio.EAN13 = Trim(.Fields("art_cb"))
                Else
                    MiEnvio.EAN13 = ""
                End If
                MiEnvio.Etq = .Fields("etq")
                If MiEnvio.Etq = 255 Then MiEnvio.Etq = 0
                MiEnvio.Pref = .Fields("pref")
                MiEnvio.vat = .Fields("tipo_iva")
                For BuclePlu = 0 To 7
                    MiEnvio.FDesc(BuclePlu) = .Fields("tip_let" & Format(BuclePlu + 2, "0"))
                Next BuclePlu
                MiEnvio.FDesc(8) = .Fields("tip_let0")
                'MiEnvio.FDesc(8) = .Fields("tip_let0")
                
                If MiEnvio.plu = 53 Then
                    MiEnvio.plu = MiEnvio.plu
                End If
                
                For BuclePlu = 0 To 19
                    Select Case BuclePlu
                        Case 0 To 7
                            s = RTrim(.Fields("des_plu" & Format(BuclePlu + 2, "0")))
                            If Len(s) < 25 Then s = s & " "
                            MiEnvio.Desc(BuclePlu) = s 'Trim(.Fields("des_plu" & Format(BuclePlu + 2, "0")))
                        Case 8
                            s = RTrim(.Fields("des_plu0"))
                            If Len(s) < 25 Then s = s & " "
                            MiEnvio.Desc(BuclePlu) = s 'Trim(.Fields("des_plu0"))
                        Case 9
                            s = RTrim(.Fields("des_plux"))
                            If Len(s) < 25 Then s = s & " "
                            MiEnvio.Desc(BuclePlu) = s 'Trim(.Fields("des_plux"))
                        Case 10 To 19
                            s = RTrim(.Fields("des_plu" & Format(BuclePlu + 1, "00")))
                            If Len(s) < 25 Then s = s & " "
                            MiEnvio.Desc(BuclePlu) = s 'Trim(.Fields("des_plu" & Format(BuclePlu + 1, "00")))
                    End Select
                Next BuclePlu
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cas.v->171
                If (versionG) Then
                    'quitar
                    'CadenadeLog "CG_EnviaArticulo"
                    '''''''
                    sSQLdesc2040 = "SELECT * FROM linTxt2040 WHERE codigo= " & CStr(MiEnvio.Code)
                    Set rstDesc2040 = Base.OpenRecordset(sSQLdesc2040)
                    If (Not rstDesc2040.EOF) Then
                        okDesc2040 = True
                        rstDesc2040.MoveFirst
                        For BuclePlu = 0 To 19
                            sFields = "txt_" & CStr(BuclePlu + 21)
                            s = RTrim(rstDesc2040.Fields(sFields).Value)
                            If Len(s) < 25 Then s = s & " "
                            MiEnvio.Desc2040(BuclePlu) = s 'Trim(rstDesc2040.Fields(sFields).Value)
                            'quitar
                            'CadenadeLog Trim(rstDesc2040.Fields(sFields).Value)
                            '''''''
                        Next BuclePlu
                    End If
                    'rstDesc2040.Close
                End If
                
                If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) Then
                    Enviar4 = True
                    For BuclePlu = 4 To 19
                        If MiEnvio.Desc(BuclePlu) <> "" Then
                            Enviar4 = False
                            Exit For
                        End If
                    Next BuclePlu
                End If
                'c2f 1.7.4
                If Dir(App.Path & "\mixmode.txt") <> "" Then
                    MiEnvio.lMix = .Fields("lmix")
                    MiEnvio.nPeso = .Fields("poid")
                End If
                ''''''''''
                '2.0.22
                'If lVL000 Then
                '    MiEnvio.lMix = False
                '    MiEnvio.nPeso = .Fields("poid")
                'End If
                '2.0.37
                If lVL000 And Dir(App.Path & "\mixmode.txt") = "" Then
                    MiEnvio.lMix = False
                    MiEnvio.nPeso = .Fields("poid")
                End If
                '''''''
                '''''''
                '************************
                ' Ahorramas
                '************************
                'If AhorraMas Then
                '    If MachineID = 0 And MiEnvio.Family = 37 Then
                '        MiEnvio.Desc(11) = "N. REFERENCIA "
                '        MiEnvio.Desc(12) = "CEBADO EN: "
                '    End If
                '    If .Fields("secc_maqui") = 40 Or .Fields("secc_maqui") = 25 Then
                '        If MiEnvio.Weight = 1 Then
                '            MiEnvio.Desc(13) = "Euros/Kg"
                '        Else
                '            MiEnvio.Desc(13) = "Euros/und"
                '        End If
                '    End If
                'End If
                If AhorraMas Then
                    If MiEnvio.Family = 37 And MachineID = 0 And MiEnvio.Etq <> 3 Then
                        MiEnvio.Desc(11) = "N. REFERENCIA "
                        MiEnvio.Desc(12) = "CEBADO EN: "
                    End If
                    If .Fields("secc_maqui") = 40 Or .Fields("secc_maqui") = 25 Then
                        If MiEnvio.Weight = 1 Then
                            'If MiEnvio.Etq <> 3 Then
                                MiEnvio.Desc(13) = "Euros/Kg"
                            'Else
                            '    MiEnvio.Desc(0) = "Euros/Kg"
                            'End If
                        Else
                            'If MiEnvio.Etq <> 3 Then
                                MiEnvio.Desc(13) = "Euros/und"
                            'Else
                            '    MiEnvio.Desc(0) = "Euros/und"
                            'End If
                        End If
                    End If
                End If
                
                '************************
                ' Grupos Conservación
                '************************
                'If UsaGruposdeConservacion Then
                '    If .Fields("grupo_conserv") <> 0 Then
                '        Set RSTgrupo = Base.OpenRecordset("select * from grpconserv where grupo=" & .Fields("grupo_conserv"))
                '        If Not RSTgrupo.EOF Then
                '            Select Case MachineID
                '                Case 0
                '                    If Not IsNull(RSTgrupo.Fields("txt1")) Then
                '                        MiEnvio.Desc(8) = RSTgrupo.Fields("txt1")
                '                    Else
                '                        MiEnvio.Desc(8) = ""
                '                    End If
                '                    If Not IsNull(RSTgrupo.Fields("txt2")) Then
                '                        MiEnvio.Desc(9) = RSTgrupo.Fields("txt2")
                '                    Else
                '                        MiEnvio.Desc(9) = ""
                '                    End If
                '                Case 1 To 2
                '                    If Not IsNull(RSTgrupo.Fields("txtsc10")) Then
                '                        MiEnvio.Desc(8) = RSTgrupo.Fields("txtsc10")
                '                    Else
                '                        MiEnvio.Desc(8) = ""
                '                    End If
                '            End Select
                '        Else
                '            MiEnvio.Desc(8) = ""
                '            MiEnvio.Desc(9) = ""
                '        End If
                '    End If
                'End If
                '************************
                ' Grupos Conservación
                '************************
                If UsaGruposdeConservacion Then
                    If .Fields("grupo_conserv") <> 0 Then
                        Set RSTgrupo = Base.OpenRecordset("select * from grpconserv where grupo=" & .Fields("grupo_conserv"))
                        If Not RSTgrupo.EOF Then
                            
                            If Not IsNull(RSTgrupo.Fields("txt1")) Then
                                If MiEnvio.Etq <> 3 Then
                                    MiEnvio.Desc(8) = RSTgrupo.Fields("txt1")
                                Else
                                    MiEnvio.Desc(1) = RSTgrupo.Fields("txt1")
                                End If
                            Else
                                If MiEnvio.Etq <> 3 Then
                                    MiEnvio.Desc(8) = ""
                                Else
                                    MiEnvio.Desc(1) = ""
                                End If
                            End If
                            If Not IsNull(RSTgrupo.Fields("txt2")) Then
                                If MiEnvio.Etq <> 3 Then
                                    MiEnvio.Desc(9) = RSTgrupo.Fields("txt2")
                                Else
                                    MiEnvio.Desc(2) = RSTgrupo.Fields("txt2")
                                End If
                            Else
                                If MiEnvio.Etq <> 3 Then
                                    MiEnvio.Desc(9) = ""
                                Else
                                    MiEnvio.Desc(2) = ""
                                End If
                            End If
                             
                        Else
                            If MiEnvio.Etq <> 3 Then
                                MiEnvio.Desc(8) = ""
                                MiEnvio.Desc(9) = ""
                            Else
                                MiEnvio.Desc(2) = ""
                                MiEnvio.Desc(3) = ""
                            
                            End If
                        End If
                    End If
                End If
                
                'If (Modificaciones = True) And (MachineID = 1) And (Not TipoC9) Then
                '    If .Fields("tran_textoSC10") = "*" And .Fields("tran_tx1SC10") = "*" Then
                '        For BuclePlu = 0 To 8
                '            MiEnvio.Desc(BuclePlu) = ""
                '        Next BuclePlu
                '    End If
                'End If
                If .Fields("borrado") = True Then
                    MiRecibe = frmControl.Balanza.Erase_Item(0, MiEnvio.sec * 10000 + MiEnvio.plu)
                    MiRecibe = frmControl.Balanza.Erase_Item(1, MiEnvio.Code)
                    If MiRecibe = 4 Or MiRecibe = 3 Or (MiRecibe = 2 And (MachineID = 1 Or MachineID = 2) And TipoC9) Then MiRecibe = 0
                Else
                    frmControl.Balanza.SC10_DescLines 10
                    frmControl.Balanza.SC10_DescLines 10
                    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And Enviar4 Then
                        'ojo ver dinosol frmControl.Balanza.SC10_DescLines 4
                    End If
                    MiRecibe = frmControl.Balanza.Send_FullItem(MiEnvio)
                End If
                If MiRecibe >= 1000 Then
                    If Not (AhorraMas) Then
                        canCelar = 1
                    End If
                    frmMonitor.MostrarDato CargaCadena(544)
                    If Not (AhorraMas) Then
                        Exit For
                    End If
                End If
                If MiRecibe = 0 Then
                    CG_EnviaArticulo = True
                    If !borrado = False Then
                        frmMonitor.MostrarDato "Cod. " & Format(MiEnvio.Code, "000000") & " ---Plu. " & Format(MiEnvio.plu, "0000") & "," & MiEnvio.TexT & Space(30 - Len(MiEnvio.TexT)) & "-->OK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(875) & " " & MiEnvio.Code & "-->OK"
                    End If
                    Edit_Record Registro
                    If Not SC10NoMarcar Then
                        Select Case MachineID
                            Case 0
                                !tran_plu = "*"
                                !tran_texto = "*"
                                !tran_tx1 = "*"
                                !tran_cb = "*"
                                If .Fields("borrado") = True And Not UsaECO1 Then .Fields("tran_en") = "*"
                                If .Fields("borrado") = True And Not UsaECO2 Then .Fields("tran_el") = "*"
                                If .Fields("borrado") = True And Not UsaEcoPrint Then .Fields("tran_ep") = String(60, "*")
                                If !borrado = True And Not usaSC10 Then
                                    !tran_pluSC10 = "*"
                                    !tran_textosc10 = "*"
                                    !tran_tx1SC10 = "*"
                                    !tran_cbsc10 = "*"
                                End If
                            Case 1 To 2
                                If .Fields("borrado") = True And Not UsaECO1 Then .Fields("tran_en") = "*"
                                If .Fields("borrado") = True And Not UsaECO2 Then .Fields("tran_el") = "*"
                                If .Fields("borrado") = True And Not UsaEcoPrint Then .Fields("tran_ep") = String(60, "*")
                                !tran_pluSC10 = "*"
                                !tran_textosc10 = "*"
                                !tran_tx1SC10 = "*"
                                !tran_cbsc10 = "*"
                                If !borrado = True And Not UsaGamaAlta Then
                                    !tran_plu = "*"
                                    !tran_texto = "*"
                                    !tran_tx1 = "*"
                                    !tran_cb = "*"
                                End If
                        End Select
                    End If
                    .Update
                Else
                    If !borrado = False Then
                        If Dir(App.Path & "\upper.txt") = "" Then
                            frmMonitor.MostrarDato "Cod. " & Format(MiEnvio.Code, "000000") & " ---Plu. " & Format(MiEnvio.plu, "0000") & "," & MiEnvio.TexT & Space(30 - Len(MiEnvio.TexT)) & "-->NOK" & "(cod) : " & MiRecibe
                            If (AhorraMas) Then
                                .Delete
                            End If
                        Else
                            .Delete
                        End If
                        CG_EnviaArticulo = False
                    Else
                        frmMonitor.MostrarDato CargaCadena(875) & " " & MiEnvio.Code & "-->NOK"
                        'c2f 25/10/05
                        'c2f 1.7.25
                        '.Delete
                        '''''''''''''
                    End If
                End If
            End If
        .Movenext
    Next Bucle
    End With
    
    If AhorraMas And usaSC10 And Not (UsaGamaAlta) Then 'c2f ver aquí
    
        StrSQL = "select * from articulo where borrado=true"
        Set Registro = Base.OpenRecordset(StrSQL)
        If Not Registro.EOF Then
            Registro.MoveFirst
            Do While Not Registro.EOF
                Registro.Delete
                Registro.Movenext
            Loop
        
        End If
        Registro.Close
    
    End If
    
    
    CerrarBase Base
End Function
'***********
'***********
'***********
Public Sub CG_EnviaTeclas(MachineID As Integer, SoloModificaciones As Boolean)
Dim Base As Database
Dim Registro As Recordset
Dim LosDatos() As TipoBarras
Dim envio As String
Dim Cadenas()
Dim Buffer As String
Dim bufferlite As String
Dim Bucle As Long
Dim Buclesep As Long
Dim BucleGrande As Integer
Dim Condicion As String
Dim Contador As Long
Dim Reg2 As Recordset
Dim Reg3 As Recordset
Dim La_Respuesta As String
Dim ElMostrador As Integer
Dim CadenaModificacion As String
Dim ElContador As Integer
Dim SC10NoMarcar As Boolean
Dim AhoraC9 As Boolean
Dim BufCut As Integer
Dim BufCut2 As Integer
Dim MyMSG As String
Dim sSQL As String
Dim resp As Long

    BufCut2 = 17
    BufCut = 11
    If (MachineID = 1 Or MachineID = 2) And TipoC9 Then
        AhoraC9 = True
        BufCut = 10
        BufCut2 = 5
    End If
    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
        SC10NoMarcar = True
    Else
        SC10NoMarcar = False
    End If
    frmControl.Switch_mode MachineID
    If SoloModificaciones Then
        If AhorraMas Then
            CadenaModificacion = " and borrado="
        Else
            CadenaModificacion = " and tran_tecSC10<>'*' and borrado="
        End If
    Else
        CadenaModificacion = " and borrado="
    End If
    ' 1º Prepara los datos
    frmMonitor.MostrarInfo CargaCadena(234)  '"ENVÍO DE TECLAS RÁPIDAS"
    Set Base = OpenDatabase(Base_General)
    FrmSC10_MostrarMensajes = False
    Condicion = "true"
    For BucleGrande = 0 To 1
        '************
        ' Por sección
        '************
'2.0.33 aho se hace plu=tecla para sección 25 balanzas self-service.
'If AhorraMas Then '2.0.36 And MiEnvio.sec = 25 Then
'    Set RSTTec25 = Base.OpenRecordset("select * from seccion where secc_maqui=" & CStr(MiEnvio.sec))
'    If Not RSTTec25.EOF Then
'        If RSTTec25.Fields("enviardatos") = "SC10" And RSTTec25.Fields("com") = 255 Then
'            RSTTec25.Close
'            Set RSTTec25 = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MiEnvio.sec) & " and codigo=" & CStr(MiEnvio.Code))
'            If Not RSTTec25.EOF Then
'                If RSTTec25.Fields("borrado") = False Then
'                    MiEnvio.plu = RSTTec25.Fields("codi_tec")
'                End If
'            End If
'
'        End If
'
'    End If
'    RSTTec25.Close
'    Set RSTTec25 = Nothing
'End If
''''''''
        
        Set Reg2 = Base.OpenRecordset _
        ("select distinct codi_ident,secc_maqui,com from seccion where (borrado=false and enviardatos='SC10')") '2.0.15
        If Not Reg2.EOF Then
            Reg2.MoveFirst
            Do Until Reg2.EOF Or canCelar = 1
                Set Reg3 = Base.OpenRecordset _
                ("select * from equipos where borrado=false and modelo=100 and secc_maqui=" & _
                Reg2!secc_Maqui)
                If Not Reg3.EOF Then
                    For Bucle = 0 To 9
                        Set Registro = Base.OpenRecordset _
                        ("select * from teclas where tabla=" & Bucle & _
                        " and (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & Reg2!codi_ident _
                        & CadenaModificacion & Condicion)
                        Buffer = ""
                        With Registro
                            If Not .EOF Then
                                .MoveFirst
                                ElContador = 0
                                Do Until .EOF
                                    If AhoraC9 Then
                                        If (!codi_tec) <= 80 Then 'c2f 2.0.15 Dialsur
                                            If BucleGrande = 1 Then
                                                Buffer = Buffer & Format(!codi_tec, "00") _
                                                & Chr(0) & Format(!plu, "000000") & Chr(0)
                                            Else
                                                Buffer = Buffer & Format(!codi_tec, "00") _
                                                & Chr(0) & "000000" & Chr(0)
                                            End If
                                        End If
                                    Else
                                        If BucleGrande = 1 Then
                                            Buffer = Buffer & Format(!codi_tec, "000") _
                                            & Chr(0) & Format(!codigo, "000000") & Chr(0)
                                            '2.0.33 aho
                                            If AhorraMas And (!secc_Maqui = 25) And (!codigo) <> 0 And (Reg2!com = 255) Then
                                                CadenadeLog "Test-->1" & "secc.: " & Format(!secc_Maqui, "00") & " Código:" & Format(!codigo, "000000")
                                                resp = CG_EnviaArticulo(1, False, (!codigo))
                                            End If
                                            '''''''''''
                                        Else
                                            Buffer = Buffer & Format(!codi_tec, "000") _
                                            & Chr(0) & "000000" & Chr(0)
                                            '2.0.33 aho
                                            If AhorraMas And (!secc_Maqui = 25) And (!codigo) <> 0 And (Reg2!com = 255) Then
                                                CadenadeLog "Test-->0" & "secc.: " & Format(!secc_Maqui, "00") & " Código:" & Format(!codigo, "000000")
                                                resp = CG_EnviaArticulo(1, False, (!codigo))
                                            End If
                                            '''''''''''
                                        End If
                                    End If
                                    If ((AhoraC9) And (!codi_tec) <= 80) Or Not (AhoraC9) Then
                                        ElContador = ElContador + 1
                                    End If
                                    .Movenext
                                Loop
                                If Buffer <> "" Then
                                    For Buclesep = 1 To Len(Buffer) Step (BufCut2 * BufCut)
                                        If (Len(Buffer) - Buclesep) < (BufCut2 * BufCut) Then
                                            bufferlite = Mid(Buffer, Buclesep, Len(Buffer) - Buclesep + 1)
                                        Else
                                            bufferlite = Mid(Buffer, Buclesep, (BufCut2 * BufCut))
                                        End If
                                        ElContador = Len(bufferlite) / BufCut
                                        Contador = Contador + 1
                                        ReDim Preserve Cadenas(Contador + 1)
                                        ReDim Preserve LosDatos(Contador + 1)
                                        LosDatos(Contador).Tipo = "Teclas de Mostrador "
                                        LosDatos(Contador).Numero = Reg2!codi_ident
                                        LosDatos(Contador).estado = Condicion
                                        If AhoraC9 Then
                                            Cadenas(Contador) = "550" & Chr(0) & Format(Reg2.Fields("secc_maqui"), "00") & _
                                            Chr(0) & Format(Bucle, "00") & Chr(0) & "0" & Chr(0) & Format(ElContador, "00") & _
                                            Chr(0) & bufferlite & Chr(3)
                                        Else
                                            Cadenas(Contador) = "570" & Chr(0) & Format(Reg2.Fields("secc_maqui"), "00") & _
                                            Chr(0) & Format(Bucle, "00") & Chr(0) & "1" & Chr(0) & Format(ElContador, "00") & _
                                            Chr(0) & bufferlite & Chr(3)
                                        End If
                                    Next Buclesep
                                End If
                            End If
                       End With
                    Next Bucle
                End If
                Reg2.Movenext
            Loop
        End If
        '***************************************
        '* Por equipo (Sólo Euroscale / Sc10 ) *
        '***************************************
        If Dir(App.Path & "\ga165.txt") = "" Then '2.0.15
        
        If True Then '(Not ((Machineid = 1 Or Machineid = 2) And TipoC9) Then
            Set Reg2 = Base.OpenRecordset _
            ("select * from equipos where borrado=false and modelo=100")
            If Not Reg2.EOF Then
                Reg2.MoveFirst
                Do Until Reg2.EOF
                    For Bucle = 0 To 9
                        Set Registro = Base.OpenRecordset _
                        ("select * from teclas where numero_eqp=" & _
                        Reg2!numero_eqp & " and tabla=" & _
                        Bucle & CadenaModificacion & Condicion)
                        Buffer = ""
                        With Registro
                            If Not .EOF Then
                                .MoveFirst
                                ElContador = 0
                                Do Until .EOF
                                    If AhoraC9 Then
                                        If BucleGrande = 1 Then
                                            Buffer = Buffer & Format(!codi_tec, "00") _
                                            & Chr(0) & Format(!plu, "000000") & Chr(0)
                                        Else
                                            Buffer = Buffer & Format(!codi_tec, "00") _
                                            & Chr(0) & "000000" & Chr(0)
                                        End If
                                    Else
                                        If BucleGrande = 1 Then
                                            Buffer = Buffer & Format(!codi_tec, "000") _
                                            & Chr(0) & Format(!codigo, "000000") & Chr(0)
                                        Else
                                            Buffer = Buffer & Format(!codi_tec, "000") _
                                            & Chr(0) & "000000" & Chr(0)
                                        End If

                                    End If
                                    ElContador = ElContador + 1
                                    .Movenext
                                Loop
                                If Buffer <> "" Then
                                    
                                    For Buclesep = 1 To Len(Buffer) Step (BufCut2 * BufCut)
                                        If (Len(Buffer) - Buclesep) < (BufCut2 * BufCut) Then
                                            bufferlite = Mid(Buffer, Buclesep, Len(Buffer) - Buclesep + 1)
                                        Else
                                            bufferlite = Mid(Buffer, Buclesep, (BufCut2 * BufCut))
                                        End If
                                        
                                        ElContador = Len(bufferlite) / BufCut
                                        Contador = Contador + 1
                                        ReDim Preserve Cadenas(Contador + 1)
                                        ReDim Preserve LosDatos(Contador + 1)
                                        LosDatos(Contador).Tipo = "Teclas de Equipo "
                                        LosDatos(Contador).Numero = Reg2!numero_eqp
                                        LosDatos(Contador).estado = Condicion
                                        If AhoraC9 Then
                                            Cadenas(Contador) = "551" & Chr(0) & Format(Reg2.Fields("numero_eqp"), "00") & _
                                            Chr(0) & Format(Bucle, "00") & Chr(0) & "0" & Chr(0) & Format(ElContador, "00") & _
                                            Chr(0) & bufferlite & Chr(3)
                                        Else
                                            Cadenas(Contador) = "571" & Chr(0) & Format(Reg2.Fields("numero_eqp"), "00") & _
                                            Chr(0) & Format(Bucle, "00") & Chr(0) & "1" & Chr(0) & Format(ElContador, "00") & _
                                            Chr(0) & bufferlite & Chr(3)

                                        End If
                                    Next Buclesep
                                End If
                            End If
                        End With
                    Next Bucle
                    Reg2.Movenext
                Loop
            End If
        End If
        
        End If
        
    Condicion = "false"
    Next BucleGrande
    '**********************
    ' Envía los datos
    '**********************
    For Bucle = 1 To Contador
        Cadenas(Bucle) = Chr(2) & Cadenas(Bucle)
        Select Case MachineID
            Case 0
            Case 1 To 2
        End Select
        If La_Respuesta = "" Then
            frmMonitor.MostrarDato CargaCadena(544)
            Exit For
        End If
        If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
            La_Respuesta = "0" & Chr(6) & "00"
        End If
        
        If canCelar = 1 Then Exit For
        If Right(LosDatos(Bucle).Tipo, 10) = "Mostrador " Then
                MyMSG = "Sec. "
            Else
                MyMSG = "Eqp. "
            End If
        If Asc(Mid(La_Respuesta, 2, 1)) <> 6 And Not (Asc(Mid(La_Respuesta, 2, 1)) = 8 And Mid(La_Respuesta, 3, 2) = "04") Then
            frmMonitor.MostrarDato MyMSG & LosDatos(Bucle).Numero & " --> NOK"
        Else
            frmMonitor.MostrarDato MyMSG & LosDatos(Bucle).Numero & " --> OK"
            If LosDatos(Bucle).estado = "true" Then
                If Right(LosDatos(Bucle).Tipo, 10) = "Mostrador " Then
                    Set Registro = Base.OpenRecordset _
                    ("select * from teclas where (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & LosDatos(Bucle).Numero & CadenaModificacion & "true")
                Else
                    Set Registro = Base.OpenRecordset _
                    ("select * from teclas where numero_eqp=" & LosDatos(Bucle).Numero & CadenaModificacion & "true")
                End If
                With Registro
                    If Not .EOF Then .MoveFirst
                    Do Until .EOF
                        If Not SC10NoMarcar Then
                            Edit_Record Registro
                            !tran_tecSC10 = "*"
                            If Not UsaGamaAlta Then !tran_tec = "*"
                            .Update
                        End If
                        If Not .EOF Then .Movenext
                    Loop
                End With
            Else
                If Right(LosDatos(Bucle).Tipo, 10) = "Mostrador " Then
                    Set Registro = Base.OpenRecordset _
                    ("select * from teclas where (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & LosDatos(Bucle).Numero & CadenaModificacion & "false")
                Else
                   Set Registro = Base.OpenRecordset _
                   ("select * from teclas where numero_eqp=" & LosDatos(Bucle).Numero & CadenaModificacion & "false")
                End If
                With Registro
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            If Not SC10NoMarcar Then
                                Edit_Record Registro
                                !tran_tecSC10 = "*"
                                .Update
                            End If
                            .Movenext
                        Loop
                    End If
                End With
            End If
        End If
    Next Bucle
    Set Registro = Base.OpenRecordset("select * from teclas where borrado=true")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If IsNull(!numero_eqp) Then
                    Set Reg3 = Base.OpenRecordset _
                    ("select * from seccion where borrado=false and codi_ident=" & !codi_ident)
                    If Not Reg3.EOF Then
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and modelo=100 and secc_maqui=" & Reg3!secc_Maqui)
                    End If
                Else
                    Set Reg2 = Base.OpenRecordset _
                    ("select * from equipos where borrado=false and modelo=100 and numero_eqp=" & !numero_eqp)
                End If
                If Reg2.EOF Then
                    If Not SC10NoMarcar Then
                        Edit_Record Registro
                        !tran_tecSC10 = "*"
                        If Not UsaGamaAlta Then
                            !tran_tec = "*"
                        End If
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
Public Sub CG_EnviaVendedor(MachineID As Integer, Modificaciones As Boolean)
        Dim MiEnvio
        Dim MiRecibe As String
        Dim Checksum As String
        Dim OrdenDat As Integer
        Dim Base As Database
        Dim rstvendedor As Recordset
        Dim Reg2 As Recordset
        Dim Contador As Long
        Dim IDV As String
        Dim txt As String
        Dim CODVEN As String
        Dim sec As String
        Dim tecla As String
        Dim Bucle As Long
        Dim NumOrden As Long
        Dim OkChange As Boolean
        Dim SC10NoMarcar As Boolean
        If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
            SC10NoMarcar = True
        Else
            SC10NoMarcar = False
        End If
        frmControl.Switch_mode MachineID
        Set Base = OpenDatabase(Base_General)
        If Modificaciones Then
            Select Case MachineID
                Case 0
                    Set rstvendedor = Base.OpenRecordset _
                    ("select * from vendedor where tran_vend<>'*' or tran_asign<>'*' order by borrado")
                Case 1 To 2
                    Set rstvendedor = Base.OpenRecordset _
                    ("select * from vendedor where tran_vendSC10<>'*' or tran_asignSC10<>'*' order by borrado")
            End Select
        Else
            Set rstvendedor = Base.OpenRecordset("select * from Vendedor order by borrado")
        End If
        '*****************************************************************
        ' Cuenta el número de vendedores
        '*****************************************************************
        Contador = 0
        If rstvendedor.EOF Then
            CerrarBase Base
            
            Exit Sub
        End If
        With rstvendedor
            .MoveLast
            Contador = .Recordcount
        End With
        frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(13)   '"ENVÍO DE VENDEDORES"
        With rstvendedor
            .MoveFirst
            NumOrden = 1
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = Contador + 2
            For Bucle = 1 To Contador
                If canCelar = 1 Then
                    CerrarBase Base
                    Exit Sub
                End If
                '****
                ' Comprueba si se tiene que enviar
                Select Case MachineID
                    Case 0
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where secc_maqui=" & !codi_ident & _
                        " and borrado=false and modelo=100")
                        If Reg2.EOF Then OkChange = True
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where secc_maqui=" & !codi_ident & _
                        " and borrado=false and modelo<100")
                    Case 1 To 2
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where secc_maqui=" & !codi_ident & _
                        " and borrado=false and modelo<100")
                        If Reg2.EOF Then OkChange = True
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where secc_maqui=" & !codi_ident & _
                        " and borrado=false and modelo=100")
                End Select
                If (Not Reg2.EOF) Or (!codi_ident = 0) Then
                    MonitorBarra
                    IDV = Format(!ident_vend, "0000")
                    CODVEN = Format(!codi_vend, "000000")
                    txt = (Trim(!Nombre)) + Space(30 - Len(!Nombre))
                    If MachineID = 0 Then
                        txt = strtranñÑ(txt)
                    Else
                        txt = ConvCad(txt)
                    End If
                    sec = Format(!codi_ident, "00")
                    tecla = Format(!tec_vend, "00")
                    ' Da de baja si está borrado
                    If !borrado = True Or !prog_vend <> "*" Then
                        MiEnvio = Chr(2) & "63" & IDV & Chr(0) & Chr(3)
                        Select Case MachineID
                            Case 0
                            Case 1 To 2
                        End Select
                        If MiRecibe = "" Then
                            frmMonitor.MostrarDato CargaCadena(544)
                            Exit For
                        End If
                        If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
                            MiRecibe = "0" & Chr(6) & "00"
                        End If
                        If Asc(Mid(MiRecibe, 2, 1)) <> 6 Then
                            frmMonitor.MostrarDato CargaCadena(414) & " " & Val(IDV) & "-->NOK"
                        Else
                            frmMonitor.MostrarDato CargaCadena(414) & " " & Val(IDV) & "-->OK"
                        End If
                    End If
                    ' Crea o elimina
                    
                    If .Fields("borrado") = False Then
                        MiEnvio = Chr(2) & "60" & IDV & Chr(0) & txt & Chr(0) & CODVEN & Chr(0) & Chr(3)
                    Else
                        MiEnvio = Chr(2) & "61" & IDV & Chr(0) & "0" & Chr(0) & Chr(3)
                    End If
                    
                    Select Case MachineID
                        Case 0
                        Case 1 To 2
                    End Select
                    
                    If MiRecibe = "" Then
                        frmMonitor.MostrarDato CargaCadena(544)
                        Exit For
                    End If
                    
                    If canCelar = 1 Then Exit For
                    If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
                            MiRecibe = "0" & Chr(6) & "00"
                        End If
                    If Asc(Mid(MiRecibe, 2, 1)) <> 6 Then
                        If !borrado = False Then
                            frmMonitor.MostrarDato CargaCadena(813) & " " & Val(IDV) & "-->NOK"
                        Else
                            frmMonitor.MostrarDato CargaCadena(814) & " " & Val(IDV) & "-->NOK"
                            If Mid(MiRecibe, 3, 2) = "04" Then
                                If Not SC10NoMarcar Then
                                    Edit_Record rstvendedor
                                    Select Case MachineID
                                        Case 0
                                            !tran_asign = "*"
                                            !tran_vend = "*"
                                            If Not usaSC10 Then
                                                !tran_asignSC10 = "*"
                                                !tran_vendSC10 = "*"
                                            End If
                                        Case 1 To 2
                                            !tran_asignSC10 = "*"
                                            !tran_vendSC10 = "*"
                                            If Not UsaGamaAlta Then
                                                !tran_asign = "*"
                                                !tran_vend = "*"
                                            End If
                                    End Select
                                    .Update
                                End If
                            End If
                        End If
                    Else
                        If !borrado = False Then
                            frmMonitor.MostrarDato CargaCadena(813) & " " & Val(IDV) & "-->OK"
                             If Not SC10NoMarcar Then
                                 Edit_Record rstvendedor
                                 Select Case MachineID
                                    Case 0
                                        !tran_asign = "*"
                                        !tran_vend = "*"
                                    Case 1 To 2
                                        !tran_asignSC10 = "*"
                                        !tran_vendSC10 = "*"
                                End Select
                                .Update
                            End If
                        Else
                            frmMonitor.MostrarDato CargaCadena(814) & " " & Val(IDV) & "-->OK"
                             If Not SC10NoMarcar Then
                                 Edit_Record rstvendedor
                                 Select Case MachineID
                                    Case 0
                                        !tran_asign = "*"
                                        !tran_vend = "*"
                                        If Not usaSC10 Then
                                            !tran_asignSC10 = "*"
                                           !tran_vendSC10 = "*"
                                        End If
                                    Case 1 To 2
                                        !tran_asignSC10 = "*"
                                        !tran_vendSC10 = "*"
                                        If Not UsaGamaAlta Then
                                            !tran_asign = "*"
                                           !tran_vend = "*"
                                        End If
                                End Select
                                .Update
                            End If
                        End If
                    End If
                    ' Da de alta
                    If (!prog_vend = "*" Or !prog_vend = "-") And !borrado = False Then
                        MiEnvio = Chr(2) & "62" & IDV & Chr(0) & sec & Chr(0) & tecla & Chr(0) & Chr(3)
                        Select Case MachineID
                            Case 0
                            Case 1 To 2
                        End Select
                        If MiRecibe = "" Then
                            frmMonitor.MostrarDato CargaCadena(544)
                            Exit For
                        End If
                        If canCelar = 1 Then Exit For
                        If (MachineID = 1 Or MachineID = 2) And (Not TipoC9) And (Not SC10Rs232) Then
                            MiRecibe = "0" & Chr(6) & "00"
                        End If
                        If Asc(Mid(MiRecibe, 2, 1)) <> 6 And Asc(Mid(MiRecibe, 3, 1)) <> 31 Then
                            frmMonitor.MostrarDato CargaCadena(328) & " " & Val(IDV) & "-->NOK"
                        Else
                            frmMonitor.MostrarDato CargaCadena(328) & " " & Val(IDV) & "-->OK"
                            If !prog_vend = "-" And OkChange Then
                                If Not SC10NoMarcar Then
                                    Edit_Record rstvendedor
                                    !prog_vend = "*"
                                    .Update
                                End If
                            End If
                        End If
                    End If
                Else
                    Edit_Record rstvendedor
                    Select Case MachineID
                        Case 0
                            !tran_vend = "*"
                            !tran_asign = "*"
                            If Not usaSC10 And !borrado = True Then
                                !tran_vendSC10 = "*"
                                !tran_asignSC10 = "*"
                            End If
                        Case 1 To 2
                            !tran_vendSC10 = "*"
                            !tran_asignSC10 = "*"
                            If Not UsaGamaAlta And !borrado = True Then
                                !tran_vend = "*"
                                !tran_asign = "*"
                            End If
                    End Select
                    .Update
                End If
                .Movenext
            Next Bucle
        End With
        CerrarBase Base
       
End Sub
Private Function ConvCad(LaCadena As String) As String
    Dim Buffer As String
    Dim Buffer2 As String
    Dim Bucle As Integer
    Buffer2 = ""
    Buffer = StrConv(LaCadena, vbUpperCase)
    For Bucle = 1 To Len(Buffer)
        Select Case Mid(Buffer, Bucle, 1)
        Case "Ñ"
            Buffer2 = Buffer2 & "@"
        Case Else
            Buffer2 = Buffer2 & Mid(Buffer, Bucle, 1)
        End Select
    Next Bucle
    ConvCad = Buffer2
End Function

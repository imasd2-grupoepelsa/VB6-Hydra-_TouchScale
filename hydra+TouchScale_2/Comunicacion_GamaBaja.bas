Attribute VB_Name = "Comunicacion_GamaBaja"
'///////////////////////////////////////////////////
' Funciones para comunicación con Gama Baja conectada,
' es decir, EcoNet/EcoLabel/V8.
'////////////////////////////////////////////////////
Option Explicit
Public GB_CONSULTAR_USUARIO As Boolean
Public Type SeccionesGB
    NSecciones As Integer
    puerto() As Integer
    seccion() As Integer
    Equipo() As Integer
    '1.7.19 (eco)
    sIp() As String
    nPort() As Long
    '''''''''''''
End Type
Public Sub CalculaSeccionesGB(SecGB As SeccionesGB, MachineID As Long)
    '////////////////////////////////////
    ' para comunicar con estas máquinas hay
    ' que hacerlo por secciones, esta
    ' función obtiene de la base de datos
    ' las secciones que existen para
    ' un determinado sistema, y las
    ' almacena en la estructura SeccionesGB.
    ' MachineID determina el sistema :
    ' 1000 --> Actualmente, V-8
    ' 1001 --> actualmente, EcoLabel/EcoNet
    Dim Base As Database
    Dim Registro As Recordset
    Dim RegBal As Recordset
    Dim MyCad As String
    Dim Bucle As Integer
    Dim NumPuertos(20) As Integer
    Dim sIp As String
    Dim nPort As Long
    
    If MachineID = 1000 Then
        MyCad = "ECON"
    Else
        MyCad = "ECOL"
    End If
    SecGB.NSecciones = 0
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select codi_ident,com,sec_ip,multiple from seccion where borrado=false and enviardatos=" & Chr(34) & MyCad & Chr(34))
    With Registro
        If Not .EOF Then
            Do Until .EOF
                If .Fields("codi_ident") > 0 Then ' Sección con valor 0? Eso no existe
                    If (.Fields("COM") > 0) Or (UsaModem) Then 'El puerto "COM0" está reservado para indicar que no se desea comunicar con la sección indicada
                        SecGB.NSecciones = SecGB.NSecciones + 1
                        ReDim Preserve SecGB.seccion(SecGB.NSecciones)
                        ReDim Preserve SecGB.Equipo(SecGB.NSecciones)
                        ReDim Preserve SecGB.puerto(SecGB.NSecciones)
                        '1.7.19 (eco)
                        ReDim Preserve SecGB.sIp(SecGB.NSecciones)
                        ReDim Preserve SecGB.nPort(SecGB.NSecciones)
                        '''''''''''''
                        SecGB.seccion(SecGB.NSecciones - 1) = .Fields("codi_ident")
                        SecGB.puerto(SecGB.NSecciones - 1) = .Fields("COM")
                        SecGB.Equipo(SecGB.NSecciones - 1) = 0
                        '1.7.19 (eco)
                        SecGB.sIp(SecGB.NSecciones - 1) = ""
                        SecGB.nPort(SecGB.NSecciones - 1) = 0
                        '''''''''''''
                        Set RegBal = Base.OpenRecordset("select numero_eqp from equipos where borrado=false and codi_ident=" & .Fields("codi_ident"))
                        If Not RegBal.EOF Then
                            SecGB.Equipo(SecGB.NSecciones - 1) = RegBal.Fields("numero_eqp")
                            NumPuertos(.Fields("COM")) = NumPuertos(.Fields("COM")) + 1
                        Else ' Si no está definido el número de equipo con el
                             ' que comunicar en la sección, esta sección
                             ' no se tendrá en cuenta
                            SecGB.NSecciones = SecGB.NSecciones - 1
                        End If
                    Else
                        '1.7.19 (eco)
                        sIp = .Fields("sec_ip")
                        If GeneraTextoIP(sIp, nPort) Then
                            SecGB.NSecciones = SecGB.NSecciones + 1
                            ReDim Preserve SecGB.seccion(SecGB.NSecciones)
                            ReDim Preserve SecGB.Equipo(SecGB.NSecciones)
                            ReDim Preserve SecGB.puerto(SecGB.NSecciones)
                            '1.7.19 (eco)
                            ReDim Preserve SecGB.sIp(SecGB.NSecciones)
                            ReDim Preserve SecGB.nPort(SecGB.NSecciones)
                            '''''''''''''
                            SecGB.seccion(SecGB.NSecciones - 1) = .Fields("codi_ident")
                            SecGB.puerto(SecGB.NSecciones - 1) = .Fields("COM")
                            SecGB.Equipo(SecGB.NSecciones - 1) = 0
                            '1.7.19 (eco)
                            SecGB.sIp(SecGB.NSecciones - 1) = sIp
                            '2.0.A
                            'SecGB.nPort(SecGB.NSecciones - 1) = nPort
                            SecGB.nPort(SecGB.NSecciones - 1) = .Fields("multiple")
                            '''''''''''''
                            
                            Set RegBal = Base.OpenRecordset("select numero_eqp from equipos where borrado=false and codi_ident=" & .Fields("codi_ident"))
                            If Not RegBal.EOF Then
                                SecGB.Equipo(SecGB.NSecciones - 1) = RegBal.Fields("numero_eqp")
                                NumPuertos(.Fields("COM")) = NumPuertos(.Fields("COM")) + 1
                            Else ' Si no está definido el número de equipo con el
                                 ' que comunicar en la sección, esta sección
                                 ' no se tendrá en cuenta
                                SecGB.NSecciones = SecGB.NSecciones - 1
                            End If
                        
                        
                        End If
                        
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    Base.Close
    '/////////////////////////////////
    ' la comunicación será automática
    ' si los puertos COM son diferentes,
    ' en otro caso se pregunta al usuario
    ' antes de cambiar de máquina para
    ' asegurarse de que cambia el switch
    ' de puertos serie
    '//////////////////////////////////
    GB_CONSULTAR_USUARIO = False
    For Bucle = 0 To 20
        If NumPuertos(Bucle) > 1 Then GB_CONSULTAR_USUARIO = True
    Next Bucle
End Sub
'*******************************************************************************
' Funciones de comunicación para máquinas de Gama Baja conectadas, que
' comprenden a las máquinas ECOLABEL y EcoNet(MachineID = 1001), y V8 (MachineID = 1000)
'*******************************************************************************
Public Sub GB_EnviaBaseMaestra(ByVal MachineID As Integer, ByVal Modificaciones As Boolean, ByVal NSec As Integer, ByVal NPUerto As Integer, ByVal NMAquina As Integer)
    canCelar = 0
    CancelHumano = False
    GB_EnviaCodBar MachineID, NSec, NPUerto, NMAquina, Modificaciones
    If (canCelar = 0) And (Not CancelHumano) Then GB_EnviaIVAS MachineID, NPUerto, NMAquina, Modificaciones
    If (canCelar = 0) And (Not CancelHumano) Then GB_EnviaCableY MachineID, NSec, NPUerto, NMAquina, Modificaciones
    If (MachineID = 1001 Or (MachineID = 1000 And Not TipoV8)) And (canCelar = 0) And (Not CancelHumano) Then GB_EnviaTextosEcoLabel NPUerto, NMAquina, NSec, Modificaciones
    If (MachineID = 1001 Or (MachineID And Not TipoV8)) And (canCelar = 0) And (Not CancelHumano) Then GB_EnviaTarasEcoLabel NPUerto, NMAquina
    If (canCelar = 0) And (Not CancelHumano) Then GB_EnviaArt MachineID, NSec, NPUerto, NMAquina, Modificaciones
End Sub
Public Function GB_Query_Tot(ByVal MachineID As Integer, NPUerto As Integer, NMAquina As Integer)
    '******************************
    ' Totales por artículo
    ' para máquinas Econet,Ecolabel,V8
    '******************************
    Dim Bucle As Long
    Dim MaxBucle As Integer
    Dim Resp As Integer
    Dim sTotArtPath As String
    Dim sTotVenPath As String
    Dim sTotMaqPath As String
    If MachineID = 1000 And TipoV8 Then
        MaxBucle = 0
        frmControl.Switch_mode 9, NPUerto, NMAquina
    Else
        MaxBucle = 1
        frmControl.Switch_mode 8, NPUerto, NMAquina
    End If
    
    If lMarteAlfa Then MaxBucle = 0
    
    If HayMulti Then
        If TiendaActual < 100 Then
        sTotArtPath = App.Path & "\T" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "gtart.tot"
        sTotVenPath = App.Path & "\T" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "gtvend.tot"
        sTotMaqPath = App.Path & "\T" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "gtmaq.tot"
        Else
        sTotArtPath = App.Path & "\T" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "gtart.tot"
        sTotVenPath = App.Path & "\T" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "gtvend.tot"
        sTotMaqPath = App.Path & "\T" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "gtmaq.tot"
        
        End If
    Else
        sTotArtPath = App.Path & "\01gtart.tot"
        sTotVenPath = App.Path & "\01gtvend.tot"
        sTotMaqPath = App.Path & "\01gtmaq.tot"
    End If
    For Bucle = MaxBucle To 0 Step -1 ' si es ecolabel, pide venta y envasado
        frmMonitor.MostrarInfo CargaCadena(165)
        Resp = frmControl.Balanza.Query_All_Totals(sTotArtPath, 0, Bucle)
        If Resp = 0 Then
            If (canCelar = 0) And (Not CancelHumano) Then
                frmMonitor.MostrarInfo CargaCadena(166)
                Resp = frmControl.Balanza.Query_All_Totals(sTotVenPath, 1, Bucle)
            End If
        
            If (canCelar = 0) And (Not CancelHumano) Then
                frmMonitor.MostrarInfo CargaCadena(164)
                Resp = frmControl.Balanza.Query_All_Totals(sTotMaqPath, 2, Bucle)
            End If
        Else
            If Resp = 5 Then frmMonitor.MostrarDato CargaCadena(1200)
            CancelHumano = True
        End If
    Next Bucle
    frmControl.Switch_mode 1000
    GB_Query_Tot = Resp
End Function
Public Sub GB_PideArticulo(Articulo As DB_Articulo, Completa As Boolean, ByVal Eqp As Integer, ByVal MyPort As Integer, ByVal MySec As Long)
    Dim MyArt As Epel_FullItem
    Dim Bucle As Integer
    Dim Resp As Long
    Dim Continuar As Boolean
    
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    
    Continuar = True
    
    frmControl.Switch_mode 8, MyPort, Eqp
    Do While Continuar
        Continuar = False
        MyArt.plu = Articulo.plu
        FrmConsultaArt.LblInfo.Caption = CargaCadena(971) & " Marte IC-ILC PLU : " & Articulo.plu
        If Completa Then
            Resp = frmControl.Balanza.Query_Item(MyArt, 1) ' Con descriptivos
        Else
            Resp = frmControl.Balanza.Query_Item(MyArt, 0) ' Sin descriptivos
        End If
        '1.7.15
        'If Resp = 0 Then
        If MyArt.plu = -1 Then '1.9.3
            Exit Do
        End If
        
        '2.0.9 If Dir(App.Path & "\marte.new") = "" Then
        If lMarteNew = False Then
        
            If MyArt.plu <= 800 Then
                Articulo.plu = MyArt.plu
                If Resp <> 0 Then
                    'CadenadeLog "Se recibe ITEM dentro de margen con respuesta <> 0"
                    Resp = 0
                Else
                    'CadenadeLog "Se recibe ITEM dentro de margen con respuesta OK"
                End If
            Else
                CadenadeLog "Artículo fuera de margen:" & CStr(MyArt.plu)
                If Resp = 0 Then
                    'CadenadeLog "la respuesta era correcta"
                Else
                    'CadenadeLog "la respuesta era incorrecta"
                End If
            End If
        Else
            If MyArt.plu <= 2000 Then
                Articulo.plu = MyArt.plu
                If Resp <> 0 Then
                    'CadenadeLog "Se recibe ITEM dentro de margen con respuesta <> 0"
                    Resp = 0
                Else
                    'CadenadeLog "Se recibe ITEM dentro de margen con respuesta OK"
                End If
            Else
                CadenadeLog "Artículo fuera de margen:" & CStr(MyArt.plu)
                If Resp = 0 Then
                    'CadenadeLog "la respuesta era correcta"
                Else
                    'CadenadeLog "la respuesta era incorrecta"
                End If
            End If
        
        End If
        'End If
        '''''''
        If Trim(MyArt.TexT) = "" And MyArt.Price = 0 And Resp = 0 Then
            Continuar = True
            Articulo.plu = Articulo.plu + 1
        End If
    
    Loop
    
    If Resp = 0 Then
        Articulo.Mostrador = MySec
        Articulo.codigo = MyArt.Code
        If Articulo.codigo = 0 Then
            Articulo.codigo = MySec * 10000 + Articulo.plu
        End If
        Articulo.Balenv = False
        Articulo.caducidad = MyArt.Caducity
        If Dir(App.Path & "\greek.txt") = "" Then
            Articulo.Descriptivos(0) = AdaptaCharPolishBPC(MyArt.TexT)
        Else
            Articulo.Descriptivos(0) = AdapTaCharGR(MyArt.TexT)
        End If
        For Bucle = 0 To 19
            If Dir(App.Path & "\greek.txt") = "" Then
                Articulo.Descriptivos(Bucle + 1) = AdaptaCharPolishBPC(MyArt.Desc(Bucle))
            Else
                Articulo.Descriptivos(Bucle + 1) = AdapTaCharGR(MyArt.Desc(Bucle))
            End If
        Next Bucle
        Articulo.EAN13 = ""
        Articulo.Etiqueta = 255
        Articulo.familia = MyArt.Family
        Articulo.GrpConserv = 0

        If MyArt.vat <= 0 Then
            Articulo.IVA = 0
        Else
            Articulo.IVA = MyArt.vat
        End If
        
        Articulo.Merma = 0
        Articulo.precio = MyArt.Price
        Articulo.Preferente = MyArt.Pref
        '¡¡¡
        'Articulo.tara = MyArt.Tare
        '!!!
        If UsaECO2 And MyArt.nPeso <> 0 Then
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from tipotara where tipo=" & CStr(MyArt.nPeso))
            If Not Registro.EOF Then
                Articulo.tara = Registro.Fields("valor")
            End If
            Registro.Close
            Set Registro = Nothing
            Base.Close
            Set Base = Nothing
        Else
            Articulo.tara = 0
        End If
        
        For Bucle = 0 To 9
            Articulo.TipoLetra(Bucle) = 0
        Next Bucle
        Articulo.Presel = 0
        Articulo.subsec = 0
        Articulo.WGH = MyArt.Weight '1.8.5
    Else
        Articulo.plu = -1
    End If
    
End Sub
Public Sub GB_EnviaTarasEcoLabel(ByVal NPUerto As Integer, ByVal NMAquina As Integer)
    '******************************
    ' Función exclusiva
    ' para máquinas ECOLABEL
    ' envía los tipos de tara
    ' definidos por el usuario
    '******************************
    Dim Base As Database
    Dim Registro As Recordset
    Dim Resp As Long
    Dim lTaras(10) As Long
    frmMonitor.MostrarInfo CargaCadena(1183)
    frmControl.Switch_mode 8, NPUerto, NMAquina
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from tipotara")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If .Fields("tipo") > 0 And .Fields("tipo") < 11 Then
                    lTaras(.Fields("tipo") - 1) = .Fields("Valor")
                End If
                .Movenext
            Loop
            Resp = frmControl.Balanza.Send_Tares(lTaras)
            If Resp = 0 Then
                frmMonitor.MostrarDato CargaCadena(1183) & " --> OK"
            Else
                frmMonitor.MostrarDato CargaCadena(1183) & " --> NOK"
            End If
        End If
    End With
    Base.Close
    frmControl.Switch_mode 1000
End Sub
Public Sub GB_EnviaTextosEcoLabel(ByVal NPUerto As Integer, ByVal NMAquina As Integer, ByVal MySec As Long, Optional ByVal lModif As Boolean)
    Dim Base As Database
    Dim Registro As Recordset
    Dim Resp As Long
    '******************************
    ' Función exclusiva
    ' para máquinas ECOLABEL
    ' envía los 2 textos libres
    ' de la etiqueta ECOLABEL
    '******************************
    frmMonitor.MostrarInfo CargaCadena(1182)
    frmControl.Switch_mode 8, NPUerto, NMAquina
    Set Base = OpenDatabase(Base_General)
    If lModif Then

        Set Registro = Base.OpenRecordset("select * from textoeco where codi_ident=" & MySec & " and (tran_el=' ' or isnull(tran_el))")
    
    Else
    
        Set Registro = Base.OpenRecordset("select * from textoeco where codi_ident=" & MySec)
        
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            'If Len(Trim(.Fields("texto1"))) > 0 And Len(Trim(.Fields("texto2"))) > 0 Then
                Resp = frmControl.Balanza.Send_LabelTexts(.Fields("texto1"), .Fields("texto2"))
                If Resp = 0 Then
                    frmMonitor.MostrarDato CargaCadena(1182) & " --> OK"
                    .Edit
                    .Fields("tran_el") = "*"
                    .Update
                Else
                    frmMonitor.MostrarDato CargaCadena(1182) & " --> NOK"
                    .Edit
                    .Fields("tran_el") = " "
                    .Update
                End If
            'End If
        End If
    End With
    Base.Close
    frmControl.Switch_mode 1000
End Sub
Private Sub GB_MarcaCodbarCableyAjenos(ByVal EsCabley As Boolean, ByVal TipoMaquina As Integer)
'////////////////////////////////////
' Los datos con flag BORRADO=true,
' deben tener marcados también
' los flags tran_en y tran_el
' para que sean eliminados realmente
' de la base. Aquí se marcan las tablas
' CODBAR y CABLEY
'//////////////////////////////////
    Dim Base As Database
    Dim Registro As Recordset
    Dim RegSec As Recordset
    Dim NombreCampo As String
    Dim Envi As String
    Dim Buf As Integer
    Set Base = OpenDatabase(Base_General)
    If TipoMaquina = 1000 Then
        NombreCampo = "tran_en"
        Envi = "ECON"
    Else
        NombreCampo = "tran_el"
        Envi = "ECOL"
    End If
    If EsCabley Then
        Set Registro = Base.OpenRecordset("select * from cabley where borrado=true")
    Else
        Set Registro = Base.OpenRecordset("select * from codbar where borrado=true")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Buf = 0
                If Not IsNull(.Fields("numero_eqp")) Then Buf = .Fields("numero_eqp")
                If Buf <> 0 Then
                    Edit_Record Registro
                    .Fields(NombreCampo) = "*"
                    .Update
                Else
                    Set RegSec = Base.OpenRecordset("select enviardatos from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
                    If Not RegSec.EOF Then
                        If RegSec.Fields("enviardatos") <> Envi Then
                            Edit_Record Registro
                            .Fields(NombreCampo) = "*"
                            .Update
                        End If
                    Else
                        Edit_Record Registro
                        .Fields(NombreCampo) = "*"
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
End Sub
Public Sub GB_EnviaCodBar(ByVal MachineID As Integer, ByVal NSec As Integer, ByVal NPUerto As Integer, ByVal NMAquina As Integer, ByVal Modificaciones As Boolean)
    '*********************************
    ' Envía los códigos de barras
    ' para máquinas Gama Baja conectadas
    '*********************************
    Dim Base As Database
    Dim Registro As Recordset
    Dim micodbar As Epel_BarCode
    Dim Resp As Long
    Dim CadMod As String
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(23)
    If MachineID = 1000 Then
        If Modificaciones Then CadMod = " and tran_en<>'*' "
    Else
        If Modificaciones Then CadMod = " and tran_el<>'*' "
    End If
    If MachineID = 1000 And TipoV8 Then
        frmControl.Switch_mode 9, NPUerto, NMAquina
    Else
        frmControl.Switch_mode 8, NPUerto, NMAquina
    End If
    GB_MarcaCodbarCableyAjenos False, MachineID
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from codbar where codi_ident=" & NSec & CadMod & " order by borrado")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If MachineID = 1000 And TipoV8 Then
                    micodbar.Net_Sale = .Fields("cb_vnt")
                    micodbar.Net_Super = ""
                    micodbar.Net_Mix = ""
                Else
                    micodbar.Net_Sale = .Fields("cb_vnt")
                    micodbar.Net_Super = .Fields("cb_ev1")
                    micodbar.Net_Mix = .Fields("cb_ev2")
                End If
                On Error Resume Next
                Resp = frmControl.Balanza.Send_BarCode(micodbar)
                On Error GoTo 0
                Select Case .Fields("borrado")
                    Case False
                        If Resp = 0 Then
                            Edit_Record Registro
                            If MachineID = 1000 Then
                                .Fields("tran_en") = "*"
                            Else
                                .Fields("tran_el") = "*"
                            End If
                            .Update
                            frmMonitor.MostrarDato "Cod. Bar --> OK"
                        Else
                            frmMonitor.MostrarDato "Cod. Bar --> NOK"
                        End If
                    Case True
                        If Resp < 1000 Then
                            Edit_Record Registro
                            frmMonitor.MostrarDato "Cod. Bar --> OK"
                            If Not UsaGamaAlta Then .Fields("tran_cb") = "*"
                            If Not usaSC10 Then .Fields("tran_cbsc10") = "*"
                            If Not UsaEcoPrint Then .Fields("tran_ep") = "*"
                            If MachineID = 1000 Then
                                .Fields("tran_en") = "*"
                                If Not UsaECO2 Then .Fields("tran_el") = "*"
                            Else
                                .Fields("tran_el") = "*"
                                If Not UsaECO1 Then .Fields("tran_en") = "*"
                            End If
                            .Update
                        Else
                            frmMonitor.MostrarDato "Cod. Bar --> NOK"
                        End If
                End Select
                .Movenext
            Loop
        End If
    End With
    Base.Close
    frmControl.Switch_mode 1000
End Sub
Public Sub GB_EnviaCableY(ByVal MachineID As Integer, ByVal NSec As Integer, ByVal NPUerto As Integer, ByVal NMAquina As Integer, ByVal Modificaciones As Boolean)
   Dim Base As Database
   Dim Resp As Long
   Dim MyModelo As Integer
   Dim Registro As Recordset
   Dim MyCab As Epel_HeadingLegend
   Dim CadMod As String
   frmMonitor.MostrarInfo CargaCadena(232)
   Select Case MachineID
        Case 1000
            If Modificaciones Then CadMod = " and tran_en<>'*'"
        Case 1001
            If Modificaciones Then CadMod = " and tran_el<>'*'"
    End Select
    If TipoV8 And MachineID = 1000 Then
        frmControl.Switch_mode 9, NPUerto, NMAquina
    Else
        frmControl.Switch_mode 8, NPUerto, NMAquina
    End If
    GB_MarcaCodbarCableyAjenos True, MachineID
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cabley where codi_ident=" & NSec & CadMod & " order by borrado")
    With Registro
        If Not .EOF Then
            Do Until .EOF
                If .Fields("borrado") = True Then
                    MyCab.Heading1 = ""
                    MyCab.Heading2 = ""
                    MyCab.Heading3 = ""
                    MyCab.Legend1 = ""
                Else
                    MyCab.Heading1 = .Fields("cabecera01")
                    MyCab.Heading2 = .Fields("cabecera02")
                    MyCab.Heading3 = .Fields("cabecera03")
                    MyCab.Legend1 = .Fields("leyenda01")
                End If
                On Error Resume Next
                Resp = frmControl.Balanza.Send_HeadingLegend(MyCab)
                On Error GoTo 0
                If .Fields("borrado") = True Then
                    If Resp >= 1000 Then
                        frmMonitor.MostrarDato CargaCadena(232) & " --> NOK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(232) & " --> OK"
                        Edit_Record Registro
                        If MachineID = 1000 Then
                            .Fields("tran_en") = "*"
                            If Not UsaECO2 Then .Fields("tran_el") = "*"
                        Else
                            .Fields("tran_el") = "*"
                            If Not UsaECO1 Then .Fields("tran_el") = "*"
                        End If
                        If Not UsaGamaAlta Then .Fields("tran_cab") = "*"
                        If Not usaSC10 Then .Fields("tran_cabsc10") = "*"
                        If Not UsaEcoPrint Then .Fields("tran_ep") = "*"
                        .Update
                    End If
                Else
                    If Resp <> 0 Then
                        frmMonitor.MostrarDato CargaCadena(232) & " --> NOK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(232) & " --> OK"
                        Edit_Record Registro
                        If MachineID = 1000 Then
                            .Fields("tran_en") = "*"
                        Else
                            .Fields("tran_el") = "*"
                        End If
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    Base.Close
    frmControl.Switch_mode 1000
End Sub
Public Sub GB_EnviaArt(ByVal MachineID As Integer, ByVal NSec As Integer, ByVal NPUerto As Integer, ByVal NMAquina As Integer, ByVal Modificaciones As Boolean)
    Dim Base As Database
    Dim Resp As Long
    Dim MyModelo As String
    Dim RegSec As Recordset
    Dim Registro As Recordset
    
    Dim nRTec As Recordset
    
    Dim MyArt As Epel_FullItem
    Dim MyControl As String
    Dim CadMod As String
    Dim RegConserv As Recordset
    Dim RegTara As Recordset
    Dim Bucle As Integer
    Dim nFich As Integer
    Dim sT As String
    Dim xPrice As Long

    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(12)
    Select Case MachineID
        Case 1000
            MyModelo = "ECON"
            MyControl = "tran_en"
            If Modificaciones Then CadMod = " and tran_en<>'*' "
            If TipoV8 Then
                frmControl.Switch_mode 9, NPUerto, NMAquina
            Else
                frmControl.Switch_mode 8, NPUerto, NMAquina
            End If
        Case 1001
            MyModelo = "ECOL"
            MyControl = "tran_el"
            If Modificaciones Then CadMod = " and tran_el<>'*' or tran_textoel<>'*' "
            frmControl.Switch_mode 8, NPUerto, NMAquina
    End Select
    Set Base = OpenDatabase(Base_General)
    '************************************
    ' Marca los que no corresponde enviar
    ' (flag BORRADO=true, debe marcarse
    ' tran_en y tran_el para que se borren
    ' de la base)
    '************************************
    Set RegSec = Base.OpenRecordset("select codi_ident from seccion where borrado=false and enviardatos<>" & Chr(34) & MyModelo & Chr(34))
    With RegSec
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & .Fields("codi_ident"))
                With Registro
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            Edit_Record Registro
                            .Fields(MyControl) = "*"
                            .Update
                            .Movenext
                        Loop
                    End If
                End With
                .Movenext
            Loop
        End If
    End With
    '************************************
    ' Envía los que proceda
    '************************************
    
    Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & NSec & CadMod & " order by borrado,plu")
    With Registro
        If Not .EOF Then
            .MoveFirst
            .MoveLast
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = .Recordcount + 2
            .MoveFirst
            Do Until .EOF
                    On Error Resume Next
                    frmMonitor.barra1.Value = frmMonitor.barra1.Value + 1
                    On Error GoTo 0
                    MyArt.Caducity = .Fields("caducidad")
                    MyArt.Code = .Fields("codigo")
                    If .Fields("codi_pes") = "U" Then
                        MyArt.Weight = 0
                    Else
                        MyArt.Weight = 1
                    End If
                    If Not (MachineID = 1000 And TipoV8) Then
                        If .Fields("prc100g") Then MyArt.Weight = 2
                        '1.7.17 se añaden espacios
                        MyArt.Desc(0) = Space(25 - Len(Trim(.Fields("des_plu2")))) & Trim(.Fields("des_plu2"))
                        MyArt.Desc(1) = Space(25 - Len(Trim(.Fields("des_plu3")))) & Trim(.Fields("des_plu3"))
                        MyArt.Desc(2) = Space(25 - Len(Trim(.Fields("des_plu4")))) & Trim(.Fields("des_plu4"))
                        MyArt.Desc(3) = Space(25 - Len(Trim(.Fields("des_plu5")))) & Trim(.Fields("des_plu5"))
                        MyArt.Desc(4) = Space(25 - Len(Trim(.Fields("des_plu6")))) & Trim(.Fields("des_plu6"))
                        MyArt.Desc(5) = Space(25 - Len(Trim(.Fields("des_plu7")))) & Trim(.Fields("des_plu7"))
                        MyArt.Desc(6) = Space(25 - Len(Trim(.Fields("des_plu8")))) & Trim(.Fields("des_plu8"))
                        MyArt.Desc(7) = Space(25 - Len(Trim(.Fields("des_plu9")))) & Trim(.Fields("des_plu9"))
                        MyArt.Desc(8) = Space(25 - Len(Trim(.Fields("des_plu0")))) & Trim(.Fields("des_plu0"))
                        MyArt.Desc(9) = Space(25 - Len(Trim(.Fields("des_plux")))) & Trim(.Fields("des_plux"))
                        '
                        
                        If id < 6000 And Not (UsaEcoPrint And TipoEcoPrint = 3) And Not (TipoV8) Then 'idiomas menores que griego pasar a mayúsculas... 2.0.a
                            MyArt.Desc(0) = UCase(MyArt.Desc(0))
                            MyArt.Desc(1) = UCase(MyArt.Desc(1))
                            MyArt.Desc(2) = UCase(MyArt.Desc(2))
                            MyArt.Desc(3) = UCase(MyArt.Desc(3))
                            MyArt.Desc(4) = UCase(MyArt.Desc(4))
                            MyArt.Desc(5) = UCase(MyArt.Desc(5))
                            MyArt.Desc(6) = UCase(MyArt.Desc(6))
                            MyArt.Desc(7) = UCase(MyArt.Desc(7))
                            MyArt.Desc(8) = UCase(MyArt.Desc(8))
                            MyArt.Desc(9) = UCase(MyArt.Desc(9))
                        
                        End If
                        
                        If UsaGruposdeConservacion Then
                            '
                            If .Fields("grupo_conserv") <> 0 Then
                                Set RegConserv = Base.OpenRecordset("select * from grpconserv where grupo=" & .Fields("grupo_conserv"))
                                If Not RegConserv.EOF Then
                                    MyArt.Desc(9) = RegConserv.Fields("txtsc10")
                                Else
                                    MyArt.Desc(9) = " "
                                End If
                            End If
                            
                        End If
                        If .Fields("tran_textoel") <> "*" Then
                            For Bucle = 0 To 9
                                If MyArt.Desc(Bucle) = "" Then MyArt.Desc(Bucle) = Space(25)
                            Next Bucle
                        End If
                        
                        'temporización marte IC-ILC nueva...
                        If Dir(App.Path & "\delayic.txt") <> "" Then
                            nFich = FreeFile()
                            Open App.Path & "\delayic.txt" For Input As #nFich
                            Line Input #nFich, sT
                            Close #nFich
                            Sleep (Val(sT))
                        End If
                    End If
                    MyArt.Family = .Fields("codi_fam")
                    If TipoV8 Then
                       If MyArt.Family > 99 Then MyArt.Family = 0
                    Else
                       If MyArt.Family > 15 Then MyArt.Family = 0
                    End If
                    If UsaEcoPrint And TipoEcoPrint = 3 Then
                        If MyArt.Family = 0 Or MyArt.Family > 15 Then
                            MyArt.Family = 1
                        End If
                    End If
                    
                    MyArt.plu = .Fields("plu")
                    MyArt.Pref = .Fields("pref")
                    If UsaEuro Then
                        MyArt.Price = .Fields("euros") * 100
                    Else
                        MyArt.Price = .Fields("precio") * (10 ^ decimales)
                    End If
                    MyArt.sec = 1
                    MyArt.subsec = .Fields("CODI_SUB")
                    
                    MyArt.Tare = 0
                    If Not (MachineID = 1000 And TipoV8) Then
                        Set RegTara = Base.OpenRecordset("select * from tipotara where valor=" & .Fields("tara"))
                        With RegTara
                            If Not .EOF Then MyArt.Tare = .Fields("tipo")
                        End With
                    End If
                    '1.7.17
                    If UsaECO2 Then
                        MyArt.TexT = Space(25 - Len(Trim(.Fields("des_plu1")))) & Trim(.Fields("des_plu1"))

                    Else
                        MyArt.TexT = .Fields("des_plu1")
                    End If
                    
                    If id < 6000 And Not (UsaEcoPrint And TipoEcoPrint = 3) And Not (TipoV8) Then 'idiomas menores que griego pasar a mayúsculas... 2.0.a
                        MyArt.TexT = Trim(MyArt.TexT)
                    End If
                    
                    If UsaEcoPrint And TipoEcoPrint = 3 Then 'IC+
                        MyArt.vat = .Fields("tipo_iva")
                    Else
                        MyArt.vat = .Fields("tipo_iva") + 1
                    End If
                    
                    'ic+
                    If IsNull(.Fields("art_cb")) Then
                        MyArt.EAN13 = Space(12)
                    Else
                        MyArt.EAN13 = Trim(.Fields("art_cb"))
                    End If
                    ''''
                    
                    If UsaECO2 And lMarteNew Then '2.0.9 Dir(App.Path & "\marte.new") <> "" Then
                        Set nRTec = Base.OpenRecordset("select * from teclas where codigo=" & CStr(MyArt.Code))
                        If Not nRTec.EOF Then
                            'c2f h+ 2.0.7 v4 alfa... se introduce tabla.
                            MyArt.nPeso = (nRTec.Fields("tabla") * 30) + nRTec.Fields("codi_tec")
                        Else
                            MyArt.nPeso = 0
                        End If
                        nRTec.Close
                        Set nRTec = Nothing
                    Else
                        MyArt.nPeso = 0
                    End If
                    
                    If .Fields("borrado") = False Then
                        '2.0.9 If Dir(App.Path & "\marte.new") <> "" Then '2.0.3 --> Polska
                        If lMarteNew Then
                            If Dir(App.Path & "\bajaprevia.txt") <> "" Then
                                Resp = frmControl.Balanza.Erase_Item(0, (10000 * 1) + .Fields("plu"))
                            End If
                            If Dir(App.Path & "\dobleenv.txt") <> "" Then
                                xPrice = MyArt.Price
                                MyArt.Price = 1
                                Resp = frmControl.Balanza.Send_FullItem(MyArt)
                                MyArt.Price = xPrice
                            End If
                        End If
                        Resp = frmControl.Balanza.Send_FullItem(MyArt)
                        If Resp = 0 Then
                            frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->OK"
                            Edit_Record Registro
                            .Fields(MyControl) = "*"
                            If (Not MachineID = 1000 And TipoV8) Then .Fields("tran_textoel") = "*"
                            .Update
                        Else
                            frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->NOK"
                        End If
                    Else
                        Resp = frmControl.Balanza.Erase_Item(0, (10000 * 1) + .Fields("plu"))
                        If Resp < 1000 Then
                            frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->OK"
                            Edit_Record Registro
                            .Fields(MyControl) = "*"
                            .Fields("tran_textoel") = "*"
                            If Not UsaGamaAlta Then .Fields("tran_plu") = "*"
                            If Not usaSC10 Then .Fields("tran_pluSC10") = "*"
                            If Not UsaEcoPrint Then .Fields("tran_ep") = "*"
                            Select Case MachineID
                                Case 1000
                                    If Not UsaECO2 Then .Fields("tran_el") = "*"
                                Case 1001
                                    If Not UsaECO1 Then .Fields("tran_en") = "*"
                            End Select
                            .Update
                        Else
                            frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->NOK"
                        End If
                    End If
           
                If Resp >= 1000 Or canCelar = 1 Then .MoveLast
                
                .Movenext
            Loop
        End If
    End With
    
    Base.Close
    frmControl.Switch_mode 1000
End Sub

Public Sub GB_EnviaIVAS(ByVal MachineID As Integer, ByVal NPUerto As Integer, ByVal NMAquina As Integer, ByVal Modificaciones As Boolean)
    '********************************
    ' Envío de tipos de IVA
    ' El flag "Modificaciones"
    ' no se tiene en cuenta por ahora
    '********************************
    Dim Base As Database
    Dim Registro As Recordset
    Dim Resp As Long
    Dim MisTipos(5) As String
    Dim MyVAT As Epel_Vat
    Select Case MachineID
        Case 1000
            If TipoV8 Then
                frmControl.Switch_mode 9, NPUerto, NMAquina
            Else
                frmControl.Switch_mode 8, NPUerto, NMAquina
            End If
        Case 1001
            frmControl.Switch_mode 8, NPUerto, NMAquina
    End Select
    frmMonitor.MostrarInfo CargaCadena(8)
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from iva")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Select Case (.Fields("Tipo"))
                    Case 1
                        MyVAT.type1 = .Fields("valor") * 100
                    Case 2
                        MyVAT.type2 = .Fields("valor") * 100
                    Case 3
                        MyVAT.type3 = .Fields("valor") * 100
                    Case 4
                        MyVAT.type4 = .Fields("valor") * 100
                    Case 5
                        MyVAT.type5 = .Fields("valor") * 100
                End Select
                .Movenext
            Loop
            .MoveFirst
        End If
    End With
    Resp = frmControl.Balanza.Send_Vat(MyVAT)
    If Resp <> 0 Then
        frmMonitor.MostrarDato CargaCadena(544)
    Else
        frmMonitor.MostrarDato CargaCadena(8) & " --> OK"
        If Not Registro.EOF Then
            Registro.MoveFirst
            Do Until Registro.EOF
                Edit_Record Registro
                Registro.Fields("tran_EN") = "*"
                Registro.Fields("tran_EL") = "*"
                Registro.Update
                Registro.Movenext
            Loop
        End If
    End If
    Base.Close
    frmControl.Switch_mode 1000
End Sub


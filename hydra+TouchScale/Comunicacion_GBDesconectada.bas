Attribute VB_Name = "Comunicacion_GBDesconectada"
Option Explicit
'***********************************************************
' Funciones de comunicación para máquinas de Gama Baja
' desconectadas, que comprenden las IV4 (todos los modelos)
' y las ECOPRINT
'***********************************************************
Public MaquinasGB(99) As Integer
Public SeccionesGBD(99) As Integer
Public Function GBD_SolicitarMaquina() As Integer
    '***********************************
    ' Función para especificar con quién se comunicará
    ' si está activada la opción de comunicar con una sóla máquina
    '***********************************
    Dim Resp As String
    Dim MyResp As Integer
    Dim Bucle As Integer
    MyResp = 0
    Resp = InputBox(CargaCadena(1201), "Hydra+")
    If IsNumeric(Resp) Then
        For Bucle = 1 To Len(Resp)
            If Not IsNumeric(Mid(Resp, Bucle, 1)) Then Resp = "0"
        Next Bucle
        MyResp = CInt(Resp)
    End If
    GBD_SolicitarMaquina = MyResp
End Function
Public Sub CalculaMaquinasGB()
    '**********************************
    ' Obtiene las máquinas existentes para
    ' realizar los ciclos de comunicación
    ' con cada máquina
    '**********************************
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim MyMaq As Integer
    Dim Bucle As Integer
    For Bucle = 0 To 99
        MaquinasGB(Bucle) = 0
    Next Bucle
    If EcoSoloUna Then
        MyMaq = GBD_SolicitarMaquina()
        If MyMaq = 0 Then Exit Sub
    End If
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select numero_eqp,codi_ident from equipos where borrado=false and modelo=400")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If EcoSoloUna Then
                    If .Fields("numero_eqp") = MyMaq Then MaquinasGB(.Fields("numero_eqp")) = 1
                Else
                    MaquinasGB(.Fields("numero_eqp")) = 1
                    SeccionesGBD(.Fields("numero_eqp")) = .Fields("codi_ident")
                End If
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
Public Sub GBD_PideArticulo(Articulo As DB_Articulo, ByVal Eqp As Integer)
    Dim MyArt As Epel_FullItem
    Dim Bucle As Integer
    Dim Resp As Long
    Dim Continuar As Boolean
    
    
    Continuar = True
    
    frmControl.Switch_mode 500 + Eqp
    If canCelar = 1 Then Exit Sub
    Do While Continuar
        Continuar = False
        MyArt.plu = Articulo.plu
        FrmConsultaArt.LblInfo.Caption = CargaCadena(971) & " EcoPrint/IV4:" & Articulo.plu
        Resp = frmControl.Balanza.Query_Item(MyArt, 0) ' Sin descriptivos
        If Trim(MyArt.TexT) = "" And MyArt.Price = 0 And Resp = 0 Then
            Continuar = True
            Articulo.plu = Articulo.plu + 1
        End If
    Loop
    If Resp = 0 Then
        Articulo.plu = MyArt.plu
        Articulo.Mostrador = 0
        Articulo.codigo = MyArt.Code
        Articulo.Balenv = False
        Articulo.caducidad = MyArt.Caducity
        Articulo.Descriptivos(0) = MyArt.TexT
        For Bucle = 0 To 19
            Articulo.Descriptivos(Bucle + 1) = MyArt.Desc(Bucle)
        Next Bucle
        Articulo.EAN13 = ""
        Articulo.Etiqueta = 255
        Articulo.familia = MyArt.Family
        Articulo.GrpConserv = 0
        If MyArt.vat <= 0 Then
            Articulo.IVA = 0
        Else
            Articulo.IVA = MyArt.vat  'MyArt.vat - 1
        End If
        Articulo.Merma = 0
        Articulo.precio = MyArt.Price
        Articulo.Preferente = MyArt.Pref
        For Bucle = 0 To 9
            Articulo.TipoLetra(Bucle) = 0
        Next Bucle
        Articulo.Presel = 0
        Articulo.subsec = 0
        If UsaEcoPrint And TipoEcoPrint = 3 Then
           If MyArt.Weight = 1 Then
               Articulo.WGH = True
           Else
               Articulo.WGH = False
           End If
        Else
           Articulo.WGH = True
        End If
    Else
        Articulo.plu = -1
    End If
    
End Sub
Public Function GBD_Query_Tot(MachineID As Boolean, FinDia As Boolean)
    '******************************
    ' Totales por artículo
    ' para máquinas desconectadas
    '******************************
    Dim Bucle As Integer
    Dim Resp As Integer
    Dim sTotArtPath As String
    Dim sTotVenPath As String
    Dim sTotMaqPath As String
    
    ' obtiene las máquinas existentes
    frmControl.Switch_mode (MachineID)
    If canCelar = 1 Then Exit Function
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
    
    frmMonitor.MostrarInfo CargaCadena(156)
    Resp = frmControl.Balanza.Query_All_Totals(sTotArtPath, 0, 0)
  
    If Resp = 0 And canCelar = 0 And Not CancelHumano Then
        frmMonitor.MostrarInfo CargaCadena(157)
        Resp = frmControl.Balanza.Query_All_Totals(sTotVenPath, 1, 0)
    End If
    
    If Resp = 0 And canCelar = 0 And Not CancelHumano Then
        frmMonitor.MostrarInfo CargaCadena(155)
        Resp = frmControl.Balanza.Query_All_Totals(sTotMaqPath, 2, 0)
    End If
    
    'c2f borrado gama baja revisar.
    'If resp = 0 And canCelar = 0 And Not CancelHumano Then
        If FinDia Then
           frmMonitor.MostrarInfo CargaCadena(282)
           frmControl.Balanza.Erase_Totals 0
        End If
    'End If


    frmControl.Switch_mode 1000
End Function
Public Sub GBD_EnviaBaseMaestra(Modificaciones As Boolean)
    '******************************
    ' Base maestra para máquinas
    ' desconectadas
    '******************************
    Dim Bucle As Integer
    ' obtiene las máquinas existentes
    CalculaMaquinasGB
    ' realiza el envío para cada
    ' máquina de la red
    For Bucle = 0 To 99
        If MaquinasGB(Bucle) = 1 Then
            canCelar = 0
            CancelHumano = False
            GBD_EnviaCodBar Bucle, Modificaciones
            If (canCelar = 0) And (CancelHumano = False) And (TipoEcoPrint <> 3) Then GBD_EnviaCableY Bucle, Modificaciones
            If (canCelar = 0) And (CancelHumano = False) And (TipoEcoPrint = 3) Then GBD_EnviaCableYIC Bucle, Modificaciones
            If (canCelar = 0) And (CancelHumano = False) Then GBD_EnviaIVAS Bucle, Modificaciones
            If (canCelar = 0) And (CancelHumano = False) Then GBD_EnviaArt Bucle, Modificaciones
            If CancelHumano Then Exit Sub
        End If
    Next Bucle
End Sub

Private Sub GBD_MarcaArticulos_Ajenos()
    '********************************************
    ' Marca los datos borrados que no
    ' corresponden a máquinas desconectadas
    ' (Artículos)
    '********************************************
    Dim Base As Database
    Dim Registro As Recordset
    Dim RegSec As Recordset
    Set Base = OpenDatabase(Base_General)
    Set RegSec = Base.OpenRecordset("select codi_ident,enviardatos from seccion where enviardatos<>" & Chr(34) & "ECOP" & Chr(34))
    If Not RegSec.EOF Then
  
        RegSec.MoveFirst
        Do Until RegSec.EOF
            Set Registro = Base.OpenRecordset("select * from articulo where borrado=true and codi_ident=" & RegSec.Fields("codi_ident"))
            With Registro
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_ep") = String(60, "*")
                        .Update
                        .Movenext
                    Loop
                End If
            End With
            RegSec.Movenext
        Loop
    End If
    Base.Close
End Sub
Private Sub GBD_MarcaCodBarCableyAjenos(EsCabley As Boolean)
    '********************************************
    ' Marca los datos borrados que no
    ' corresponden a máquinas desconectadas
    ' (Códigos de Barras y Cabeceras y Leyendas)
    '********************************************
    Dim Base As Database
    Dim Buf As Integer
    Dim RegMaq As Recordset
    Dim Registro As Recordset
    Set Base = OpenDatabase(Base_General)
    If EsCabley Then
        Set Registro = Base.OpenRecordset("select * from cabley where borrado=true")
    Else
        Set Registro = Base.OpenRecordset("select * from codbar where borrado=true")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                ' marca los códigos de barras por mostrador
                Buf = 0
                If Not IsNull(.Fields("numero_eqp")) Then
                    If IsNumeric(.Fields("numero_eqp")) Then Buf = .Fields("numero_eqp")
                End If
                If Buf = 0 Then
                    Edit_Record Registro
                    .Fields("tran_ep") = "*"
                    .Update
                Else
                    ' marca los de máquinas no desconectadas
                    Set RegMaq = Base.OpenRecordset("select modelo from equipos where numero_eqp=" & .Fields("numero_eqp"))
                    If Not .EOF Then
                        If RegMaq.Fields("modelo") <> 400 Then
                            Edit_Record Registro
                            .Fields("tran_ep") = "*"
                            .Update
                        End If
                    Else
                        Edit_Record Registro
                        .Fields("tran_ep") = "*"
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    Base.Close
End Sub

Public Sub GBD_EnviaCodBar(NumMaq As Integer, Modificaciones As Boolean)
    Dim Base As Database
    Dim Registro As Recordset
    Dim micodbar As Epel_BarCode
    Dim Resp As Long
    Dim CadMod As String
    '******************************************
    ' envío de códigos de barras a una máquina
    '*******************************************
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(23)
    frmControl.Switch_mode 500 + NumMaq
    If canCelar = 1 Then Exit Sub
    If Modificaciones Then CadMod = " and tran_ep<>'*'"
    
    '********************************************
    ' 1º marca los datos borrados que no
    ' corresponden a máquinas desconectadas
    '********************************************
    GBD_MarcaCodBarCableyAjenos False
    '********************************************
    ' 2º envía los datos que corresponden con la
    ' máquina elegida
    '********************************************
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from codbar where numero_eqp=" & NumMaq & CadMod & " order by borrado")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                micodbar.Net_Sale = .Fields("cb_vnt")
                micodbar.Net_Super = .Fields("cb_ev1")
                micodbar.Net_Mix = .Fields("cb_ev2")
                On Error Resume Next
                Resp = frmControl.Balanza.Send_BarCode(micodbar)
                On Error GoTo 0
                Select Case .Fields("borrado")
                    Case False
                        If Resp = 0 Then
                            Edit_Record Registro
                            .Fields("tran_ep") = "*"
                            .Update
                            frmMonitor.MostrarDato "Cod. Bar --> OK"
                        Else
                            frmMonitor.MostrarDato "Cod. Bar --> NOK"
                        End If
                    Case True
                        If Resp < 1000 Then
                            Edit_Record Registro
                            frmMonitor.MostrarDato "Cod. Bar --> OK"
                            .Fields("tran_ep") = "*"
                            .Fields("tran_bar") = "*"
                            .Fields("tran_barsc10") = "*"
                            .Fields("tran_en") = "*"
                            .Fields("tran_el") = "*"
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
Public Sub GBD_EnviaCableY(NumMaq As Integer, Modificaciones As Boolean)
   Dim Base As Database
   Dim Resp As Long
   Dim MyModelo As Integer
   Dim Registro As Recordset
   Dim MyCab As Epel_HeadingLegend
   Dim CadMod As String
   frmMonitor.MostrarInfo CargaCadena(232)
   frmControl.Switch_mode 500 + NumMaq
   If canCelar = 1 Then Exit Sub
   If Modificaciones Then CadMod = " and tran_ep<>'*'"
    '********************************************
    ' 1º marca los datos borrados que no
    ' corresponden a máquinas desconectadas
    '********************************************
    GBD_MarcaCodBarCableyAjenos True
    '********************************************
    ' 2º envía los datos que corresponden con la
    ' máquina elegida
    '********************************************
    Resp = 0
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cabley where numero_eqp=" & NumMaq & CadMod & " order by borrado")
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
                        .Fields("tran_ep") = "*"
                        .Fields("tran_cab") = "*"
                        .Fields("tran_cabsc10") = "*"
                        .Fields("tran_en") = "*"
                        .Fields("tran_el") = "*"
                        .Update
                    End If
                Else
                    If Resp <> 0 Then
                        frmMonitor.MostrarDato CargaCadena(232) & " --> NOK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(232) & " --> OK"
                        Edit_Record Registro
                        .Fields("tran_ep") = "*"
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    Base.Close
    If Resp <> 0 Then canCelar = 1   '2.0.23
    frmControl.Switch_mode 1000
End Sub
'c2f iv4 ic+ ... errores en tipo de letra.
Public Sub GBD_EnviaCableYIC(NumMaq As Integer, Modificaciones As Boolean)
   Dim Base As Database
   Dim Resp As Long
   Dim MyModelo As Integer
   Dim Registro As Recordset
   Dim MyCab As Epel_HeadingLegend
   Dim CadMod As String
   frmMonitor.MostrarInfo CargaCadena(232)
   frmControl.Switch_mode 500 + NumMaq
   If canCelar = 1 Then Exit Sub
   If Modificaciones Then CadMod = " and tran_ep<>'*'"
    '********************************************
    ' 1º marca los datos borrados que no
    ' corresponden a máquinas desconectadas
    '********************************************
    GBD_MarcaCodBarCableyAjenos True
    '********************************************
    ' 2º envía los datos que corresponden con la
    ' máquina elegida
    '********************************************
    Resp = 0
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cabley where numero_eqp=" & NumMaq & CadMod & " order by borrado")
    With Registro
        If Not .EOF Then
            Do Until .EOF
                If .Fields("borrado") = True Then
                    MyCab.Heading1 = ""
                    MyCab.Heading2 = ""
                    MyCab.Heading3 = ""
                    MyCab.Heading4 = ""
                    MyCab.Legend1 = ""
                    MyCab.Legend2 = ""
                    MyCab.FHeading1 = "0"
                    MyCab.FHeading2 = "0"
                    MyCab.FHeading3 = "0"
                    MyCab.FHeading4 = "0"
                    MyCab.FLegend1 = "0"
                    MyCab.FLegend2 = "0"
                Else
                    MyCab.Heading1 = .Fields("cabecera01")
                    MyCab.Heading2 = .Fields("cabecera02")
                    MyCab.Heading3 = .Fields("cabecera03")
                    MyCab.Heading4 = .Fields("cabecera04")
                    MyCab.Legend1 = .Fields("leyenda01")
                    MyCab.Legend2 = .Fields("leyenda02")
                    MyCab.FHeading1 = .Fields("tletracab1")
                    MyCab.FHeading2 = .Fields("tletracab2")
                    MyCab.FHeading3 = .Fields("tletracab3")
                    MyCab.FHeading4 = .Fields("tletracab4")
                    MyCab.FLegend1 = .Fields("tletraley1")
                    MyCab.FLegend2 = .Fields("tletraley2")
                
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
                        .Fields("tran_ep") = "*"
                        .Fields("tran_cab") = "*"
                        .Fields("tran_cabsc10") = "*"
                        .Fields("tran_en") = "*"
                        .Fields("tran_el") = "*"
                        .Update
                    End If
                Else
                    If Resp <> 0 Then
                        frmMonitor.MostrarDato CargaCadena(232) & " --> NOK"
                    Else
                        frmMonitor.MostrarDato CargaCadena(232) & " --> OK"
                        Edit_Record Registro
                        .Fields("tran_ep") = "*"
                        .Update
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
    Base.Close
    If Resp <> 0 Then canCelar = 1   '2.0.23
    frmControl.Switch_mode 1000
End Sub

Public Sub GBD_EnviaArt(NumMaq As Integer, Modificaciones As Boolean)
    Dim Base As Database
    Dim Resp As Long
    Dim MyModelo As String
    Dim RegSec As Recordset
    Dim Registro As Recordset
    Dim MyArt As Epel_FullItem
    Dim MyControl As String
    Dim SeEnvia As Boolean
    Dim CadenaMod As String
    Dim RegTara As Recordset
    Dim Bucle As Integer
    frmMonitor.MostrarInfo CargaCadena(738) & " " & CargaCadena(12)
    
    
    frmControl.Switch_mode 500 + NumMaq
    If canCelar = 1 Then Exit Sub
    
    '************************************
    ' 1º Marca los datos borrados que no
    '    proceda enviar
    '************************************
    'C2F 1.8.0
    'GBD_MarcaArticulos_Ajenos -->MUCHO OJO SE HA ANULADO
    '************************************
    ' 2º Envía los que proceda
    '************************************
    Resp = 0
    Set Base = OpenDatabase(Base_General)
    Set RegSec = Base.OpenRecordset("select codi_ident from equipos where borrado=false and numero_eqp=" & NumMaq)
    If Not RegSec.EOF Then
        
        Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & RegSec.Fields("codi_ident") & " order by borrado,plu")
        '................
        'Set Registro = Base.OpenRecordset("select * from articulo where tran_plu<>'*'")
        '................
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
                        SeEnvia = False
                        CadenaMod = ""
                        If Not IsNull(.Fields("tran_ep")) Then CadenaMod = .Fields("tran_ep")
                        If Len(CadenaMod) <> 60 Then CadenaMod = Space(60)
                        If Modificaciones Then
                            'C2F 1.8.0 If Mid(CadenaMod, NumMaq, 1) <> "*" Then SeEnvia = True
                            If (Mid(CadenaMod, NumMaq, 1) <> "*" Or .Fields("TRAN_PLU") <> "*") Then SeEnvia = True
                        Else
                            SeEnvia = True
                        End If
                        If SeEnvia Then
                            MyArt.Caducity = .Fields("caducidad")
                            MyArt.Code = .Fields("codigo")
                            
                            MyArt.Family = .Fields("codi_fam")
                            'ver aquí c2f
                            If TipoV8 <> 9 Then
                                If MyArt.Family > 15 Then MyArt.Family = 0
                            Else
                                If MyArt.Family > 99 Then MyArt.Family = 0
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
                            MyArt.TexT = .Fields("des_plu1")
                            If UsaEcoPrint And TipoEcoPrint = 3 Then 'ic+
                                MyArt.vat = .Fields("tipo_iva")
                            Else
                                MyArt.vat = .Fields("tipo_iva") + 1
                            End If
                            If .Fields("codi_pes") = "U" Then
                                MyArt.Weight = 0
                            Else
                                MyArt.Weight = 1
                            End If
                            
                            'ic+
                            If IsNull(.Fields("art_cb")) Then
                                MyArt.EAN13 = Space(12) '2.0.23
                            Else
                                MyArt.EAN13 = Trim(.Fields("art_cb"))
                            End If
                            ''''
                            
                            If .Fields("borrado") = False Then
                                Resp = frmControl.Balanza.Send_FullItem(MyArt)
                                If Resp = 0 Then
                                    frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->OK"
                                    Edit_Record Registro
                                    
                                    .Fields("TRAN_PLU") = "*"
                                    .Fields("TRAN_CB") = "*"
                                    
                                    If NumMaq = 1 Then
                                        .Fields("tran_ep") = "*" & Mid(CadenaMod, 2)
                                    Else
                                        .Fields("tran_ep") = Left(CadenaMod, NumMaq - 1) & "*" & Mid(CadenaMod, NumMaq + 1)
                                    End If
                                    .Update
                                Else
                                    frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->NOK"
                                End If
                            Else
                                Resp = frmControl.Balanza.Erase_Item(0, NumMaq * 10000 + .Fields("plu"))
                                If Resp < 1000 Then
                                    frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->OK"
                                    Edit_Record Registro
                                    .Fields("tran_en") = "*"
                                    .Fields("tran_el") = "*"
                                    
                                    .Fields("TRAN_PLU") = "*"
                                    .Fields("TRAN_CB") = "*"
                                    
                                    If Not UsaGamaAlta Then .Fields("tran_plu") = "*"
                                    If Not usaSC10 Then .Fields("tran_plusc10") = "*"
                                    If NumMaq = 1 Then
                                        .Fields("tran_ep") = "*" & Mid(CadenaMod, 2)
                                    Else
                                        .Fields("tran_ep") = Left(CadenaMod, NumMaq - 1) & "*" & Mid(CadenaMod, NumMaq + 1)
                                    End If
                                    .Update
                                Else
                                    frmMonitor.MostrarDato "Cod. " & Format(MyArt.Code, "000000") & " ---Plu. " & Format(MyArt.plu, "0000") & "," & MyArt.TexT & Space(30 - Len(MyArt.TexT)) & "-->NOK"
                                End If
                            End If
                        End If
                    If Resp >= 1000 Or canCelar = 1 Then .MoveLast
                    
                    If Resp >= 1000 Then canCelar = 1 '2.0.23
                    
                    .Movenext
                Loop
            End If
        End With
    End If
    Base.Close
    frmControl.Switch_mode 1000
End Sub

Public Sub GBD_EnviaIVAS(NumMaq As Integer, Modificaciones As Boolean)
    Dim Base As Database
    Dim Registro As Recordset
    Dim Resp As Long
    Dim MisTipos(5) As String
    Dim MyVAT As Epel_Vat
    frmControl.Switch_mode 500 + NumMaq
    If canCelar = 1 Then Exit Sub
    Resp = 0
    frmMonitor.MostrarInfo "Envío de Tipos de IVA"
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
                'Registro.Fields("tran_ep") = "*"
                Registro.Update
                Registro.Movenext
            Loop
        End If
    End If
    Base.Close
    If Resp <> 0 Then canCelar = 1   '2.0.23
    frmControl.Switch_mode 1000
End Sub
Public Sub GBD_SolicitaPedidoInventario(ByVal NumMaq As Integer)
    '///////////////////////////////////////
    '// Solicitud de pedido e inventario  //
    '// para IV4TQ                        //
    '///////////////////////////////////////
    Dim RutaPedido As String
    Dim Resp As Long
    frmControl.Switch_mode 500 + NumMaq
    If canCelar = 1 Then Exit Sub
    Resp = 0
    If Not HayMulti Then
        RutaPedido = App.Path & "\01pedinv.bal"
    Else
        If TiendaActual < 100 Then
        RutaPedido = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "pedinv.bal"
        Else
        RutaPedido = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "pedinv.bal"
        End If
    End If
    If Dir(RutaPedido) <> "" Then
        Kill RutaPedido
    End If
    If frmControl.EstadoBal < 1000 Then
        
        frmMonitor.MostrarInfo CargaCadena(1107)
        Resp = frmControl.Balanza.Query_Order_Inventory(RutaPedido)
        
        frmMonitor.Label1.Caption = CargaCadena(445)
        frmMonitor.cmdmonitorstop.Enabled = False
        frmMonitor.Command1.Enabled = True
    End If
    
    If Resp >= 1000 Then canCelar = 1 '2.0.23
    
    If canCelar = 0 Then
        frmMonitor.MostrarDato CargaCadena(1097) & " " & Mid(RutaPedido, 1, Len(RutaPedido) - 3) & "sdf"
    Else
        On Error Resume Next
        Kill RutaPedido
    End If
End Sub
Public Sub GBD_SolicitaFichIC()
    '///////////////////////////////////////
    '// Solicitud de Memoria de tique int.//
    '// para IV4 IC+                      //
    '///////////////////////////////////////
    Dim RutaCab As String
    Dim RutaLin As String
    Dim Resp As Long

    If canCelar = 1 Then Exit Sub
    If Not HayMulti Then
        RutaCab = App.Path & "\" & Format(Now, "ddmmyy") & "-01cabic.bal"
        RutaLin = App.Path & "\" & Format(Now, "ddmmyy") & "-01linic.bal"
    Else
        If TiendaActual < 100 Then
        RutaCab = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(Now, "ddmmyy") & "-" & Format(TiendaActual, "00") & "cabic.bal"
        RutaLin = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(Now, "ddmmyy") & "-" & Format(TiendaActual, "00") & "linic.bal"
        Else
        RutaCab = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(Now, "ddmmyy") & "-" & Format(TiendaActual, "000") & "cabic.bal"
        RutaLin = App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(Now, "ddmmyy") & "-" & Format(TiendaActual, "000") & "linic.bal"
        
        End If
    End If
        
    frmMonitor.MostrarInfo "Tickets IC+ Discharging..."
    Resp = frmControl.Balanza.Query_All_Ticket_IC(RutaCab, RutaLin)
    
    frmMonitor.Label1.Caption = CargaCadena(445)
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.Command1.Enabled = True

    If canCelar = 0 Then
        frmMonitor.MostrarDato CargaCadena(1226) & " " & Mid(RutaCab, 1, Len(RutaCab) - 3)
        frmMonitor.MostrarDato CargaCadena(1226) & " " & Mid(RutaLin, 1, Len(RutaLin) - 3)
    End If

End Sub

Public Sub Totaliza_Pedido_Inventario(ByVal RutaPedido As String)
    '/////////////////////////////////////////////////////
    ' IV4 TQ, los datos llegan con número de equipo
    ' más número de PLU. Aquí se obtiene el mostrador
    ' al que se refiere así como el dato de Pesado/No pesado
    ' Además ,se suman los pedidos e inventarios
    ' de varias máquinas para dar los totales por sección
    '//////////////////////////////////////////////////////
    Dim Arch1 As Integer
    Dim Arch2 As Integer
    Dim Vistos() As Long
    Dim NVistos As Long
    Dim RegArt As Recordset
    Dim Base As Database
    Dim Bucle As Long
    Dim Registro As Recordset
    Dim BufPedidos() As String
    Dim NBufPedidos As Long
    Dim Buf As String
    Dim Nuevo As Boolean
    Dim MyPed As Double
    Dim MyInv As Double
    Dim MyWHG As String
    Set Base = OpenDatabase(Base_General)
    NVistos = 0
    Arch1 = FreeFile()
    Open RutaPedido For Input As #Arch1
    
    Do Until EOF(Arch1)
        Line Input #Arch1, Buf
        If Len(Buf) = 28 And IsNumeric(Buf) Then
            Nuevo = True
            Set Registro = Base.OpenRecordset _
            ("select codi_ident from equipos where modelo=400 and borrado=false and numero_eqp=" & _
            Val(Left(Buf, 2)))
            Set RegArt = Base.OpenRecordset _
            ("select codi_pes from articulo where borrado=false and codigo=" & _
            Val(Mid(Buf, 7, 6)))
            MyWHG = "2"
            With RegArt
                If Not .EOF Then
                    If .Fields("codi_pes") = "W" Then
                        MyWHG = "1"
                    Else
                        MyWHG = "0"
                    End If
                End If
            End With
            With Registro
                If Not .EOF Then
                    Buf = Format(.Fields("codi_ident"), "000") & Mid(Buf, 3)
                    If NVistos > 0 Then
                        For Bucle = 0 To NVistos - 1
                            If Vistos(Bucle) = Val(Mid(Buf, 1, 7)) Then
                                Nuevo = False
                                Exit For
                            End If
                        Next Bucle
                    End If
                    If Nuevo Then
                        NBufPedidos = NBufPedidos + 1
                        ReDim Preserve BufPedidos(NBufPedidos)
                        BufPedidos(NBufPedidos - 1) = Buf & MyWHG
                        NVistos = NVistos + 1
                        ReDim Preserve Vistos(NVistos)
                        Vistos(NVistos - 1) = Val(Left(Buf, 7))
                    Else
                        For Bucle = 0 To NBufPedidos - 1
                            If Mid(BufPedidos(Bucle), 7) = Mid(Buf, 7) Then
                                MyInv = CDbl(Mid(BufPedidos(Bucle), 14, 8))
                                MyPed = CDbl(Mid(BufPedidos(Bucle), 22, 8))
                                MyInv = MyInv + CDbl(Mid(Buf, 14, 8))
                                MyPed = MyPed + CDbl(Mid(Buf, 22, 8))
                                BufPedidos(Bucle) = Mid(BufPedidos(Bucle), 1, 13) & Format(MyInv, "00000000") & Format(MyPed, "00000000") & Right(BufPedidos(Bucle), 1)
                                Exit For
                            End If
                        Next Bucle
                    End If
                End If
            End With
        End If
    Loop
    Close #Arch1
    Arch2 = FreeFile()
    Open Mid(RutaPedido, 1, Len(RutaPedido) - 3) & "sdf" For Output As #Arch2
    For Bucle = 1 To NBufPedidos
        Print #Arch2, BufPedidos(Bucle - 1)
    Next Bucle
    Close #Arch2
    Base.Close
End Sub


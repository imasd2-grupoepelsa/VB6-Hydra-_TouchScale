Attribute VB_Name = "Informes"
Option Explicit
Private LongitudTotal As Long
Public TipoOrden As Integer
Public L_WEB As Boolean
Private Sub AsignaEtiqueta(sNombre, sCampo As String, Ancho As Integer)
    Dim Bucle As Integer
    For Bucle = 1 To 15
        If Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Visible = False Then Exit For
    Next Bucle
    Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Left = LongitudTotal
    Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Width = 142 * Ancho
    Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).Left = LongitudTotal
    Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).Width = 142 * Ancho
    LongitudTotal = LongitudTotal + Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Width
    Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Caption = sNombre
    Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).Visible = True
    Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).DataField = sCampo
    Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).Visible = True
    If Dir(App.Path & "\etc\l_web.txt") <> "" Then
        Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).BackColor = &HFFFFFF
        Reporte.Sections("Sección2").Controls("Etiqueta" & Format(Bucle, "0")).ForeColor = 0
    End If
    For Bucle = 1 To 15
        If Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).Visible = False Then
            Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).DataField = sCampo
            Reporte.Sections("Sección1").Controls("Texto" & Format(Bucle, "0")).CanGrow = True
        End If
    Next Bucle
End Sub
Private Sub MuestraInforme()
    If Dir(App.Path & "\etc\l_web.txt") = "" Then
        DescargaReport = False
        Reporte.Show
        Do Until DescargaReport
            Sleep (100)
            Do_Events
        Loop
    Else
        Unload FrmSplash
        L_WEB = False
        FrmListOptions.Show
        FrmListOptions.Tag = 1
        Do Until L_WEB
            Sleep (100)
            DoEvents
        Loop
        Unload Reporte
    End If
End Sub
Private Sub MuestraInforme2L()
    If Dir(App.Path & "\etc\l_web.txt") = "" Then
        DescargaReport = False
        Reporte2L.Show
        Do Until DescargaReport
            Sleep (100)
            Do_Events
        Loop
    Else
        Unload FrmSplash
        L_WEB = False
        FrmListOptions.Show
        FrmListOptions.Tag = 2
        Do Until L_WEB
            Sleep (100)
            DoEvents
        Loop
        Unload Reporte2L
    End If
End Sub
Private Sub AsignaEtiqueta2L(sNombre, sCampo As String, Ancho As Integer)
Dim Bucle As Integer
    For Bucle = 2 To 16
        If Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Visible = False Then Exit For
    Next Bucle
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Left = LongitudTotal
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Width = 142 * Ancho
    Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).Left = LongitudTotal
    Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).Width = 142 * Ancho
    LongitudTotal = LongitudTotal + Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Width
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Caption = sNombre
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).Visible = True
    Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).DataField = sCampo
    Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).Visible = True
    If Dir(App.Path & "\etc\l_web.txt") <> "" Then
        Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).BackColor = &HFFFFFF
        Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta" & Format(Bucle, "0")).ForeColor = 0
    End If
    For Bucle = 2 To 16
        If Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).Visible = False Then
            Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).DataField = sCampo
        End If
    Next Bucle
End Sub
Public Sub Informe_CambiosPrecio()
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select fecha,codi_ident as agrupador,plu,codigo,codi_sub,codi_fam,des_plu1,precio,tipo_ven,comentario from log order by codi_ident,plu,codigo", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(302)
    AsignaEtiqueta2L CargaCadena(434), "fecha", 8
    AsignaEtiqueta2L CargaCadena(69), "codigo", 6
    AsignaEtiqueta2L CargaCadena(68), "plu", 3
    AsignaEtiqueta2L CargaCadena(10), "codi_sub", 4
    AsignaEtiqueta2L CargaCadena(67), "codi_fam", 4
    
    AsignaEtiqueta2L CargaCadena(71), "precio", 6
    AsignaEtiqueta2L CargaCadena(79), "tipo_ven", 5
    AsignaEtiqueta2L CargaCadena(70), "des_plu1", 12
    AsignaEtiqueta2L CargaCadena(48), "comentario", 12
    AsignaFuncion2l 1, 4, "codigo", False
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(24)
    Etiqueta_General2l CargaCadena(1189)
    Reporte2L.Refresh
    MuestraInforme2L
End Sub
Public Sub Informe_Etq8()
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select fecha,codi_ident as agrupador,plu,codigo,codi_sub,codi_fam,des_plu1,precio,tipo_ven,comentario from log order by codi_ident,plu,codigo", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load ReporE8
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(302)
    AsignaEtiqueta2L CargaCadena(434), "fecha", 8
    AsignaEtiqueta2L CargaCadena(69), "codigo", 6
    AsignaEtiqueta2L CargaCadena(68), "plu", 3
    AsignaEtiqueta2L CargaCadena(10), "codi_sub", 4
    AsignaEtiqueta2L CargaCadena(67), "codi_fam", 4
    
    AsignaEtiqueta2L CargaCadena(71), "precio", 6
    AsignaEtiqueta2L CargaCadena(79), "tipo_ven", 5
    AsignaEtiqueta2L CargaCadena(70), "des_plu1", 12
    AsignaEtiqueta2L CargaCadena(48), "comentario", 12
    AsignaFuncion2l 1, 4, "codigo", False
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(24)
    Etiqueta_General2l CargaCadena(1189)
    ReporE8.Orientation = rptOrientLandscape
    On Error Resume Next
    ReporE8.Refresh
    On Error GoTo 0
    DescargaReport = False
    ReporE8.Show
    Do Until DescargaReport
        Sleep (100)
        Do_Events
    Loop
    
End Sub


Public Sub Inicia_Funciones_2l(sCampo As String)
    Dim Bucle As Integer
    For Bucle = 1 To 6
        Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("Funcion" & Format(Bucle, "0")).Visible = False
        Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("Funcion" & Format(Bucle, "0")).DataField = sCampo
        Reporte2L.Sections("ReportFooter").Controls("General" & Format(Bucle, "0")).Visible = False
        Reporte2L.Sections("ReportFooter").Controls("General" & Format(Bucle, "0")).DataField = sCampo
    Next Bucle
End Sub
Public Sub AsignaFuncion2l(Numero As Integer, Tipo As Integer, sCampo As String, bGeneral As Boolean)
    Dim Bucle As Integer
    Dim MyVal As Integer
    Dim MySec As String
    Dim MyNom As String
    
   
    If bGeneral = False Then
        MySec = "Consulta2l_Grouping_Footer"
        MyNom = "Funcion"
        Reporte2L.Sections(MySec).Controls("Línea3").Visible = True
    Else
        MySec = "ReportFooter"
        MyNom = "General"
        Reporte2L.Sections(MySec).Controls("Etiqueta22").Visible = True
        Reporte2L.Sections(MySec).Controls("Forma3").Visible = True
    End If
    Reporte2L.Sections(MySec).Controls(MyNom & Format(Numero, "0")).Visible = True
    Reporte2L.Sections(MySec).Controls(MyNom & Format(Numero, "0")).DataField = sCampo
    Reporte2L.Sections(MySec).Controls(MyNom & Format(Numero, "0")).FunctionType = Tipo
    For Bucle = 2 To 16
        If Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(Bucle, "0")).Visible = False Then
            MyVal = Bucle - 1
            MyVal = MyVal - (6 - Numero)
            Exit For
        End If
    Next Bucle
    If Not lngt Then
        Reporte2L.Sections(MySec).Controls(MyNom & Format(Numero, "0")).Left = Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(MyVal, "0")).Left
        Reporte2L.Sections(MySec).Controls(MyNom & Format(Numero, "0")).Width = Reporte2L.Sections("consulta2l_detail").Controls("Texto" & Format(MyVal, "0")).Width
    End If
End Sub
Public Sub Etiqueta_Total2l(sCAd As String)
    Dim Bucle As Integer
    'Dim aRch As Integer
    'Dim sParte As String
    'Dim nGT As Integer
    
    'If Dir(App.Path & "\ngt.txt") <> "" And sOrdenNGT = "gtsecs" Then
    '    If Dir(sPathNGT & "ngtord") = "" Then
    '        nGT = 1
    '        nlastgt = 1
    '    Else
    '        aRch = FreeFile()
    '        Open sPathNGT & "ngtord" For Input As #aRch
    '        Line Input #aRch, sParte
    '        Close #aRch
    '        nlastgt = Val(sParte)
    '        Load frmnGT
    '        frmnGT.Visible = True
    '        Do While frmnGT.Visible = True
    '            Sleep (10)
    '            DoEvents
    '        Loop
    '    End If
    '    Reporte2L.Sections("ReportFooter").Controls("etiqueta22").Caption = sCAd & " N.G.T.:" & Format(nlastgt, "0000")
    '
    'Else
    'If Not lngt Then
    
        Reporte2L.Sections("ReportFooter").Controls("etiqueta22").Caption = sCAd
        
        Reporte2L.Sections("ReportFooter").Controls("etiqueta22").Visible = True
        For Bucle = 1 To 6
            If Reporte2L.Sections("ReportFooter").Controls("General" & Format(Bucle, "0")).Visible = True Then
                Reporte2L.Sections("ReportFooter").Controls("etiqueta22").Left = Reporte2L.Sections("ReportFooter").Controls("General" & Format(Bucle, "0")).Left - Reporte2L.Sections("ReportFooter").Controls("etiqueta22").Width - 100
                Exit For
            End If
        Next Bucle
    
    'End If
End Sub
Public Sub Etiqueta_General2l(sCAd As String)
    Dim Bucle As Integer
    Dim nNumMax As Integer
    
    If lngt Then
        nNumMax = 3
    Else
        nNumMax = 6
    End If
    
    Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("etiqueta21").Caption = sCAd
    Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("etiqueta21").Visible = True
    
    For Bucle = 1 To nNumMax
        If Reporte2L.Sections("consulta2l_grouping_footer").Controls("Funcion" & Format(Bucle, "0")).Visible = True Then
            Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("etiqueta21").Left = Reporte2L.Sections("consulta2l_grouping_footer").Controls("Funcion" & Format(Bucle, "0")).Left - Reporte2L.Sections("Consulta2l_Grouping_Footer").Controls("etiqueta21").Width - 100
            Exit For
        End If
    Next Bucle
End Sub
Public Sub Informe_Rechazos()
    Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
   
    Set Informacion = AbreRecordset("select tipo_ven,codigo,codi_sub,codi_fam,plu,precio,des_plu1,tara,comentario,codi_ident as agrupador," & Chr(34) & CargaCadena(285) & Chr(34) & " as tipo  from rechazos where tipo_ven=true  order by codi_ident,plu,codigo" & _
    " union select tipo_ven,codigo,codi_sub,codi_fam,plu,precio,des_plu1,tara,comentario,codi_ident as agrupador," & Chr(34) & CargaCadena(286) & Chr(34) & " as tipo  from rechazos where tipo_ven=false  order by agrupador,plu,codigo")
    
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(302)
    AsignaEtiqueta2L CargaCadena(69), "codigo", 4
    AsignaEtiqueta2L CargaCadena(10), "codi_sub", 6
    AsignaEtiqueta2L CargaCadena(67), "codi_fam", 4
    AsignaEtiqueta2L CargaCadena(68), "plu", 3
    AsignaEtiqueta2L CargaCadena(71), "precio", 6
    AsignaEtiqueta2L CargaCadena(79), "tipo", 4
    AsignaEtiqueta2L CargaCadena(70), "des_plu1", 10
    AsignaEtiqueta2L CargaCadena(74), "tara", 4
    AsignaEtiqueta2L CargaCadena(48), "comentario", 18
    AsignaFuncion2l 4, 4, "codigo", False
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(689)
    Etiqueta_General2l CargaCadena(1188)
    Reporte2L.Refresh
    MuestraInforme2L
    
End Sub
Public Sub Informe_Articulos_Seccion(ByVal Sec As Integer)
    Dim Informacion As New ADODB.Recordset
    Dim Tipo_Orden As String
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    FrmOrdenar.Show vbModal
    Select Case TipoOrden
        Case 0
            Tipo_Orden = "codi_ident,codigo"
        Case 1
            Tipo_Orden = "codi_ident,plu"
        Case 2
            Tipo_Orden = "codi_ident,des_plu1"
    End Select
    If Sec <> 0 Then
        If UsaEuro Then
            Informacion.Open "select codigo,plu,des_plu1,euros,codi_ident,tara,pref,caducidad,codi_pes,tipo_iva from articulo where codi_ident=" & Sec & " and borrado=false order by " & Tipo_Orden, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
        Else
            Informacion.Open "select codigo,plu,des_plu1,precio,codi_ident,tara,pref,caducidad,codi_pes,tipo_iva from articulo where codi_ident=" & Sec & " and borrado=false order by " & Tipo_Orden, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
        End If
    Else
        If UsaEuro Then
            Informacion.Open "select codigo,plu,des_plu1,euros,codi_ident,tara,pref,caducidad,codi_pes,tipo_iva from articulo where  borrado=false order by " & Tipo_Orden, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
        Else
            Informacion.Open "select codigo,plu,des_plu1,precio,codi_ident,tara,pref,caducidad,codi_pes,tipo_iva from articulo where borrado=false order by " & Tipo_Orden, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
        End If
    End If
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta
    Entorno.rsConsulta.Requery
    Load Reporte
    AsignaEtiqueta CargaCadena(69), "codigo", 4
    AsignaEtiqueta CargaCadena(68), "plu", 3
    AsignaEtiqueta CargaCadena(70), "des_plu1", 13
    If UsaEuro Then
        AsignaEtiqueta CargaCadena(71), "euros", 6
    Else
        AsignaEtiqueta CargaCadena(71), "precio", 6
    End If
    AsignaEtiqueta CargaCadena(1185), "codi_ident", 6
    AsignaEtiqueta CargaCadena(74), "tara", 4
    AsignaEtiqueta CargaCadena(73), "pref", 6
    AsignaEtiqueta CargaCadena(72), "caducidad", 6
    AsignaEtiqueta CargaCadena(79), "codi_pes", 4
    AsignaEtiqueta CargaCadena(553), "tipo_iva", 4
    Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(894)
    Reporte.Refresh
    MuestraInforme
End Sub
Public Sub Informe_CambiosTemporales()
    Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Set Informacion = AbreRecordset("select date,hour,nume,numbal,secc,code,plu,vendor,price,amount,codi_pes from tickets where codi_pes=" & Chr(34) & "0" & Chr(34))
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta
    Entorno.rsConsulta.Requery
    Load Reporte
    AsignaEtiqueta CargaCadena(434), "date", 8
    AsignaEtiqueta CargaCadena(435), "hour", 8
    AsignaEtiqueta CargaCadena(868), "nume", 5
    AsignaEtiqueta CargaCadena(303), "numbal", 5
    AsignaEtiqueta CargaCadena(1185), "secc", 5
    AsignaEtiqueta CargaCadena(69), "code", 6
    AsignaEtiqueta CargaCadena(68), "plu", 5
    AsignaEtiqueta CargaCadena(13), "vendor", 6
    AsignaEtiqueta CargaCadena(71), "price", 6
    AsignaEtiqueta CargaCadena(132), "amount", 5
    Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1186)
    Reporte.Refresh
    MuestraInforme
End Sub
Public Sub Informe_Articulos_Balanzas()
    Dim Informacion As New ADODB.Recordset
    Dim Bucle As Integer
    
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select * from consart order by sec,cod", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta
    Entorno.rsConsulta.Requery
    Load Reporte
    
    AsignaEtiqueta CargaCadena(69), "Cod", 4
    AsignaEtiqueta CargaCadena(68), "Plu", 3
    AsignaEtiqueta CargaCadena(54), "Sec", 3
    AsignaEtiqueta CargaCadena(66), "Sub", 3
    AsignaEtiqueta CargaCadena(67), "Fam", 4
    AsignaEtiqueta CargaCadena(71), "Prc", 4
    AsignaEtiqueta CargaCadena(72), "Cad", 4
    AsignaEtiqueta CargaCadena(74), "Tara", 4
    AsignaEtiqueta CargaCadena(73), "Pref", 4
    AsignaEtiqueta CargaCadena(553), "IVA", 3
    AsignaEtiqueta CargaCadena(79), "WGH", 4
    AsignaEtiqueta CargaCadena(623), "ETQ", 3
    AsignaEtiqueta CargaCadena(23), "EAN", 8
    AsignaEtiqueta CargaCadena(70), "DESC0", 8
    Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1187)
    Reporte.Refresh
    MuestraInforme
End Sub
Public Sub Informe_Tiendas()
    '*******************************************
    ' Genera un informe con los datos de las
    ' tiendas existentes, que previamente
    ' se han introducido en las base
    ' por la función "Lista_Las_Tiendas"
    '*******************************************
    Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select * from tiendas order by numero", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta
    Entorno.rsConsulta.Requery
    Load Reporte
    AsignaEtiqueta CargaCadena(1138), "numero", 5
    AsignaEtiqueta CargaCadena(552), "nombre", 12
    AsignaEtiqueta CargaCadena(1130), "provincia", 7
    AsignaEtiqueta CargaCadena(1128), "poblacion", 7
    AsignaEtiqueta CargaCadena(1129), "CP", 6
    AsignaEtiqueta CargaCadena(1127), "direccion", 14
    AsignaEtiqueta CargaCadena(1131), "telefono", 8
    Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1125)
    Reporte.Refresh
    MuestraInforme
End Sub
Public Sub Informe_Pedido(Ruta As String)
    '**************************************
    ' Genera un informe de pedidos a partir
    ' de los datos del fichero "XXped.txt"
    ' generado por Hydra+
    '**************************************
    Dim Reg As New ADODB.Recordset
    Dim Arch As Integer
    Dim Buf As String
    Dim IsOpen As Boolean
    If Dir(Ruta) = "" Then
        MsgBox "No hay ficheros de pedido generados. Solicite pedido a las balanzas", vbCritical
        End
    Else
        Arch = FreeFile()
        On Error Resume Next
        Open Ruta For Input As #Arch
        If Err.Number = 0 Then
            Reg.Close
            On Error GoTo 0
            Reg.Open "select * from pedido", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False", , adLockOptimistic
            
            Do Until EOF(Arch)
                Line Input #Arch, Buf
                If Len(Buf) >= 16 Then
                    Reg.AddNew
                    Reg.Fields("codigo") = Val(Left(Buf, 6))
                    If Mid(Buf, 7, 1) = "W" Then
                        Reg.Fields("tipo") = "Peso"
                        Reg.Fields("pedido") = Val(Mid(Buf, 8, 8)) / 1000
                    Else
                        Reg.Fields("tipo") = "Unidades"
                        Reg.Fields("pedido") = Val(Mid(Buf, 8, 8))
                    End If
                    Reg.Fields("descriptivo") = Mid(Buf, 16)
                    Reg.Update
                End If
            Loop
            Close #Arch
            CreaTablaPropia "informe", Reg
            IsOpen = False
            If Not Reg.EOF And Reg.BOF Then Reg.MoveFirst
            
            '
            Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
            Entorno.Consulta
            Entorno.rsConsulta.Requery
            Load Reporte
            AsignaEtiqueta CargaCadena(69), "Codigo", 6
            AsignaEtiqueta CargaCadena(70), "Descriptivo", 14
            AsignaEtiqueta CargaCadena(55), "Tipo", 8
            AsignaEtiqueta CargaCadena(108), "Pedido", 8
            Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1091)
            Reporte.Refresh
            MuestraInforme
        Else
            MsgBox "Error al acceder al fichero de pedido. Está bloqueado", vbCritical
            End
        End If
        If IsOpen Then
            Reg.Close
        
        End If
    End If
End Sub
Public Sub Informe_Inventario(Ruta As String)
    '*********************************************
    ' Genera un informe de inventario a partir
    ' de los datos del fichero "XXinv.txt"
    ' generado por Hydra+
    '*********************************************
    Dim Reg As New ADODB.Recordset
    Dim Arch As Integer
    Dim Buf As String
    Dim IsOpen As Boolean
    If Dir(Ruta) = "" Then
        MsgBox "No hay ficheros de pedido generados. Solicite pedido a las balanzas", vbCritical
        End
    Else
        Arch = FreeFile()
        On Error Resume Next
        Open Ruta For Input As #Arch
        If Err.Number = 0 Then
            Reg.Close
            On Error GoTo 0
            Reg.Open "select * from inventario", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False", , adLockOptimistic
            
            Do Until EOF(Arch)
                Line Input #Arch, Buf
                If Len(Buf) >= 16 Then
                    Reg.AddNew
                    Reg.Fields("codigo") = Val(Left(Buf, 6))
                    If Mid(Buf, 7, 1) = "W" Then
                        Reg.Fields("tipo") = "Peso"
                        Reg.Fields("inventario") = Val(Mid(Buf, 8, 8)) / 1000
                    Else
                        Reg.Fields("tipo") = "Unidades"
                        Reg.Fields("inventario") = Val(Mid(Buf, 8, 8))
                    End If
                    Reg.Fields("descriptivo") = Mid(Buf, 16)
                    Reg.Update
                End If
            Loop
            Close #Arch
            CreaTablaPropia "informe", Reg
            IsOpen = False
            Reg.Close
            Reg.Open "select codigo,descriptivo,tipo,inventario from inventario", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb"
            
            '
            Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
            Entorno.Consulta
            Entorno.rsConsulta.Requery
            Load Reporte
            AsignaEtiqueta CargaCadena(69), "Codigo", 6
            AsignaEtiqueta CargaCadena(70), "Descriptivo", 14
            AsignaEtiqueta CargaCadena(55), "Tipo", 8
            AsignaEtiqueta CargaCadena(109), "Inventario", 8
            Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1092)
            Reporte.Refresh
            MuestraInforme
        Else
            MsgBox "Error al acceder al fichero de pedido. Está bloqueado", vbCritical
            End
        End If
        If IsOpen Then
            Reg.Close
        End If
    End If
End Sub
Public Sub Informe_PedidoInventario(Ruta As String)
    '*********************************************
    ' Genera un informe de inventario a partir
    ' de los datos del fichero "XXinv.txt"
    ' generado por Hydra+
    '*************************
    Dim Reg As New ADODB.Recordset
    Dim Arch As Integer
    Dim Buf As String
    Dim IsOpen As Boolean
    If Dir(Ruta) = "" Then
        MsgBox "No hay ficheros de pedido/inventario generados. Solicite pedido/inventario a las balanzas", vbCritical
        End
    Else
        Arch = FreeFile()
        On Error Resume Next
        Open Ruta For Input As #Arch
        If Err.Number = 0 Then
            Reg.Close
            On Error GoTo 0
            Reg.Open "select * from pedinv order by mostrador,plu", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False", , adLockOptimistic
            With Reg
                If Not .EOF Then
                    Do Until .EOF
                        If Not .EOF Then .Delete
                        If Not .EOF Then .MoveNext
                    Loop
                    Reg.Close
                    Reg.Open "select * from pedinv order by mostrador,plu", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False", , adLockOptimistic
                End If
            End With
            Do Until EOF(Arch)
                Line Input #Arch, Buf
                If Len(Buf) >= 16 Then
                    Reg.AddNew
                    Reg.Fields("Mostrador") = Val(Left(Buf, 3))
                    Reg.Fields("plu") = Val(Mid(Buf, 4, 4))
                    Reg.Fields("codigo") = Val(Mid(Buf, 8, 6))
                    Reg.Fields("pedido") = Val(Mid(Buf, 14, 8))
                    Reg.Fields("inventario") = Val(Mid(Buf, 22, 8))
                    Select Case Right(Buf, 1)
                        Case "1"
                            Reg.Fields("Tipo") = "Pesado"
                            Reg.Fields("inventario") = Round(Val(Mid(Buf, 14, 8)) / 1000, 3)
                            Reg.Fields("pedido") = Round(Val(Mid(Buf, 22, 8)) / 1000, 3)
                        Case "0"
                            Reg.Fields("Tipo") = "No Pesado"
                            Reg.Fields("inventario") = Val(Mid(Buf, 14, 8))
                            Reg.Fields("pedido") = Val(Mid(Buf, 22, 8))
                        Case Else
                            Reg.Fields("Tipo") = "Sin Datos"
                            Reg.Fields("inventario") = Val(Mid(Buf, 14, 8))
                            Reg.Fields("pedido") = Val(Mid(Buf, 22, 8))
                    End Select
                    Reg.Update
                End If
            Loop
            Close #Arch
            CreaTablaPropia "informe", Reg
            IsOpen = False
            Reg.Close
            Reg.Open "select codigo,descriptivo,tipo,inventario from inventario", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb"
            Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
            Entorno.Consulta
            Entorno.rsConsulta.Requery
            Load Reporte
            AsignaEtiqueta "Mostrador", "mostrador", 6
            AsignaEtiqueta "Plu", "plu", 6
            AsignaEtiqueta "Código", "codigo", 8
            AsignaEtiqueta "Pedido", "pedido", 8
            AsignaEtiqueta "Inventario", "Inventario", 8
            AsignaEtiqueta "Tipo", "tipo", 6
            Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = "Listado de Pedido e Inventario"
            Reporte.Refresh
            MuestraInforme
        Else
            MsgBox "Error al acceder al fichero de pedido. Está bloqueado", vbCritical
            End
        End If
        If IsOpen Then
            Reg.Close
        End If
    End If
End Sub
Public Sub Informe_Equipos()
    '***********************************************
    ' Genera un informe de las máuinas existentes
    ' ordenadas por mostrador
    '***********************************************
    Dim Informacion As New ADODB.Recordset
    Dim strSeguridad As String
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = " and codi_ident=" & UsuarioActual.Asociado & " "
    End If
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Set Informacion = AbreRecordset("select numero_eqp,descripcio,codi_ident as agrupador," & Chr(34) & "TouchScale" & Chr(34) & ",secc_maqui,prog_eqp from equipos where modelo<100 " & strSeguridad & " and borrado=false" & _
    " union select numero_eqp,descripcio,codi_ident as agrupador," & Chr(34) & "V-12" & Chr(34) & ",secc_maqui," & Chr(34) & "N/A" & Chr(34) & " from equipos where modelo=100 " & strSeguridad & " and borrado=false " & _
    " union select numero_eqp,descripcio,codi_ident as agrupador," & Chr(34) & "Marte IC-ILC / V-8" & Chr(34) & ",secc_maqui," & Chr(34) & "N/A" & Chr(34) & " from equipos where modelo=200 " & strSeguridad & " and borrado=false " & _
    " union select numero_eqp,descripcio,codi_ident as agrupador," & Chr(34) & "IV4 LCD / IV4" & Chr(34) & ",secc_maqui," & Chr(34) & "N/A" & Chr(34) & " from equipos where modelo=400 " & strSeguridad & " and borrado=false")
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(302)
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(579)
    AsignaEtiqueta2L CargaCadena(323), Informacion.Fields(0).Name, 12
    AsignaEtiqueta2L CargaCadena(48), Informacion.Fields(1).Name, 12
    AsignaEtiqueta2L CargaCadena(55), Informacion.Fields(3).Name, 12
    AsignaEtiqueta2L CargaCadena(54), Informacion.Fields(4).Name, 12
    AsignaEtiqueta2L CargaCadena(569), Informacion.Fields(5).Name, 12
    Reporte2L.Refresh
    MuestraInforme2L
End Sub
Public Sub Informe_Vendedores()
    Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select ident_vend,nombre,tec_vend,codi_ident as agrupador from vendedor where borrado=false", _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(1185)
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(13)
    AsignaEtiqueta2L CargaCadena(69), Informacion.Fields(0).Name, 12
    AsignaEtiqueta2L CargaCadena(552), Informacion.Fields(1).Name, 12
    AsignaEtiqueta2L CargaCadena(325), Informacion.Fields(2).Name, 12
    Reporte2L.Refresh
    MuestraInforme2L
End Sub
Public Sub informe_control_vendedores()
Dim sSQL As String
Dim sCon As String
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    sSQL = "SELECT ident_vend,secc_maqui,imporCod,operCod,imporDir,operDir,imporCan,operCan,imporNeg,operNeg,codi_ident AS agrupador " & _
           "FROM gtInfVen " & _
           "ORDER BY ident_vend"
    sCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";" & _
           "Persist Security Info=False"
    Informacion.Open sSQL, sCon
    CreaTablaPropia "informe", Informacion
    sCon = "Provider=MSDataShape.1;" & _
           "Persist Security Info=False;" & _
           "Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";" & _
           "Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Conexion.Open sCon
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(1185)
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(13)
    AsignaEtiqueta2L CargaCadena(69), Informacion.Fields(0).Name, 5
    AsignaEtiqueta2L CargaCadena(1277), Informacion.Fields(2).Name, 6
    AsignaEtiqueta2L CargaCadena(1278), Informacion.Fields(3).Name, 7
    AsignaEtiqueta2L CargaCadena(1279), Informacion.Fields(4).Name, 6
    AsignaEtiqueta2L CargaCadena(1280), Informacion.Fields(5).Name, 7
    AsignaEtiqueta2L CargaCadena(122), Informacion.Fields(6).Name, 6
    AsignaEtiqueta2L CargaCadena(121), Informacion.Fields(7).Name, 7
    AsignaEtiqueta2L CargaCadena(150), Informacion.Fields(8).Name, 6
    AsignaEtiqueta2L CargaCadena(120), Informacion.Fields(9).Name, 7
    Reporte2L.Refresh
    MuestraInforme2L
End Sub
Public Sub Informe_Teclas(MyNum As String)
    Dim Informacion As New ADODB.Recordset
    Dim MyCad As String
    Dim NumBase As Integer
    If Left(MyNum, 1) = "0" Then
        MyCad = "numero_eqp="
    Else
        MyCad = "isnull(numero_eqp) and codi_ident="
    End If
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select tabla,codi_tec,codigo,plu,desc_plu from teclas where borrado=false and " & MyCad & Val(Mid(MyNum, 2)), _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Load Reporte
    AsignaEtiqueta CargaCadena(223), Informacion.Fields(0).Name, 8
    AsignaEtiqueta CargaCadena(325), Informacion.Fields(1).Name, 8
    AsignaEtiqueta CargaCadena(69), Informacion.Fields(2).Name, 8
    AsignaEtiqueta "plu", Informacion.Fields(3).Name, 8
    AsignaEtiqueta CargaCadena(539), Informacion.Fields(4).Name, 12
    If Left(MyNum, 1) = "0" Then
        Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = _
        CargaCadena(321) & "-" & CargaCadena(303) & " " & Mid(MyNum, 2)
    Else
        Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = _
        CargaCadena(321) & "-" & CargaCadena(302) & " " & Mid(MyNum, 2)
    End If
    Reporte.Refresh
    MuestraInforme
End Sub

Public Sub Informe_Mostradores()
    Dim Bucle As Integer
    Dim Informacion As New ADODB.Recordset
    Dim strSeguridad As String
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = " and codi_ident=" & UsuarioActual.Asociado & " "
    End If
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "select codi_ident,secc_maqui,descripcio," & Chr(34) & "TouchScale" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "GA" & Chr(34) & "union " & _
    "select codi_ident,secc_maqui,descripcio," & Chr(34) & "SC10" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "SC10" & Chr(34) & "union " & _
    "select codi_ident,secc_maqui,descripcio," & Chr(34) & "V-8" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "ECON" & Chr(34) & "union " & _
    "select codi_ident,secc_maqui,descripcio," & Chr(34) & "Marte IC-ILC" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "ECOL" & Chr(34) & "union " & _
    "select codi_ident,secc_maqui,descripcio," & Chr(34) & "IV4 LCD / IV4" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "ECOP" & Chr(34) & "union " & _
    "select codi_ident,secc_maqui,descripcio," & Chr(34) & "Euroscale+SC10" & Chr(34) & " as agrupador from seccion where borrado=false " & strSeguridad & " and enviardatos=" & Chr(34) & "TODOS" & Chr(34), _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(55)
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(578)
    AsignaEtiqueta2L CargaCadena(1185), Informacion.Fields(0).Name, 12
    AsignaEtiqueta2L CargaCadena(54), Informacion.Fields(1).Name, 12
    AsignaEtiqueta2L CargaCadena(48), Informacion.Fields(2).Name, 12
    Reporte2L.Refresh
    MuestraInforme2L
End Sub
'\\\\\\\\\\\
'///////////
Public Sub Informe_Totales(ByVal NombreInforme As String, sCadena As String, Etiquetas() As String, Largo() As Integer, nEtiquetas As Integer, sTots() As String, nTots As Integer)
Dim Bucle As Integer
Dim Informacion As New ADODB.Recordset
    
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open sCadena, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
    CreaTablaPropia "Informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    If Not lngt Then
        Entorno.Consulta2L_Grouping
        Entorno.rsConsulta2L_Grouping.Requery
    End If
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(47)
    If Filtro0(FrmTotales.CmbFiltro(0).ListIndex) = 12 Then
        Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = NombreInforme & " Dscart."
    Else
        Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = NombreInforme
    End If
    For Bucle = 1 To nEtiquetas - 1
        'If lngt Then
        '    'If Bucle <> 3 Then
        '    '    AsignaEtiqueta2L Etiquetas(Bucle), Informacion.Fields(Bucle - 1).Name, Largo(Bucle)
        '    'Else
        '        AsignaEtiqueta2L Etiquetas(Bucle), FrmTotales.Lista.ColumnHeaders(Bucle), Largo(Bucle)
        '
        '    'End If
        'Else
        AsignaEtiqueta2L Etiquetas(Bucle), Informacion.Fields(Bucle).Name, Largo(Bucle)
        'End If
    Next Bucle
    For Bucle = 1 To nTots
        AsignaFuncion2l Bucle + (6 - nTots), 0, sTots(Bucle), False
        AsignaFuncion2l Bucle + (6 - nTots), 0, sTots(Bucle), True
    Next Bucle
    Etiqueta_General2l CargaCadena(153)
    If Filtro0(FrmTotales.CmbFiltro(0).ListIndex) = 12 Then
        Etiqueta_Total2l CargaCadena(114) & " Descart."
    Else
        Etiqueta_Total2l CargaCadena(114)
    End If
    'Reporte2L.Refresh
    MuestraInforme2L
End Sub
'\\\\\\\\\\\
'///////////
Public Function informe_clientes()
Dim cnt As Long
Dim MyCad As String
Dim NumBase As Integer
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    Informacion.Open "SELECT * FROM Cliente WHERE borrado=false ", _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    CreaTablaPropia "informe", Informacion
    Entorno.Conexion.Open "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Load Reporte
    Reporte.Sections("Sección4").Controls("Etiqueta18").Caption = CargaCadena(1056)
    AsignaEtiqueta CargaCadena(69), Informacion.Fields(0).Name, 8
    AsignaEtiqueta CargaCadena(552), Informacion.Fields(1).Name, 8
    For cnt = 1 To 4
        AsignaEtiqueta CargaCadena(539) & " " & CStr(cnt), Informacion.Fields(cnt + 1).Name, 10
    Next cnt
    Reporte.Refresh
    MuestraInforme
End Function
'\\\\\\\\\\\
'///////////
Public Function informe_descuentos()
Dim sSQL As String
Dim sCon As String
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    sSQL = "SELECT seccion AS agrupador,fecha,nume,vended,importeBruto,importeNeto " & _
           "FROM descuentos " & _
           "ORDER BY nume"
    sCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";" & _
           "Persist Security Info=False"
    Informacion.Open sSQL, sCon
    CreaTablaPropia "informe", Informacion
    sCon = "Provider=MSDataShape.1;" & _
           "Persist Security Info=False;" & _
           "Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";" & _
           "Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Conexion.Open sCon
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(1185)
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(1276)
    AsignaEtiqueta2L CargaCadena(434), Informacion.Fields(1).Name, 11 'fecha
    AsignaEtiqueta2L CargaCadena(123), Informacion.Fields(2).Name, 11 'numero ticket
    AsignaEtiqueta2L CargaCadena(13), Informacion.Fields(3).Name, 11 'Vendedores
    AsignaEtiqueta2L CargaCadena(1274), Informacion.Fields(4).Name, 11 'Importe bruto
    AsignaEtiqueta2L CargaCadena(1275), Informacion.Fields(5).Name, 11 'Importe neto
    AsignaFuncion2l 5, 0, "ImporteBruto", False
    AsignaFuncion2l 5, 0, "ImporteBruto", True
    AsignaFuncion2l 6, 0, "ImporteNeto", False
    AsignaFuncion2l 6, 0, "ImporteNeto", True
    Etiqueta_General2l CargaCadena(153)
    Etiqueta_Total2l CargaCadena(114)
    Reporte2L.Refresh
    MuestraInforme2L
End Function
'\\\\\\\\\\\
'///////////
Public Function informe_horario()
Dim sSQL As String
Dim sCon As String
Dim Informacion As New ADODB.Recordset
    LongitudTotal = 0
    If Entorno.Conexion.State <> adStateClosed Then
        Entorno.Conexion.Close
    End If
    sSQL = "SELECT FECHA AS AGRUPADOR, HORA, " & _
                   "OPERACIONES,OPERACIONES_CAN,OPERACIONES_NEG, " & _
                   "IMPORTE, IMPORTE_CAN, IMPORTE_NEG, " & _
                   "PESO, PESO_CAN, PESO_NEG, " & _
                   "UNIDADES, UNIDADES_CAN, UNIDADES_NEG " & _
           "FROM gthora " & _
           "ORDER BY HORA"
    sCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";" & _
           "Persist Security Info=False"
    Informacion.Open sSQL, sCon
    CreaTablaPropia "informe", Informacion
    sCon = "Provider=MSDataShape.1;" & _
           "Persist Security Info=False;" & _
           "Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";" & _
           "Data Provider=MICROSOFT.JET.OLEDB.4.0"
    Entorno.Conexion.Open sCon
    Entorno.Consulta2L_Grouping
    Entorno.rsConsulta2L_Grouping.Requery
    Load Reporte2L
    Reporte2L.Sections("consulta2l_grouping_header").Controls("Etiqueta1").Caption = CargaCadena(434) 'FECHA
    Reporte2L.Sections("ReportHeader").Controls("Etiqueta19").Caption = CargaCadena(140) 'TOTAL HORARIO
    AsignaEtiqueta2L CargaCadena(1441), Informacion.Fields(1).Name, 4    'HORA
    AsignaEtiqueta2L CargaCadena(1442), Informacion.Fields(2).Name, 4 'OPERACIONES
    AsignaEtiqueta2L CargaCadena(1443), Informacion.Fields(3).Name, 4  'OPERACIONES_CAN
    AsignaEtiqueta2L CargaCadena(1444), Informacion.Fields(4).Name, 4 'OPERACIONES_NEG
    AsignaEtiqueta2L CargaCadena(1445), Informacion.Fields(5).Name, 4 'IMPORTE
    AsignaEtiqueta2L CargaCadena(1446), Informacion.Fields(6).Name, 4 'IMPORTE_CAN
    AsignaEtiqueta2L CargaCadena(1447), Informacion.Fields(7).Name, 4 'IMPORTE_NEG
    AsignaEtiqueta2L CargaCadena(1448), Informacion.Fields(8).Name, 4 'PESO
    AsignaEtiqueta2L CargaCadena(1449), Informacion.Fields(9).Name, 4 'PESO_CAN
    AsignaEtiqueta2L CargaCadena(1450), Informacion.Fields(10).Name, 4 'PESO_NEG
    AsignaEtiqueta2L CargaCadena(1451), Informacion.Fields(11).Name, 4 'UNIDADES
    AsignaEtiqueta2L CargaCadena(1452), Informacion.Fields(12).Name, 4 'UNIDADES_CAN
    AsignaEtiqueta2L CargaCadena(1453), Informacion.Fields(13).Name, 4 'UNIDADES_NEG
    'AsignaFuncion2l 1, 0, "OPERACIONES", True
    'AsignaFuncion2l 2, 0, "OPERACIONES_CAN", True
    'AsignaFuncion2l 3, 0, "OPERACIONES_NEG", True
    'AsignaFuncion2l 4, 0, "IMPORTE", True
    'AsignaFuncion2l 5, 0, "IMPORTE_CAN", True
    'AsignaFuncion2l 6, 0, "IMPORTE_NEG", True
    'AsignaFuncion2l 7, 0, "PESO", True
    'AsignaFuncion2l 8, 0, "PESO_CAN", True
    'AsignaFuncion2l 9, 0, "PESO_NEG", True
    'AsignaFuncion2l 10, 0, "UNIDADES", True
    'AsignaFuncion2l 11, 0, "UNIDADES_CAN", True
    'AsignaFuncion2l 12, 0, "UNIDADES_NEG", True
    'Etiqueta_General2l CargaCadena(153)
    'Etiqueta_Total2l CargaCadena(114)
    Reporte2L.Refresh
    MuestraInforme2L
End Function


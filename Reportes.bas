Attribute VB_Name = "Reportes"
Option Explicit
Public Rep_Orden As Integer
Public Sub Reportes_Mostradores()
    Dim P As ReporteMostradores
    Load ReporteMostradores
    If DataEnvironment1.rsListadoMostradores_Grouping.State = 1 Then
        DataEnvironment1.rsListadoMostradores_Grouping.Close
    End If
    
    If ReporteMostradores.Visible = True Then Exit Sub
    If DataEnvironment1.Connection1.State = 0 Then DataEnvironment1.Connection1.Open
    If DataEnvironment1.Connection2.State = 0 Then DataEnvironment1.Connection2.Open
    ReporteMostradores.Sections("Sección4").Controls("etiqueta2").Caption = LoadResString(id + 578)
    ReporteMostradores.Sections("Sección6").Controls("etiqueta1").Caption = LoadResString(id + 555)
    ReporteMostradores.Sections("Sección6").Controls("etiqueta3").Caption = LoadResString(id + 554)
    ReporteMostradores.Sections("Sección6").Controls("etiqueta4").Caption = LoadResString(id + 54)
    ReporteMostradores.Sections("Sección6").Controls("etiqueta5").Caption = LoadResString(id + 57)
    '**************
    If DataEnvironment1.rsListadoMostradores_Grouping.State = 0 Then
        DataEnvironment1.rsListadoMostradores_Grouping.Open
    End If
    If DataEnvironment1.rsListadoMostradores_Grouping.State <> 0 Then
        DataEnvironment1.rsListadoMostradores_Grouping.Requery
    End If
    ReporteMostradores.Sections("sección4").Controls("etqempresa").Caption = Empresa
    ReporteMostradores.Caption = LoadResString(id + 9) '"Mostradores"

    ReporteMostradores.Visible = True
    Unload ReporteMostradores
    'On Error Resume Next

    Do_Events
End Sub

Public Sub Reportes_Equipos()
    If ReporteBalanzas.Visible = True Then Exit Sub
    
    If DataEnvironment1.Connection1.State = 0 Then DataEnvironment1.Connection1.Open
    If DataEnvironment1.Connection2.State = 0 Then DataEnvironment1.Connection2.Open
    
    ReporteBalanzas.Sections("Sección4").Controls("etiqueta2").Caption = LoadResString(id + 579)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta1").Caption = LoadResString(id + 554)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta3").Caption = LoadResString(id + 57)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta7").Caption = LoadResString(id + 54)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta6").Caption = LoadResString(id + 567)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta5").Caption = LoadResString(id + 569)
    ReporteBalanzas.Sections("Sección6").Controls("etiqueta10").Caption = LoadResString(id + 302)
    '**************
    If DataEnvironment1.rsListadoBalanzas_Grouping.State <> 0 Then
        DataEnvironment1.rsListadoBalanzas_Grouping.Requery
    End If
    ReporteBalanzas.Sections("sección4").Controls("etqempresa").Caption = Empresa
    ReporteBalanzas.Caption = LoadResString(id + 11) '"Equipos"
    ReporteBalanzas.Show
    
End Sub
Public Sub Reportes_Articulos(Optional MySec As Integer)
    
    Dim MiOrd As String
    Dim MiMos As String
    If ReporteArticulos.Visible = True Then Exit Sub

    If DataEnvironment1.Connection1.State = 0 Then DataEnvironment1.Connection1.Open
    If DataEnvironment1.Connection2.State = 0 Then DataEnvironment1.Connection2.Open
    FrmOrdenar.Show
    Do While FrmOrdenar.Visible = True
        Do_Events
    Loop
    Unload FrmOrdenar
    Select Case Rep_Orden
        Case 0
            MiOrd = "codigo"
        Case 1
            MiOrd = "plu"
        Case 2
            MiOrd = "Des_plu1"
    End Select
    If MySec = 0 Then
        MiMos = ""
    Else
        MiMos = " and codi_ident=" & MySec
    End If
    
    ReporteArticulos.Sections("sección4").Controls("etiqueta9").Caption = Empresa
    ReporteArticulos.Caption = LoadResString(id + 12)
    If UsaEuro Then
        ReporteArticulos.Sections("sección1").Controls("txtprc").DataField = "EUROS"
    Else
        ReporteArticulos.Sections("sección1").Controls("txtprc").DataField = "PRC"
    End If
    '**********************
    ' idioma
    '**********************
    ReporteArticulos.Sections("sección4").Controls("Etiqueta10").Caption = LoadResString(id + 894)
    If MySec = 0 Then
        ReporteArticulos.Sections("sección4").Controls("Etiqueta16").Caption = LoadResString(id + 895)
    Else
        ReporteArticulos.Sections("sección4").Controls("Etiqueta16").Caption = LoadResString(id + 302) & " : " & MySec
    End If
    ReporteArticulos.Sections("sección2").Controls("Etiqueta1").Caption = LoadResString(id + 89)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta2").Caption = LoadResString(id + 68)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta3").Caption = LoadResString(id + 70)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta4").Caption = LoadResString(id + 71)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta5").Caption = LoadResString(id + 302)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta7").Caption = LoadResString(id + 74)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta6").Caption = LoadResString(id + 79)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta8").Caption = LoadResString(id + 72)
    ReporteArticulos.Sections("sección2").Controls("Etiqueta17").Caption = LoadResString(id + 553)
    '**********************
    ' /idioma
    '**********************
    'If DataEnvironment1.rslistadoarticulos.State <> 0 Then DataEnvironment1.rslistadoarticulos.Close
    'DataEnvironment1.rslistadoarticulos.Open
    DataEnvironment1.Commands("listadoarticulos").CommandText = _
    "SELECT CODI_FAM, CODI_IDENT, CODI_SUB, CODIGO, DES_PLU1, PLU, PRECIO  AS PRC,EUROS, TARA, CADUCIDAD, CODI_PES ,TIPO_IVA FROM Articulo where borrado=false " & MiMos & " ORDER BY codi_ident," & MiOrd
    On Error Resume Next
    DataEnvironment1.rslistadoarticulos.Close
    ReporteArticulos.Show
    
End Sub
Public Sub Reportes_CambioPrecio()
    
    If DataEnvironment1.Connection1.State = 0 Then DataEnvironment1.Connection1.Open
    If DataEnvironment1.Connection2.State = 0 Then DataEnvironment1.Connection2.Open
    If DataReport10.Visible = True Then
        lCogeTiquet = True
        Exit Sub
    End If
    If DataEnvironment1.rsseclog.State <> 1 Then
        DataEnvironment1.rsseclog.Open
        DataEnvironment1.rsseclog.Filter = usersecvenflt
    Else
        DataEnvironment1.rsseclog.Requery
    End If
    DataReport10.Sections("ReportHeader").Controls("Etiqueta9").Caption = Empresa
    DataReport10.Sections("ReportHeader").Controls("Etiqueta10").Caption = LoadResString(id + 24)
    '*************
    ' Idioma
    '*************
    DataReport10.Sections("seclog_header").Controls("etiqueta1").Caption = LoadResString(id + 302)
    DataReport10.Sections("seclog_header").Controls("etiqueta19").Caption = LoadResString(id + 434)
    DataReport10.Sections("seclog_header").Controls("etiqueta11").Caption = LoadResString(id + 70)
    DataReport10.Sections("seclog_header").Controls("etiqueta5").Caption = LoadResString(id + 71)
    DataReport10.Sections("seclog_header").Controls("etiqueta6").Caption = LoadResString(id + 79)
    DataReport10.Sections("seclog_header").Controls("etiqueta8").Caption = ""
    DataReport10.Sections("seclog_header").Controls("etiqueta7").Caption = "User"
    DataReport10.Sections("seclog_footer").Controls("etiqueta12").Caption = LoadResString(id + 24)
    DataReport10.Show
    
End Sub
Public Sub Reportes_Vendedores()
    
    If DataEnvironment1.Connection1.State = 0 Then DataEnvironment1.Connection1.Open
    If DataEnvironment1.Connection2.State = 0 Then DataEnvironment1.Connection2.Open
    If ReporteVendedores.Visible = True Then Exit Sub
    ReporteVendedores.Sections("Sección4").Controls("EtqEmpresa").Caption = Empresa
    ReporteVendedores.Sections("Sección4").Controls("etiqueta2").Caption = LoadResString(id + 13)
    ReporteVendedores.Sections("Sección6").Controls("etiqueta10").Caption = LoadResString(id + 302)
    ReporteVendedores.Sections("Sección6").Controls("etiqueta1").Caption = LoadResString(id + 69)
    ReporteVendedores.Sections("Sección6").Controls("etiqueta3").Caption = LoadResString(id + 552)
    ReporteVendedores.Sections("Sección6").Controls("etiqueta7").Caption = LoadResString(id + 325)
    If DataEnvironment1.rsvendedores_Grouping.State <> 0 Then
        DataEnvironment1.rsvendedores_Grouping.Requery
    End If
    ReporteVendedores.Show
End Sub

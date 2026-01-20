Attribute VB_Name = "odb"
Option Explicit
'''''''''''''''
'''''''''''''''
Public Function copia_dbase(ByRef ayyyymm As String, ByRef aTabla As String, ByRef arstsrc As DAO.Recordset, ByRef apntRini As Long, ByRef apntRfin As Long) As Integer
    Dim dbBCK As DAO.Database
    Dim rstBCK As DAO.Recordset
    Dim rstSYS As DAO.Recordset
    Dim YYYYMM_MDB As String
    Dim sSQL As String
    Dim cntF As Long
    Dim cntr As Long
    On Error GoTo ret
    Dim lVer As Boolean
    Dim sx As String
    Dim nx As Variant
    
    If LCase(aTabla) = "cabecera" Or LCase(aTabla) = "tickets" Then lVer = True
    
    YYYYMM_MDB = var.rutaSYS & "\S\" & ayyyymm & ".mdb"
    If (Dir(YYYYMM_MDB) = "") Then
        odb.crea_base_historica YYYYMM_MDB
    Else
        If lVer Then
            On Error Resume Next
            Set dbBCK = OpenDatabase(YYYYMM_MDB)
            dbBCK.Execute "alter table cabecera alter column factura TEXT(25)"
            dbBCK.Close
            Set dbBCK = OpenDatabase(YYYYMM_MDB)
            dbBCK.Execute "alter table tickets alter column factura TEXT(25)"
            dbBCK.Close
            On Error GoTo 0
        End If
    End If
    Set dbBCK = OpenDatabase(YYYYMM_MDB)
    If aTabla = "cliente" Then
        dbBCK.Execute "delete from cliente"
    End If
    On Error GoTo ret
    
    sSQL = "SELECT *" & _
           " FROM " & aTabla
    Set rstBCK = dbBCK.OpenRecordset(sSQL)
    arstsrc.MoveFirst
    arstsrc.Move apntRini
    For cntr = apntRini To apntRfin
        rstBCK.AddNew
        For cntF = 0 To rstBCK.Fields.Count - 1
            'If lVer Then
            '    If (LCase(rstBCK.Fields(rstBCK.Fields(cntF).Name).Name) = "factura") And (rstBCK.Fields(rstBCK.Fields(cntF).Name).Size < 25) Then
            '
            '    End If
            'End If
            sx = rstBCK.Fields(cntF).Name
            nx = arstsrc.Fields(rstBCK.Fields(cntF).Name).Value
            rstBCK.Fields(rstBCK.Fields(cntF).Name).Value = arstsrc.Fields(rstBCK.Fields(cntF).Name).Value
        Next cntF
        rstBCK.Update
        arstsrc.MoveNext
    Next cntr
    rstBCK.Close
    Set rstBCK = Nothing
ret:
    If (Err.Number <> 0) Then
        'MsgBox aTabla & "-->" & sx & "=" & nx & " --> " & Err.Description
        copia_dbase = 1
    End If
    dbBCK.Close
    
    Err.Clear
    On Error GoTo 0
End Function
''''''''''''
''''''''''''
Public Function crea_base_historica(ByRef ayyyymm As String) As Integer
    Dim dbS As DAO.Database
    Dim rstS As DAO.Recordset
    On Error GoTo ret
    Set dbS = CreateDatabase(ayyyymm, dbLangGeneral, dbVersion40)
    crea_tablas_cabecera dbS
    crea_tablas_ticket dbS
    crea_tabla_totales dbS
    crea_tabla_infor_vendedor dbS
    crea_tabla_horario dbS
    If UsaFichasVacuno Then crea_tabla_vacuno dbS
    crea_tabla_cliente dbS
ret:
    dbS.Close
    If (Err.Number <> 0) Then MsgBox Err.Description, vbCritical
    Err.Clear
    On Error GoTo 0
End Function
''''''''''''
''''''''''''
Private Function crea_tablas_cabecera(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim rst As DAO.Recordset
    Dim sSQL As String
    Dim cnt As Long
    Dim cntr As Long
    Dim sTabla As String
    sTabla = "cabecera"
    Set tdf = adb.CreateTableDef(sTabla)
    With tdf
            
        .Fields.Append .CreateField("NUME", dbText)
        .Fields.Append .CreateField("SECCION", dbText)
        .Fields.Append .CreateField("BALANZA", dbInteger)
        .Fields.Append .CreateField("VENDED", dbText)
        .Fields.Append .CreateField("NUMLIN", dbText)
        .Fields.Append .CreateField("LINCAN", dbText)
        .Fields.Append .CreateField("IMPORTE", dbDouble, 15)
        .Fields.Append .CreateField("HORA", dbText)
        .Fields.Append .CreateField("ONOFF", dbText)
        .Fields.Append .CreateField("STPETIC", dbText)
        .Fields.Append .CreateField("CODCLI", dbText)
        .Fields.Append .CreateField("FECHA", dbText)
        .Fields.Append .CreateField("TYPTIC", dbText)
        .Fields.Append .CreateField("CODBAR", dbText)
        .Fields.Append .CreateField("RESERV1", dbText)
        .Fields.Append .CreateField("RESERV2", dbText)
        .Fields.Append .CreateField("RESERV3", dbText)
        .Fields.Append .CreateField("d_fecha", dbDate)
        .Fields.Append .CreateField("d_hora", dbDate)
        .Fields.Append .CreateField("TIPOPAGO", dbInteger)
        .Fields.Append .CreateField("cantidadEntre", dbDouble, 15)
        .Fields.Append .CreateField("cambio", dbDouble, 15)
        .Fields.Append .CreateField("importe_bruto", dbDouble)
        .Fields.Append .CreateField("Cobrado", dbBoolean, 1)
        .Fields.Append .CreateField("numlote", dbText)
        .Fields.Append .CreateField("abonado", dbBoolean)
        .Fields.Append .CreateField("ticketabono", dbLong)
        .Fields.Append .CreateField("ticketorigen", dbLong)
        .Fields.Append .CreateField("abono", dbBoolean)
        .Fields.Append .CreateField("descuento", dbDouble)
        .Fields.Append .CreateField("base0", dbDouble)
        .Fields.Append .CreateField("por0", dbDouble)
        .Fields.Append .CreateField("imp0", dbDouble)
        .Fields.Append .CreateField("base1", dbDouble)
        .Fields.Append .CreateField("por1", dbDouble)
        .Fields.Append .CreateField("imp1", dbDouble)
        .Fields.Append .CreateField("base2", dbDouble)
        .Fields.Append .CreateField("por2", dbDouble)
        .Fields.Append .CreateField("imp2", dbDouble)
        .Fields.Append .CreateField("base3", dbDouble)
        .Fields.Append .CreateField("por3", dbDouble)
        .Fields.Append .CreateField("imp3", dbDouble)
        .Fields.Append .CreateField("base4", dbDouble)
        .Fields.Append .CreateField("por4", dbDouble)
        .Fields.Append .CreateField("imp4", dbDouble)
        .Fields.Append .CreateField("factura", dbText, 25)

        For cntr = 0 To .Fields.Count - 1
            .Fields(cntr).AllowZeroLength = True
        Next cntr
    
    End With
    adb.TableDefs.Append tdf
    Set tdf = Nothing

End Function
''''''''''''
''''''''''''
Private Function crea_tablas_ticket(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim cnt As Long
    Dim cntr As Long
    Dim sTabla As String
    'For cnt = 0 To 2
        'If cnt = 0 Then
        sTabla = "tickets"
        'If cnt = 1 Then sTabla = "ticketsSC10"
        'If cnt = 2 Then sTabla = "ticketsTQ"
        Set tdf = adb.CreateTableDef(sTabla)
        With tdf
            .Fields.Append .CreateField("DATE", dbText)
            .Fields.Append .CreateField("HOUR", dbText)
            .Fields.Append .CreateField("NUMBAL", dbInteger)
            .Fields.Append .CreateField("VENDOR", dbText)
            .Fields.Append .CreateField("NUME", dbText)
            .Fields.Append .CreateField("NUMLIN", dbText)
            .Fields.Append .CreateField("SECC", dbText)
            .Fields.Append .CreateField("PLU", dbText)
            .Fields.Append .CreateField("CODE", dbText)
            .Fields.Append .CreateField("ITEM", dbText)
            .Fields.Append .CreateField("VAR", dbText)
            .Fields.Append .CreateField("PRICE", dbDouble)
            .Fields.Append .CreateField("AMOUNT", dbDouble)
            .Fields.Append .CreateField("WEIGHT", dbDouble)
            .Fields.Append .CreateField("UNITS", dbDouble)
            .Fields.Append .CreateField("STPETIC", dbText)
            .Fields.Append .CreateField("CODI_PES", dbText)
            .Fields.Append .CreateField("LINCAN", dbText)
            .Fields.Append .CreateField("IVA", dbInteger)
            .Fields.Append .CreateField("TYPTIC", dbText)
            .Fields.Append .CreateField("RESERV1", dbText)
            .Fields.Append .CreateField("RESERV2", dbText)
            .Fields.Append .CreateField("RESERV3", dbText)
            .Fields.Append .CreateField("TIPOPAGO", dbInteger)
            .Fields.Append .CreateField("d_fecha", dbDate)
            .Fields.Append .CreateField("d_hora", dbDate)
            .Fields.Append .CreateField("ivapercent", dbDouble)
            .Fields.Append .CreateField("numlote", dbText)
            .Fields.Append .CreateField("cliente", dbText, 15)
            .Fields.Append .CreateField("clienteT", dbText, 15)
            .Fields.Append .CreateField("factura", dbText, 25)
            .Fields.Append .CreateField("porcentaje", dbDouble)
            .Fields.Append .CreateField("descuento", dbDouble)
            .Fields.Append .CreateField("DPT", dbText, 4)
            .Fields.Append .CreateField("importe_bruto", dbDouble)
            For cntr = 0 To .Fields.Count - 1
                .Fields(cntr).AllowZeroLength = True
            Next cntr
        End With
        adb.TableDefs.Append tdf
        Set tdf = Nothing
    'Next cnt
End Function
''''''''''''
''''''''''''
Private Function crea_tabla_horario(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim rst As DAO.Recordset
    Dim cnt As Long
    Set tdf = adb.CreateTableDef("gthora")
    With tdf
        .Fields.Append .CreateField("FECHA", dbDate)
        .Fields.Append .CreateField("HORA", dbText)
        .Fields.Append .CreateField("SECCION", dbLong)
        .Fields.Append .CreateField("EQUIPO", dbLong)
        .Fields.Append .CreateField("VENDEDOR", dbLong)
        .Fields.Append .CreateField("TIQUET", dbLong)
        .Fields.Append .CreateField("OPERACIONES", dbLong)
        .Fields.Append .CreateField("OPERACIONES_CAN", dbLong)
        .Fields.Append .CreateField("OPERACIONES_NEG", dbLong)
        .Fields.Append .CreateField("IMPORTE", dbDouble, 9)
        .Fields.Append .CreateField("IMPORTE_CAN", dbDouble, 9)
        .Fields.Append .CreateField("IMPORTE_NEG", dbDouble, 9)
        .Fields.Append .CreateField("PESO", dbDouble)
        .Fields.Append .CreateField("PESO_CAN", dbDouble)
        .Fields.Append .CreateField("PESO_NEG", dbDouble)
        .Fields.Append .CreateField("UNIDADES", dbLong)
        .Fields.Append .CreateField("UNIDADES_CAN", dbLong)
        .Fields.Append .CreateField("UNIDADES_NEG", dbLong)
        adb.TableDefs.Append tdf
    End With
End Function
''''''''''''
''''''''''''
Private Function crea_tabla_totales(ByRef adb As DAO.Database)
    Dim tdfcabt As TableDef
    Dim tdftot As TableDef
    Dim Bucle As Integer
    Dim TextoAdicional As String
    Dim i As Integer
    Set tdftot = adb.CreateTableDef("gtsecs")
    With tdftot
        .Fields.Append .CreateField("Modo", dbInteger)
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("codi_ident", dbInteger)
        .Fields.Append .CreateField("codi_sub", dbInteger)
        .Fields.Append .CreateField("codi_fam", dbInteger)
        .Fields.Append .CreateField("tiquets", dbInteger)
        .Fields.Append .CreateField("operacion", dbDouble)
        .Fields.Append .CreateField("peso", dbDouble)
        .Fields.Append .CreateField("unidades", dbDouble)
        .Fields.Append .CreateField("pesetas", dbDouble)
        .Fields.Append .CreateField("secc_maqui", dbInteger)
        .Fields.Append .CreateField("txt_fam", dbText, 30)
        .Fields.Append .CreateField("txt_sub", dbText, 30)
        On Error Resume Next
        For i = 0 To .Fields.Count - 1
          .Fields(i).AllowZeroLength = True
        Next i
        adb.TableDefs.Append tdftot
    End With
    ''''''''
    Set tdftot = adb.CreateTableDef("gtvend")
    With tdftot
        .Fields.Append .CreateField("Modo", dbInteger)
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("codi_ident", dbInteger)
        .Fields.Append .CreateField("ident_vend", dbInteger)
        .Fields.Append .CreateField("nombre", dbText)
        .Fields.Append .CreateField("supertot", dbDouble) 'totcance+total
        .Fields.Append .CreateField("totcance", dbDouble)
        .Fields.Append .CreateField("total", dbDouble)
        .Fields.Append .CreateField("tiquets", dbDouble)
        .Fields.Append .CreateField("operacion", dbDouble)
        .Fields.Append .CreateField("peso", dbDouble)
        .Fields.Append .CreateField("cancela", dbDouble) 'cancelaciones
        .Fields.Append .CreateField("tiempo", dbDouble) 'tiempo de trabajo en minutos
        .Fields.Append .CreateField("unidad", dbDouble)
        .Fields.Append .CreateField("pesetas", dbDouble)
        .Fields.Append .CreateField("desembolso", dbDouble)
        .Fields.Append .CreateField("reembolso", dbDouble)
        .Fields.Append .CreateField("impor_neg", dbDouble)
        .Fields.Append .CreateField("credito", dbDouble)
        .Fields.Append .CreateField("secc_maqui", dbInteger)
        For i = 0 To .Fields.Count - 1
          .Fields(i).AllowZeroLength = True
        Next i
        adb.TableDefs.Append tdftot
    End With
    ''''''''
    Set tdftot = adb.CreateTableDef("gtarti")
    With tdftot
        .Fields.Append .CreateField("Modo", dbInteger)
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("plu", dbInteger)
        .Fields.Append .CreateField("des_plu1", dbText)
        .Fields.Append .CreateField("codi_ident", dbInteger)
        .Fields.Append .CreateField("codi_sub", dbInteger)
        .Fields.Append .CreateField("codi_fam", dbInteger)
        .Fields.Append .CreateField("codigo", dbDouble)
        .Fields.Append .CreateField("precio", dbDouble)
        .Fields.Append .CreateField("unidades", dbDouble)
        .Fields.Append .CreateField("peso", dbDouble)
        .Fields.Append .CreateField("pesetas", dbDouble)
        .Fields.Append .CreateField("tara", dbInteger)
        .Fields.Append .CreateField("caducidad", dbInteger)
        .Fields.Append .CreateField("operacion", dbDouble)
        .Fields.Append .CreateField("tipo_iva", dbDouble)
        .Fields.Append .CreateField("modificado", dbBoolean)
        .Fields.Append .CreateField("secc_maqui", dbInteger)
        .Fields.Append .CreateField("txt_fam", dbText, 30)
        .Fields.Append .CreateField("txt_sub", dbText, 30)
        For i = 0 To .Fields.Count - 1
          .Fields(i).AllowZeroLength = True
        Next i
        adb.TableDefs.Append tdftot
    End With
End Function
''''''''''''
''''''''''''
Private Function crea_tabla_infor_vendedor(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim rst As DAO.Recordset
    Dim cnt As Long
    Set tdf = adb.CreateTableDef("gtInfVen")
    With tdf
        .Fields.Append .CreateField("ident_vend", dbLong)
        .Fields.Append .CreateField("codi_ident", dbLong)
        .Fields.Append .CreateField("secc_maqui", dbLong)
        .Fields.Append .CreateField("imporCod", dbDouble, 9) 'cas.n->importe codificado (con PLU)
        .Fields.Append .CreateField("operCod", dbDouble, 9) 'cas.n->operaciones codificadas (con PLU)
        .Fields.Append .CreateField("imporDir", dbDouble, 9) 'cas.n->importe directo (con PLU 0)
        .Fields.Append .CreateField("operDir", dbDouble, 9) 'cas.n->operaciones directas (con PLU 0)
        .Fields.Append .CreateField("imporCan", dbDouble, 9) 'cas.n->importe anulado
        .Fields.Append .CreateField("operCan", dbDouble, 9) 'cas.n->operaciones anuladas
        .Fields.Append .CreateField("imporNeg", dbDouble, 9) 'cas.n->importe negativo
        .Fields.Append .CreateField("operNeg", dbDouble, 9) 'cas.n->operaciones negativas
        .Fields.Append .CreateField("modo", dbLong)
        .Fields.Append .CreateField("fecha", dbDate)
        adb.TableDefs.Append tdf
    End With
End Function
''''''''''''
''''''''''''
Private Function crea_tabla_vacuno(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim rst As DAO.Recordset
    Dim cnt As Long
    Set tdf = adb.CreateTableDef("GTVacuno")
    With tdf
        .Fields.Append .CreateField("Ficha", dbDouble, 2)
        .Fields.Append .CreateField("LNumOperaciones", dbDouble, 2)
        .Fields.Append .CreateField("LNumUnidades", dbDouble, 2)
        .Fields.Append .CreateField("LImporte", dbDouble, 2)
        .Fields.Append .CreateField("LPeso", dbDouble, 2)
        .Fields.Append .CreateField("VNumOperaciones", dbDouble, 2)
        .Fields.Append .CreateField("VNumUnidades", dbDouble, 2)
        .Fields.Append .CreateField("VImporte", dbDouble, 2)
        .Fields.Append .CreateField("VPeso", dbDouble, 2)
        .Fields.Append .CreateField("Bloqueo", dbBoolean, 2)
        .Fields.Append .CreateField("Total", dbBoolean, 2)
        .Fields.Append .CreateField("Fecha", dbText, 10)
        For cnt = 0 To .Fields.Count - 1
            .Fields(cnt).AllowZeroLength = True
        Next cnt
        adb.TableDefs.Append tdf
    End With
End Function

''''''''''''

Private Function crea_tabla_cliente(ByRef adb As DAO.Database)
    Dim tdf As TableDef
    Dim rst As DAO.Recordset
    Dim cnt As Long
    Set tdf = adb.CreateTableDef("Cliente")
    With tdf
        .Fields.Append .CreateField("cod_cli", dbDouble)
        .Fields.Append .CreateField("nom_cli", dbText)
        .Fields.Append .CreateField("data1", dbText)
        .Fields.Append .CreateField("data2", dbText)
        .Fields.Append .CreateField("data3", dbText)
        .Fields.Append .CreateField("data4", dbText)
        For cnt = 0 To .Fields.Count - 1
            .Fields(cnt).AllowZeroLength = True
        Next cnt
        adb.TableDefs.Append tdf
    End With
End Function

''''''''''''


Attribute VB_Name = "importar"
Option Explicit
Sub importar_vinculadas()

Dim dbsjet As Database
Dim dbsdbaseIV As Database

Dim miarti As Recordset
Dim miarti_d As Recordset
Dim miseccion As Recordset
Dim miseccion_d As Recordset

Dim misubsec As Recordset
Dim misubsec_d As Recordset
Dim mifam As Recordset
Dim mifam_d As Recordset
Dim micodbar As Recordset
Dim micodbar_d As Recordset
Dim micabley As Recordset
Dim micabley_d As Recordset
Dim mivendedor As Recordset
Dim mivendedor_d As Recordset
Dim miequip As Recordset
Dim miequip_d As Recordset
Dim mitecla As Recordset
Dim mitecla_d As Recordset
Dim mipubli As Recordset
Dim mipubli_d As Recordset
Dim numero As Long
Dim num_reg As Long

'****************************variables articulos************************
Dim tienda(6000) 'As Double
Dim seccion(6000) 'As Double
Dim subsec(6000) 'As Double
Dim familia(6000) 'As Double
Dim codigo(6000) 'As Double
Dim plu(6000) 'As Double
Dim precio(6000) 'As Double
Dim cadu(6000) 'As Double
Dim tara(6000) 'As Double
Dim codi_pes(6000)
Dim des_plu1(6000) 'As String
Dim des_plu2(6000) 'As String
Dim des_plu3(6000) 'As String
Dim des_plu4(6000) 'As String
Dim des_plu5(6000) 'As String
Dim des_plu6(6000) 'As String
Dim des_plu7(6000) 'As String
Dim des_plu8(6000) 'As String
Dim des_plu9(6000) 'As String
Dim des_plu0(6000) 'As String
Dim des_plux(6000) 'As String
Dim art_cb(6000) 'As String
Dim tran_plu(6000) 'As String
Dim tran_texto(6000) 'As String
Dim tran_cb(6000) 'As String
Dim tran_tx1(6000)
'****************************Fin variables articulos************************

'****************************variables mostrador************************
Dim descripcio(999)
Dim tran_sec(999)
Dim tran_var(999)
'****************************Fin variables mostrador************************

'****************************variables subseccion***************************
Dim des_sub(999)
Dim tran_sub(999)
Dim numero_fam(999)
Dim codi_sub(9999)

'****************************Fin variables subseccion************************

'****************************variables familia*******************************
Dim des_fam(9999)
Dim tran_fam(9999)
Dim codi_fam(9999)

'****************************Fin variables familia************************

'****************************variables equipos****************************
Dim des_eqp(9999)
Dim tran_eqp(9999)
Dim numero_eqp(999)
Dim estado_eqp(999)
Dim modelo(999)
Dim prioridad(999)
Dim tipo_eqp(999)
Dim secc_maqui(999)
'****************************Fin variables equipos***************************

'****************************variables codbars****************************
Dim sec_eqp(999)
Dim tran_bar(999)
Dim cb_vnt(999)
Dim cb_sup(999)
Dim cb_mix(999)
Dim cb_ev1(999)
Dim cb_ev2(999)
'****************************Fin variables codbars***************************

'****************************variables cabley********************************
Dim cabecera01(999)
Dim cabecera02(999)
Dim cabecera03(999)
Dim cabecera04(999)
Dim cabecera05(999)
Dim leyenda01(999)
Dim leyenda02(999)
Dim leyenda03(999)
Dim leyenda04(999)
Dim leyenda05(999)

Dim TLETRACAB1(999)
Dim TLETRACAB2(999)
Dim TLETRACAB3(999)
Dim TLETRACAB4(999)
Dim TLETRACAB5(999)

Dim TLETRALEY1(999)
Dim TLETRALEY2(999)
Dim TLETRALEY3(999)
Dim TLETRALEY4(999)
Dim TLETRALEY5(999)
Dim TRAN_CAB(999)
Dim TRAN_LEY(999)

'****************************Fin variables cabley***************************

'****************************variables teclas*******************************
Dim tabla(9999)
Dim tran_tec(9999)
Dim codi_tec(9999)

'****************************Fin variables teclas***************************

'****************************variables vendedores*******************************
Dim ident_vend(999)
Dim codi_vend(999)
Dim tec_vend(999)
Dim nombre_ven(999)
Dim tran_vend(999)
Dim tran_asign(999)
Dim fecha_vend(999)
'****************************Fin variables vendedores***************************

'****************************variables publicidad*******************************
Dim publicidad(999)
Dim oferta01(999)
Dim oferta02(999)
Dim oferta03(999)
Dim oferta04(999)
Dim oferta05(999)
Dim oferta06(999)
Dim oferta07(999)
Dim oferta08(999)
Dim oferta09(999)
Dim oferta10(999)
Dim tipo_publ(999)
Dim tran_pub(999)
Dim tran_ofe(999)
Dim tran_act(999)
'****************************Fin variables publicidad***************************

Dim i As Integer
'Dim prpBucle As Property
'
'Dim NewFld As Field

Set dbsdbaseIV = OpenDatabase(miruta, False, 0, "DBASE IV")
Set dbsjet = OpenDatabase(miruta & "\" & "dbase.mdb")

Set miarti = dbsjet.OpenRecordset("Articulo")
Set miseccion = dbsjet.OpenRecordset("Seccion")
Set misubsec = dbsjet.OpenRecordset("subsec")
Set mifam = dbsjet.OpenRecordset("familia")
Set miequip = dbsjet.OpenRecordset("equipos")
Set micodbar = dbsjet.OpenRecordset("codbar")
Set micabley = dbsjet.OpenRecordset("cabley")
Set mitecla = dbsjet.OpenRecordset("teclas")
Set mivendedor = dbsjet.OpenRecordset("vendedor")
Set mipubli = dbsjet.OpenRecordset("publi")

Set miarti_d = dbsdbaseIV.OpenRecordset("Articulo")
Set miseccion_d = dbsdbaseIV.OpenRecordset("Seccion")
Set misubsec_d = dbsdbaseIV.OpenRecordset("subsec")
Set mifam_d = dbsdbaseIV.OpenRecordset("familia")
Set miequip_d = dbsdbaseIV.OpenRecordset("equipos")
Set micodbar_d = dbsdbaseIV.OpenRecordset("codbar")
Set micabley_d = dbsdbaseIV.OpenRecordset("cabley")
Set mitecla_d = dbsdbaseIV.OpenRecordset("teclas")
Set mivendedor_d = dbsdbaseIV.OpenRecordset("vendedor")
Set mipubli_d = dbsdbaseIV.OpenRecordset("publi")

With miarti_d

      numero = 0
      Do While Not .EOF
      tienda(numero) = !tienda
      seccion(numero) = !codi_ident
      subsec(numero) = !codi_sub
      familia(numero) = !codi_fam
      codigo(numero) = !codigo
      plu(numero) = !plu
      precio(numero) = !precio
      codi_pes(numero) = !codi_pes
      cadu(numero) = !caducidad
      tara(numero) = !tara
      des_plu1(numero) = !des_plu1
      des_plu2(numero) = !des_plu2
      des_plu3(numero) = !des_plu3
      des_plu4(numero) = !des_plu4
      des_plu5(numero) = !des_plu5
      des_plu6(numero) = !des_plu6
      des_plu7(numero) = !des_plu7
      des_plu8(numero) = !des_plu8
      des_plu9(numero) = !des_plu9
      des_plu0(numero) = !des_plu0
      des_plux(numero) = !des_plux
      art_cb(numero) = !art_cb
      tran_texto(numero) = !tran_texto
      tran_plu(numero) = !tran_plu
      tran_cb(numero) = !tran_cb
      tran_tx1(numero) = !tran_tx1
'      .Update
      .MoveNext
      numero = numero + 1
      Loop
      num_reg = numero
      .Close
'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_SECC", dbDouble, 6)
'      .Fields.Append .CreateField("CODI_SUB", dbDouble, 3)
'      .Fields.Append .CreateField("CODI_FAM", dbDouble, 4)
'      .Fields.Append .CreateField("PLU", dbDouble, 4)
'      .Fields.Append .CreateField("CODIGO", dbDouble, 6)
'      .Fields.Append .CreateField("CODI_PES", dbText, 1)
'      .Fields.Append .CreateField("CODI_TEC", dbDouble, 3)
'      .Fields.Append .CreateField("TIPO_IVA", dbDouble, 1)
'      .Fields.Append .CreateField("PRECIO_SIN", dbDouble, 12)
'      .Fields.Append .CreateField("PRECIO", dbDouble, 7)
'      .Fields.Append .CreateField("TIPO_DESC", dbDouble, 1)
'      .Fields.Append .CreateField("CADUCIDAD", dbDouble, 3)
'      .Fields.Append .CreateField("TARA", dbDouble, 5)
'      .Fields.Append .CreateField("TIP_LET1", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET2", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET3", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET4", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET5", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET6", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET7", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET8", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET9", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LET0", dbDouble, 1)
'      .Fields.Append .CreateField("TIP_LETX", dbDouble, 1)
'      .Fields.Append .CreateField("DES_PLU1", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU2", dbText, 0)
'      .Fields.Append .CreateField("DES_PLU3", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU4", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU5", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU6", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU7", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU8", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU9", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU0", dbText, 25)
'      .Fields.Append .CreateField("DES_PLUX", dbText, 25)
'
'      .Fields.Append .CreateField("DES_PLU11", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU12", dbText, 0)
'      .Fields.Append .CreateField("DES_PLU13", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU14", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU15", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU16", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU17", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU18", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU19", dbText, 25)
'      .Fields.Append .CreateField("DES_PLU20", dbText, 25)
'
'      .Fields.Append .CreateField("ENVIO", dbText, 1)
'      .Fields.Append .CreateField("PROG_PLU", dbText, 130)
'      .Fields.Append .CreateField("TRAN_PLU", dbText, 1)
'      .Fields.Append .CreateField("PROG_TEXTO", dbText, 242)
'      .Fields.Append .CreateField("PROG_TX1", dbText, 242)
'      .Fields.Append .CreateField("TRAN_TEXTO", dbText, 1)
'      .Fields.Append .CreateField("TRAN_TX1", dbText, 1)
'      .Fields.Append .CreateField("ART_CB", dbText, 12)
'      .Fields.Append .CreateField("TRAN_CB", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'
''      For i = 0 To .Fields.Count - 1
''      .Fields(i).AllowZeroLength = True
''      Next i
'
''     For Each prpBucle In .Properties
''            On Error Resume Next
''            If prpBucle <> "" Then Debug.Print "    " & prpBucle.Name & " = " & prpBucle
''     Next prpBucle
'
'      dbsmostrador.TableDefs.Append mitabledef
      
End With

With miarti
'On Error Resume Next
      numero = 0
      Do Until numero = num_reg
      .AddNew
      !tienda = tienda(numero)
      !codi_ident = seccion(numero)
      !codi_sub = subsec(numero)
      !codi_fam = familia(numero)
      !codigo = codigo(numero)
      !plu = plu(numero)
      !precio = precio(numero)
      !caducidad = cadu(numero)
      !tara = tara(numero)
      !codi_pes = codi_pes(numero)
      !des_plu1 = des_plu1(numero)
      !des_plu2 = des_plu2(numero)
      !des_plu3 = des_plu3(numero)
      !des_plu4 = des_plu4(numero)
      !des_plu5 = des_plu5(numero)
      !des_plu6 = des_plu6(numero)
      !des_plu7 = des_plu7(numero)
      !des_plu8 = des_plu8(numero)
      !des_plu9 = des_plu9(numero)
      !des_plu0 = des_plu0(numero)
      !art_cb = art_cb(numero)
      !tran_texto = tran_texto(numero)
      !tran_cb = tran_cb(numero)
      !tran_tx1 = tran_tx1(numero)
      !tran_plu = tran_plu(numero)
      .Update
       numero = numero + 1
      Loop
      .Close
End With

With miseccion_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        seccion(numero) = !codi_ident
        tran_sec(numero) = !tran_sec
        tran_var(numero) = !tran_var
        descripcio(numero) = !descripcio
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_SEC", dbDouble, 6)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("DESCRIPCIO", dbText, 21)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'      .Fields.Append .CreateField("NUMERO_SUB", dbDouble, 2)
'      .Fields.Append .CreateField("NUMERO_VND", dbDouble, 4)
'      .Fields.Append .CreateField("PROG_SEC", dbText, 43)
'      .Fields.Append .CreateField("TRAN_SEC", dbText, 1)
'      .Fields.Append .CreateField("PROG_VAR", dbText, 17)
'      .Fields.Append .CreateField("TRAN_VAR", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append miseccion
'
End With

With miseccion
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !codi_ident = seccion(numero)
        !tran_sec = tran_sec(numero)
        !tran_var = tran_var(numero)
        !descripcio = descripcio(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With

With misubsec_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        des_sub(numero) = !descripcio
        tran_sub(numero) = !tran_sub
        numero_fam(numero) = !numero_fam
        codi_sub(numero) = !codi_sub
        seccion(numero) = !codi_ident
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_SUB", dbDouble, 3)
'      .Fields.Append .CreateField("DESCRIPCIO", dbText, 21)
'      .Fields.Append .CreateField("NUMERO_FAM", dbDouble, 2)
'      .Fields.Append .CreateField("PROG_SUB", dbText, 36)
'      .Fields.Append .CreateField("TRAN_SUB", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append misubsec
 End With
 
With misubsec
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !descripcio = des_sub(numero)
        !tran_sub = tran_sub(numero)
        !numero_fam = numero_fam(numero)
        !codi_sub = codi_sub(numero)
        !codi_ident = seccion(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With
 
 With mifam_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        des_fam(numero) = descripcio
        tran_fam(numero) = !tran_fam
        codi_fam(numero) = !codi_fam
        codi_sub(numero) = !codi_sub
        seccion(numero) = !codi_ident
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_SUB", dbDouble, 3)
'      .Fields.Append .CreateField("CODI_FAM", dbDouble, 4)
'      .Fields.Append .CreateField("DESCRIPCIO", dbText, 21)
'      .Fields.Append .CreateField("NUMERO_PLU", dbDouble, 5)
'      .Fields.Append .CreateField("PROG_FAM", dbText, 36)
'      .Fields.Append .CreateField("TRAN_FAM", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append mifam
 End With

With mifam
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        If !descripcio <> "" Then
        !descripcio = des_fam(numero)
        End If
        
        !tran_fam = tran_fam(numero)
        !codi_fam = codi_fam(numero)
        !codi_sub = codi_sub(numero)
        !codi_ident = seccion(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With

With miequip_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        seccion(numero) = !codi_ident
        des_eqp(numero) = !descripcio
        tran_eqp(numero) = !tran_eqp
        numero_eqp(numero) = !numero_eqp
        estado_eqp(numero) = !estado
        modelo(numero) = !modelo
        prioridad(numero) = !prioridad
        tipo_eqp(numero) = !tipo_eqp
        secc_maqui(numero) = !secc_maqui
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
       
'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("TIPO_EQP", dbDouble, 1)
'      .Fields.Append .CreateField("PRIORIDAD", dbDouble, 1)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'      .Fields.Append .CreateField("ESTADO", dbText, 1)
'      .Fields.Append .CreateField("PROG_PO", dbText, 1)
'      .Fields.Append .CreateField("PROG_CL", dbText, 1)
'      .Fields.Append .CreateField("MODELO", dbDouble, 1)
'      .Fields.Append .CreateField("DESCRIPCIO", dbText, 21)
'      .Fields.Append .CreateField("PROG_EQP", dbText, 17)
'      .Fields.Append .CreateField("TRAN_EQP", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append miequip
 End With
 
With miequip
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !codi_ident = seccion(numero)
        !descripcio = des_eqp(numero)
        !tran_eqp = tran_eqp(numero)
        !numero_eqp = numero_eqp(numero)
        !estado = estado_eqp(numero)
        !modelo = modelo(numero)
        !prioridad = prioridad(numero)
        !tipo_eqp = tipo_eqp(numero)
        !secc_maqui = secc_maqui(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With
 
 With micodbar_d
        numero = 0
        Do While Not .EOF
         sec_eqp(numero) = !secc_eqp  'local o red
         numero_eqp(numero) = !numero_eqp
         secc_maqui(numero) = !secc_maqui
         seccion(numero) = !codi_ident
         tran_bar(numero) = !tran_bar
         cb_vnt(numero) = !cb_vnt
         cb_sup(numero) = !cb_sup
         cb_mix(numero) = !cb_mix
         cb_ev1(numero) = !cb_ev1
         cb_ev2(numero) = !cb_ev2
         
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'
'      .Fields.Append .CreateField("SECC_EQP", dbText, 1)
'
'      .Fields.Append .CreateField("CB_VNT", dbText, 12)
'
'      .Fields.Append .CreateField("CB_SUP", dbText, 12)
'
'      .Fields.Append .CreateField("CB_MIX", dbText, 12)
'      .Fields.Append .CreateField("CB_EV1", dbText, 12)
'      .Fields.Append .CreateField("CB_EV2", dbText, 12)
'
'      .Fields.Append .CreateField("PROG_BAR", dbText, 72)
'      .Fields.Append .CreateField("TRAN_BAR", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append micodbar
 End With

With micodbar
        numero = 0
        Do Until numero = num_reg
        .AddNew
         !secc_eqp = sec_eqp(numero)  'local o red
         !numero_eqp = numero_eqp(numero)
         !secc_maqui = secc_maqui(numero)
         !codi_ident = seccion(numero)
         !tran_bar = tran_bar(numero)
         !cb_vnt = cb_vnt(numero)
         !cb_sup = cb_sup(numero)
         !cb_mix = cb_mix(numero)
         !cb_ev1 = cb_ev1(numero)
         !cb_ev2 = cb_ev2(numero)
         
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With

With micabley_d
 
        numero = 0
        Do While Not .EOF
                
         seccion(numero) = !codi_ident
         secc_maqui(numero) = !secc_maqui
         TRAN_CAB(numero) = !TRAN_CAB
         TRAN_LEY(numero) = !TRAN_LEY
         
         cabecera01(numero) = !cabecera01
         cabecera02(numero) = !cabecera02
         cabecera03(numero) = !cabecera03
         cabecera04(numero) = !cabecera04
         cabecera05(numero) = !cabecera05
         
         leyenda01(numero) = !leyenda01
         leyenda02(numero) = !leyenda02
         leyenda03(numero) = !leyenda03
         leyenda04(numero) = !leyenda04
         leyenda05(numero) = !leyenda05
         
         TLETRACAB1(numero) = !TLETRACAB1
         TLETRACAB2(numero) = !TLETRACAB2
         TLETRACAB3(numero) = !TLETRACAB3
         TLETRACAB4(numero) = !TLETRACAB4
         TLETRACAB5(numero) = !TLETRACAB5
         
         TLETRALEY1(numero) = !TLETRALEY1
         TLETRALEY2(numero) = !TLETRALEY2
         TLETRALEY3(numero) = !TLETRALEY3
         TLETRALEY4(numero) = !TLETRALEY4
         TLETRALEY5(numero) = !TLETRALEY5
         
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
        

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_EQP", dbBoolean, 1)
'
'      .Fields.Append .CreateField("CABECERA01", dbText, 25)
'      .Fields.Append .CreateField("CABECERA02", dbText, 25)
'      .Fields.Append .CreateField("CABECERA03", dbText, 25)
'      .Fields.Append .CreateField("CABECERA04", dbText, 25)
'      .Fields.Append .CreateField("CABECERA05", dbText, 25)
'
'      .Fields.Append .CreateField("LEYENDA01", dbText, 25)
'      .Fields.Append .CreateField("LEYENDA02", dbText, 25)
'      .Fields.Append .CreateField("LEYENDA03", dbText, 25)
'      .Fields.Append .CreateField("LEYENDA04", dbText, 25)
'      .Fields.Append .CreateField("LEYENDA05", dbText, 25)
'
'      .Fields.Append .CreateField("TLETRACAB1", dbText, 1)
'      .Fields.Append .CreateField("TLETRACAB2", dbText, 1)
'      .Fields.Append .CreateField("TLETRACAB3", dbText, 1)
'      .Fields.Append .CreateField("TLETRACAB4", dbText, 1)
'      .Fields.Append .CreateField("TLETRACAB5", dbText, 1)
'
'      .Fields.Append .CreateField("TLETRALEY1", dbText, 1)
'      .Fields.Append .CreateField("TLETRALEY2", dbText, 1)
'      .Fields.Append .CreateField("TLETRALEY3", dbText, 1)
'      .Fields.Append .CreateField("TLETRALEY4", dbText, 1)
'      .Fields.Append .CreateField("TLETRALEY5", dbText, 1)
'
'
'      .Fields.Append .CreateField("PROG_CAB", dbText, 151)
'      .Fields.Append .CreateField("TRAN_CAB", dbText, 1)
'      .Fields.Append .CreateField("PROG_LEY", dbText, 151)
'      .Fields.Append .CreateField("TRAN_LEY", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append micabley
 End With

With micabley
        numero = 0
        Do Until numero = num_reg
        .AddNew
         !codi_ident = seccion(numero)
         !secc_maqui = secc_maqui(numero)
         !TRAN_CAB = TRAN_CAB(numero)
         !TRAN_LEY = TRAN_LEY(numero)
         
         !cabecera01 = cabecera01(numero)
         !cabecera02 = cabecera02(numero)
         !cabecera03 = cabecera03(numero)
         !cabecera04 = cabecera04(numero)
         !cabecera05 = cabecera05(numero)
         
         !leyenda01 = leyenda01(numero)
         !leyenda02 = leyenda02(numero)
         !leyenda03 = leyenda03(numero)
         !leyenda04 = leyenda04(numero)
         !leyenda05 = leyenda05(numero)
         
         !TLETRACAB1 = TLETRACAB1(numero)
         !TLETRACAB2 = TLETRACAB2(numero)
         !TLETRACAB3 = TLETRACAB3(numero)
         !TLETRACAB4 = TLETRACAB4(numero)
         !TLETRACAB5 = TLETRACAB5(numero)
         
         !TLETRALEY1 = TLETRALEY1(numero)
         !TLETRALEY2 = TLETRALEY2(numero)
         !TLETRALEY3 = TLETRALEY3(numero)
         !TLETRALEY4 = TLETRALEY4(numero)
         !TLETRALEY5 = TLETRALEY5(numero)
         
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With

With mitecla_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        codigo(numero) = !codigo
        plu(numero) = !plu
        des_plu1(numero) = !DESC_PLU
        
        tabla(numero) = !tabla
        tran_tec(numero) = !tran_tec
        codi_tec(numero) = !codi_tec
        numero_eqp(numero) = !numero_eqp
        seccion(numero) = !codi_ident
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

End With

With mitecla
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !codigo = codigo(numero)
        !plu = plu(numero)
        !DESC_PLU = des_plu1(numero)
        
        !tabla = tabla(numero)
        !tran_tec = tran_tec(numero)
        !codi_tec = codi_tec(numero)
        !numero_eqp = numero_eqp(numero)
        !codi_ident = seccion(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("CODIGO", dbDouble, 6)
'      .Fields.Append .CreateField("PLU", dbDouble, 4)
'      .Fields.Append .CreateField("DESC_PLU", dbText, 25)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_EQP", dbBoolean, 1)
'      .Fields.Append .CreateField("TABLA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_TEC", dbDouble, 3)
'      .Fields.Append .CreateField("PROG_TEC", dbText, 18)
'      .Fields.Append .CreateField("TRAN_TEC", dbText, 1)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append mitecla
 End With

With mivendedor_d
        numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        seccion(numero) = !codi_ident
        ident_vend(numero) = !ident_vend
        codi_vend(numero) = !codi_vend
        tec_vend(numero) = !tec_vend
        nombre_ven(numero) = !Nombre
        tran_vend(numero) = !tran_vend
        tran_asign(numero) = !tran_asign
        fecha_vend(numero) = !fecha
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close

'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("IDENT_VEND", dbDouble, 4)
'      .Fields.Append .CreateField("CODI_VEND", dbDouble, 6)
'      .Fields.Append .CreateField("TEC_VEND", dbDouble, 2)
'      .Fields.Append .CreateField("NOMBRE", dbText, 30)
'      .Fields.Append .CreateField("PROG_VEND", dbText, 45)
'      .Fields.Append .CreateField("TRAN_VEND", dbText, 1)
'      .Fields.Append .CreateField("ASIGN_VEND", dbText, 13)
'      .Fields.Append .CreateField("TRAN_ASIGN", dbText, 1)
'      .Fields.Append .CreateField("FECHA", dbDate, 8)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append mivendedor
 End With
 
With mivendedor
        numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !codi_ident = seccion(numero)
        !ident_vend = ident_vend(numero)
        !codi_vend = codi_vend(numero)
        !tec_vend = tec_vend(numero)
        !Nombre = nombre_ven(numero)
        !tran_vend = tran_vend(numero)
        !tran_asign = tran_asign(numero)
        !fecha = fecha_vend(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
End With

With mipubli_d
       numero = 0
        Do While Not .EOF
        tienda(numero) = !tienda
        seccion(numero) = !codi_ident
        secc_maqui(numero) = !secc_maqui
        numero_eqp(numero) = !numero_eqp
        publicidad(numero) = !publicidad
        oferta01(numero) = !oferta01
        oferta02(numero) = !oferta02
        oferta03(numero) = !oferta03
        oferta04(numero) = !oferta04
        oferta05(numero) = !oferta05
        oferta06(numero) = !oferta06
        oferta07(numero) = !oferta07
        oferta08(numero) = !oferta08
        oferta09(numero) = !oferta09
        oferta10(numero) = !oferta10
        tipo_publ(numero) = !tipo_publ
        tran_pub(numero) = !tran_pub
        tran_ofe(numero) = !tran_ofe
        tran_act(numero) = !tran_act
        .MoveNext
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
        
'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
'      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
'      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
'      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
'      .Fields.Append .CreateField("PUBLICIDAD", dbText, 100)
'      .Fields.Append .CreateField("OFERTA01", dbText, 16)
'      .Fields.Append .CreateField("OFERTA02", dbText, 16)
'      .Fields.Append .CreateField("OFERTA03", dbText, 16)
'      .Fields.Append .CreateField("OFERTA04", dbText, 16)
'      .Fields.Append .CreateField("OFERTA05", dbText, 16)
'      .Fields.Append .CreateField("OFERTA06", dbText, 16)
'      .Fields.Append .CreateField("OFERTA07", dbText, 16)
'      .Fields.Append .CreateField("OFERTA08", dbText, 16)
'      .Fields.Append .CreateField("OFERTA09", dbText, 16)
'      .Fields.Append .CreateField("OFERTA10", dbText, 16)
'      .Fields.Append .CreateField("TIPO_PUBL", dbText, 1)
'      .Fields.Append .CreateField("PROG_PUB", dbText, 113)
'      .Fields.Append .CreateField("TRAN_PUB", dbText, 1)
'      .Fields.Append .CreateField("PROG_OFE", dbText, 180)
'      .Fields.Append .CreateField("TRAN_OFE", dbText, 1)
'      .Fields.Append .CreateField("PROG_ACT", dbText, 9)
'      .Fields.Append .CreateField("TRAN_ACT", dbText, 1)
'       .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      dbsmostrador.TableDefs.Append mipubli
 End With

With mipubli
       numero = 0
        Do Until numero = num_reg
        .AddNew
        !tienda = tienda(numero)
        !codi_ident = seccion(numero)
        !secc_maqui = secc_maqui(numero)
        !numero_eqp = numero_eqp(numero)
        !publicidad = publicidad(numero)
        !oferta01 = oferta01(numero)
        !oferta02 = oferta02(numero)
        !oferta03 = oferta03(numero)
        !oferta04 = oferta04(numero)
        !oferta05 = oferta05(numero)
        !oferta06 = oferta06(numero)
        !oferta07 = oferta07(numero)
        !oferta08 = oferta08(numero)
        !oferta09 = oferta09(numero)
        !oferta10 = oferta10(numero)
        !tipo_publ = tipo_publ(numero)
        !tran_pub = tran_pub(numero)
        !tran_ofe = tran_ofe(numero)
        !tran_act = tran_act(numero)
        .Update
        numero = numero + 1
        Loop
        num_reg = numero
        .Close
        
End With
'***************************Borrar todas la tablas .dbf*******************

dbsdbaseIV.Execute "DROP TABLE Articulo.dbf;"
dbsdbaseIV.Execute "DROP TABLE seccion.dbf;"
dbsdbaseIV.Execute "DROP TABLE subsec.dbf;"
dbsdbaseIV.Execute "DROP TABLE familia.dbf;"
dbsdbaseIV.Execute "DROP TABLE codbar.dbf;"
dbsdbaseIV.Execute "DROP TABLE cabley.dbf;"
dbsdbaseIV.Execute "DROP TABLE publi.dbf;"
dbsdbaseIV.Execute "DROP TABLE teclas.dbf;"

On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE Gtarti.dbf;"
On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE Gthora.dbf;"
On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE Gtsecs.dbf;"
On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE Gtvend.dbf;"
On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE Tiquetsc.dbf;"
On Error Resume Next
dbsdbaseIV.Execute "DROP TABLE tipetq.dbf;"
dbsdbaseIV.Execute "DROP TABLE Equipos.dbf;"
dbsdbaseIV.Execute "DROP TABLE vendedor.dbf;"

'****************************Fin de borrado de las tablas .dbf*************

dbsdbaseIV.Close
dbsjet.Close


'*****************************Prueba de campos nuevos******************************************

'Dim dbsmostrador As Database
'Dim mitabledef As TableDef
'
'
'Set dbsmostrador = OpenDatabase(miruta, False, 0, "Dbase IV")
''Set dbsmostrador = OpenDatabase(miruta, False, 0, "dbase.mdb")
'
'Set mitabledef = dbsmostrador.TableDefs("Articulo")
''On Error Resume Next
''With mitabledef
''      .Fields.Append .CreateField("des_plu11", dbText, 25)
'
'      dbsmostrador.Execute "ALTER TABLE Articulo " & "ADD COLUMN des_plu20 TEXT;"
'      dbsmostrador.Execute "ALTER TABLE Articulo " & "ADD COLUMN des_plu19 TEXT;"
'
''End With
'dbsmostrador.Close


'**********************************************************************************************
End Sub

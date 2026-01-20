Attribute VB_Name = "s"
Option Explicit
Global SUPECO As Boolean
Public Codigos() As Long
Public NCodigos As Long
Public Mostradores() As Integer
Public Secciones() As Integer
Public NMostradores As Integer
Public Buscado As Boolean
Private Function Record_Action(Registro As Recordset, lAction As Long) As Boolean
    Dim RetVal As Boolean
    Dim ErrorVal As Long
    Dim Continuar As Boolean
    Continuar = True
    RetVal = False
    Do While Continuar
        On Error Resume Next
        If lAction = 0 Then
            Registro.Edit
            ErrorVal = Err.Number
        End If
        If lAction = 1 Then
            Registro.Delete
            ErrorVal = Err.Number
        End If
        On Error GoTo 0
        Select Case ErrorVal
            Case 0
                RetVal = True
                Continuar = False
            Case 3260 ' otro está bloqueando
                Sleep (10)
            Case 3197 ' el registro ha sido borrado
                RetVal = False
                Continuar = False
            Case Else ' base dañada
                CadenadeLog "Critical error while updating recordset"
                
        End Select
    Loop
    Record_Action = RetVal
End Function
Public Function Edit_Record(Registro As Recordset) As Boolean
    Edit_Record = Record_Action(Registro, 0)
End Function
Public Function Delete_Record(Registro As Recordset) As Boolean
    Delete_Record = Record_Action(Registro, 1)
End Function
Public Sub Buscar_Mostradores()
    Dim MyBase As Database
    Dim Registro As Recordset
    NMostradores = 0
    Set MyBase = OpenDatabase(Base_General)
    Set Registro = MyBase.OpenRecordset("select * from seccion where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                NMostradores = NMostradores + 1
                ReDim Preserve Mostradores(NMostradores)
                ReDim Preserve Secciones(NMostradores)
                Mostradores(NMostradores - 1) = .Fields("codi_ident")
                Secciones(NMostradores - 1) = .Fields("secc_maqui")
                If Not .EOF Then .MoveNext
            Loop
        End If
    End With
    Buscado = True
    MyBase.Close
End Sub
Public Sub Buscar_Codigos()
    Dim MyBase As Database
    Dim Registro As Recordset
    NCodigos = 0
    Set MyBase = OpenDatabase(Base_General)
    Set Registro = MyBase.OpenRecordset("select codigo from articulo where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                NCodigos = NCodigos + 1
                ReDim Preserve Codigos(NCodigos)
                Codigos(NCodigos - 1) = !codigo
                If Not .EOF Then .MoveNext
            Loop
        End If
    End With
    Buscado = True
    MyBase.Close
End Sub
Public Function Existe_Mostrador(NumMos As Long) As Integer
    Dim Resp As Integer
    Dim bucle As Integer
    Resp = -1
    For bucle = 0 To NMostradores - 1
        If NumMos = Mostradores(bucle) Then
            Resp = Secciones(bucle)
            Exit For
        End If
    Next bucle
    Existe_Mostrador = Resp
End Function
Public Function Existe_Codigo(NumCod As Long) As Long
    Dim Resp As Long
    Dim bucle As Integer
    Resp = -1
    For bucle = 0 To NCodigos - 1
        If NumCod = Codigos(bucle) Then
            Resp = Codigos(bucle)
            Exit For
        End If
    Next bucle
    Existe_Codigo = Resp
End Function

Public Function Alta_Articulo(Balenv As Boolean, Base As Database, RegArticulo As Recordset, Optional desc As Integer) As Integer
    Dim ModificacionTramo As Boolean
    Dim ModificacionBasica As Boolean
    Dim Modificaciontexto As Boolean
    Dim ModificacionTx1 As Boolean
    Dim ModificacionCodBar As Boolean
    Dim ModificacionPrecio As Boolean
    Dim bucle As Integer
    Dim Buclelite As Integer
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim Reg4 As Recordset
    Dim Reg5 As Recordset
    Dim RegTramo As Recordset
    Dim Retorno As Integer
    Dim RegFam As Recordset
    Dim RegSubsec As Recordset
    Dim RstRechazos As Recordset
    Dim RegPlu As Recordset
    Dim ChkTramo As Boolean
    Dim MiPrecio As Double
    Dim MiEuro As Double
    Dim CierraDespues As Boolean
    Dim SeCambioPrecio As Boolean
    Dim PluBuffer As Integer
    Dim Alcampo As Boolean
    Dim MiWgh As Boolean
    Dim MiNombre As String
    Dim MiAlta As Boolean
    Dim BucleCad As Integer
    Dim lDarBajaAlta As Boolean
    Dim nPosiPlu As Double
    
    If Buscado = False Then
        Buscar_Mostradores
        Buscar_Codigos
    End If
    MiAlta = False
    SeCambioPrecio = False
    If articulo.prc = -1 Then articulo.prc = 0
    If articulo.cad = -1 Then articulo.cad = 0
    If articulo.tar = -1 Then articulo.tar = 0
    If articulo.sub = -1 Then articulo.sub = 0
    If articulo.fam = -1 Then articulo.fam = 0
    If articulo.pre = -1 Then articulo.pre = 0
    If articulo.etq = -1 Then articulo.etq = 255
    If articulo.gpc = -1 Then articulo.gpc = 0
    '*****************
    ' Pesado/No pesado
    ' se acepta para pesado :
    ' "W","P","0"
    ' para no pesado :
    ' "U","1"
    '*****************
    MiWgh = False
    If articulo.wgh = "0" Then
        MiWgh = True
    Else
        If articulo.wgh = "P" Then
            MiWgh = True
        Else
            If articulo.wgh = "W" Then
                MiWgh = True
            Else
                MiWgh = False
            End If
        End If
    End If
    If articulo.trm1 <> -1 Or articulo.trm1 <> -1 Or _
    articulo.prc2 <> -1 Or articulo.trm2 <> -1 Then
        ChkTramo = True
    Else
        ChkTramo = False
    End If
    'c2f 26/07/05
    'If Trim(articulo.cba) <> "" Then
    '    If Len(articulo.cba) <> 12 Then
    '        Retorno = 5
    '        GoTo fin
    '    End If
    'End If
    If RechazarPlu0 And articulo.prc = 0 Then
        Retorno = 50
        GoTo fin
    End If
    If SUPECO And ((articulo.plu < 1000) Or (articulo.plu > 9999)) Then
        Retorno = 14
        GoTo fin
    End If
    If SUPECO And ((articulo.plu >= 2000) And (articulo.plu <= 2999)) Then
        Retorno = 14
        GoTo fin
    End If
    If (VarPublics.C9Cod <> 0) And (articulo.plu > 6999) Then    '///cas.v118
        Retorno = 14
        GoTo fin
    End If
    If articulo.cod = 0 Then
        Retorno = 14
        GoTo fin
    End If
    If articulo.sec = 0 Then
        If Dir(App.Path & "\sec0a1.txt") <> "" Then
            articulo.sec = 1
        Else
           Retorno = 15
           GoTo fin
        End If
    End If
    If (VarPublics.C9Cod <> 0) And (articulo.sec > 1) Then   '///cas.v118
        Retorno = 15
        GoTo fin
    End If
    If articulo.Merma > 99 Or articulo.Presec > 99 Then
        Retorno = 22
        GoTo fin
    End If
    If ChkTramo = True Then
        If articulo.trm1 = -1 And articulo.prc1 = -1 Then
            Retorno = 16
            GoTo fin
        End If
        If (articulo.trm2 <> -1 And articulo.prc2 = -1) Or _
        (articulo.prc2 <> -1 And articulo.trm2 = -1) Then
            Retorno = 16
            GoTo fin
        End If
        If (articulo.trm1 <> -1 And articulo.prc1 = -1) Or _
        (articulo.prc1 <> -1 And articulo.trm1 = -1) Then
            Retorno = 16
            GoTo fin
        End If
    End If
    '***************
    ' Precio / Euros
    '***************
    If UsaEuro Then
        MiEuro = articulo.prc / 100
        MiPrecio = EUROtoPTA(articulo.prc / 100)
    Else
        MiPrecio = articulo.prc / (10 ^ Decimales)
        MiEuro = PTAtoEURO(articulo.prc)
    End If
    
    If TypeName(RegSubsec) = "Nothing" Then
        Set RegSubsec = Base.OpenRecordset("select codi_ident,borrado,codi_sub,tran_sub from subsec")
    End If
    If ChkTramo Then
        Set RegTramo = Base.OpenRecordset("select count(codigo) from tramos")
        With RegTramo
            If .Fields(0) >= 200 Then
                Retorno = 17
                
                GoTo fin
            End If
        End With
    End If
    Retorno = 4
  
    If Existe_Mostrador(articulo.sec) < 0 Then
        Retorno = 3
        GoTo fin
    End If

    If FamiliasFijas Then
        Set RegFam = Base.OpenRecordset("select codi_fam from fam_code where codi_fam=" & articulo.fam & " and codi_ident=" & articulo.sec)
        If RegFam.EOF Then
            Retorno = 3
            GoTo fin
        End If
    End If
    
    If SubseccionesFijas Then
        Set RegFam = Base.OpenRecordset("select codi_sub from sub_code where codi_sub=" & articulo.sub & " and codi_ident=" & articulo.sec)
        If RegFam.EOF Then
            Retorno = 3
            GoTo fin
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''
    'Carrefour
    ''''''''''''''''''''''''''''''''''''
    If Dir(App.Path & "\hiper") <> "" Then
    
        '********************
        ' comprueba plu repetido
        '********************
        If articulo.plu <> 0 Then
            'posible hueco
            nPosiPlu = 0
            For bucle = 1 To 9999
                Set RegPlu = Base.OpenRecordset _
                ("select plu from articulo where secc_maqui=" & _
                Existe_Mostrador(articulo.sec) & " and plu=" & bucle & " and borrado=false")
                If RegPlu.EOF Then
                   nPosiPlu = bucle
                   bucle = 10000
                End If
            Next bucle
            
            Set RegPlu = Base.OpenRecordset _
            ("select * from articulo where secc_maqui=" & _
            Existe_Mostrador(articulo.sec) & " and plu=" & articulo.plu & " and codigo<>" & articulo.cod & " and borrado=false")
            If Not RegPlu.EOF Then 'se cambia el PLU a PosiPlu
                RegPlu.Edit
                RegPlu!plu = nPosiPlu
                RegPlu!prc3 = Val(Format(Now, "yymmddhhmm"))
                RegPlu!tran_en = ""
                RegPlu!tran_el = ""
                RegPlu!tran_plu = ""
                RegPlu!tran_pluSC10 = ""
                RegPlu!tran_textoel = ""
                RegPlu!tran_texto = ""
                RegPlu!tran_textosc10 = ""
                RegPlu!tran_tx1 = ""
                RegPlu!tran_tx1SC10 = ""
                RegPlu!tran_cb = ""
                RegPlu!tran_cbsc10 = ""
                RegPlu.Update
            End If
        End If
   
    End If
    ''''''''''''''''''''''''''''''''''''
    
    With RegArticulo
        ' Nuevo dato
        If Existe_Codigo(articulo.cod) < 0 Then
            '********************
            ' comprueba plu repetido
            '********************
            If articulo.plu <> 0 Then
                Set RegPlu = Base.OpenRecordset _
                ("select plu from articulo where secc_maqui=" & _
                Existe_Mostrador(articulo.sec) & " and plu=" & articulo.plu & " and borrado=false")
                If Not RegPlu.EOF Then
                    Retorno = 21
                    GoTo fin
                End If
            End If
            '********************
            .AddNew
            NCodigos = NCodigos + 1
            ReDim Preserve Codigos(NCodigos)
            Codigos(NCodigos - 1) = articulo.cod
            MiAlta = True
            Retorno = 0
            ModificacionBasica = True
            !prc3 = Val(Format(Now, "yymmddhhmm"))
            !tran_en = ""
            !tran_el = ""
            !tran_textoel = "*"
            !tran_plu = ""
            !tran_pluSC10 = ""
            !tran_texto = ""
            !tran_textosc10 = ""
            !tran_tx1 = ""
            !tran_tx1SC10 = ""
            !tran_cb = ""
            !tran_cbsc10 = ""
            !tipo_iva = articulo.iva
            !codigo = articulo.cod
            !codi_ident = articulo.sec
            !pref = articulo.pre
            !etq = articulo.etq
            !grupo_conserv = articulo.gpc
            !fcb = articulo.Merma
            !ning = articulo.Presec
            .Fields("prc100g") = articulo.prc100g
            !secc_maqui = Existe_Mostrador(articulo.sec)
            If articulo.plu = 0 And Dir(App.Path & "\hiper") = "" Then
                Set Reg3 = Base.OpenRecordset _
                ("select max(plu) from articulo where borrado=false and secc_maqui=" & Existe_Mostrador(articulo.sec))
                If Not IsNull(Reg3.Fields(0)) And Not Reg3.EOF Then
                    If Reg3.Fields(0) < 9999 Then
                        !plu = Reg3.Fields(0) + 1
                        PluBuffer = Reg3.Fields(0) + 1
                    Else
                        
                        For bucle = 1 To 9999
                            'RegArticulo.Filter = "borrado = False And secc_maqui = """ _
                            '& Reg2!secc_maqui & " and plu=" & bucle
                            RegArticulo.Filter = "borrado = False And secc_maqui = " _
                            & Existe_Mostrador(articulo.sec) & " and plu=" & bucle
                            
                            Set Reg4 = RegArticulo.OpenRecordset
                            If Reg4.EOF Then
                                !plu = bucle
                                PluBuffer = bucle
                                Exit For
                            End If
                        Next bucle
                    End If
                Else
                    !plu = 1
                    PluBuffer = 1
                End If
            Else
                !plu = articulo.plu
                PluBuffer = articulo.plu
            End If
            If Trim(articulo.cba) <> "" Then
                !art_cb = articulo.cba
            Else
                .Fields("art_cb") = ""
            End If
            !codi_sub = articulo.sub
            !codi_fam = articulo.fam
            !precio = MiPrecio
            !Euros = MiEuro
            !Caducidad = articulo.cad
            '********************************
            ' datos de caducidad programados
            ' en fichero
            '
            If DataCaducidad.Numero > 0 Then
                For bucle = 0 To DataCaducidad.Numero - 1
                    If DataCaducidad.Seccion(bucle) = articulo.sec Then
                        If DataCaducidad.Subseccion(bucle) = articulo.sub Then
                            If DataCaducidad.Familia(bucle) = articulo.fam Then
                                !Caducidad = DataCaducidad.Caducidad(bucle)
                            End If
                        End If
                    End If
                Next bucle
            End If
            '********************************
            !tara = articulo.tar
             
            If MiWgh = True Then
                !codi_pes = "W"
            Else
                !codi_pes = "U"
            End If
            If Balenv = False Then
                !Balenv = False
            Else
                !Balenv = True
                Set Reg5 = Base.OpenRecordset _
                ("select enviardatos,multiple,tran_secSC10,tran_sec from seccion where borrado=false and secc_maqui=" & Existe_Mostrador(articulo.sec))
                With Reg5
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                        If !enviardatos <> "TODOS" Then
                            Edit_Record Reg5
                            If !multiple <> 1 Then
                            
                                If !enviardatos = "GA" Then !tran_secSC10 = ""
                                If !enviardatos = "SC10" Then !tran_sec = ""
                            End If
                            !multiple = 1
                            .Update
                        End If
                        If Not .EOF Then .MoveNext
                        Loop
                    End If
                End With
            End If
           .Fields("des_plu1") = articulo.txt(0)
           .Fields("des_plu2") = articulo.txt(1)
           .Fields("des_plu3") = articulo.txt(2)
           .Fields("des_plu4") = articulo.txt(3)
           .Fields("des_plu5") = articulo.txt(4)
           .Fields("des_plu6") = articulo.txt(5)
           .Fields("des_plu7") = articulo.txt(6)
           .Fields("des_plu8") = articulo.txt(7)
           .Fields("des_plu9") = articulo.txt(8)
           .Fields("des_plu0") = articulo.txt(9)
           .Fields("des_plux") = articulo.txt(10)
           .Fields("des_plu11") = articulo.txt(11)
           .Fields("des_plu12") = articulo.txt(12)
           .Fields("des_plu13") = articulo.txt(13)
           .Fields("des_plu14") = articulo.txt(14)
           .Fields("des_plu15") = articulo.txt(15)
           .Fields("des_plu16") = articulo.txt(16)
           .Fields("des_plu17") = articulo.txt(17)
           .Fields("des_plu18") = articulo.txt(18)
           .Fields("des_plu19") = articulo.txt(19)
           .Fields("des_plu20") = articulo.txt(20)
            For bucle = 0 To 9
                'If bucle < 10 Then
                    Select Case bucle
                        Case 0 To 8
                            MiNombre = "tip_let" & Val(bucle + 1)
                        Case 9
                            MiNombre = "tip_let0"
                    End Select
                    .Fields(MiNombre) = articulo.tlt(bucle)
                'End If
            Next bucle
            '*****
            ' Tramos
            '*****
            If ChkTramo Then
                !usatramos = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=false and codigo=" & !codigo)
                With RegTramo
                    If .EOF Then
                        .AddNew
                        
                    Else
                        Edit_Record RegTramo
                    End If
                        !borrado = False
                        !tran_tramo = ""
                        !codigo = articulo.cod
                        If articulo.trm1 <> -1 And articulo.prc1 <> -1 Then
                            
                            !tramo1 = articulo.trm1 / 1000
                            If UsaEuro Then
                                !euros1 = articulo.prc1 / 100
                                !precio1 = EUROtoPTA(articulo.prc1 / 100)
                            Else
                                !euros1 = PTAtoEURO(articulo.prc1)
                                !precio1 = articulo.prc1
                            End If
                                If articulo.trm2 <> -1 And articulo.prc2 <> -1 Then
                                    !tramo2 = articulo.trm2 / 1000
                                    If UsaEuro Then
                                        !euros2 = articulo.prc2 / 100
                                        !precio2 = EUROtoPTA(articulo.prc2 / 100)
                                    Else
                                        !precio2 = articulo.prc2
                                        !euros2 = PTAtoEURO(articulo.prc2)
                                    End If
                                Else
                                    !tramo2 = -1
                                    !precio2 = -1
                                    !euros2 = -1
                                End If
                        End If
                    .Update
                End With
            Else
                !usatramos = False
            End If
                        
            .Update
            If articulo.sub <> "0" Then
                RegSubsec.FindFirst "borrado=false and codi_ident=" & _
                articulo.sec & " and codi_sub=" & articulo.sub
                With RegSubsec
                If .EOF Or .NoMatch Then
                    .AddNew
                    !codi_sub = articulo.sub
                    !codi_ident = articulo.sec
                    '!secc_maqui = Reg2!secc_maqui
                    !borrado = False
                    !tran_sub = ""
                    .Update
                End If
                End With
            End If
            If articulo.fam <> "0" Then
                Set Reg3 = Base.OpenRecordset _
                ("select codi_fam,codi_sub,codi_ident,secc_maqui,borrado,tran_fam from familia where borrado=false and codi_ident=" & _
                articulo.sec & " and codi_sub=" & articulo.sub & _
                " and codi_fam=" & articulo.fam)
                With Reg3
                If .EOF Then
                    .AddNew
                    !codi_fam = articulo.fam
                    !codi_sub = articulo.sub
                    !codi_ident = articulo.sec
                    !secc_maqui = Existe_Mostrador(articulo.sec)
                    !borrado = False
                    !tran_fam = ""
                    .Update
                End If
                End With
            End If
        Else
            '************************
            ' Modificación
            '************************
            RegArticulo.FindFirst "borrado=false and codigo=" & CStr(articulo.cod)
            
            If lEsAho Then
               If !codi_ident <> articulo.sec Then
                  CadenadeLog "Baja de artículo" & CStr(articulo.cod)
                  Retorno = 69
                  GoTo fin
               End If
            End If
            
            If articulo.plu <> 0 Then
            
                If !plu <> articulo.plu Then
                    If Dir(App.Path & "\hiper") <> "" Then
                        lDarBajaAlta = True
                    Else
                        Retorno = 2
                        GoTo fin
                    End If
                End If
            End If
            Edit_Record RegArticulo
            ''''''''''''''''''''''''''' en principio para carrefour
            If lDarBajaAlta = True Then
            '    !borrado = True
            '    'articulo.txt(0) = .Fields("des_plu1")
            '    articulo.txt(1) = .Fields("des_plu2")
            '    articulo.txt(2) = .Fields("des_plu3")
            '    articulo.txt(3) = .Fields("des_plu4")
            '    articulo.txt(4) = .Fields("des_plu5")
            '    articulo.txt(5) = .Fields("des_plu6")
            '    articulo.txt(6) = .Fields("des_plu7")
            '    articulo.txt(7) = .Fields("des_plu8")
            '    articulo.txt(9) = .Fields("des_plu0")
            '    articulo.txt(10) = .Fields("des_plux")
            '    articulo.txt(11) = .Fields("des_plu11")
            '    articulo.txt(12) = .Fields("des_plu12")
            '    articulo.txt(13) = .Fields("des_plu13")
            '    articulo.txt(14) = .Fields("des_plu14")
            '    articulo.txt(15) = .Fields("des_plu15")
            '    articulo.txt(16) = .Fields("des_plu16")
            '    articulo.txt(17) = .Fields("des_plu17")
            '    articulo.txt(18) = .Fields("des_plu18")
            '    articulo.txt(19) = .Fields("des_plu19")
            '    articulo.txt(20) = .Fields("des_plu20")

            '    articulo.tlt(0) = .Fields("tip_let1")
            '    articulo.tlt(1) = .Fields("tip_let2")
            '    articulo.tlt(2) = .Fields("tip_let3")
            '    articulo.tlt(3) = .Fields("tip_let4")
            '    articulo.tlt(4) = .Fields("tip_let5")
            '    articulo.tlt(5) = .Fields("tip_let6")
            '    articulo.tlt(6) = .Fields("tip_let7")
            '    articulo.tlt(7) = .Fields("tip_let8")
            '    articulo.tlt(8) = .Fields("tip_let9")
            '    articulo.tlt(9) = .Fields("tip_let0")
            '    articulo.tlt(10) = 0

            '    .Update
            '    .AddNew
            '    !prc3 = Val(Format(Now, "yymmddhhmm"))
            '    !tran_en = ""
            '    !tran_el = ""
            '    !tran_textoel = "*"
            '    !tran_plu = ""
            '    !tran_pluSC10 = ""
            '    !tran_texto = ""
            '    !tran_textosc10 = ""
            '    !tran_tx1 = ""
            '    !tran_tx1SC10 = ""
            '    !tran_cb = ""
            '    !tran_cbsc10 = ""
            '    !tipo_iva = articulo.iva
            '    !codigo = articulo.cod
            '    !codi_ident = articulo.sec
            '    !pref = articulo.pre
            '    !etq = articulo.etq
            '    !grupo_conserv = articulo.gpc
            '    !fcb = articulo.Merma
            '    !ning = articulo.Presec
            '    .Fields("prc100g") = articulo.prc100g
            '    !secc_maqui = Existe_Mostrador(articulo.sec)
            '    !plu = articulo.plu
                !plu = articulo.plu
                PluBuffer = articulo.plu
            Else
                PluBuffer = !plu
            End If
                
                'PluBuffer = !plu
                ModificacionBasica = False
                Modificaciontexto = False
                ModificacionTx1 = False
                ModificacionCodBar = False
                ModificacionTramo = False
                If TodoModificaciones Then
                    ModificacionBasica = True
                    Modificaciontexto = True
                    ModificacionTx1 = True
                    ModificacionCodBar = True
                    ModificacionTramo = True
                End If
                If (!grupo_conserv <> articulo.gpc) And (articulo.m_gpc) Then
                    Modificaciontexto = True
                    ModificacionTx1 = True
                End If
                If IsNull(!tipo_iva) Then !tipo_iva = 0
                If (.Fields("prc100g") <> articulo.prc100g) And articulo.m_prc100g Then ModificacionBasica = True
                If (!tipo_iva <> articulo.iva) And articulo.m_iva Then ModificacionBasica = True
                If (!pref <> articulo.pre) And articulo.m_pre Then ModificacionBasica = True
                If (!etq <> articulo.etq) And articulo.m_etq Then ModificacionBasica = True
                If (!codi_sub <> articulo.sub) And articulo.m_sub Then ModificacionBasica = True
                If (!codi_fam <> articulo.fam) And articulo.m_fam Then ModificacionBasica = True
                If (!fcb <> articulo.Merma) And articulo.m_merma Then ModificacionBasica = True
                If (!ning <> articulo.Presec) And articulo.m_presec Then ModificacionBasica = True
                If Not UsaEuro Then
                    If !precio <> MiPrecio Then
                        ModificacionBasica = True
                        ModificacionPrecio = True
                        SeCambioPrecio = True
                    End If
                Else
                    If !Euros <> MiEuro Then
                        ModificacionBasica = True
                        ModificacionPrecio = True
                        SeCambioPrecio = True
                    End If
                End If
                If articulo.cad <> 0 Or lEsAho = True Then
                    If (!Caducidad <> articulo.cad) And articulo.m_cad Then ModificacionBasica = True
                End If
                If SUPECO Or Dir(App.Path & "\hiper") <> "" And (articulo.tar = 0) Then articulo.tar = !tara
                If Dir(App.Path & "\hiper") <> "" And (articulo.cad = 0) Then articulo.cad = !Caducidad
                If (!tara <> articulo.tar) And articulo.m_tar Then ModificacionBasica = True
                If (!art_cb <> articulo.cba) And articulo.m_cba Then ModificacionCodBar = True
                '****************************
                ' Alcampo : en las modificaciones
                ' no se altera pesado/no pesado
                ' porque llegan datos de AS400
                ' a corregir por el usuario
                If Not Conf_Al Then
                    If (MiWgh = True And !codi_pes = "U") Or _
                    (MiWgh = False And !codi_pes = "W") Then ModificacionBasica = True
                End If
                '/Alcampo
                '***********************************
                If (Balenv = False And !Balenv = True) Or _
                (Balenv = True And !Balenv = False) Then ModificacionBasica = True
                
                If Dir(App.Path & "\hiper") = "" Then
                
                    For Buclelite = 0 To 10
                        Select Case Buclelite
                            Case 0 To 8
                                MiNombre = "des_plu" & Val(Buclelite + 1)
                            Case 9
                                MiNombre = "des_plu0"
                            Case 10
                                MiNombre = "des_plux"
                        End Select
                        If Not AHOLD Or (AHOLD And Trim(articulo.txt(Buclelite)) <> "") Then
                            If articulo.m_txt(Buclelite) Then
                                If (.Fields(MiNombre) <> articulo.txt(Buclelite)) Or _
                                (.Fields(MiNombre) = "" And articulo.txt(Buclelite) <> "") Or _
                                (.Fields(MiNombre) <> "" And articulo.txt(Buclelite) = "") _
                                Then Modificaciontexto = True
                            End If
                        End If
                        If Buclelite < 10 Then
                            MiNombre = "des_plu" & Val(Buclelite + 11)
                            If Not AHOLD Or (AHOLD And Trim(articulo.txt(Buclelite + 11)) <> "") Then
                                If articulo.m_txt(Buclelite + 11) Then
                                    If (.Fields(MiNombre) <> articulo.txt(Buclelite + 11)) Or _
                                    (.Fields(MiNombre) = "" And articulo.txt(Buclelite + 11) <> "") Or _
                                    (.Fields(MiNombre) <> "" And articulo.txt(Buclelite + 11) = "") _
                                    Then Modificaciontexto = True
                                End If
                            End If
                        End If
                        If Buclelite < 10 Then
                            Select Case Buclelite
                                Case 0 To 8
                                    MiNombre = "tip_let" & Val(Buclelite + 1)
                                Case 9
                                    MiNombre = "tip_let0"
                            End Select
                            If Not AHOLD Or (AHOLD And articulo.tlt(Buclelite) <> 0) Then
                                If articulo.m_tlt(Buclelite) Then
                                    If (.Fields(MiNombre) <> articulo.tlt(Buclelite)) _
                                    Then Modificaciontexto = True
                                End If
                            End If
                        End If
                    Next Buclelite
                
                Else
                    Modificaciontexto = False
                    Modificaciontexto = False
                End If
                
                If ModificacionBasica Then
                    !prc3 = Val(Format(Now, "yymmddhhmm"))
                    !tran_en = ""
                    !tran_el = ""
                    !tran_plu = ""
                    !tran_pluSC10 = ""
                End If
                If Modificaciontexto Then
                    !prc3 = Val(Format(Now, "yymmddhhmm"))
                    !tran_textoel = ""
                    !tran_texto = ""
                    !tran_textosc10 = ""
                End If
                If ModificacionTx1 Then
                    !prc3 = Val(Format(Now, "yymmddhhmm"))
                    !tran_tx1 = ""
                    !tran_tx1SC10 = ""
                End If
                If ModificacionCodBar Then
                    !prc3 = Val(Format(Now, "yymmddhhmm"))
                    !tran_cb = ""
                    !tran_cbsc10 = ""
                End If
                If articulo.m_prc100g Then .Fields("PRC100G") = articulo.prc100g
                If articulo.m_presec Then !ning = articulo.Presec
                If articulo.m_merma Then !fcb = articulo.Merma
                If articulo.m_cba Then !art_cb = articulo.cba
                If articulo.m_sub Then !codi_sub = articulo.sub
                If articulo.m_fam Then !codi_fam = articulo.fam
                If Not UsaEuro Then
                    !precio = MiPrecio
                    If ModificacionPrecio Or IsNull(!Euros) Then !Euros = MiEuro
                Else
                    !Euros = MiEuro
                    If ModificacionPrecio Or IsNull(!precio) Then !precio = MiPrecio
                End If
                If articulo.m_iva Then !tipo_iva = articulo.iva
                If articulo.cad <> 0 Or lEsAho = True Then
                    !Caducidad = articulo.cad
                End If
                If articulo.m_tar Then !tara = articulo.tar
                If articulo.m_pre Then !pref = articulo.pre
                If articulo.m_etq Then !etq = articulo.etq
                If articulo.m_gpc Then !grupo_conserv = articulo.gpc
                '********************************
                ' Alcampo : en las modificaciones
                ' no se altera pesado/no pesado
                ' porque llegan datos de AS400
                ' a corregir por el usuario
                If Not Conf_Al Then
                    If MiWgh = True Then
                        !codi_pes = "W"
                    Else
                        !codi_pes = "U"
                    End If
                End If
                '/Alcampo
                '*********************
                If Balenv = False Then
                    !Balenv = False
                    Set Reg2 = Base.OpenRecordset _
                    ("select secc_maqui from seccion where borrado=false and codi_ident=" & !codi_ident)
                    Set Reg5 = Base.OpenRecordset _
                    ("select count(codigo) from articulo where borrado=false and secc_maqui=" & _
                    Reg2!secc_maqui & " and balenv=true")
                    
                    If Reg5.Fields(0) = 1 Then
                        Reg5.Close
                        Set Reg5 = Base.OpenRecordset _
                        ("select multiple,enviardatos,tran_sec,tran_secSC10 from seccion where borrado=false and secc_maqui=" & _
                        Reg2!secc_maqui)
                        With Reg5
                            If Not .EOF Then
                                .MoveFirst
                                Do Until .EOF
                                    If !multiple = 1 Then
                                        Edit_Record Reg5
                                        !multiple = -1
                                        If !enviardatos = "GA" Then !tran_secSC10 = ""
                                        If !enviardatos = "TODOS" Then !tran_sec = ""
                                        .Update
                                    End If
                                    If Not .EOF Then .MoveNext
                                Loop
                            End If
                        End With
                    End If
                                
                Else
                    !Balenv = True
                    Set Reg2 = Base.OpenRecordset _
                    ("select secc_maqui from seccion where borrado=false and codi_ident=" & !codi_ident)
                    Set Reg5 = Base.OpenRecordset _
                    ("select enviardatos,multiple,tran_sec,tran_secSC10 from seccion where borrado=false and secc_maqui=" & Reg2!secc_maqui)
                    With Reg5
                        If Not .EOF Then
                            .MoveFirst
                            Do Until .EOF
                            If !enviardatos <> "TODOS" Then
                                Edit_Record Reg5
                                If !multiple <> 1 Then
                                
                                    If !enviardatos = "GA" Then !tran_secSC10 = ""
                                    If !enviardatos = "SC10" Then !tran_sec = ""
                                End If
                                !multiple = 1
                                .Update
                            End If
                            If Not .EOF Then .MoveNext
                            Loop
                        End If
                    End With
                End If
                If Dir(App.Path & "\hiper") = "" Then
                    
                    For bucle = 0 To 20
                        Select Case bucle
                            Case 0 To 8
                                MiNombre = "des_plu" & Val(bucle + 1)
                            Case 9
                                MiNombre = "des_plu0"
                            Case 10
                                MiNombre = "des_plux"
                            Case 11 To 20
                                MiNombre = "des_plu" & Val(bucle)
                        End Select
                        If Not AHOLD Then
                            If articulo.m_txt(bucle) Then .Fields(MiNombre) = articulo.txt(bucle)
                        Else
                            If articulo.m_txt(bucle) And Trim(articulo.txt(bucle)) <> "" Then .Fields(MiNombre) = articulo.txt(bucle)
                        End If
                        If bucle < 10 Then
                            Select Case bucle
                                Case 0 To 8
                                    MiNombre = "tip_let" & Val(bucle + 1)
                                Case 9
                                    MiNombre = "tip_let0"
                            End Select
                            If Not AHOLD Then
                                If articulo.m_tlt(bucle) Then .Fields(MiNombre) = articulo.tlt(bucle)
                            Else
                                If articulo.m_tlt(bucle) And articulo.tlt(bucle) <> 0 Then .Fields(MiNombre) = articulo.tlt(bucle)
                            End If
                        End If
                    Next bucle
                
                End If
                '*****
                ' Tramos
                '*****
                If ChkTramo And !usatramos = False Then
                    !usatramos = True
                    !tran_plu = ""
                    ModificacionTramo = True
                    Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=true")
                    With RegTramo
                        If .EOF Then
                            .AddNew
                        Else
                            Edit_Record RegTramo
                        End If
                            !borrado = False
                            !tran_tramo = ""
                            !codigo = articulo.cod
                            If articulo.trm1 <> -1 And articulo.prc1 <> -1 Then
                                !tramo1 = articulo.trm1 / 1000
                                If UsaEuro Then
                                    !euros1 = articulo.prc1 / 100
                                    !precio1 = EUROtoPTA(articulo.prc1 / 100)
                                Else
                                    !euros1 = EUROtoPTA(articulo.prc1)
                                    !precio1 = articulo.prc1
                                End If
                                    If articulo.trm2 <> -1 And articulo.prc2 <> -1 Then
                                        !tramo2 = articulo.trm2 / 1000
                                        If UsaEuro Then
                                            !euros2 = articulo.prc2 / 100
                                            !precio2 = EUROtoPTA(articulo.prc2 / 100)
                                        Else
                                            !euros2 = PTAtoEURO(articulo.prc2)
                                            !precio2 = articulo.prc2
                                        End If
                                    Else
                                        !tramo2 = -1
                                        !precio2 = -1
                                        !euros2 = -1
                                    End If
                            End If
                        .Update
                    End With
                End If
                If ChkTramo And !usatramos = True Then
                    !usatramos = True
                    Set RegTramo = Base.OpenRecordset _
                    ("select * from tramos where borrado=false and codigo=" & !codigo)
                    With RegTramo
                        If Not .EOF Then
                            Edit_Record RegTramo
                            !tran_tramo = ""
                            !codigo = articulo.cod
                            If !precio1 <> articulo.prc1 Or !tramo1 <> articulo.trm1 Then
                                
                                If UsaEuro Then
                                    !euros1 = articulo.prc1 / 100
                                    !precio1 = EUROtoPTA(articulo.prc2 / 100)
                                Else
                                    !euros1 = PTAtoEURO(articulo.prc2)
                                    !precio1 = articulo.prc1
                                End If
                                !tramo1 = articulo.trm1 / 1000
                                RegArticulo!tran_plu = ""
                                ModificacionTramo = True
                            End If
                            If (articulo.prc2 = -1 And !precio2 >= 0) Then
                                !euros2 = -1
                                !precio2 = -1
                                !tramo2 = -1
                                RegArticulo!tran_plu = ""
                                ModificacionTramo = True
                            Else
                                If articulo.prc2 <> -1 And _
                                (!precio2 <> articulo.prc2 Or !tramo2 <> articulo.trm2) Then
                                    If UsaEuro Then
                                        !euros2 = articulo.prc2 / 100
                                        !precio2 = EUROtoPTA(articulo.prc2 / 100)
                                    Else
                                        !euros2 = PTAtoEURO(articulo.prc2)
                                        !precio2 = articulo.prc2
                                    End If
                                    !tramo2 = articulo.trm2 / 1000
                                    RegArticulo!tran_plu = ""
                                    ModificacionTramo = True
                                End If
                            End If
                        .Update
                        End If
                    End With
                End If
                If ChkTramo = False And !usatramos = True Then
                    !usatramos = False
                    !tran_plu = ""
                    Set RegTramo = Base.OpenRecordset _
                    ("select * from tramos where borrado=false and codigo=" & !codigo)
                    With RegTramo
                        If Not .EOF Then
                            Edit_Record RegTramo
                            !borrado = True
                            .Update
                            ModificacionTramo = True
                        End If
                    End With
                End If
                
                'If lDarBajaAlta = True Then
                
                '    '.Fields("des_plu1") = articulo.txt(0)
                '    .Fields("des_plu2") = articulo.txt(1)
                '    .Fields("des_plu3") = articulo.txt(2)
                '    .Fields("des_plu4") = articulo.txt(3)
                '    .Fields("des_plu5") = articulo.txt(4)
                '    .Fields("des_plu6") = articulo.txt(5)
                '    .Fields("des_plu7") = articulo.txt(6)
                '    .Fields("des_plu8") = articulo.txt(7)
                '    .Fields("des_plu9") = articulo.txt(8)
                '    .Fields("des_plu0") = articulo.txt(9)
                '    .Fields("des_plux") = articulo.txt(10)
                '    .Fields("des_plu11") = articulo.txt(11)
                '    .Fields("des_plu12") = articulo.txt(12)
                '    .Fields("des_plu13") = articulo.txt(13)
                '    .Fields("des_plu14") = articulo.txt(14)
                '    .Fields("des_plu15") = articulo.txt(15)
                '    .Fields("des_plu16") = articulo.txt(16)
                '    .Fields("des_plu17") = articulo.txt(17)
                '    .Fields("des_plu18") = articulo.txt(18)
                '    .Fields("des_plu19") = articulo.txt(19)
                '    .Fields("des_plu20") = articulo.txt(20)
                '
                '    .Fields("tip_let1") = articulo.tlt(0)
                '    .Fields("tip_let2") = articulo.tlt(1)
                '    .Fields("tip_let3") = articulo.tlt(2)
                '    .Fields("tip_let4") = articulo.tlt(3)
                '    .Fields("tip_let5") = articulo.tlt(4)
                '    .Fields("tip_let6") = articulo.tlt(5)
                '    .Fields("tip_let7") = articulo.tlt(6)
                '    .Fields("tip_let8") = articulo.tlt(7)
                '    .Fields("tip_let9") = articulo.tlt(9)
                '    .Fields("tip_let0") = articulo.tlt(10)
                '    .Fields("tip_letx") = articulo.tlt(11)
                '
                'End If
                
                .Update
                
                
                If articulo.sub <> "0" Then
                    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
                    articulo.sec & " and codi_sub=" & articulo.sub
                    With RegSubsec
                    If .EOF Or .NoMatch Then
                        .AddNew
                        !codi_sub = articulo.sub
                        !codi_ident = articulo.sec
                        '!secc_maqui = Reg2!secc_maqui
                        !borrado = False
                        !tran_sub = ""
                        .Update
                    End If
                    End With
                    
                End If
                If articulo.fam <> "0" Then
                    Set Reg3 = Base.OpenRecordset _
                    ("select codi_fam,codi_sub,codi_ident,secc_maqui,borrado,tran_fam from familia where borrado=false and codi_ident=" & _
                    articulo.sec & " and codi_sub=" & articulo.sub & _
                    " and codi_fam=" & articulo.fam)
                    With Reg3
                    If .EOF Then
                        .AddNew
                        !codi_fam = articulo.fam
                        !codi_sub = articulo.sub
                        !codi_ident = articulo.sec
                        !secc_maqui = Reg2!secc_maqui
                        !borrado = False
                        !tran_fam = ""
                        .Update
                    End If
                    End With
                    
                    
                End If
            
      End If
    End With
    
    If ModificacionTramo Or ModificacionBasica Or Modificaciontexto Or _
    ModificacionTx1 Or ModificacionCodBar Or ModificacionPrecio Then
        
        Retorno = 1
    Else
        Retorno = 4
    End If
fin:
Select Case Retorno
    Case 1
        If MiAlta Then
            frmimportar.Lista.AddItem "Cód : " & Format(articulo.cod, "000000") & "    Plu : " & Format(articulo.plu, "0000") & "    ALTA"
        Else
            frmimportar.Lista.AddItem "Cód : " & Format(articulo.cod, "000000") & "    Plu : " & Format(articulo.plu, "0000") & "    MODIFICACION"
        End If
    Case 4
        frmimportar.Lista.AddItem "Cód : " & Format(articulo.cod, "000000") & "    Plu : " & Format(articulo.plu, "0000") & "    SIN MODIFICACION"
    Case Else
        frmimportar.Lista.AddItem "Cód : " & Format(articulo.cod, "000000") & "    Plu : " & Format(articulo.plu, "0000") & "    RECHAZO"
End Select
If LogActivo Then
    If Retorno = 1 And SeCambioPrecio Then
        Set RstRechazos = Base.OpenRecordset("select * from log")
        With RstRechazos
            .AddNew
            If UsaEuro Then
                !Texto = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(articulo.cod, "000000") & "," & _
                "000" & "," & Format(articulo.sec, "000") & "," & Format(articulo.prc * 100, "0000000") & "," & Trim(articulo.txt(0)) & "," & Format(Existe_Mostrador(articulo.sec), "00")
            Else
                !Texto = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(articulo.cod, "000000") & "," & _
                "000" & "," & Format(articulo.sec, "000") & "," & Format(articulo.prc * 10 ^ Decimales, "0000000") & "," & Trim(articulo.txt(0)) & "," & Format(Existe_Mostrador(articulo.sec), "00")
            End If
            !codi_ident = articulo.sec
            !codi_sub = articulo.sub
            !codigo = articulo.cod
            !plu = PluBuffer
            If UsaEuro Then
                !precio = articulo.prc / 100
            Else
                !precio = !precio / (10 ^ Decimales)
            End If
            !codi_fam = articulo.fam
            !Caducidad = articulo.cad
            !tara = articulo.tar
            !des_plu1 = articulo.txt(0)
            !fecha = Date
            If MiWgh = True Then
                !tipo_ven = "W"
            Else
                !tipo_ven = "U"
            End If
            If Alcampo Then
                !comentario = "cambio precio AS400"
            Else
                !comentario = "Cambio Precio"
            End If
            .Update
        End With
    End If
          
    If (Retorno <> 0 And Retorno <> 4 And Retorno <> 1) Then
        Set RstRechazos = Base.OpenRecordset("select * from rechazos")
        With RstRechazos
            .AddNew
            !codi_ident = articulo.sec
            !codi_sub = articulo.sub
            !codigo = articulo.cod
            !plu = articulo.plu
            If UsaEuro Then
                !precio = !precio / 100
            Else
                !precio = !precio / (10 ^ Decimales)
            End If
            !tipo_ven = MiWgh
            !codi_fam = articulo.fam
            !Caducidad = articulo.cad
            !tara = articulo.tar
            !des_plu1 = articulo.txt(0)
            If Retorno <> 1 Then
                Select Case Retorno
                Case 2
                    !comentario = "Intento de cambio de PLU del artículo"
                Case 3
                    !comentario = "No existe sección"
                Case 5
                    !comentario = "Código de Barras con menos de 12 caracteres"
                Case 14
                    !comentario = "Falta código de Artículo"
                Case 15
                    !comentario = "Falta número de Mostrador"
                Case 16
                    !comentario = "Error en precios por tramos"
                Case 17
                    !comentario = "Base de datos de Tramos llena"
                Case 21
                    !comentario = "PLU repetido"
                Case 22
                    !comentario = "Datos Incorrectos"
                End Select
                .Update
            Else
                If Not Alcampo Then
                    !comentario = "Modificación de Precio"
                Else
                    !comentario = "Cambio Precio AS400"
                End If
            End If
        End With
    End If
End If
If RESPONSE And Retorno <> 2 Then Print #desc, Format(Retorno, "00")
Alta_Articulo = Retorno
End Function
Public Sub Borrar_Articulo(Balenv As Boolean, Base As Database, Optional desc As Integer, Optional NoText As Boolean, Optional PorPlu As Boolean)

    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim RegTeclas As Recordset
    
    Set Registro = Base.OpenRecordset _
    ("select * from seccion where borrado=false and codi_ident=" & articulo.sec)
    If Registro.EOF Then
        'Base.Close
        If RESPONSE And (Not NoText) Then Print #desc, "02"
        Exit Sub
    End If
    If PorPlu Then
        Set Registro = Base.OpenRecordset _
        ("select * from articulo where borrado=false and plu=" & articulo.plu & " and codi_ident=" & articulo.sec)
    Else
        Set Registro = Base.OpenRecordset _
        ("select * from articulo where borrado=false and codigo=" & articulo.cod)
    End If
    With Registro
        If Not .EOF Then
            Edit_Record Registro
            frmimportar.Lista.AddItem "Cód : " & Format(articulo.cod, "000000") & "    Plu : " & Format(articulo.plu, "0000") & "    BAJA"
            !borrado = True
            !tran_plu = ""
            !tran_pluSC10 = ""
            .Update
            Set Reg2 = Base.OpenRecordset _
            ("select * from seccion where borrado=false and codi_ident=" & !codi_ident)
            Set Reg3 = Base.OpenRecordset _
            ("select * from articulo where borrado=false and balenv=true and secc_maqui=" & Reg2!secc_maqui)
            If Reg3.EOF Then
                Reg3.Close
                Set Reg3 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & Reg2!secc_maqui)
                With Reg3
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            If !multiple = 1 Then
                                Edit_Record Reg3
                                !multiple = -1
                                If !enviardatos = "GA" Then !tran_secSC10 = ""
                                If !enviardatos = "SC10" Then !tran_sec = ""
                                .Update
                                Set RegTeclas = Base.OpenRecordset _
                                ("select * from teclas where codigo=" & articulo.cod)
                                With RegTeclas
                                    If Not .EOF Then
                                        .MoveFirst
                                        Do Until .EOF
                                            Edit_Record RegTeclas
                                            !borrado = True
                                            !tran_tec = ""
                                            !tran_tecSC10 = ""
                                            .Update
                                        Loop
                                    End If
                                End With
                            End If
                            If Not .EOF Then .MoveNext
                        Loop
                    End If
                End With
            End If
                    
                        
            'Base.Close
            
        Else
            'Base.Close
        End If
    End With
    If RESPONSE And (Not NoText) Then Print #desc, "01"
End Sub
Public Sub Borrar_todos_Articulo(Base As Database)

    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Reg3 As Recordset
    Dim RegTeclas As Recordset
    
    Set Registro = Base.OpenRecordset _
    ("select * from articulo")

    With Registro
        If Not .EOF Then
            
            .MoveFirst
            Do Until .EOF
            
                frmimportar.Lista.AddItem "Cód : " & Format(!codigo, "000000") & "    Plu : " & Format(!plu, "0000") & "    BAJA"
                If Not .EOF Then .Delete
                If Not .EOF Then .MoveNext


            
            Loop
            
            Set RegTeclas = Base.OpenRecordset _
            ("select * from teclas")
            With RegTeclas
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        If Not .EOF Then .Delete
                        If Not .EOF Then .MoveNext
                    
                    Loop
                End If
            End With
        
        End If
        
    End With
    'If RESPONSE And (Not NoText) Then Print #desc, "01"
End Sub

Public Function Alta_Teclas(LaTecla As DB_Teclas, Base As Database) As Integer
    '********************
    ' Sólo puede enviarse por plu + sec o por codigo
    ' pero no acepta los tres valores simultáeos
    '********************
    'Dim Base As Database
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Retorno As Integer
    '*******************
    ' Condiciones de error
    '*******************

    If LaTecla.CodArticulo = 0 Then
        If LaTecla.PluArticulo = 0 Or LaTecla.SecArticulo = 0 Then
            Retorno = 3
            'Base.Close
            GoTo fin
        Else
            Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & _
            LaTecla.SecArticulo & " and plu=" & LaTecla.PluArticulo & " and borrado=false")
            If Registro.EOF Then
                'Base.Close
                Retorno = 2
                GoTo fin
            Else
                LaTecla.CodArticulo = Registro!codigo
            End If
        End If
    Else
        If LaTecla.PluArticulo <> 0 Or LaTecla.SecArticulo <> 0 Then
            'Base.Close
            Retorno = 18
            GoTo fin
        Else
            Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & _
            LaTecla.CodArticulo & " and borrado=false")
            If Registro.EOF Then
                'Base.Close
                Retorno = 2
                GoTo fin
            Else
                LaTecla.PluArticulo = Registro!plu
                LaTecla.SecArticulo = Registro!codi_ident
            End If
        End If
    End If
    '*******************
    If LaTecla.Destino_NBalanza = 0 Then
        Set Registro = Base.OpenRecordset("select * from teclas where codi_tec=" & LaTecla.Destino_CodiTec & _
        " and codi_ident=" & LaTecla.Destino_NMostrador & " and tabla=" & LaTecla.Destino_Tabla & _
        " and isnull(numero_eqp) and borrado=false")
    Else
        Set Registro = Base.OpenRecordset("select * from teclas where codi_tec=" & LaTecla.Destino_CodiTec & _
        " and numero_eqp=" & LaTecla.Destino_NBalanza & " and tabla=" & LaTecla.Destino_Tabla & _
        " and borrado=false")
    End If
    With Registro
        If .EOF Then
            Retorno = 0
            ' dato nuevo
            .AddNew
        Else
            Retorno = 1
            Edit_Record Registro
        End If
        Set Reg2 = Base.OpenRecordset("select * from articulo where codigo=" & _
        LaTecla.CodArticulo & " and borrado=false")
        If LaTecla.Destino_NBalanza = 0 Then
            !codi_ident = LaTecla.Destino_NMostrador
        Else
            !numero_eqp = LaTecla.Destino_NBalanza
        End If
        !secc_maqui = Reg2!secc_maqui
        !codigo = Reg2!codigo
        !plu = Reg2!plu
        !desc_plu = Reg2!des_plu1
        !secc_eqp = 0
        !Tabla = LaTecla.Destino_Tabla
        !codi_tec = LaTecla.Destino_CodiTec
        !tran_tec = ""
        !tran_tecSC10 = ""
        !borrado = False
        .Update
    End With
    'Base.Close
fin:
    Alta_Teclas = Retorno
        
End Function
Public Function Baja_Teclas(LaTecla As DB_Teclas, Base As Database) As Integer
    
  
    Dim Registro As Recordset
    Dim Reg2 As Recordset
    Dim Retorno As Integer
    
    '*******************
    ' Condiciones de error
    '*******************
    If LaTecla.Destino_NBalanza <> 0 And LaTecla.Destino_NMostrador <> 0 Then
        Retorno = 3
        GoTo fin
    End If
    '*******************
    If TypeName(Base) = "Nothing" Then Set Base = OpenDatabase(Base_General)
    
    Set Registro = Base.OpenRecordset("select * from teclas where codi_ident=" & _
    LaTecla.Destino_NMostrador & " and codigo=" & articulo.cod & _
    " and isnull(numero_eqp) and borrado=false")
    
    With Registro
        If .EOF Then
            Retorno = 4
        
            GoTo fin
        Else
            .MoveFirst
            Do Until .EOF
                Edit_Record Registro
                !borrado = True
                !tran_tec = ""
                !tran_tecSC10 = ""
                !codigo = 0
                .Update
                If Not .EOF Then .MoveNext
            Loop
            Retorno = 0
        End If
    End With
    
fin:
    
    Baja_Teclas = Retorno
        
End Function


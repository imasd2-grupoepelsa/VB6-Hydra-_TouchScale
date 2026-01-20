Attribute VB_Name = "s"
Option Explicit
Global SUPECO As Boolean
Public Codigos() As Long
Public NCodigos As Long
Public Mostradores() As Integer
Public Secciones() As Integer
Public NMostradores As Integer
Public Buscado As Boolean
Public lEsTest As Boolean
Public lCOCOBIO As Boolean
Public lFamPluPos As Boolean
Public lTransPluAsc As Boolean
Public lEsGestMag As Boolean
Public lBelros As Boolean

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
Public Function Busca_Codigos_Sec_plu(ByVal nS As Integer, ByVal nPLu As Integer) As Long
    Dim MyBase As Database
    Dim Registro As Recordset
    Dim nResp As Long
    Set MyBase = OpenDatabase(Base_General)
    Set Registro = MyBase.OpenRecordset("select * from articulo where borrado=false and plu=" & CStr(nPLu) & " and codi_ident=" & CStr(nS))
    If Not Registro.EOF Then
        nResp = Registro.Fields("codigo")
    Else
        nResp = 0
    End If
    Registro.Close
    MyBase.Close
    Set Registro = Nothing
    Set MyBase = Nothing
    'CadenadeLog "Busca_codigos_sec_plu:" & "Sec.:" & CStr(nS) & " PLU:" & CStr(nPLu) & " Codigo=" & CStr(nResp)
    Busca_Codigos_Sec_plu = nResp
End Function

Public Function Existe_Mostrador(NumMos As Long) As Integer
    Dim Resp As Integer
    Dim Bucle As Integer
    Resp = -1
    For Bucle = 0 To NMostradores - 1
        If NumMos = Mostradores(Bucle) Then
            Resp = Secciones(Bucle)
            Exit For
        End If
    Next Bucle
    Existe_Mostrador = Resp
End Function
Public Function Existe_Codigo(NumCod As Long) As Long
    Dim Resp As Long
    Dim Bucle As Integer
    Resp = -1
    For Bucle = 0 To NCodigos - 1
        If NumCod = Codigos(Bucle) Then
            Resp = Codigos(Bucle)
            Exit For
        End If
    Next Bucle
    Existe_Codigo = Resp
End Function

Public Function Alta_Articulo(Balenv As Boolean, ByVal Base As Database, ByVal RegArticulo As Recordset, Optional desc As Integer) As Integer
    Dim ModificacionTramo As Boolean
    Dim ModificacionBasica As Boolean
    'Dim Modificaciontexto As Boolean
    'Dim ModificacionTx1 As Boolean
    'Dim ModificacionCodBar As Boolean
    'Dim ModificacionPrecio As Boolean
    Dim Bucle As Integer
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
    'Dim lEsGestMag As Boolean
    Dim nTmpFam As Integer
    
    'CadenadeLog "Alta_Articulo in..."
    
    'datos gestmag
    'Código de Artículo
    'código de Barras
    'descriptivo principal
    'tipo de venta.
    'precio unitario....
    ' En modificación resto de datos no alterar...
    
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
    'hydratouch 1.0.10
    'If articulo.fam = -1 Then articulo.fam = 0
    '''''''''''
    If (articulo.fam = -1) Or (articulo.m_fam = False) Then
        articulo.fam = 1
    End If
    'belros
    If articulo.sec = 0 Then articulo.sec = 1
    '''''''''''
    If articulo.pre = -1 Then articulo.pre = 0
    If articulo.etq = -1 Then articulo.etq = 0
    If articulo.gpc = -1 Then articulo.gpc = 0
    
    'CadenadeLog "Posición 0.1"
    
    '*****************
    ' Pesado/No pesado
    ' se acepta para pesado :
    ' "W","P","0"
    ' para no pesado :
    ' "U","1"
    '*****************
    MiWgh = False
    If Dir(App.Path & "\plrm.txt") <> "" Then
        If articulo.wgh = "1" Then
            MiWgh = True
        Else
            MiWgh = False
        End If
    Else
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
    End If
    
    'CadenadeLog "Posición 0.2"
    
    'MsgBox "sobrepasada verificación tipo venta..."
    If articulo.trm1 <> 0 Or articulo.trm2 <> 0 Or _
    articulo.prc2 <> 0 Or articulo.prc1 <> 0 Then
        CadenadeLog "Posición 5...Tramo. TRM1: " & CStr(articulo.trm1) & " TRM2: " & CStr(articulo.trm2) & " PRC1:" & CStr(articulo.prc1) & " PRC2:" & CStr(articulo.prc2)
    
        ChkTramo = True
        CadenadeLog "activado chktramo..."
    Else
        ChkTramo = False
    End If
    
    
    'CadenadeLog "Posición 0.3"
        
    'If articulo.trm1 <> -1 Or articulo.trm2 <> -1 Or _
    'articulo.prc2 <> -1 Or articulo.prc1 <> -1 Then
    '    ChkTramo = True
    'Else
    '    ChkTramo = False
    'End If
    
    
    'If (articulo.trm1 < articulo.trm2) And (articulo.prc2 < articulo.prc1) And (articulo.prc1 < articulo.prc) Then
    '    ChkTramo = True
    'Else
    '    ChkTramo = False
    'End If
        
    'c2f 26/07/05
    'If Trim(articulo.cba) <> "" Then
    '    If Len(articulo.cba) <> 12 Then
    '        Retorno = 5
    '        GoTo fin
    '    End If
    'End If
    'c2f 18/12/06 --> Codisys se agregan 0s a la derecha para completar 12 caracteres...
    'If Len(Trim(articulo.cba)) < 12 And Len(Trim(articulo.cba)) = 8 Then
    '    articulo.cba = Mid(Trim(articulo.cba), 1, 7)
    '    articulo.cba = Mid("000000000000", 1, 12 - Len(Trim(articulo.cba))) & Trim(articulo.cba)
    'End If
    articulo.cba = Trim(articulo.cba)
    If Len(articulo.cba) = 12 Then
        articulo.cba = articulo.cba & Chr(48 + DigitoControl(articulo.cba))
    End If
    'MsgBox "Tramos y CB..."
    
    '''''''''''''
    If RechazarPlu0 And articulo.prc = 0 And (lEsAho = False) And UCase(FicheroImport) <> UCase("baletl.dat") Then

        Retorno = 50
        'CadenadeLog "Retorno:" & CStr(Retorno)
        
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
        ''''''''''BELROS
        Retorno = 4
        GoTo fin
        '''''''''''
        If Dir(App.Path & "\tiendamt.txt") = "" Then
            Retorno = 14
            GoTo fin
        Else
            articulo.cod = Busca_Codigos_Sec_plu(articulo.sec, articulo.plu)
            If articulo.cod = 0 Then
                Retorno = 14
                GoTo fin
            End If
        End If
    End If
    
    'CadenadeLog "Posición 0.4"
    
    'If articulo.sec = 0 And (lEsAho = False) And UCase(FicheroImport) <> UCase("baletl.dat") Then
    '    If Dir(App.Path & "\sec0a1.txt") <> "" Or lEsGestMag Then
    '        If Dir(App.Path & "\sec0a1.txt") <> "" Then
    '            articulo.sec = 1
    '        End If
    '    Else
    '       Retorno = 15
    '       GoTo fin
    '    End If
    'End If
    If (VarPublics.C9Cod <> 0) And (articulo.sec > 1) Then   '///cas.v118
        Retorno = 15
        GoTo fin
    End If
    If articulo.Merma > 99 Or articulo.Presec > 99 Then
        Retorno = 22
        GoTo fin
    End If
    'If ChkTramo = True Then
    '    If articulo.trm1 = -1 And articulo.prc1 = -1 Then
    '        Retorno = 16
    '        GoTo fin
    '    End If
    '    If (articulo.trm2 <> -1 And articulo.prc2 = -1) Or _
    '    (articulo.prc2 <> -1 And articulo.trm2 = -1) Then
    '        Retorno = 16
    '        GoTo fin
    '    End If
    '    If (articulo.trm1 <> -1 And articulo.prc1 = -1) Or _
    '    (articulo.prc1 <> -1 And articulo.trm1 = -1) Then
    '        Retorno = 16
    '        GoTo fin
    '    End If
    'End If
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
    
    'If TypeName(RegSubsec) = "Nothing" Then
    '    Set RegSubsec = Base.OpenRecordset("select codi_ident,borrado,codi_sub,tran_sub from subsec")
    'End If
    'If ChkTramo Then
    '    Set RegTramo = Base.OpenRecordset("select count(codigo) from tramos")
    '    With RegTramo
    '        If .Fields(0) >= 400 Then
    '            Retorno = 17
    '
    '            GoTo fin
    '        End If
    '    End With
    'End If
    
    If articulo.fam = Val(FamiliaVacuno) Then
        articulo.etq = Val(EtiquetaEsp)
    End If
    
    Retorno = 4
  
    'CadenadeLog "Posición 0.5"
    
    If lEsGestMag = False Then
    
        If Existe_Mostrador(articulo.sec) < 0 And UCase(FicheroImport) <> UCase("baletl.dat") Then
            'CadenadeLog "No existe sección del Artículo"
            Retorno = 3
            GoTo fin
        End If
    
        'CadenadeLog "Familia artículo:" & CStr(articulo.fam)
        
        'If FamiliasFijas Then
        '    Set RegFam = Base.OpenRecordset("select codi_fam from fam_code where codi_fam=" & articulo.fam)  ' 1.0. 3& " and codi_ident=" & articulo.sec)
        '    If RegFam.EOF Then
        '        Retorno = 3
        '        GoTo fin
        '    End If
        '    RegFam.Close
        '    Set RegFam = Nothing
        'End If
        
        'If SubseccionesFijas Then
        '    Set RegFam = Base.OpenRecordset("select codi_sub from sub_code where codi_sub=" & articulo.sub & " and codi_ident=" & articulo.sec)
        '    If RegFam.EOF Then
        '        Retorno = 3
        '        GoTo fin
        '    End If
        'End If
    
    End If
    
    ''''''''''''''''''''''''''''''''''''
    'Carrefour
    ''''''''''''''''''''''''''''''''''''
    If Dir(App.Path & "\hiper") <> "" Then
    
        '********************
        ' comprueba plu repetido
        '********************
'CadenadeLog "<-01->"
        If articulo.plu <> 0 Then
            'posible hueco
            nPosiPlu = 0
            For Bucle = 1 To 9999
                Set RegPlu = Base.OpenRecordset _
                ("select plu from articulo where secc_maqui=" & _
                Existe_Mostrador(articulo.sec) & " and plu=" & Bucle & " and borrado=false")
                If RegPlu.EOF Then
                   nPosiPlu = Bucle
                   Bucle = 10000
                End If
            Next Bucle
'CadenadeLog "<-02->"
            Set RegPlu = Base.OpenRecordset _
            ("select * from articulo where secc_maqui=" & _
            Existe_Mostrador(articulo.sec) & " and plu=" & articulo.plu & " and codigo<>" & articulo.cod & " and borrado=false")
            If Not RegPlu.EOF Then 'se cambia el PLU a PosiPlu
                
                RegPlu.Edit
                RegPlu!plu = nPosiPlu
                RegPlu!PRC3 = Val(Format(Now, "yymmddhhmm"))
                RegPlu!tran_en = "" '
                RegPlu!tran_el = "" '
                RegPlu!tran_plu = "" '
                RegPlu!tran_pluSC10 = "" '
                RegPlu!tran_textoel = ""  '
                RegPlu!tran_texto = "" '
                RegPlu!tran_textosc10 = "" '
                RegPlu!tran_tx1 = "" '
                RegPlu!tran_tx1SC10 = "" '
                RegPlu!tran_cb = "" '
                RegPlu!tran_cbsc10 = "" '
                RegPlu!tran_ep = Space(60) '
                RegPlu.Update
            End If
        End If
'CadenadeLog "<-03->"
    End If
    ''''''''''''''''''''''''''''''''''''
    'CadenadeLog "Posición 0.6"
    
    With RegArticulo
        ' Nuevo dato
        If Existe_Codigo(articulo.cod) < 0 Then
            '********************
            ' comprueba plu repetido
            '********************
            
            'GestMag
            If lEsGestMag Then 'se dá de alta en la sección más baja
                'Retorno = 14
                'GoTo fin
                Set RegPlu = Base.OpenRecordset("select * from seccion order by secc_maqui")
                If Not RegPlu.EOF Then
                    RegPlu.MoveFirst
                    articulo.sec = RegPlu.Fields("secc_maqui")
                    
                Else
                    Retorno = 14
                    GoTo fin
                End If
                RegPlu.Close
            End If
            
            If (lEsAho) And UCase(FicheroImport) = UCase("baletl.dat") Then
                Retorno = 14
                GoTo fin
            End If
'CadenadeLog "<-04->"
            'If articulo.plu <> 0 Then
            '    Set RegPlu = Base.OpenRecordset _
            '    ("select plu from articulo where secc_maqui=" & _
            '    Existe_Mostrador(articulo.sec) & " and plu=" & articulo.plu & " and borrado=false")
            '
            '    If Not RegPlu.EOF Then
            '        If Dir(App.Path & "\transpluasc.exe") <> "" Then
            '            articulo.plu = 0
            '        Else
            '
            '            Retorno = 21
            '            GoTo fin
            '
            '        End If
            '    End If
            'End If
            '********************
            
            'CadenadeLog "Posición 0.7"
            'hiperusera, si ya existe el plu de un articulo nuevo poner a cero el plu del que se va a dar de alta... se anula de momento
            'If articulo.plu <> 0 And ltranspluasc Then
            '    Set RegPlu = Base.OpenRecordset _
            '    ("select plu from articulo where secc_maqui=" & _
            '    Existe_Mostrador(articulo.sec) & " and plu=" & articulo.plu & " and borrado=false")
            '    If Not RegPlu.EOF Then
            '        articulo.plu = 0
            '    End If
            'End If
            '''''''''''''''''''''''''''''''
            
'CadenadeLog "<-05->"
            .AddNew
            NCodigos = NCodigos + 1
            ReDim Preserve Codigos(NCodigos)
            Codigos(NCodigos - 1) = articulo.cod
            MiAlta = True
            Retorno = 0
            ModificacionBasica = True
            If articulo.m_onkey Then
                If articulo.onkey = 1 Then !onkey = True
                If articulo.onkey = 0 Then !onkey = False
            Else
                !onkey = True
            End If
            !PRC3 = Val(Format(Now, "yymmddhhmm"))
            !tran_en = ""
            !tran_el = ""
            'eliminado !tran_ep = Space(60)
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
            !plu = articulo.plu
            !codi_ident = articulo.sec
            !pref = articulo.pre
            !etq = articulo.etq
            !grupo_conserv = articulo.gpc
            !fcb = articulo.Merma
            !ning = articulo.Presec
            !tip_letx = 0
            .Fields("prc100g") = articulo.prc100g
            !secc_maqui = Existe_Mostrador(articulo.sec)
            If articulo.m_image Then
                'CadenadeLog "IMAGEN:" & articulo.imagen
                'articulo.imagen = separaFich(articulo.imagen)
                'CadenadeLog "IMAGEN:" & articulo.imagen
                If Trim(articulo.imagen) = "" Then
                    articulo.imagen = "" '"#FF646464"
                End If
                !imagen = Trim(articulo.imagen)
            Else
                !imagen = "" '"#FF646464"
            End If
            'If articulo.m_image Then
            '    articulo.m_image = False
            '    If IsNumeric(Val(articulo.imagen)) Then
            '        !imagen = APaleta(Val(articulo.imagen))
            '    Else
            '        If Trim(articulo.imagen) <> "NO IMAGE" Then
            '            If Trim(articulo.imagen) <> "" Then
            '                If Dir(Trim(articulo.imagen)) <> "" Then
            '                    FileCopy Trim(articulo.imagen), App.Path & "\images\items\" & CStr(articulo.cod)
            '                    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
            '                Else
            '                    If Dir(App.Path & "\itmimg.jpg") <> "" Then
            '                        FileCopy App.Path & "\itmimg.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
            '                        !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
            '                    Else
            '                        !imagen = "#FF646464"
            '                    End If
            '                End If
            '            Else
            '                If Dir(App.Path & "\itmimg.jpg") <> "" Then
            '                    FileCopy App.Path & "\itmimg.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
            '                    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
            '                Else
            '                    !imagen = "#FF646464"
            '                End If
            '            End If
            '        'Else
            '        '    FileCopy App.Path & "\vacio.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
            '        '    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
            '        End If
            '    End If
            'End If
            'ScalePos --> Familia sin sección
            'If articulo.sec = 0 Then articulo.plu = 0
            '''''''''''''''''''''''''''''''''
            If lFamPluPos Then
            !posicion = articulo.plu
            Else
            !posicion = 0
            End If
'CadenadeLog "<-06->"
            'CadenadeLog "Posición 0.8"
                
            'c2f ahorramas trabajará tanto para SC10 como para EuroScale con PLU=0
            'c2f 2.0.21
            'If articulo.plu = 0 And Dir(App.Path & "\hiper") = "" And Not (lEsAho) And Not (lEsGestMag) Then
            '    Set Reg3 = Base.OpenRecordset _
            '    ("select max(plu) from articulo where borrado=false and secc_maqui=" & Existe_Mostrador(articulo.sec))
            '    If Not IsNull(Reg3.Fields(0)) And Not Reg3.EOF Then
            '        If Reg3.Fields(0) < 9999 Then
            '            !plu = Reg3.Fields(0) + 1
            '            PluBuffer = Reg3.Fields(0) + 1
            '        Else
            '
            '            For bucle = 1 To 9999
            '                'RegArticulo.Filter = "borrado = False And secc_maqui = """ _
            '                '& Reg2!secc_maqui & " and plu=" & bucle
            '                RegArticulo.Filter = "borrado = False And secc_maqui = " _
            '                & Existe_Mostrador(articulo.sec) & " and plu=" & bucle
            '
            '                Set Reg4 = RegArticulo.OpenRecordset
            '                If Reg4.EOF Then
            '                    !plu = bucle
            '                    PluBuffer = bucle
            '                    Exit For
            '                End If
            '            Next bucle
            '        End If
            '    Else
            '        !plu = 1
            '        PluBuffer = 1
            '    End If
            'Else
            '    If lEsAho Or lEsGestMag Then 'c2f 2.0.20
            '        !plu = 0
            '    Else
            '        !plu = articulo.plu
            '    End If
            '    PluBuffer = articulo.plu
            'End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Trim(articulo.cba) <> "" Then
                !art_cb = articulo.cba
            Else
                .Fields("art_cb") = ""
            End If
            !codi_sub = articulo.sub
            'If articulo.fam <> 0 Then 'c2f/caspiunza
                !codi_Fam = articulo.fam
            'Else
            '    !codi_fam = 0
            'End If
            !precio = MiPrecio
            !Euros = MiEuro
            !Caducidad = articulo.cad
            '********************************
            ' datos de caducidad programados
            ' en fichero
            '
'CadenadeLog "<01>"
            If articulo.cad = 0 Then
                If DataCaducidad.Numero > 0 Then
                    For Bucle = 0 To DataCaducidad.Numero - 1
                        If DataCaducidad.Seccion(Bucle) = articulo.sec Then
                            'If DataCaducidad.Subseccion(bucle) = articulo.sub Then
                            '    If DataCaducidad.Familia(bucle) = articulo.fam Then
                                    !Caducidad = DataCaducidad.Caducidad(Bucle)
                            '    End If
                            'End If
                        End If
                    Next Bucle
                End If
            End If
            '********************************
            !tara = articulo.tar
            !tara_envasado = articulo.tar
            'If Dir(App.Path & "\bdp.txt") = "" Then
                If MiWgh = True Then
                    !codi_pes = "W"
                Else
                    !codi_pes = "U"
                End If
            'Else
            '    !codi_pes = "W"
            'End If
            
            'prueba hiperusera se anula...
            'If Balenv = False Then
            '    !Balenv = False
            'Else
            '    !Balenv = True
            '    Set Reg5 = Base.OpenRecordset _
            '    ("select enviardatos,multiple,tran_secSC10,tran_sec from seccion where borrado=false and secc_maqui=" & Existe_Mostrador(articulo.sec))
            '    With Reg5
            '        If Not .EOF Then
            '            .MoveFirst
            '            Do Until .EOF
            '            If !enviardatos <> "TODOS" Then
            '                Edit_Record Reg5
            '                If !multiple <> 1 Then
            '
            '                    If !enviardatos = "GA" Then !tran_secSC10 = ""
            '                    If !enviardatos = "SC10" Then !tran_sec = ""
            '                End If
            '                !multiple = 1
            '                .Update
            '            End If
            '            If Not .EOF Then .MoveNext
            '            Loop
            '        End If
            '    End With
            '    Reg5.Close
            '    Set Reg5 = Nothing
            'End If
            ''''''''''''''''''''''''
            
           'CadenadeLog "Posición 0.9"
'CadenadeLog "<02>"
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
           If (lTransPluAsc = False) Or ((articulo.sec <> 14) And (articulo.sec <> 16)) Then
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
           Else
           
            If ((articulo.sec <> 14) And (articulo.sec <> 16)) Then
            If Trim(articulo.txt(10)) <> "" Then .Fields("des_plux") = articulo.txt(10)
            If Trim(articulo.txt(11)) <> "" Then .Fields("des_plu11") = articulo.txt(11)
            If Trim(articulo.txt(12)) <> "" Then .Fields("des_plu12") = articulo.txt(12)
            If Trim(articulo.txt(13)) <> "" Then .Fields("des_plu13") = articulo.txt(13)
            If Trim(articulo.txt(14)) <> "" Then .Fields("des_plu14") = articulo.txt(14)
            If Trim(articulo.txt(15)) <> "" Then .Fields("des_plu15") = articulo.txt(15)
            If Trim(articulo.txt(16)) <> "" Then .Fields("des_plu16") = articulo.txt(16)
            If Trim(articulo.txt(17)) <> "" Then .Fields("des_plu17") = articulo.txt(17)
            If Trim(articulo.txt(18)) <> "" Then .Fields("des_plu18") = articulo.txt(18)
            If Trim(articulo.txt(19)) <> "" Then .Fields("des_plu19") = articulo.txt(19)
            If Trim(articulo.txt(20)) <> "" Then .Fields("des_plu20") = articulo.txt(20)
            End If
           
           End If
'CadenadeLog "<03>"
            'CadenadeLog "Posición 0.A"
            
            For Bucle = 0 To 9
                'If bucle < 10 Then
                    Select Case Bucle
                        Case 0 To 8
                            MiNombre = "tip_let" & Val(Bucle + 1)
                        Case 9
                            MiNombre = "tip_let0"
                    End Select
                    .Fields(MiNombre) = articulo.tlt(Bucle)
                'End If
            Next Bucle
            
            'CadenadeLog "Posición 0.B"
                
            '*****
            ' Tramos
            '*****
            '1.7.5
            !lmix = articulo.lmix
            !poid = articulo.npeso
            '''''''
            
            'CadenadeLog "Posición 0.C"
            'If lTransPluAsc = False Then
            If ChkTramo Then
                !usatramos = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=false and codigo=" & !codigo)
                With RegTramo
                    If .EOF Then
                        .AddNew
                        CadenadeLog "Se añade artículo a tabla tramos:" & !codigo
                    Else
                        Edit_Record RegTramo
                        CadenadeLog "Se modifica en tabla tramos artículo:" & !codigo
                    End If
                        !borrado = False
                        !tran_tramo = ""
                        RegArticulo.Fields("tran_plu") = ""
                        !codigo = articulo.cod
                        'If articulo.trm1 <> -1 And articulo.prc1 <> -1 Then
                            If articulo.trm1 = -1 Then articulo.trm1 = 0
                            If articulo.trm2 = -1 Then articulo.trm2 = 0
                            If articulo.prc1 = -1 Then articulo.prc1 = 0
                            If articulo.prc2 = -1 Then articulo.prc2 = 0
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
                            'Else
                            '    !tramo2 = -1
                            '    !precio2 = -1
                            '    !euros2 = -1
                            End If
                        'End If
                    .Update
                End With
                RegTramo.Close
                Set RegTramo = Nothing
            
            Else
                If (lTransPluAsc = False) Then
                    !usatramos = False
                End If
            End If
            
            'End If
'CadenadeLog "<04>"
            If UsaEuro Then
                If (articulo.m_tarifa(1)) Then !tar1 = articulo.tarifa(1) / (100)
                If (articulo.m_tarifa(2)) Then !tar2 = articulo.tarifa(2) / (100)
                If (articulo.m_tarifa(3)) Then !tar3 = articulo.tarifa(3) / (100)
                If (articulo.m_tarifa(4)) Then !tar4 = articulo.tarifa(4) / (100)
                If (articulo.m_tarifa(5)) Then !tar5 = articulo.tarifa(5) / (100)
                If (articulo.m_tarifa(6)) Then !tar6 = articulo.tarifa(6) / (100)
                If (articulo.m_tarifa(7)) Then !tar7 = articulo.tarifa(7) / (100)
                If (articulo.m_tarifa(8)) Then !tar8 = articulo.tarifa(8) / (100)
                If (articulo.m_tarifa(9)) Then !tar9 = articulo.tarifa(9) / (100)
                If (articulo.m_tarifa(10)) Then !tar10 = articulo.tarifa(10) / (100)
            
            Else
                If (articulo.m_tarifa(1)) Then !tar1 = articulo.tarifa(1) / (10 ^ Decimales)
                If (articulo.m_tarifa(2)) Then !tar2 = articulo.tarifa(2) / (10 ^ Decimales)
                If (articulo.m_tarifa(3)) Then !tar3 = articulo.tarifa(3) / (10 ^ Decimales)
                If (articulo.m_tarifa(4)) Then !tar4 = articulo.tarifa(4) / (10 ^ Decimales)
                If (articulo.m_tarifa(5)) Then !tar5 = articulo.tarifa(5) / (10 ^ Decimales)
                If (articulo.m_tarifa(6)) Then !tar6 = articulo.tarifa(6) / (10 ^ Decimales)
                If (articulo.m_tarifa(7)) Then !tar7 = articulo.tarifa(7) / (10 ^ Decimales)
                If (articulo.m_tarifa(8)) Then !tar8 = articulo.tarifa(8) / (10 ^ Decimales)
                If (articulo.m_tarifa(9)) Then !tar9 = articulo.tarifa(9) / (10 ^ Decimales)
                If (articulo.m_tarifa(10)) Then !tar10 = articulo.tarifa(10) / (10 ^ Decimales)
            
            End If
            
            .Update
            
            'CadenadeLog "Posición 0.D"
            
            'If articulo.sub <> "0" Then
            '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
            '    articulo.sec & " and codi_sub=" & articulo.sub
            '    With RegSubsec
            '    If .EOF Or .NoMatch Then
            '        .AddNew
            '        !codi_sub = articulo.sub
            '        !codi_ident = articulo.sec
            '        '!secc_maqui = Reg2!secc_maqui
            '        !borrado = False
            '        !tran_sub = ""
            '        .Update
            '    End If
            '    End With
            'End If
'.Fields("codi_fam") = nFam
'.Fields("codi_ident") = 0
'.Fields("secc_maqui") = 0
'.Fields("txt_fam") = Left(TxtFam, 30)
'.Fields("posicion") = Value
'.Fields("imagen") = Trim(TxtEtq)
            
            'If articulo.fam <> "0" Then
            If lCOCOBIO = False Then
                Set Reg3 = Base.OpenRecordset("select max(index) from familias where secc_maqui=" & CStr(articulo.sec))
                If Reg3.EOF Then
                    nTmpFam = 1
                Else
                    If Not IsNull(Reg3.Fields(0)) Then
                        nTmpFam = Reg3.Fields(0) + 1
                    Else
                        nTmpFam = 1
                    End If
                End If
                Set Reg3 = Base.OpenRecordset _
                ("select * from fam_code where codi_fam=" & CStr(articulo.fam))
                With Reg3
                    If .EOF Then
                        .AddNew
                        !codi_Fam = articulo.fam
                        !codi_ident = 0
                        !secc_maqui = 0
                        !txt_fam = "DPT. " & CStr(articulo.fam)
                        !posicion = 0
                        !imagen = ""
                        'hiperusera If lEsTest Then !Index = ((articulo.sec) * 100 + articulo.fam)
                        !Index = 0
                        .Update
                    End If
                End With
                If lEsTest Then
                    Set Reg3 = Base.OpenRecordset _
                    ("select * from familias where codi_fam=" & CStr(articulo.fam) & " and secc_maqui=" & CStr(articulo.sec))
                    With Reg3
                    If .EOF Then
                        .AddNew
                        !codi_Fam = articulo.fam
                        !codi_ident = articulo.sec
                        !secc_maqui = articulo.sec
                        !tran_fam = ""
                        '!txt_fam = "DPT. " & CStr(articulo.fam)
                        !Index = nTmpFam
                        !borrado = False
                        .Update
                    End If
                    End With
                End If
            End If
            'End If
        Else
            '************************
            ' Modificación
            '************************
            RegArticulo.FindFirst "borrado=false and codigo=" & CStr(articulo.cod)
            
            'If lEsAho And UCase(FicheroImport) <> UCase("baletl.dat") Then
            '   If !codi_ident <> articulo.sec Then
            '      CadenadeLog "Baja de artículo" & CStr(articulo.cod)
            '      Retorno = 69
            '      GoTo fin
            '   End If
            'Else
            'End If
            'If (lEsAho = True) And UCase(FicheroImport) = UCase("baletl.dat") Then
            '
            '        Edit_Record RegArticulo
            '        If lEsGestMag Then
            '            .Fields("des_plu1") = articulo.txt(0)
            '        Else
            '            For bucle = 0 To 20
            '                Select Case bucle
            '                    Case 0 To 8
            '                        MiNombre = "des_plu" & CStr(bucle + 1)
            '                    Case 9
            '                        MiNombre = "des_plu0"
            '                    Case 10
            '                        MiNombre = "des_plux"
            '                    Case 11 To 20
            '                        MiNombre = "des_plu" & CStr(bucle)
            '                End Select
            '                .Fields(MiNombre) = articulo.txt(bucle)
            '                If bucle < 10 Then
            '                    Select Case bucle
            '                        Case 0 To 8
            '                            MiNombre = "tip_let" & Val(bucle + 1)
            '                        Case 9
            '                            MiNombre = "tip_let0"
            '                    End Select
            '                    .Fields(MiNombre) = articulo.tlt(bucle)
            '                End If
            '            Next bucle
            '        End If
            '
             '       ModificacionTx1 = True
            '        Modificaciontexto = True
            '        !PRC3 = Val(Format(Now, "yymmddhhmm"))
            '        !tran_textoel = ""
            '        !tran_texto = ""
            '        !tran_textosc10 = ""
            '        !PRC3 = Val(Format(Now, "yymmddhhmm"))
            '        !tran_tx1 = ""
            '        !tran_tx1SC10 = ""
            '
            '        If lEsGestMag = False Then
            '            !etq = articulo.etq
            '        End If
            '
            '        .Update
            '        GoTo etl
            '
            'End If
            
            
            'If articulo.plu <> 0 Then
            '    If !plu <> articulo.plu Then
            '        If Dir(App.Path & "\hiper") <> "" Then
            '            lDarBajaAlta = True
            '        Else
            '            Retorno = 2
            '            GoTo fin
            '        End If
            '    End If
           '
            'End If
            
            'If articulo.sec = 0 Then articulo.plu = 0
            
            Edit_Record RegArticulo
            ''''''''''''''''''''''''''' en principio para carrefour
            If !codi_ident <> articulo.sec Then
                ModificacionBasica = True
                !codi_ident = articulo.sec
                !secc_maqui = Existe_Mostrador(articulo.sec)
            End If
            !tip_letx = 0
            
            'CadenadeLog "Posición M.0"
            
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
                If IsNull(!plu) Then
                    !plu = 0
                    PluBuffer = 0
                Else
                    PluBuffer = !plu
                End If
            End If
                
                'CadenadeLog "Posición M.1"
                
                'PluBuffer = !plu
                ModificacionBasica = False
                'Modificaciontexto = False
                'ModificacionTx1 = False
                'ModificacionCodBar = False
                ModificacionTramo = False
                'If TodoModificaciones Then
                '    ModificacionBasica = True
                '    Modificaciontexto = True
                '    ModificacionTx1 = True
                '    ModificacionCodBar = True
                '    ModificacionTramo = True
                'End If
                'If (!grupo_conserv <> articulo.gpc) And (articulo.m_gpc) Then
                '    Modificaciontexto = True
                '    ModificacionTx1 = True
                'End If
                If lEsGestMag = False Then
                    If IsNull(!tipo_iva) Then !tipo_iva = 0
                End If
                If (.Fields("prc100g") <> articulo.prc100g) And articulo.m_prc100g Then ModificacionBasica = True
                If (!tipo_iva <> articulo.iva) And articulo.m_iva Then ModificacionBasica = True
                If (!pref <> articulo.pre) And articulo.m_pre Then ModificacionBasica = True
                If (!etq <> articulo.etq) And articulo.m_etq Then ModificacionBasica = True
                'If (!codi_sub <> articulo.sub) And articulo.m_sub Then ModificacionBasica = True
                If (!codi_Fam <> articulo.fam) And articulo.m_fam Then ModificacionBasica = True
                If (!fcb <> articulo.Merma) And articulo.m_merma Then ModificacionBasica = True
                If (!ning <> articulo.Presec) And articulo.m_presec Then ModificacionBasica = True
                If articulo.m_onkey Then
                    If articulo.onkey = 0 And !onkey Then ModificacionBasica = True
                    If articulo.onkey = 1 And !onkey = False Then ModificacionBasica = True
                End If
                'CadenadeLog "Posición M.2"
                If Not UsaEuro Then
                    If !precio <> MiPrecio Then
                        ModificacionBasica = True
                        'ModificacionPrecio = True
                        SeCambioPrecio = True
                    End If
                Else
                    If !Euros <> MiEuro Then
                        ModificacionBasica = True
                        'ModificacionPrecio = True
                        SeCambioPrecio = True
                    End If
                End If
                If articulo.m_onkey Then
                    If articulo.onkey = 1 Then !onkey = True
                    If articulo.onkey = 0 Then !onkey = False
                Else
                    !onkey = True
                End If
                'CadenadeLog "Posición M.3"
                'If articulo.cad <> 0 Or lEsAho = True Then
                '    If (!Caducidad <> articulo.cad) And articulo.m_cad Then ModificacionBasica = True
                'End If
                'If articulo.cad = 0 And Dir(App.Path & "\transpluasc.exe") <> "" Then
                '    If (!Caducidad <> articulo.cad) Then ModificacionBasica = True
                'End If
                
                'If SUPECO Or Dir(App.Path & "\hiper") <> "" And (articulo.tar = 0) Then articulo.tar = !tara
                'If (!tara <> articulo.tar) Then ModificacionBasica = True
                'If Dir(App.Path & "\hiper") <> "" And (articulo.cad = 0) Then articulo.cad = !Caducidad
                If (!Caducidad <> articulo.cad) And articulo.m_cad Then ModificacionBasica = True
                If (!tara <> articulo.tar) And articulo.m_tar Then ModificacionBasica = True
                If (!art_cb <> articulo.cba) And articulo.m_cba Then ModificacionBasica = True
                'CadenadeLog "Posición M.4"
                '****************************
                ' Alcampo : en las modificaciones
                ' no se altera pesado/no pesado
                ' porque llegan datos de AS400
                ' a corregir por el usuario
                If Not Conf_Al Then
                    'If Dir(App.Path & "\bdp.txt") = "" Then
                        If (MiWgh = True And !codi_pes = "U") Or _
                        (MiWgh = False And !codi_pes = "W") Then ModificacionBasica = True
                    'End If
                End If
                '/Alcampo
                '***********************************
                'If (Balenv = False And !Balenv = True) Or _
                '(Balenv = True And !Balenv = False) Then ModificacionBasica = True
                
                ''2.0.29
                'If lEsAho Then ModificacionBasica = True
                ''''''''
                If articulo.m_image Then
                    'CadenadeLog "IMAGEN:" & articulo.imagen
                    'articulo.imagen = separaFich(articulo.imagen)
                    'CadenadeLog "IMAGEN:" & articulo.imagen
                    If Trim(articulo.imagen) = "" Then
                        articulo.imagen = "" ' "#FF646464"
                    End If
                    !imagen = Trim(articulo.imagen)
                'Else
                '    !imagen = "#FF646464"
                End If
                
                
                'CadenadeLog "Posición M.5"
                'If articulo.m_image Then
                '    articulo.m_image = False
                '    If IsNumeric(Val(articulo.imagen)) Then
                '        If !imagen <> APaleta(Val(articulo.imagen)) Then
                '            !imagen = APaleta(Val(articulo.imagen))
                '            ModificacionBasica = True
                '        End If
                '    Else
                '        If Trim(articulo.imagen) <> "NO IMAGE" Then
                '            If Trim(articulo.imagen) <> "" Then
                '                If Trim(articulo.imagen) <> Trim(!magen) Then
                '                    ModificacionBasica = True
                '                End If
                '                If Dir(Trim(articulo.imagen)) <> "" Then
                '                    FileCopy Trim(articulo.imagen), App.Path & "\images\items\" & CStr(articulo.cod)
                '                    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '                Else
                '                    If Dir(App.Path & "\itmimg.jpg") <> "" Then
                '                        FileCopy App.Path & "\itmimg.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
                '                        !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '                    Else
                '                        !imagen = "#FF646464"
                '                    End If
                 '               End If
                 '           Else
                '                If Trim(articulo.imagen) <> Trim(!magen) Then
                '                    ModificacionBasica = True
                '                End If
                '
                '                If Dir(App.Path & "\itmimg.jpg") <> "" Then
                '                    FileCopy App.Path & "\itmimg.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
                '                    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '                Else
                '                    !imagen = "#FF646464"
                '                End If
                 '           End If
                 '       'Else
                '        '    FileCopy App.Path & "\vacio.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
                '        '    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '        End If
                '    End If
                '
                '    'If IsNumeric(Val(articulo.imagen)) Then
                '    '    !imagen = APaleta(Val(articulo.imagen))
                '    'Else
                '    '    If Trim(articulo.imagen) <> "NO IMAGE" Then
                '    '        If Dir(Trim(articulo.imagen)) <> "" Then
                '    '            FileCopy Trim(articulo.imagen), App.Path & "\images\items\" & CStr(articulo.cod)
                '    '            !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '    '        Else
                '    '            FileCopy App.Path & "\itmimg.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
                '    '            !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '    '        End If
                '    '    'Else
                '    '    '    FileCopy App.Path & "\vacio.jpg", App.Path & "\images\items\" & CStr(articulo.cod)
                '    '    '    !imagen = App.Path & "\images\items\" & CStr(articulo.cod)
                '    '    End If
                '    'End If
               '
                'End If
                
                'CadenadeLog "Posición M.6"
                
                If lFamPluPos Then
                    If !posicion <> articulo.plu Then
                        !posicion = articulo.plu
                        ModificacionBasica = True
                    End If
                End If
                If Dir(App.Path & "\tiendamt.txt") <> "" Then
                    ModificacionBasica = True
                End If
                
                
                If Dir(App.Path & "\hiper") = "" Then
                
                    If lEsGestMag Then
                        .Fields("des_plu1") = articulo.txt(0)
                    Else
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
                                Then ModificacionBasica = True 'Modificaciontexto = True
                            End If
                        End If
                        If Buclelite < 10 Then
                            MiNombre = "des_plu" & Val(Buclelite + 11)
                            If Not AHOLD Or (AHOLD And Trim(articulo.txt(Buclelite + 11)) <> "") Then
                                If articulo.m_txt(Buclelite + 11) Then
                                    If (.Fields(MiNombre) <> articulo.txt(Buclelite + 11)) Or _
                                    (.Fields(MiNombre) = "" And articulo.txt(Buclelite + 11) <> "") Or _
                                    (.Fields(MiNombre) <> "" And articulo.txt(Buclelite + 11) = "") _
                                    Then ModificacionBasica = True 'Modificaciontexto = True
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
                                    Then ModificacionBasica = True 'Modificaciontexto = True
                                End If
                            End If
                        End If
                    Next Buclelite
                    End If
                    
                Else
                    'Modificaciontexto = False
                    'Modificaciontexto = False
                End If
                
                'CadenadeLog "Posición M.7"
                
                If ModificacionBasica Then
                    !PRC3 = Val(Format(Now, "yymmddhhmm"))
                    !tran_en = ""
                    !tran_el = ""
                    !tran_ep = Space(60)
                    !tran_plu = ""
                    !tran_pluSC10 = ""
                    !tran_texto = ""
                    !tran_textosc10 = ""
                    !tran_tx1 = ""
                    !tran_tx1SC10 = ""
                    !tran_cb = ""
                    !tran_cbsc10 = ""
                End If
                'If Modificaciontexto Then
                '    !PRC3 = Val(Format(Now, "yymmddhhmm"))
                '    !tran_textoel = ""
                '    !tran_ep = Space(60)
                '    !tran_texto = ""
                '    !tran_textosc10 = ""
                'End If
                'If ModificacionTx1 Then
                '    !PRC3 = Val(Format(Now, "yymmddhhmm"))
                '    !tran_tx1 = ""
                '    !tran_ep = Space(60)
                '    !tran_tx1SC10 = ""
                'End If
                'If ModificacionCodBar Then
                '    !PRC3 = Val(Format(Now, "yymmddhhmm"))
                '    !tran_cb = ""
                '    !tran_cbsc10 = ""
                '    !tran_ep = Space(60)
                'End If
                If articulo.m_prc100g Then .Fields("PRC100G") = articulo.prc100g
                If articulo.m_presec Then !ning = articulo.Presec
                If articulo.m_merma Then !fcb = articulo.Merma
                If articulo.m_cba Then !art_cb = articulo.cba
                If articulo.m_sub Then !codi_sub = articulo.sub
                If Dir(App.Path & "\etc\ekobal.dat.fia") = "" Then
                    If articulo.fam <> 0 Then !codi_Fam = articulo.fam
                End If
                'c2f/caspiunza If articulo.m_fam Then !codi_fam = articulo.fam
                '1.7.5
                'CadenadeLog "Posición M.8"
                !lmix = articulo.lmix
                !poid = articulo.npeso
                '''''''
                'CadenadeLog "Posición M.9"
                If Not UsaEuro Then
                    !precio = MiPrecio
                    'If ModificacionPrecio Or IsNull(!Euros) Then
                    !Euros = MiEuro
                Else
                    !Euros = MiEuro
                    'If ModificacionPrecio Or IsNull(!precio) Then
                    !precio = MiPrecio
                End If
                If articulo.m_iva Then !tipo_iva = articulo.iva
                'If articulo.cad <> 0 Or lEsAho = True Or Dir(App.Path & "\transpluasc.exe") <> "" Then
                '    !Caducidad = articulo.cad
                'End If
                If articulo.m_cad Or articulo.cad <> 0 Then
                    !Caducidad = articulo.cad
                End If
                
                If articulo.m_tar Then
                    !tara = articulo.tar
                    !tara_envasado = articulo.tar
                End If
                If articulo.m_pre Then !pref = articulo.pre
                If articulo.m_etq Then !etq = articulo.etq
                If articulo.m_gpc Then !grupo_conserv = articulo.gpc
                If lEsAho Then
                    !etq = 0
                End If
                'CadenadeLog "Posición M.A"
                '********************************
                ' datos de caducidad programados
                ' en fichero
                'c2f/caspiunza también en modificación
                If lEsGestMag = False Then
                    If articulo.cad = 0 Then
                        If DataCaducidad.Numero > 0 Then
                            For Bucle = 0 To DataCaducidad.Numero - 1
                                If DataCaducidad.Seccion(Bucle) = articulo.sec Then
                                    If DataCaducidad.Subseccion(Bucle) = articulo.sub Then
                                        If DataCaducidad.Familia(Bucle) = articulo.fam Then
                                            !Caducidad = DataCaducidad.Caducidad(Bucle)
                                        End If
                                    End If
                                End If
                            Next Bucle
                        End If
                    End If
                End If
                '********************************
                
                '********************************
                ' Alcampo : en las modificaciones
                ' no se altera pesado/no pesado
                ' porque llegan datos de AS400
                ' a corregir por el usuario
                'If Dir(App.Path & "\bdp.txt") = "" Then
                    If Not Conf_Al And articulo.m_wgh = True Then
                        If MiWgh = True Then
                            !codi_pes = "W"
                        Else
                            !codi_pes = "U"
                        End If
                    End If
                'End If
                
                'CadenadeLog "Posición M.B"
                
                '/Alcampo
                '*********************
                'pruebas hiperusera se anula...
                'If Balenv = False Then
                '    !Balenv = False
                '    Set Reg2 = Base.OpenRecordset _
                '    ("select secc_maqui from seccion where borrado=false and codi_ident=" & !codi_ident)
                '    Set Reg5 = Base.OpenRecordset _
                '    ("select count(codigo) from articulo where borrado=false and secc_maqui=" & _
                '    Reg2!secc_maqui & " and balenv=true")
                '
                '    If Reg5.Fields(0) = 1 Then
                '        Reg5.Close
                '        Set Reg5 = Base.OpenRecordset _
                '        ("select multiple,enviardatos,tran_sec,tran_secSC10 from seccion where borrado=false and secc_maqui=" & _
                '        Reg2!secc_maqui)
                '        With Reg5
                '            If Not .EOF Then
                '                .MoveFirst
                '                Do Until .EOF
                '                    If !multiple = 1 Then
                '                        Edit_Record Reg5
                '                        !multiple = -1
                '                        If !enviardatos = "GA" Then !tran_secSC10 = ""
                '                        If !enviardatos = "TODOS" Then !tran_sec = ""
                '                        .Update
                '                    End If
                '                    If Not .EOF Then .MoveNext
                '                Loop
                '            End If
                '        End With
                '    End If
                '
                'Else
                '    !Balenv = True
                '    Set Reg2 = Base.OpenRecordset _
                '    ("select secc_maqui from seccion where borrado=false and codi_ident=" & !codi_ident)
                '    Set Reg5 = Base.OpenRecordset _
                '    ("select enviardatos,multiple,tran_sec,tran_secSC10 from seccion where borrado=false and secc_maqui=" & Reg2!secc_maqui)
                '    With Reg5
                '        If Not .EOF Then
                '            .MoveFirst
                '            Do Until .EOF
                '            If !enviardatos <> "TODOS" Then
                '                Edit_Record Reg5
                '                If !multiple <> 1 Then
                '
                '                    If !enviardatos = "GA" Then !tran_secSC10 = ""
                '                    If !enviardatos = "SC10" Then !tran_sec = ""
                '                End If
                '                !multiple = 1
                '                .Update
                '            End If
                '            If Not .EOF Then .MoveNext
                '            Loop
                '        End If
                '    End With
                'End If
                'Reg2.Close
                'Set Reg2 = Nothing
                'Reg5.Close
                'Set Reg5 = Nothing
                
                If Dir(App.Path & "\hiper") = "" Then
                    If lEsGestMag = False Then
                    
                    For Bucle = 0 To 20
                        Select Case Bucle
                            Case 0 To 8
                                MiNombre = "des_plu" & Val(Bucle + 1)
                            Case 9
                                MiNombre = "des_plu0"
                            Case 10
                                MiNombre = "des_plux"
                            Case 11 To 20
                                MiNombre = "des_plu" & Val(Bucle)
                        End Select
                        If Not AHOLD Then
                            If lTransPluAsc = False Then
                                If articulo.m_txt(Bucle) Then .Fields(MiNombre) = articulo.txt(Bucle)
                            Else
                               
                               If ((articulo.sec = 14) Or (articulo.sec = 16)) And Bucle >= 10 Then
                                   MiNombre = MiNombre
                               Else
                                If MiNombre = "des_plu2" And Mid(articulo.txt(Bucle), 1, 1) = Chr(94) Then
                                    On Error Resume Next
                                    articulo.txt(Bucle) = Mid(articulo.txt(Bucle), 2)
                                    If Err.Number <> 0 Then
                                        On Error GoTo 0
                                        articulo.txt(Bucle) = ""
                                    End If
                                    On Error GoTo 0
                                End If
                                
                                If Mid(articulo.txt(Bucle), 1, 1) = "|" Then
                                    articulo.txt(Bucle) = " " & articulo.txt(Bucle)
                                End If
                                
                                If articulo.m_txt(Bucle) Then
                                    If Trim(articulo.txt(Bucle)) <> "" Then
                                        .Fields(MiNombre) = articulo.txt(Bucle)
                                    End If
                                End If
                               End If
                            End If
                        Else
                            If articulo.m_txt(Bucle) And Trim(articulo.txt(Bucle)) <> "" Then .Fields(MiNombre) = articulo.txt(Bucle)
                        End If
                        If Bucle < 10 Then
                            Select Case Bucle
                                Case 0 To 8
                                    MiNombre = "tip_let" & Val(Bucle + 1)
                                Case 9
                                    MiNombre = "tip_let0"
                            End Select
                            If Not AHOLD Then
                                If articulo.m_tlt(Bucle) Then .Fields(MiNombre) = articulo.tlt(Bucle)
                            Else
                                If articulo.m_tlt(Bucle) And articulo.tlt(Bucle) <> 0 Then .Fields(MiNombre) = articulo.tlt(Bucle)
                            End If
                        End If
                    Next Bucle
                    
                    End If
                    
                End If
                '*****
                ' Tramos
                '*****
                
                'If Dir(App.Path & "\transpluasc.exe") = "" Then
                'If lTransPluAsc = False Then
                If ChkTramo And !usatramos = True Then
                    '!usatramos = True
                    CadenadeLog "modificación tramo:" & articulo.cod
                    Set RegTramo = Base.OpenRecordset _
                    ("select * from tramos where borrado=false and codigo=" & !codigo)
                    With RegTramo
                        If Not .EOF Then
                            Edit_Record RegTramo
                            !tran_tramo = ""
                            RegArticulo.Fields("tran_plu") = ""
                            !codigo = articulo.cod
                            If articulo.trm1 = -1 Then articulo.trm1 = 0
                            If articulo.trm2 = -1 Then articulo.trm2 = 0
                            If articulo.prc1 = -1 Then articulo.prc1 = 0
                            If articulo.prc2 = -1 Then articulo.prc2 = 0
                            
                            'If !precio1 <> articulo.prc1 Or !tramo1 <> articulo.trm1 Then
                                
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
                            'End If
                            'If (articulo.prc2 = -1 And !precio2 >= 0) Then
                            '    !euros2 = -1
                            '    !precio2 = -1
                            '    !tramo2 = -1
                            '    RegArticulo!tran_plu = ""
                            '    ModificacionTramo = True
                            'Else
                                'If articulo.prc2 <> -1 And _
                                '(!precio2 <> articulo.prc2 Or !tramo2 <> articulo.trm2) Then
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
                                'End If
                            'End If
                        .Update
                        End If
                    End With
                
                    RegTramo.Close
                    Set RegTramo = Nothing
                    
                End If
                
                If ChkTramo And !usatramos = False Then
                    CadenadeLog "alta tramo desde modificación:" & articulo.cod
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
                            RegArticulo.Fields("tran_plu") = ""
                            !codigo = articulo.cod
                            'If articulo.trm1 <> -1 And articulo.prc1 <> -1 Then
                                If articulo.trm1 = -1 Then articulo.trm1 = 0
                                If articulo.trm2 = -1 Then articulo.trm2 = 0
                                If articulo.prc1 = -1 Then articulo.prc1 = 0
                                If articulo.prc2 = -1 Then articulo.prc2 = 0
                            
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
                                'Else
                                '    !tramo2 = -1
                                '    !precio2 = -1
                                '    !euros2 = -1
                                End If
                            'End If
                        .Update
                    End With
                    RegTramo.Close
                    Set RegTramo = Nothing
                    
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
                    RegTramo.Close
                    Set RegTramo = Nothing
                    
                End If
                
                'End If
                
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
                '1.7.5
                
                If UsaEuro Then
                    If (articulo.m_tarifa(1)) Then !tar1 = articulo.tarifa(1) / (100)
                    If (articulo.m_tarifa(2)) Then !tar2 = articulo.tarifa(2) / (100)
                    If (articulo.m_tarifa(3)) Then !tar3 = articulo.tarifa(3) / (100)
                    If (articulo.m_tarifa(4)) Then !tar4 = articulo.tarifa(4) / (100)
                    If (articulo.m_tarifa(5)) Then !tar5 = articulo.tarifa(5) / (100)
                    If (articulo.m_tarifa(6)) Then !tar6 = articulo.tarifa(6) / (100)
                    If (articulo.m_tarifa(7)) Then !tar7 = articulo.tarifa(7) / (100)
                    If (articulo.m_tarifa(8)) Then !tar8 = articulo.tarifa(8) / (100)
                    If (articulo.m_tarifa(9)) Then !tar9 = articulo.tarifa(9) / (100)
                    If (articulo.m_tarifa(10)) Then !tar10 = articulo.tarifa(10) / (100)
                
                Else
                    If (articulo.m_tarifa(1)) Then !tar1 = articulo.tarifa(1) / (10 ^ Decimales)
                    If (articulo.m_tarifa(2)) Then !tar2 = articulo.tarifa(2) / (10 ^ Decimales)
                    If (articulo.m_tarifa(3)) Then !tar3 = articulo.tarifa(3) / (10 ^ Decimales)
                    If (articulo.m_tarifa(4)) Then !tar4 = articulo.tarifa(4) / (10 ^ Decimales)
                    If (articulo.m_tarifa(5)) Then !tar5 = articulo.tarifa(5) / (10 ^ Decimales)
                    If (articulo.m_tarifa(6)) Then !tar6 = articulo.tarifa(6) / (10 ^ Decimales)
                    If (articulo.m_tarifa(7)) Then !tar7 = articulo.tarifa(7) / (10 ^ Decimales)
                    If (articulo.m_tarifa(8)) Then !tar8 = articulo.tarifa(8) / (10 ^ Decimales)
                    If (articulo.m_tarifa(9)) Then !tar9 = articulo.tarifa(9) / (10 ^ Decimales)
                    If (articulo.m_tarifa(10)) Then !tar10 = articulo.tarifa(10) / (10 ^ Decimales)
                
                End If
                
                .Update
                
                
                'If articulo.sub <> "0" Then
                '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
                '    articulo.sec & " and codi_sub=" & articulo.sub
                '    With RegSubsec
                '    If .EOF Or .NoMatch Then
                '        .AddNew
                '        !codi_sub = articulo.sub
                '        !codi_ident = articulo.sec
                '        '!secc_maqui = Reg2!secc_maqui
                '        !borrado = False
                '        !tran_sub = ""
                '        .Update
                '    End If
                '    End With
                '
                'End If
                'If articulo.fam <> "0" Then
                '    Set Reg3 = Base.OpenRecordset _
                '    ("select codi_fam,codi_sub,codi_ident,secc_maqui,borrado,tran_fam from familia where borrado=false and codi_ident=" & _
                '    articulo.sec & " and codi_sub=" & articulo.sub & _
                '    " and codi_fam=" & articulo.fam)
                '    With Reg3
                '    If .EOF Then
                '        .AddNew
                '        !codi_Fam = articulo.fam
                '        !codi_sub = articulo.sub
                '        !codi_ident = articulo.sec
                '        !secc_maqui = Reg2!secc_maqui
                '        !borrado = False
                '        !tran_fam = ""
                '        .Update
                '    End If
                '    End With
                '
                '
                'End If
                'If articulo.fam <> "0" Then
                If lCOCOBIO = False Then
                    Set Reg3 = Base.OpenRecordset _
                    ("select * from fam_code where codi_fam=" & CStr(articulo.fam))
                    With Reg3
                        If .EOF Then
                            .AddNew
                            !codi_Fam = articulo.fam
                            !codi_ident = 0
                            !secc_maqui = 0
                            !txt_fam = "DPT. " & CStr(articulo.fam)
                            !posicion = 0
                            !imagen = ""
                            .Update
                        End If
                    End With
                    If lEsTest Then
                        Set Reg3 = Base.OpenRecordset _
                        ("select * from familias where codi_fam=" & CStr(articulo.fam) & " and secc_maqui=" & CStr(articulo.sec))
                        With Reg3
                        If .EOF Then
                            .AddNew
                            !codi_Fam = articulo.fam
                            !codi_ident = articulo.sec
                            !secc_maqui = articulo.sec
                            !tran_fam = ""
                            '!txt_fam = "DPT. " & CStr(articulo.fam)
                            !Index = nTmpFam
                            !borrado = False
                            .Update
                        End If
                        End With
                    End If
                End If
                'End If
            
      End If
    
etl:

    End With
    
    'If ModificacionTramo Or ModificacionBasica Or Modificaciontexto Or _
    'ModificacionTx1 Or ModificacionCodBar Or ModificacionPrecio Then
    If ModificacionBasica Or ModificacionTramo Then
        Retorno = 1
    Else
        Retorno = 4
    End If
fin:

If lEsBacK = False Then
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
End If
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
            !codi_Fam = articulo.fam
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
        'c2fcaspiunza
        RstRechazos.Close
        Set RstRechazos = Nothing
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
            !codi_Fam = articulo.fam
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
        'c2f/caspiunza
        RstRechazos.Close
        Set RstRechazos = Nothing
    End If
End If
'c2f/caspiunza
'Base.Close
'RegArticulo.Close
'''''''''''''''

'If RESPONSE And Retorno <> 2 Then Print #desc, Format(Retorno, "00")
'CadenadeLog "Alta_articulo--> Retorno:" & CStr(Retorno)
    
    'CadenadeLog "Alta_Articulo out..."
    
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
        'If RESPONSE And (Not NoText) Then Print #desc, "02"
        Exit Sub
    End If
    If Dir(App.Path & "\tiendamt.txt") <> "" Then
        If articulo.cod = 0 Then
            articulo.cod = Busca_Codigos_Sec_plu(articulo.sec, articulo.plu)
        End If
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
            !tran_ep = Space(60)
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
    'If RESPONSE And (Not NoText) Then Print #desc, "01"
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
        !tran_tecSC10 = CStr(LaTecla.Destino_NMostrador)
        !borrado = False
        !prog_tec = "TOUCH"
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
    
    'On Error Resume Next
    'Registro.Close 'c2f/caspiunza
    'Base.Close 'c2f/caspiunza
    'On Error GoTo 0
    
fin:
    
    Baja_Teclas = Retorno
        
End Function

''''''''''''
''''''''''''
'Public Function alta_txt2040(abase As Database, aart As registroarticulos) As Integer
'    Dim rst As Recordset
'    Dim cnt As Long
'    Dim sSQL As String
'    Dim nPLu As Long
'    Dim nSec As Long
'
'    sSQL = "SELECT * FROM articulo WHERE codigo=" & CStr(aart.cod)
'    Set rst = abase.OpenRecordset(sSQL)
'
'    If Not rst.EOF Then
'
'        nPLu = rst.Fields("plu")
'        nSec = rst.Fields("codi_ident")
'
'        sSQL = "SELECT * FROM linTxt2040 WHERE codigo=" & CStr(aart.cod)
'        Set rst = abase.OpenRecordset(sSQL)
'        With rst
'            If .BOF Then
'                .AddNew
'            Else
'                .Edit
'            End If
'            .Fields("codigo") = aart.cod
'            .Fields("plu") = nPLu
'            .Fields("mostrador") = nSec
'            .Fields("txt_21") = aart.txt2040(0)
'            .Fields("txt_22") = aart.txt2040(1)
'            .Fields("txt_23") = aart.txt2040(2)
 '           .Fields("txt_24") = aart.txt2040(3)
 '           .Fields("txt_25") = aart.txt2040(4)
 '           .Fields("txt_26") = aart.txt2040(5)
 '           .Fields("txt_27") = aart.txt2040(6)
 '           .Fields("txt_28") = aart.txt2040(7)
 '           .Fields("txt_29") = aart.txt2040(8)
 '           .Fields("txt_30") = aart.txt2040(9)
 '           .Fields("txt_31") = aart.txt2040(10)
 '           .Fields("txt_32") = aart.txt2040(11)
 '           .Fields("txt_33") = aart.txt2040(12)
 '           .Fields("txt_34") = aart.txt2040(13)
 '           .Fields("txt_35") = aart.txt2040(14)
 '           .Fields("txt_36") = aart.txt2040(15)
 '           .Fields("txt_37") = aart.txt2040(16)
 '           .Fields("txt_38") = aart.txt2040(17)
 '           .Fields("txt_39") = aart.txt2040(18)
 '           .Fields("txt_40") = aart.txt2040(19)
 '           .Fields("tran_txt2") = ""
 '       End With
 '       rst.Update
 '   'ret:
 '       rst.Close
 '       Set rst = Nothing
 '       If Err.Number <> 0 Then
 '           alta_txt2040 = 1
 '       Else
 '           alta_txt2040 = 0
 '       End If
 '       'Err.Clear
 '
 '   Else
 '       rst.Close
 '       alta_txt2040 = 0
 '   End If
 '
 '   'On Error GoTo 0'

'End Function
Public Function alta_txt2040(abase As Database, aart As registroarticulos) As Integer
    Dim rst As Recordset
    Dim cnt As Long
    Dim sSQL As String
    Dim nPLu As Long
    Dim nSec As Long
    Dim sF As String
    Dim nC As Integer
    Dim resto As String
    
'CadenadeLog "alta_txt entrada Código:" & CStr(aart.cod)

    sSQL = "SELECT * FROM articulo WHERE codigo=" & CStr(aart.cod)
    Set rst = abase.OpenRecordset(sSQL)

'CadenadeLog "alta_txt select Código:" & CStr(aart.cod)
    
    If Not rst.EOF Then
        
        nPLu = rst.Fields("plu")
        nSec = rst.Fields("codi_ident")
        
        sSQL = "SELECT * FROM linTxt2040 WHERE codigo=" & CStr(aart.cod)
        Set rst = abase.OpenRecordset(sSQL)
        With rst
            If .BOF Then
                .AddNew
            Else
                .Edit
            End If
            .Fields("codigo") = aart.cod
            .Fields("plu") = nPLu
            .Fields("mostrador") = nSec
'CadenadeLog "<-10->"
            For cnt = 0 To 59 '59 --> mayor que 53 error base de datos registro muy grande...
                'sF = "txt_" & CStr(21 + cnt)
                'If lN And aart.lTraza And (cnt + 21) >= 40 Then '(cnt + 21) = 40 Then
                '    'If Len(Trim(aart.txt2040(39))) = 34 Then
                '        'Corte Palabra
                '
                '    'End If
                '    If (Len(Trim(aart.txt2040(cnt))) >= 34) Then 'And Left(aart.txt2040(cnt + 1), 1) <> " " And Trim(aart.txt2040(cnt + 1)) <> "" Then
                '        '    For nC = Len(aart.txt2040(39)) To 1 Step -1
                '        '        If Mid(aart.txt2040(39), nC, 1) = " " Then
                '        '
                '        '        End If
                '        '    Next nC
                '        '
                '        'End If
                '        resto = ""
                '        'CadenadeLog "RETRANQUEO:" & CStr(aart.cod)
                '        For nC = Len(aart.txt2040(cnt)) To 1 Step -1
                '            If Mid(aart.txt2040(cnt), nC, 1) = " " Then
                '                If (Len(aart.txt2040(cnt)) - Len(Mid(aart.txt2040(cnt), nC + 1))) <= 34 Then
                '                    resto = Mid(aart.txt2040(cnt), nC + 1)
                '                    aart.txt2040(cnt) = Mid(aart.txt2040(cnt), 1, nC - 1)
                '                    'CadenadeLog "Queda:" & aart.txt2040(cnt) & " LEN:" & CStr(Len(aart.txt2040(cnt)))
                '                    Exit For
                '                End If
                '            End If
                '        Next nC
                '
                '        'CadenadeLog "Resto:" & resto
                '        aart.txt2040(cnt + 1) = resto & aart.txt2040(cnt + 1)
                '        'CadenadeLog "Siguiente:" & aart.txt2040(cnt + 1)
                '        'For nC = cnt To 53
                '        '    aart.txt2040(nC + 1) = aart.txt2040(nC)
                '        'Next nC
                '        'aart.txt2040(19) = ""
                '    End If
                'End If
                sF = "txt_" & CStr(21 + cnt)
                If Len(RTrim(aart.txt2040(cnt))) <= 34 Then
''CadenadeLog "<-10->" & CStr(cnt)
                    If lTransPluAsc Then

''cadenadelog "...lin:" & CStr(cnt) & " :" & aart.txt2040(cnt)

                        'If Not IsNull(aart.txt2040(cnt)) Then
                        '    If Mid(aart.txt2040(cnt), 1, 1) = "|" Then
                        '
                        '        'cadenadelog "...Detectado PIPE..."
                        '
                        '        aart.txt2040(cnt) = " " & aart.txt2040(cnt)
                        '
                        '        'cadenadelog "...lin:" & CStr(cnt) & " :" & aart.txt2040(cnt) & " Len:" & Len(aart.txt2040(cnt))
                        '
                        '    Else
                        '        'cadenadelog "...lin:" & CStr(cnt) & " :" & aart.txt2040(cnt) & " Len:" & Len(aart.txt2040(cnt))
                        '    End If
                        'End If
                        'cadenadelog "...Campo:" & sF
                        
                        .Fields(sF) = aart.txt2040(cnt)
                        
                        ''cadenadelog "...Introducido en campo..."
                        
                    Else
                        .Fields(sF) = RTrim(aart.txt2040(cnt))
                    End If
                Else
                    CadenadeLog "Supera 34 caracteres Código:" & CStr(aart.cod) & " " & RTrim(aart.txt2040(cnt))
                End If
                
                
                'CadenadeLog "txt2040(" & CStr(cnt) & ")"
            Next cnt
            
            '.Fields("txt_21") = Trim(aart.txt2040(0))
            '.Fields("txt_22") = Trim(aart.txt2040(1))
            '.Fields("txt_23") = Trim(aart.txt2040(2))
            '.Fields("txt_24") = Trim(aart.txt2040(3))
            '.Fields("txt_25") = Trim(aart.txt2040(4))
            '.Fields("txt_26") = Trim(aart.txt2040(5))
            '.Fields("txt_27") = Trim(aart.txt2040(6))
            '.Fields("txt_28") = Trim(aart.txt2040(7))
            '.Fields("txt_29") = Trim(aart.txt2040(8))
            '.Fields("txt_30") = Trim(aart.txt2040(9))
            '.Fields("txt_31") = Trim(aart.txt2040(10))
            '.Fields("txt_32") = Trim(aart.txt2040(11))
            '.Fields("txt_33") = Trim(aart.txt2040(12))
            '.Fields("txt_34") = Trim(aart.txt2040(13))
            '.Fields("txt_35") = Trim(aart.txt2040(14))
            '.Fields("txt_36") = Trim(aart.txt2040(15))
            '.Fields("txt_37") = Trim(aart.txt2040(16))
            '.Fields("txt_38") = Trim(aart.txt2040(17))
            '.Fields("txt_39") = Trim(aart.txt2040(18))
            '.Fields("txt_40") = Trim(aart.txt2040(19))
            '
            '.Fields("txt_41") = Trim(aart.txt2040(20))
            '.Fields("txt_42") = Trim(aart.txt2040(21))
            '.Fields("txt_43") = Trim(aart.txt2040(22))
            '.Fields("txt_44") = Trim(aart.txt2040(23))
            '.Fields("txt_45") = Trim(aart.txt2040(24))
            '.Fields("txt_46") = Trim(aart.txt2040(25))
            '.Fields("txt_47") = Trim(aart.txt2040(26))
            '.Fields("txt_48") = Trim(aart.txt2040(27))
            '.Fields("txt_49") = Trim(aart.txt2040(28))
            '.Fields("txt_50") = Trim(aart.txt2040(29))
           '
            '.Fields("txt_51") = Trim(aart.txt2040(30))
            '.Fields("txt_52") = Trim(aart.txt2040(31))
            '.Fields("txt_53") = Trim(aart.txt2040(32))
            '.Fields("txt_54") = Trim(aart.txt2040(33))
            '.Fields("txt_55") = Trim(aart.txt2040(34))
            '.Fields("txt_56") = Trim(aart.txt2040(35))
            '.Fields("txt_57") = Trim(aart.txt2040(36))
            '.Fields("txt_58") = Trim(aart.txt2040(37))
            '.Fields("txt_59") = Trim(aart.txt2040(38))
            '.Fields("txt_60") = Trim(aart.txt2040(39))
            
            '.Fields("txt_61") = Trim(aart.txt2040(40))
            '.Fields("txt_62") = Trim(aart.txt2040(41))
            '.Fields("txt_63") = Trim(aart.txt2040(42))
            '.Fields("txt_64") = Trim(aart.txt2040(43))
            '.Fields("txt_65") = Trim(aart.txt2040(44))
            '.Fields("txt_66") = Trim(aart.txt2040(45))
            '.Fields("txt_67") = Trim(aart.txt2040(46))
            '.Fields("txt_68") = Trim(aart.txt2040(47))
            '.Fields("txt_69") = Trim(aart.txt2040(48))
            '.Fields("txt_70") = Trim(aart.txt2040(49))
            
            '.Fields("txt_71") = Trim(aart.txt2040(50))
            '.Fields("txt_72") = Trim(aart.txt2040(51))
            '.Fields("txt_73") = Trim(aart.txt2040(52))
            '.Fields("txt_74") = Trim(aart.txt2040(53))
            '.Fields("txt_75") = Trim(aart.txt2040(54))
            '.Fields("txt_76") = Trim(aart.txt2040(55))
            '.Fields("txt_77") = Trim(aart.txt2040(56))
            '.Fields("txt_78") = Trim(aart.txt2040(57))
            '.Fields("txt_79") = Trim(aart.txt2040(58))
            '.Fields("txt_80") = Trim(aart.txt2040(59))
            
''cadenadelog "...Marcado"

            .Fields("tran_txt2") = ""
            
            
'CadenadeLog "<-11->"
        End With
        rst.Update
    'ret:
        
''cadenadelog "...Amacenado"

        rst.Close
        Set rst = Nothing
        
        If Err.Number <> 0 Then
            alta_txt2040 = 1
        Else
            alta_txt2040 = 0
        End If
        'Err.Clear
    
    Else
        rst.Close
        alta_txt2040 = 0
    End If
    
'CadenadeLog "alta_txt salida Código:" & CStr(aart.cod)
    'On Error GoTo 0
'End If
End Function

''''''''''''
''''''''''''
Public Function DB_Consulta_Art(ByVal nC As Long) As Boolean
    Dim nR As Recordset
    Dim RegTramo As Recordset
    Dim nMybase As Database
    Dim Resp As Boolean
    Dim bucle2 As Integer
    Dim sCampo As String
    
    Resp = False
    'CadenadeLog "Base de datos:" & Base_General
    Set nMybase = OpenDatabase(Base_General)
    On Error Resume Next
    Set nR = nMybase.OpenRecordset("select * from articulo where codigo=" & CStr(nC))
    If Err.Number <> 0 Then
        CadenadeLog "Error en: DB_consulta_Art:" & CStr(Err.Number) & " Descripción:" & Err.Description
        On Error GoTo 0
    End If
    If nR.EOF Then
       Resp = False
    Else
        Resp = True
        'With nR
'    cod As Long
'    plu As Long
'    sec As Long
'    sub As Long
'    fam As Long
'    prc As Double
'    wgh As String
'    cad As Long
'    pre As Long
'    tar As Long
'    cba As String
'    tlt(11) As Long
'    txt(21) As String
'    typ As String
'    etq As Long
'    gpc As Long
'    trm1 As Long
'    trm2 As Long
'    prc1 As Double
'    prc2 As Double
'    iva As Integer
'    tecla As Integer
'    JuegoTecla As Integer
'    Merma As Integer
'    Presec As Integer
'    prc100g As Boolean
            articulo.plu = nR.Fields("plu")
            articulo.sec = nR.Fields("codi_ident")
            articulo.fam = nR.Fields("codi_fam")
            articulo.sub = nR.Fields("codi_sub")
            articulo.cad = nR.Fields("caducidad")
            articulo.pre = nR.Fields("Pref")
            articulo.tar = nR.Fields("tara")
            articulo.prc100g = nR.Fields("prc100g")
            articulo.gpc = nR.Fields("grupo_conserv")
            articulo.iva = nR.Fields("tipo_iva")
            articulo.Merma = nR.Fields("fcb")
            articulo.Presec = nR.Fields("ning")
            articulo.etq = nR.Fields("etq")
            'articulo.Balenv = nr.fields("balenv") '...Ver
            If Not IsNull(nR.Fields("art_cb")) Then
                articulo.cba = Trim(nR.Fields("art_cb"))
            Else
                articulo.cba = ""
            End If
            articulo.wgh = nR.Fields("codi_pes")
            If UsaEuro Then
                articulo.prc = nR.Fields("euros")
            Else
                articulo.prc = nR.Fields("precio")
            End If
            articulo.txt(0) = nR.Fields("des_plu1")
            articulo.txt(1) = nR.Fields("des_plu2")
            articulo.txt(2) = nR.Fields("des_plu3")
            articulo.txt(3) = nR.Fields("des_plu4")
            articulo.txt(4) = nR.Fields("des_plu5")
            articulo.txt(5) = nR.Fields("des_plu6")
            articulo.txt(6) = nR.Fields("des_plu7")
            articulo.txt(7) = nR.Fields("des_plu8")
            articulo.txt(8) = nR.Fields("des_plu9")
            articulo.txt(9) = nR.Fields("des_plu0")
            If lTransPluAsc = False Then
                articulo.txt(10) = nR.Fields("des_plux")
                articulo.txt(11) = nR.Fields("des_plu11")
                articulo.txt(12) = nR.Fields("des_plu12")
                articulo.txt(13) = nR.Fields("des_plu13")
                articulo.txt(14) = nR.Fields("des_plu14")
                articulo.txt(15) = nR.Fields("des_plu15")
                articulo.txt(16) = nR.Fields("des_plu16")
                articulo.txt(17) = nR.Fields("des_plu17")
                articulo.txt(18) = nR.Fields("des_plu18")
                articulo.txt(19) = nR.Fields("des_plu19")
                articulo.txt(20) = nR.Fields("des_plu20")
            End If
            articulo.tlt(0) = nR.Fields("tip_let1")
            articulo.tlt(1) = nR.Fields("tip_let2")
            articulo.tlt(2) = nR.Fields("tip_let3")
            articulo.tlt(3) = nR.Fields("tip_let4")
            articulo.tlt(4) = nR.Fields("tip_let5")
            articulo.tlt(5) = nR.Fields("tip_let6")
            articulo.tlt(6) = nR.Fields("tip_let7")
            articulo.tlt(7) = nR.Fields("tip_let8")
            articulo.tlt(8) = nR.Fields("tip_let9")
            articulo.tlt(9) = nR.Fields("tip_let0")
            articulo.prc1 = 0
            articulo.prc2 = 0
            articulo.trm1 = 0
            articulo.trm2 = 0
            
            articulo.lmix = nR.Fields("lmix")
            articulo.npeso = nR.Fields("poid")
            
            'articulo.prc3 = nr.fields("prc3") ...Ver
            
            If nR.Fields("usatramos") = True Then
                RegTramo.OpenRecordset ("select * from tramos where borrado=false and codigo=" & CStr(nC))
                If Not RegTramo.EOF Then
                    articulo.trm1 = RegTramo.Fields("tramo1")
                    articulo.trm2 = RegTramo.Fields("tramo2")
                    If UsaEuro Then
                        articulo.prc1 = RegTramo.Fields("euros1")
                        articulo.prc2 = RegTramo.Fields("euros2")
                    Else
                        articulo.prc1 = RegTramo.Fields("precio1")
                        articulo.prc2 = RegTramo.Fields("precio2")
                    End If
                End If
            End If
            
        'End With
            ''''''''''''''''''''''''
    End If
    
    nR.Close
    Set nR = Nothing
    nMybase.Close
    Set nMybase = Nothing
    'CadenadeLog "DB_Consulta_Art CÓDIGO:" & CStr(nC) & CStr(Resp)
    DB_Consulta_Art = Resp
End Function

Public Function separaFich(ByVal s As String) As String
   Dim Ls As Integer
   Dim nLs As Integer
   Dim Bucle As Integer
   Ls = Len(s)
   For Bucle = Ls To 1 Step -1
       If Mid(s, Bucle, 1) = "\" Then
           s = Mid(s, Bucle + 1)
           Bucle = Ls + 1
       End If
   Next Bucle
   separaFich = s
End Function


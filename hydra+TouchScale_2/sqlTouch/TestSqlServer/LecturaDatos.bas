Attribute VB_Name = "LecturaDatos"
Public Sub LeeImagenesArticulos()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    Dim strStream As New ADODB.Stream
    Dim MyBase As DAO.Database
    Dim MyRecArt As DAO.Recordset
    Dim s As String
    Dim n As Long
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.articuloimagen order by idarticulo", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset imágenes artículo..."
    Else
   
        strStream.Type = adTypeBinary
        strStream.Open
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                strStream.Write MyRecSet.Fields("imagen").Value
                strStream.SaveToFile App.Path & "\images\items\" & CStr(MyRecSet.Fields("idarticulo")), adSaveCreateOverWrite
                
                
                Set MyRecArt = MyBase.OpenRecordset("select * from articulo where codigo=" & CStr(MyRecSet.Fields("idarticulo")))
                If Not MyRecArt.EOF Then
                    MyRecArt.Edit
                    MyRecArt.Fields("imagen") = App.Path & "\images\items\" & CStr(MyRecSet.Fields("idarticulo"))
                    MyRecArt.Update
                End If
                n = n + 1
                MyRecSet.MoveNext
                
                
                strStream.Flush
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    MyRecArt.Close
    Set MyRecArt = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeImagenesFamilias()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    Dim strStream As New ADODB.Stream
    Dim n As Long
    Dim MyBase As DAO.Database
    Dim MyRecFam As DAO.Recordset
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.familia order by idfamilia", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset imágenes familias..."
    Else
   
        strStream.Type = adTypeBinary
        strStream.Open
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                strStream.Write MyRecSet.Fields("imagen").Value
                strStream.SaveToFile App.Path & "\images\families\" & CStr(MyRecSet.Fields("idfamilia")), adSaveCreateOverWrite
                Set MyRecFam = MyBase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(MyRecSet.Fields("idfamilia")))
                If Not MyRecFam.EOF Then
                    MyRecFam.Edit
                    MyRecFam.Fields("imagen") = App.Path & "\images\families\" & CStr(MyRecSet.Fields("idfamilia"))
                    MyRecFam.Update
                End If
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    MyRecFam.Close
    Set MyRecFam = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeImagenesVendedores()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    Dim strStream As New ADODB.Stream
    Dim n As Long
    Dim MyBase As DAO.Database
    Dim MyRecVnd As DAO.Recordset
    
    On Error Resume Next

    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.usuarioimagen order by idusuario", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset imágenes Vendedores..."
    Else
   
        strStream.Type = adTypeBinary
        strStream.Open
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                strStream.Write MyRecSet.Fields("imagen").Value
                strStream.SaveToFile App.Path & "\images\vendors\" & CStr(MyRecSet.Fields("idusuario")), adSaveCreateOverWrite
                Set MyRecVnd = MyBase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("idusuario")))
                If Not MyRecVnd.EOF Then
                    MyRecVnd.Edit
                    MyRecVnd.Fields("imagen") = App.Path & "\images\vendors\" & CStr(MyRecSet.Fields("idusuario"))
                    MyRecVnd.Update
                End If
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    
    MyRecVnd.Close
    Set MyRecVnd = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeArticulos()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim nPrc As Double
    Dim n As Long
    Dim sDesc As String
    Dim nLen As Long
    Dim nCont As Double
    Dim sParte As String
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.articulo order by idarticulo", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset Artículos..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from articulo where codigo=" & CStr(MyRecSet.Fields("idarticulo")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                nFam = GetFamByCodeArt(MyRecSet.Fields("idarticulo"))
                If nFam = 999999999 Then
                    nFam = 1
                End If
                nSec = GetSecByCodeFam(nFam)
                If nSec = 999999999 Then
                    nSec = 1
                End If
                nPlu = GetPluByCodeArtSec(MyRecSet.Fields("idarticulo"), nSec)
                If nPlu = 999999999 Then
                    nPlu = 0
                End If
                MyRec.Fields("codi_ident") = nSec
                MyRec.Fields("secc_maqui") = nSec
                MyRec.Fields("codi_sub") = 0
                MyRec.Fields("codi_fam") = nFam
                MyRec.Fields("plu") = nPlu
                MyRec.Fields("codigo") = MyRecSet.Fields("idarticulo")
                If MyRecSet.Fields("tipoventa") = 1 Then
                    MyRec.Fields("codi_pes") = "U"
                Else
                    MyRec.Fields("codi_pes") = "W"
                End If
                MyRec.Fields("tipo_iva") = MyRecSet.Fields("idimpuestoventa") - 155
                nPrc = GetPrcByCodeArt(MyRecSet.Fields("idarticulo"))
                If nPrc <> 999999999 Then
                    MyRec.Fields("euros") = nPrc
                Else
                    MyRec.Fields("euros") = 0
                End If
                MyRec.Fields("prc3") = Format(Now, "yymmddhhmm")
                MyRec.Fields("tara") = MyRecSet.Fields("taraventa")
                MyRec.Fields("tran_ep") = CStr(MyRecSet.Fields("taraenvasado"))
                MyRec.Fields("caducidad") = MyRecSet.Fields("caducidad")
                MyRec.Fields("pref") = MyRecSet.Fields("consumopreferente")
                MyRec.Fields("poid") = MyRecSet.Fields("pesoaproximado")
                MyRec.Fields("fcb") = MyRecSet.Fields("merma")
                MyRec.Fields("ning") = MyRecSet.Fields("unidadespreseleccion")
                MyRec.Fields("etq") = MyRecSet.Fields("tipoetiqueta")
                MyRec.Fields("des_plu1") = Mid(MyRecSet.Fields("descripcion"), 1, 25)
                
                If Not IsNull(Trim(MyRecSet.Fields("ingredientes"))) Then
                
                    sDesc = Trim(MyRecSet.Fields("ingredientes"))
                    nLen = Len(sDesc)
                    nCont = nLen / 25
                    
                    If nCont <> Int(nLen / 25) Then nCont = Int(nCont) + 1
                        
                    For nLen = 1 To nCont
                        
                        Select Case nLen
                        Case 1
                            MyRec.Fields("des_plu2") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 2
                            MyRec.Fields("des_plu3") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 3
                            MyRec.Fields("des_plu4") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 4
                            MyRec.Fields("des_plu5") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 5
                            MyRec.Fields("des_plu6") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 6
                            MyRec.Fields("des_plu7") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 7
                            MyRec.Fields("des_plu8") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 8
                            MyRec.Fields("des_plu9") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 9
                            MyRec.Fields("des_plu0") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 10
                            MyRec.Fields("des_plux") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 11
                            MyRec.Fields("des_plu11") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 12
                            MyRec.Fields("des_plu12") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 13
                            MyRec.Fields("des_plu13") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 14
                            MyRec.Fields("des_plu14") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 15
                            MyRec.Fields("des_plu15") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 16
                            MyRec.Fields("des_plu16") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 17
                            MyRec.Fields("des_plu17") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 18
                            MyRec.Fields("des_plu18") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 19
                            MyRec.Fields("des_plu19") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        Case 20
                            MyRec.Fields("des_plu20") = Mid(sDesc, ((nLen - 1) * 25) + 1, 25)
                        
                        
                        End Select
                        
                    
                    Next nLen
                        
                End If
                
                sParte = GetEANByCodeArt(MyRecSet.Fields("idarticulo"))
                If sParte <> "999999999" Then
                    MyRec.Fields("art_cb") = Mid(sParte, 1, 12)
                Else
                    MyRec.Fields("art_cb") = Space(12)
                End If
                
               
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub
Public Sub LeeRestoDatosFamilias()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    Dim nSec As Long
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.familia order by idfamilia", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset Familia..."
    Else
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(MyRecSet.Fields("idfamilia")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                
                MyRec.Fields("codi_fam") = MyRecSet.Fields("idfamilia")
                MyRec.Fields("txt_fam") = Mid(MyRecSet.Fields("descripcion"), 1, 25)
                nSec = GetSecByCodeFam(MyRecSet.Fields("idfamilia"))
                If nSec <> 999999999 Then
                    MyRec.Fields("codi_ident") = nSec
                Else
                    MyRec.Fields("codi_ident") = 1
                End If
               
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeFamilias()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.familiasseccion order by idfamilia", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset FamiliasSeccion-Familias..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from familia where codi_fam=" & CStr(MyRecSet.Fields("idfamilia")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                
                MyRec.Fields("codi_fam") = MyRecSet.Fields("idfamilia")
                MyRec.Fields("codi_ident") = MyRecSet.Fields("idseccion")
                MyRec.Fields("secc_maqui") = MyRecSet.Fields("idseccion")
                MyRec.Fields("codi_sub") = 0
               
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
    Call LeeRestoDatosFamilias
    
End Sub

Public Sub LeeVendedores()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.usuario order by idusuario", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset Vendedores..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("idusuario")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                
                MyRec.Fields("ident_vend") = MyRecSet.Fields("idusuario")
                MyRec.Fields("codi_vend") = MyRecSet.Fields("codigo")
                MyRec.Fields("nombre") = Mid(MyRecSet.Fields("nombre"), 1, 30)
               
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
    
End Sub

Public Sub LeeRestoDatosSeccion()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "select * from dbo.seccion order by IdSeccion", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset Seccion..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from seccion where codi_ident=" & CStr(MyRecSet.Fields("idseccion")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                
                MyRec.Fields("codi_ident") = MyRecSet.Fields("idseccion")
                MyRec.Fields("secc_maqui") = MyRecSet.Fields("idseccion")
                MyRec.Fields("descripcio") = MyRecSet.Fields("nombre")
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeSecciones()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    Dim cm As String
    
    'Dim Ffecha As Variant
    'Dim IdDoc As Variant
    'quitar
    'Ffecha = Now()
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If


'''quitar
'    MyRecSet.Open "select * from dbo.documento where SUBSTRING(convert(varchar(23),fecha,121),12,8)='00:00:00'", MyConnObj, adOpenKeyset, adLockOptimistic
'    If Not MyRecSet.EOF Then
'        MyRecSet.MoveFirst
'        Do While Not MyRecSet.EOF
'            Ffecha = MyRecSet.Fields("fecha")
'            IdDoc = MyRecSet.Fields("iddocumento")
'            'recoger datos
'
'
'
'
'
'            '''''''''''''''
'            'marcar...
'            cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
'            MyConnObj.Execute cm
'            ''''''''''
'
'            MyRecSet.Update
'
'            MyRecSet.MoveNext
'        Loop
'    End If
'
'''''''''
    
    
    MyRecSet.Open "select distinct(IdSeccion) from dbo.FamiliasSeccion order by IdSeccion", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset FamiliasSeccion-Secciones..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from seccion where codi_ident=" & CStr(MyRecSet.Fields("idseccion")))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                
                MyRec.Fields("codi_ident") = MyRecSet.Fields("idseccion")
                MyRec.Fields("secc_maqui") = MyRecSet.Fields("idseccion")
               
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    Call LeeRestoDatosSeccion
End Sub

Public Sub LeeIVAS()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim n As Long
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.tipoimpuesto order by idtipoimpuesto", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset TipoImpuestos..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from iva where tipo=" & CStr(MyRecSet.Fields("idtipoimpuesto") - 155))
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                MyRec.Fields("tipo") = MyRecSet.Fields("idtipoimpuesto") - 155
                MyRec.Fields("valor") = MyRecSet.Fields("porcentaje")
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

Public Sub LeeTramos()
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim MyBase As DAO.Database
    Dim MyRec As DAO.Recordset 'tabla articulo
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim n As Long
    
    On Error Resume Next
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbase.mdb")

    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "No es posible la Conexión..."
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.precio order by idarticulo", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "error en recordset Precios..."
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveFirst
            n = 1
            Do While Not MyRecSet.EOF
                Set MyRec = MyBase.OpenRecordset("select * from tramos where codigo=" & CStr(MyRecSet.Fields("idarticulo")) & " and tramo1=" & _
                CStr(MyRecSet.Fields("unidadesminimo")) & " and tramo2=" & CStr(MyRecSet.Fields("unidadesmaximo")))
                                            
                If MyRec.EOF Then
                    MyRec.AddNew
                Else
                    MyRec.Edit
                End If
                MyRec.Fields("codigo") = MyRecSet.Fields("idarticulo")
                MyRec.Fields("tramo1") = MyRecSet.Fields("unidadesminimo")
                MyRec.Fields("tramo2") = MyRecSet.Fields("unidadesmaximo")
                MyRec.Fields("euros1") = MyRecSet.Fields("preciodefecto")
                
                MyRec.Update
                
                n = n + 1
                MyRecSet.MoveNext
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing

    MyRec.Close
    Set MyRec = Nothing
    MyBase.Close
    Set MyBase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub
Public Function GetEANByCodeArt(ByVal nCode As Long) As String
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim nResul As String
    
    Dim n As Long
    
    On Error Resume Next
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = "999999999"
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.codigobarra where idarticulo=" & CStr(nCode), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = "999999999"
        GoTo FIN
    Else
   
   
        If Not MyRecSet.EOF Then
            nResul = MyRecSet.Fields("codigobarra")
        Else
            nResul = "999999999"
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
FIN:
    On Error GoTo 0
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    On Error GoTo 0
    
    GetEANByCodeArt = nResul
End Function

Public Function GetPrcByCodeArt(ByVal nCode As Long) As Double
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim nResul As Double
    
    Dim n As Long
    
    On Error Resume Next
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    End If
    
    MyRecSet.Open "Select * from dbo.precio where idarticulo=" & CStr(nCode) & " and unidadesminimo=0", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    Else
   
   
        If Not MyRecSet.EOF Then
            nResul = MyRecSet.Fields("preciodefecto")
        Else
            nResul = 999999999
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
FIN:
    On Error GoTo 0
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    On Error GoTo 0
    
    GetPrcByCodeArt = nResul
End Function

Public Function GetFamByCodeArt(ByVal nCode As Long) As Long
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim nResul As Long
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    End If
    MyRecSet.Open "Select * from dbo.articulosfamilia where idarticulo=" & CStr(nCode), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    Else
        If Not MyRecSet.EOF Then
            nResul = MyRecSet.Fields("idfamilia")
        Else
            nResul = 999999999
        End If
    End If

FIN:
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    On Error GoTo 0
    
    GetFamByCodeArt = nResul

End Function

Public Function GetSecByCodeFam(ByVal nFam As Long) As Long
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim nResul As Long
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    End If
    MyRecSet.Open "Select * from dbo.familiasseccion where idfamilia=" & CStr(nFam), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    Else
        If Not MyRecSet.EOF Then
            nResul = MyRecSet.Fields("idseccion")
        Else
            nResul = 999999999
        End If
    End If

FIN:
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    On Error GoTo 0
    
    GetSecByCodeFam = nResul

End Function

Public Function GetPluByCodeArtSec(ByVal nCode As Long, ByVal nSec As Long) As Long
    Dim MyConnObj As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset 'tabla articulo SqlServer
    Dim nResul As Long
    MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    End If
    MyRecSet.Open "Select * from dbo.articuloreferencia where idarticulo=" & CStr(nCode) & " and idseccion=" & CStr(nSec), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FIN
    Else
        If Not MyRecSet.EOF Then
            nResul = MyRecSet.Fields("referencia")
        Else
            nResul = 999999999
        End If
    End If

FIN:
    
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    MyConnObj.Close
    Set MyConnObj = Nothing
    On Error GoTo 0
    
    GetPluByCodeArtSec = nResul

End Function


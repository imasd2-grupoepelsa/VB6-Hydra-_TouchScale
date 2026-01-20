Attribute VB_Name = "Global_Dat"
'//////////////////////////////////////////////
' este módulo importa el fichero "GLOBAL.DAT"
' original de las herramientas "envía" y "recibe"
' dicho fichero contiene oredenes con formato
' fijo para importar:
' Secciones, subsecciones y familias
' Equipos
' Artículos con líneas de descriptivo
' Códigos de Barras
' Cabeceras y leyendas
' Publicidad y Ofertas
' Teclas rápidas
Option Explicit
Dim sTxtTmp As String

Public lUpper As Boolean
Public lTransPlu As Boolean
Public lCOCOBIO As Boolean
Public lAcumulaSS As Boolean
Public lAcumulaET As Boolean
Public lhydraexe As Boolean
Public lAlba As Boolean

Public Type Tipo_Buf
    NSecciones As Integer
    NEquipos As Integer
    NVendedores As Integer
    NArticulos As Integer
    NTeclas As Integer
    NCodBar As Integer
    NPubli As Integer
    NOfer As Integer
    NTipoPubli As Integer
    NCabeceras As Integer
    NLeyendas As Integer
    NTextos As Integer
    NPaises As Integer
    NFichas As Integer
    nEAN As Integer
    nImage As Integer
    nTarifas As Integer
    nEtiqueta As Integer
    nText15 As Integer
    nFichNutri As Integer
    nVitaminas As Integer
    nFichaArt As Integer
    BufSecciones() As String
    BufEquipos() As String
    BufVendedores() As String
    BufArticulos() As String
    BufTeclas() As String
    BufCodBar() As String
    BufPubli() As String
    BufOfer() As String
    BufTipoPubli() As String
    BufCabeceras() As String
    BufLeyendas() As String
    BufTextos() As String
    BufPaises() As String
    BufFichas() As String
    BufEAN() As String
    BufImagen() As String
    BufTarifas() As String
    BufEtiqueta() As String
    BufText15() As String
    BufFichaNutri() As String
    BufVitaminas() As String
    BufNutriArt() As String
End Type
Public Sub ImportTeclas(ByVal sS As String)
    Dim Arch1 As Integer
    Dim Buffer As String
    Dim Base As dao.Database
    Dim MyCod As Long
    Dim MySec As Integer
    Dim MyTabla As Integer
    Dim myTec As Integer
    Dim Registro As dao.Recordset
    Dim sDes As String
        
    If Dir(sS) <> "" Then
        If FileLen(sS) > 0 Then
            FrmBalGloDat.Caption = "Leyendo Teclas..."
            Load FrmBalGloDat
            FrmBalGloDat.Show
            FrmBalGloDat.dato "Leyendo Archivo ..."
        
            Set Base = dao.OpenDatabase(App.Path & "\dbasetouch.mdb")
            Arch1 = FreeFile()
            Open sS For Input As #Arch1
            Do While Not EOF(Arch1)
            Line Input #Arch1, Buffer
                If Len(Buffer) = 13 Then
                    MySec = Mid(Buffer, 1, 2)
                    MyTabla = Mid(Buffer, 3, 2)
                    myTec = Mid(Buffer, 5, 3)
                    MyCod = Mid(Buffer, 8)
                    If MyCod <> 0 Then
                        Set Registro = Base.OpenRecordset("select codigo,des_plu1 from articulo where codigo=" & CStr(MyCod))
                        If Not Registro.EOF Then
                            sDes = Registro.Fields("des_plu1")
                        Else
                            sDes = ""
                        End If
                        Registro.Close
                        Set Registro = Nothing
                        Base.Execute "delete from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'"
                        Set Registro = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'")
                        If Registro.EOF Then
                            Registro.AddNew
                        Else
                            Registro.Edit
                        End If
                        
                        FrmBalGloDat.dato "Sec.:" & CStr(MySec) & " Tabla:" & CStr(MyTabla) & " Tecla:" & CStr(myTec) & " Cod.:" & CStr(MyCod)
                        
                        Registro.Fields("codi_ident") = MySec
                        Registro.Fields("secc_maqui") = MySec
                        Registro.Fields("tran_tecsc10") = MySec
                        Registro.Fields("codigo") = MyCod
                        Registro.Fields("plu") = 0
                        Registro.Fields("desc_plu") = sDes
                        Registro.Fields("secc_eqp") = False
                        Registro.Fields("numero_eqp") = False
                        Registro.Fields("tabla") = MyTabla
                        Registro.Fields("codi_tec") = myTec
                        Registro.Fields("prog_tec") = "TOUCH"
                        Registro.Fields("tran_tec") = ""
                        Registro.Fields("borrado") = False
                        
                        Registro.Update
                        Registro.Close
                        Set Registro = Nothing
                    Else
                        'Base.Execute "delete from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'"
                        Set Registro = Base.OpenRecordset("select * from teclas where secc_maqui=" & CStr(MySec) & " and codi_tec=" & CStr(myTec) & " and tabla=" & CStr(MyTabla) & " and TRAN_TECSC10='" & CStr(MySec) & "'")
                        If Not Registro.EOF Then
                            Registro.Edit
                            Registro.Fields("borrado") = True
                            Registro.Update
                        End If
                        Registro.Close
                        Set Registro = Nothing
                        
                        FrmBalGloDat.dato "Sec.:" & CStr(MySec) & " Tabla:" & CStr(MyTabla) & " Tecla:" & CStr(myTec) & " Baja."
                    End If
                End If
            Loop
            Close #Arch1
            Sleep (10)
            Kill sS
            Sleep (10)
            Base.Close
            Set Base = Nothing
            Unload FrmBalGloDat
        End If
    End If

End Sub
Public Function Importa_Global_Dat()
    Dim Arch1 As Integer
    Dim Arch2 As Integer
    Dim Buffer As String
    Dim MyBuf As String
    Dim Tipo As String ' tipo : artivulos, vendedores...
    Dim Modo As String ' alta, baja ....
    Dim MyOK As Boolean
    Dim nCli As Integer
    Dim nFCli As Integer

    
On Error GoTo fin
    nCli = 0

    If lUpper = False Then
        Load FrmBalGloDat
        FrmBalGloDat.Show
        FrmBalGloDat.dato "Leyendo Archivo ..."
    End If
        
    Dim BufDatos As Tipo_Buf
    Dim sS As Variant
    Dim ffff As Variant
    Dim pppp As Integer
    Dim sdesfam As String
    Dim simgfam As String
    Dim ssSecs(1) As Long
    Dim nResp As Variant
    
    If Dir(App.Path & "\global.dat") <> "" Or (lUpper And Dir(App.Path & "\globaltouch.dat") <> "") Then
        Arch1 = FreeFile()
        If (lUpper And Dir(App.Path & "\globaltouch.dat") <> "") Then
            Open App.Path & "\globaltouch.dat" For Input As #Arch1
        Else
            Open App.Path & "\global.dat" For Input As #Arch1
        End If
        Arch2 = FreeFile()
        Open App.Path & "\globaltouchno.dat" For Output As #Arch2
        Do Until EOF(Arch1)
            Line Input #Arch1, Buffer
            If Len(Buffer) >= 2 Then
                Tipo = left(Buffer, 1)
                Modo = Mid(Buffer, 2, 1)
                Buffer = Mid(Buffer, 2)
                MyOK = False
                If (Modo = "0" Or Modo = "1" Or Modo = "2") Or (Tipo = "Z" Or Tipo = "=" Or Tipo = "M" Or Tipo = "O") Or Tipo = "Y" Or Tipo = "C" Or Tipo = "U" Or Tipo = "T" Then
                    Select Case Tipo
                        Case "0" ' secciones
                            If lUpper Then
                            Else
                                If Len(Buffer) = 9 Or Len(Buffer) = 34 Then
                                    MyOK = True
                                    BufDatos.NSecciones = BufDatos.NSecciones + 1
                                    ReDim Preserve BufDatos.BufSecciones(BufDatos.NSecciones)
                                    BufDatos.BufSecciones(BufDatos.NSecciones - 1) = Buffer
                                End If
                            End If
                        Case "1" ' subsecciones (ignorado)
                        Case "2" ' familias (ignorado)
                            'If lUpperNW Then
                            'If lUpper Then 'AtencionFamilias

'"2" SS FF Decripción Imágen
'donde:
'
' sS = Sección
' FF = familia
' Descripción=30 caracteres
' Imágen=Nombre del fichero
'20101FRUTAS                        fruta.png

                                If Len(Buffer) >= 40 Then
                                    MyOK = True
                                    If Len(Buffer) > 36 Then
                                        simgfam = Trim(Mid(Buffer, 35))
                                    Else
                                        simgfam = ""
                                    End If
                                    ssSecs(0) = Mid(Buffer, 1, 2)
                                    sS = Mid(Buffer, 1, 2)
                                    ffff = Mid(Buffer, 3, 2)
                                    sdesfam = Trim(Mid(Buffer, 5, 30))
                                    pppp = ffff
                                    nResp = Alta_Familias(ffff, sdesfam, 0, "", False, , pppp, simgfam)
                                    nResp = Alta_Familias_Counter(ffff, sdesfam, 0, ssSecs(), False, , simgfam, sS)
                                End If
                            'End If
                        Case "3" ' vendedores (por sección, de momento)
                            'If Len(Buffer) = 16 Then
                            '    MyOK = True
                            '    BufDatos.NVendedores = BufDatos.NVendedores + 1
                            '    ReDim Preserve BufDatos.BufVendedores(BufDatos.NVendedores)
                            '    BufDatos.BufVendedores(BufDatos.NVendedores - 1) = Buffer
                            'End If
                        Case "4" ' Equipos
                            If lUpper Then
                            Else
                            
                                If Len(Buffer) = 5 Or Len(Buffer) = 22 Then
                                    MyOK = True
                                    BufDatos.NEquipos = BufDatos.NEquipos + 1
                                    ReDim Preserve BufDatos.BufEquipos(BufDatos.NEquipos)
                                    BufDatos.BufEquipos(BufDatos.NEquipos - 1) = Buffer
                                End If
                            End If
                        Case "5" ' Articulos'
                            If Len(Buffer) = 60 Or Len(Buffer) = 62 Or Len(Buffer) = 86 Or Len(Buffer) = 91 Or Len(Buffer) = 92 Then
                                MyOK = True
                                BufDatos.NArticulos = BufDatos.NArticulos + 1
                                ReDim Preserve BufDatos.BufArticulos(BufDatos.NArticulos)
                                BufDatos.BufArticulos(BufDatos.NArticulos - 1) = Buffer
Else
MyOK = MyOK
                            End If
                         'Case "6" ' teclas
                        '    If Len(Buffer) = 10 Then
                        '        MyOK = True
                        '        BufDatos.NTeclas = BufDatos.NTeclas + 1
                        '        ReDim Preserve BufDatos.BufTeclas(BufDatos.NTeclas)
                        '        BufDatos.BufTeclas(BufDatos.NTeclas - 1) = Buffer
                        '    End If
                        Case "7" 'cod. barras
                            If Len(Buffer) = 64 Or Len(Buffer) = 124 Then
                                'If lAgora Then
                                '    MyOK = True
                                'Else
                                '    MyOK = True
                                '    BufDatos.NCodBar = BufDatos.NCodBar + 1
                                '    ReDim Preserve BufDatos.BufCodBar(BufDatos.NCodBar)
                                '    BufDatos.BufCodBar(BufDatos.NCodBar - 1) = Buffer
                                'End If
                                MyOK = True
                                If lAgora = False Then
                                    BufDatos.NCodBar = BufDatos.NCodBar + 1
                                    ReDim Preserve BufDatos.BufCodBar(BufDatos.NCodBar)
                                    BufDatos.BufCodBar(BufDatos.NCodBar - 1) = Buffer
                                Else
                                    On Error Resume Next
                                    'Close #Arch1
                                    'End
                                End If
                                
                            End If
                        'Case "8" 'publi. continua
                        '    If Len(Buffer) = 104 Then
                        '        MyOK = True
                        '        BufDatos.NPubli = BufDatos.NPubli + 1
                        '        ReDim Preserve BufDatos.BufPubli(BufDatos.NPubli)
                        '        BufDatos.BufPubli(BufDatos.NPubli - 1) = Buffer
                        '    End If
                        'Case "9" 'publi. ofertas
                        '    If Len(Buffer) = 164 Then
                        '        MyOK = True
                        '        BufDatos.NOfer = BufDatos.NOfer + 1
                        '        ReDim Preserve BufDatos.BufOfer(BufDatos.NOfer)
                        '        BufDatos.BufOfer(BufDatos.NOfer - 1) = Buffer
                        '    End If
                        'Case ":" 'tipo publi
                        '    If Len(Buffer) = 5 Then
                        '        MyOK = True
                        '        BufDatos.NTipoPubli = BufDatos.NTipoPubli + 1
                        '        ReDim Preserve BufDatos.BufTipoPubli(BufDatos.NTipoPubli)
                        '        BufDatos.BufTipoPubli(BufDatos.NTipoPubli - 1) = Buffer
                        '    End If
                        Case ";" 'cabeceras
                            If Len(Buffer) = 135 Or Len(Buffer) = 134 Then
                                MyOK = True
                                BufDatos.NCabeceras = BufDatos.NCabeceras + 1
                                ReDim Preserve BufDatos.BufCabeceras(BufDatos.NCabeceras)
                                BufDatos.BufCabeceras(BufDatos.NCabeceras - 1) = Buffer
                            End If
                        Case "<" 'leyendas
                            If Len(Buffer) = 135 Or Len(Buffer) = 134 Then
                                MyOK = True
                                BufDatos.NLeyendas = BufDatos.NLeyendas + 1
                                ReDim Preserve BufDatos.BufLeyendas(BufDatos.NLeyendas)
                                BufDatos.BufLeyendas(BufDatos.NLeyendas - 1) = Buffer
                            End If
                        Case "@" ' alta teclas
                            If Len(Buffer) = 16 Or Len(Buffer) = 15 Then
                                MyOK = True
                                BufDatos.NTeclas = BufDatos.NTeclas + 1
                                ReDim Preserve BufDatos.BufTeclas(BufDatos.NTeclas)
                                BufDatos.BufTeclas(BufDatos.NTeclas - 1) = Buffer
                            End If
                        
                         Case "C" ' Clientes
                            If nCli = 0 Or Dir(App.Path & "\clientes.dat") = "" Then
                                If Dir(App.Path & "\clientes.dat") <> "" And nCli = 0 Then
                                    Kill App.Path & "\clientes.dat"
                                    Sleep (200)
                                End If
                                If Dir(App.Path & "\clientes.dat") = "" Then
                                    nFCli = FreeFile()
                                    Open App.Path & "\clientes.dat" For Output As #nFCli
                                    Close #nFCli
                                    Sleep (200)
                                End If
                                nCli = 255
                            End If
                            nFCli = FreeFile()
                            Open App.Path & "\clientes.dat" For Append As #nFCli
                            Print #nFCli, Mid(Buffer, 2)
                            Close #nFCli
                            MyOK = True
                       
                        Case "E" 'Alta etiquetas
                            MyOK = True
                            BufDatos.nEtiqueta = BufDatos.nEtiqueta + 1
                            ReDim Preserve BufDatos.BufEtiqueta(BufDatos.nEtiqueta)
                            BufDatos.BufEtiqueta(BufDatos.nEtiqueta - 1) = Buffer
                       
                        Case "F" '1.9.7
                            'If lUpperNW Then
                            'If lUpper Then 'AtencionFamilias
                                'If Len(Buffer) >= 39 Then
                                    MyOK = True
                                    If Len(Buffer) >= 40 Then
                                        simgfam = Trim(Mid(Buffer, 40))
                                    Else
                                        simgfam = ""
                                    End If
                                    ssSecs(0) = Mid(Buffer, 2, 2)
                                    sS = Mid(Buffer, 2, 2)
                                    ffff = Mid(Buffer, 4, 3)
                                    pppp = Mid(Buffer, 7, 3)
                                    sdesfam = Trim(Mid(Buffer, 10, 30))
                                    If Mid(Buffer, 1, 1) = "0" Then
                                        nResp = Alta_Familias(ffff, sdesfam, 0, "", False, , pppp, simgfam)
                                        nResp = Alta_Familias_Counter(ffff, sdesfam, 0, ssSecs(), False, , simgfam, sS, pppp)
                                        'If nResp <> 0 And nResp <> 1 And nResp <> 4 Then
                                            Print #Arch2, "F" & Buffer
                                            FrmBalGloDat.dato "FAM.            " & Buffer
                                        'Else
                                        '    FrmBalGloDat.dato "OK -->   - Fam. " & Buffer
                                        'End If
                                    Else
                                        nResp = Baja_Familias(ffff)
                                        'If nResp <> 0 And nResp <> 1 And nResp <> 4 Then
                                            Print #Arch2, "F" & Buffer
                                            FrmBalGloDat.dato "FAM.            " & Buffer
                                        'Else
                                        '    FrmBalGloDat.dato "OK -->   - Fam. " & Buffer
                                        'End If
                                    
                                    End If
                                'End If
                        
                        Case "G" 'Vendedores
                            MyOK = True
                            BufDatos.NVendedores = BufDatos.NVendedores + 1
                            ReDim Preserve BufDatos.BufVendedores(BufDatos.NVendedores)
                            BufDatos.BufVendedores(BufDatos.NVendedores - 1) = Buffer
''''''''''''''''''''''''''''''''
'sCard compuesta por 6 dig. número de ficha y a continuación valores con 6 dig. (sin signo decimal, dos dígitos menos significativos serán los decimales) para cada uno de los campos:
'carbohydrate_100//carbohydrate_portion//energy_kcal_100//energy_kcal_portion//
'energy_kj_100//energy_kj_portion//fat_100//fat_portion//fibre_100//fibre_portion//mono_unsaturates_100//mono_unsaturates_portion//polyols_100//
'polyols_portion//polyunsaturates_100//polyunsaturates_portion//portion_weight//protein_100//
'protein_portion//salt_100//salt_portion//'saturates_100//saturates_portion//
'starch_100//starch_portion//sugars_100//sugars_portion
                        Case "I"
                            MyOK = True
                            BufDatos.nFichNutri = BufDatos.nFichNutri + 1
                            ReDim Preserve BufDatos.BufFichaNutri(BufDatos.nFichNutri)
                            BufDatos.BufFichaNutri(BufDatos.nFichNutri - 1) = Buffer
'sVita_Mine compuesto por 6 dig. número de ficha,
'2 dig. identificador de vitamina o mineral, 6 dig. cantidad por 100g (sin signo decimal,
'dos dígitos menos significativos serán los decimales) y 6 dig. cantidad por porción (sin signo decimal,
'dos dígitos menos significativos serán los decimales)
'ejemplo:     “00002008002500001300”
'            Ficha: 20
'            Identificador: 8
'            cantidad 100g: 25
'            cantidad porción: 13
                        Case "J"
                            MyOK = True
                            BufDatos.nVitaminas = BufDatos.nVitaminas + 1
                            ReDim Preserve BufDatos.BufVitaminas(BufDatos.nVitaminas)
                            BufDatos.BufVitaminas(BufDatos.nVitaminas - 1) = Buffer
'    sItemCard compuesta por 6 dig. Código de artículo y 6 dig. número de Ficha
'    ejemplo: “000001000020” --> Código artículo 1, Ficha 20
                        Case "K"
                            MyOK = True
                            BufDatos.nFichaArt = BufDatos.nFichaArt + 1
                            ReDim Preserve BufDatos.BufNutriArt(BufDatos.nFichaArt)
                            BufDatos.BufNutriArt(BufDatos.nFichaArt - 1) = Buffer
''''''''''''''''''''''''''''''''
                        'Case "=" 'adicionales artículo antiguo (ignorado)
                        '2.0.34
                        Case "M"
                            'bloque 2
                            'If Len(Buffer) = 266 Or Len(Buffer) = 256 Or lUpper Then
                                MyOK = True
                                BufDatos.NTextos = BufDatos.NTextos + 1
                                ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                                BufDatos.BufTextos(BufDatos.NTextos - 1) = "M" & Buffer
                            'End If
                        Case "O"
                            'bloque3
                            'If Len(Buffer) = 266 Or Len(Buffer) = 256 Or lUpper Then
                                MyOK = True
                                BufDatos.NTextos = BufDatos.NTextos + 1
                                ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                                BufDatos.BufTextos(BufDatos.NTextos - 1) = "O" & Buffer
                            'End If
                        ''''''''''
                        Case "P" ' paises (nueva)
                            If Len(Buffer) = 24 Then
                                MyOK = True
                                BufDatos.NPaises = BufDatos.NPaises + 1
                                ReDim Preserve BufDatos.BufPaises(BufDatos.NPaises)
                                BufDatos.BufPaises(BufDatos.NPaises - 1) = Buffer
                            End If
                        
                        Case "T" 'Alta text15
                            MyOK = True
                            BufDatos.nText15 = BufDatos.nText15 + 1
                            ReDim Preserve BufDatos.BufText15(BufDatos.nText15)
                            BufDatos.BufText15(BufDatos.nText15 - 1) = Buffer
                        
                        Case "U" 'Alta imágenes
                            MyOK = True
                            BufDatos.nImage = BufDatos.nImage + 1
                            ReDim Preserve BufDatos.BufImagen(BufDatos.nImage)
                            BufDatos.BufImagen(BufDatos.nImage - 1) = Buffer
                        
                        Case "V" ' fichas vacuno (nueva)
                            If Len(Buffer) = 159 Then
                                MyOK = True
                                BufDatos.NFichas = BufDatos.NFichas + 1
                                ReDim Preserve BufDatos.BufFichas(BufDatos.NFichas)
                                BufDatos.BufFichas(BufDatos.NFichas - 1) = Buffer
                            End If
                        
                        Case "W" ' EAN Artículo (nueva)
                            'If Len(Buffer) = 19 Then
                                MyOK = True
                                BufDatos.nEAN = BufDatos.nEAN + 1
                                ReDim Preserve BufDatos.BufEAN(BufDatos.nEAN)
                                BufDatos.BufEAN(BufDatos.nEAN - 1) = Buffer
                            'End If
                        
                        Case "X" 'tarifas
                            MyOK = True
                            BufDatos.nTarifas = BufDatos.nTarifas + 1
                            ReDim Preserve BufDatos.BufTarifas(BufDatos.nTarifas)
                            BufDatos.BufTarifas(BufDatos.nTarifas - 1) = Buffer
                        
                        Case "Y" 'adicionales artículo (11-20) (Nueva)
                            ''bloque 1
                            ''If Len(Buffer) = 256 Or lUpper Then
                            '    MyOK = True
                            '    BufDatos.NTextos = BufDatos.NTextos + 1
                            '    ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                            '    BufDatos.BufTextos(BufDatos.NTextos - 1) = "Y" & Buffer
                            ''End If
                            MyOK = True
                            If lAgora = False Then
                        
                                If Len(Buffer) = 256 Or Len(Buffer) > 266 Then
                                    MyOK = True
                                    'If Len(Buffer) > 266 Then l34 = True
                                    BufDatos.NTextos = BufDatos.NTextos + 1
                                    ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                                    BufDatos.BufTextos(BufDatos.NTextos - 1) = "Y" & Buffer
                                End If
                            
                            End If
                            
                        Case "Z", "=" 'adicionales artículo
                            ''bloque 0
                            ''If Len(Buffer) = 266 Or Len(Buffer) = 256 Or lUpper Then
                            '    MyOK = True
                            '    BufDatos.NTextos = BufDatos.NTextos + 1
                            '    ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                            '    BufDatos.BufTextos(BufDatos.NTextos - 1) = "Z" & Buffer
                            ''Else
                            ''    'If Len(Buffer) = 215 Then c2f 1.8.2
                            ''    If Len(Buffer) = 214 Then
                            ''        MyOK = True
                            ''       BufDatos.NTextos = BufDatos.NTextos + 1
                            ''        ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                            ''        BufDatos.BufTextos(BufDatos.NTextos - 1) = "=" & Buffer & "0" & Space(25) & "0" & Space(25)
                            ''    End If
                            ''End If
                            If Len(Buffer) = 266 Or Len(Buffer) = 256 Or Len(Buffer) > 266 Then
                                MyOK = True
                                If lAgora = False Then
                                    'If Len(Buffer) > 266 Then l34 = True
                                    BufDatos.NTextos = BufDatos.NTextos + 1
                                    ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                                    BufDatos.BufTextos(BufDatos.NTextos - 1) = "Z" & Buffer
                                End If
                            Else
                                'If Len(Buffer) = 215 Then c2f 1.8.2
                                If Len(Buffer) = 214 Then
                                    MyOK = True
                                    If lAgora = False Then
                                    
                                        BufDatos.NTextos = BufDatos.NTextos + 1
                                        ReDim Preserve BufDatos.BufTextos(BufDatos.NTextos)
                                        BufDatos.BufTextos(BufDatos.NTextos - 1) = "=" & Buffer & "0" & Space(25) & "0" & Space(25)
                                    
                                    End If
                                End If
                            End If
                        
                        Case "*" 'comentario
                            MyOK = True
                            
                            
                    End Select
                    If MyOK = False Then Print #Arch2, Buffer
                Else
                    If lUpper = False Then
                        FrmBalGloDat.dato "NOK --> " & Buffer
                    End If
                    Print #Arch2, Buffer
                End If
            Else
                If lUpper = False Then
                    FrmBalGloDat.dato "NOK --> " & Buffer
                End If
                Print #Arch2, Buffer
            End If
        Loop
        Close #Arch1
        Close #Arch2
        Procesa_Datos BufDatos
        If nCli <> 0 Then ' procesar clientes...
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "clientes.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                frmControl.ControlImport1.Importar_Background "clientes.dat"
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
                Corregir_Nulos
        End If
    Else
        If lUpper = False Then
            FrmBalGloDat.dato "No se encontró el archivo"
        End If
        Do_Events
    End If
    
fin:


If Err.Number <> 0 Then
    Buffer = "Error:" & CStr(Err.Number) & " --> " & Err.Description
    Buffer = Buffer
    CadenadeLog Buffer
    'If Err.Number = 55 Then
        For Arch1 = 1 To 255
            Close #Arch1
        Next Arch1
    'End If
End If
On Error GoTo 0

    Call Corregir_Nulos
    If Dir(App.Path & "\asign.tst") <> "" Then
        Call AsignTec
    End If
    
    Sleep (2000)
    If lUpper = False Then
        Unload FrmBalGloDat
    End If
End Function
Private Sub Procesa_Datos(BufDatos As Tipo_Buf)
    '1 º secciones
    If BufDatos.NSecciones > 0 Then
        Procesa_Secciones BufDatos
    End If
    '2 º equipos
    If BufDatos.NEquipos > 0 Then
        Procesa_Equipos BufDatos
    End If
    '3 º códigos de barras
    If BufDatos.NCodBar > 0 Then
        Procesa_CodBar BufDatos
    End If
    '4 º cabeceras y leyendas
    'If BufDatos.NCabeceras > 0 Then
    '    Procesa_Cabeceras BufDatos
    'End If
    'If BufDatos.NLeyendas > 0 Then
    '    Procesa_Leyendas BufDatos
    'End If
    '5 º publicidad y ofertas
    'If BufDatos.NPubli > 0 Then
    '    Procesa_Publi BufDatos
    'End If
    'If BufDatos.NOfer > 0 Then
    '    Procesa_Ofer BufDatos
    'End If
    'If BufDatos.NTipoPubli > 0 Then
    '    Procesa_TipoPubli BufDatos
    'End If
    '6 º vendedores (por sección, de momento)
    If BufDatos.NVendedores > 0 Then
        Procesa_Vendedores BufDatos
    End If
    '7 º artículos
    If BufDatos.NArticulos > 0 Then
        Procesa_Articulos BufDatos
    End If
    'W EAN artículos
    If BufDatos.nEAN > 0 Then
        Procesa_EanArticulos BufDatos
    End If
    'Textos adicionales...2.0.34...Se modifica para que se atienda "M" y "O" 21 a 30 y 31 a 40.
    If BufDatos.NTextos > 0 Then
        Procesa_Textos BufDatos
    End If
    '8 º teclas
    If BufDatos.NTeclas > 0 Then
        Procesa_Teclas BufDatos
    End If
    'Imagenes artículos
    If BufDatos.nImage > 0 Then
        Procesa_Imagen BufDatos
    End If
    'Etiqueta artículos
    If BufDatos.nEtiqueta > 0 Then
        Procesa_Etiqueta BufDatos
    End If
    'TEXT15
    If BufDatos.nText15 > 0 Then
        Procesa_TEXT15 BufDatos
    End If
    
    '9 º paises
    If BufDatos.NPaises > 0 Then
        Procesa_Paises BufDatos
    End If
    '10 º fichas
    If BufDatos.NFichas > 0 Then
        Procesa_Fichas BufDatos
    End If
    If BufDatos.nTarifas > 0 Then
        Procesa_Tarifas BufDatos
    End If
''''''''''''''''''''''''
    If BufDatos.nFichNutri > 0 Then
        Procesa_FichaNutri BufDatos
    End If
    If BufDatos.nVitaminas > 0 Then
        Procesa_Vitaminas BufDatos
    End If
    If BufDatos.nFichaArt > 0 Then
        Procesa_FichaArt BufDatos
    End If


'''''''''''''''''''''''
End Sub
Private Sub Procesa_FichaNutri(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from nutrition where card<>0")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nFichNutri - 1

        MyCod = Val(Mid(BufDatos.BufFichaNutri(bucle), 2, 6))
        If MyCod = 0 Then
            Print #Arch2, BufDatos.BufFichaNutri(bucle)
            FrmBalGloDat.dato "NOK -->  + Ficha Nutricional: " & BufDatos.BufFichaNutri(bucle)
        End If
        Set Registro = Base.OpenRecordset("select * from nutrition where card=" & MyCod)
        
        If Mid(BufDatos.BufFichaNutri(bucle), 1, 1) = "0" Then
            If Registro.EOF Then
                Registro.AddNew
            Else
                Registro.Edit
            End If
                
            With Registro
                .Fields("card") = MyCod
                .Fields("cadena") = Mid(BufDatos.BufFichaNutri(bucle), 2)
                .Fields("borrado") = " "
                .Fields("tran_nutri") = " "
                
                .Update
                FrmBalGloDat.dato "OK -->  + Ficha Nutricional: " & BufDatos.BufFichaNutri(bucle)
            End With

        Else
            If Not Registro.EOF Then
                Registro.Edit
                Registro.Fields("borrado") = "*"
                Registro.Fields("tran_nutri") = " "
                Registro.Update
            End If
        End If
    Next bucle
    Close #Arch2
    Base.Close

End Sub
Private Sub Procesa_Vitaminas(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from nutrition where card<>0")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nVitaminas - 1

        MyCod = Val(Mid(BufDatos.BufVitaminas(bucle), 2, 6))
        If MyCod = 0 Then
            Print #Arch2, BufDatos.BufVitaminas(bucle)
            FrmBalGloDat.dato "NOK -->  + Vitaminas: " & BufDatos.BufVitaminas(bucle)
        End If
        Set Registro = Base.OpenRecordset("select * from vitamins where card=" & MyCod)
        
        If Mid(BufDatos.BufVitaminas(bucle), 1, 1) = "0" Then
            If Registro.EOF Then
                Registro.AddNew
            Else
                Registro.Edit
            End If
                
            With Registro
                .Fields("card") = MyCod
                .Fields("cadena") = Mid(BufDatos.BufVitaminas(bucle), 2)
                .Fields("borrado") = " "
                .Fields("tran_vita") = " "
                
                .Update
                FrmBalGloDat.dato "OK -->  + Vitaminas: " & BufDatos.BufVitaminas(bucle)
            End With

        Else
            If Not Registro.EOF Then
                Registro.Edit
                Registro.Fields("borrado") = "*"
                Registro.Fields("tran_vita") = " "
                Registro.Update
            End If
        End If
    Next bucle
    Close #Arch2
    Base.Close
End Sub
Private Sub Procesa_FichaArt(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo where codigo<>0")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nFichaArt - 1

        MyCod = Val(Mid(BufDatos.BufNutriArt(bucle), 2, 6))
        
        Registro.FindFirst "codigo=" & MyCod & " and borrado=false"

        If Registro.EOF Then
            Print #Arch2, BufDatos.BufNutriArt(bucle)
            FrmBalGloDat.dato "NOK -->  + Ficha Art.: " & BufDatos.BufNutriArt(bucle)
        Else
            With Registro
                Edit_Record Registro
                .Fields("posicion") = Mid(BufDatos.BufNutriArt(bucle), 8, 6)
                .Fields("tran_plu") = ""
                .Update
                FrmBalGloDat.dato "OK -->  + Ficha Art.: " & BufDatos.BufNutriArt(bucle)
            End With
        End If
    Next bucle
    Close #Arch2
    Base.Close
End Sub

Private Sub Procesa_Teclas(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim myTec As DB_Teclas
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim Reg1 As dao.Recordset
    Dim Reg2 As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        
        For bucle = 0 To BufDatos.NTeclas
            
            If left(BufDatos.BufTeclas(bucle), 1) = "0" And BGrande = 1 Then
                ' Alta
                '0109055072510
                '0002080625
                'Upper
                '@010100103039700
                If Len(BufDatos.BufTeclas(bucle)) = 10 Then
                    myTec.CodArticulo = 0
                    myTec.Destino_Tabla = 0
                    '1.8.3
                    'If lUpper Then
                    '    MyTEC.Section = 1
                    'Else
                        myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    'End If
                    ''''''
                    If Val(Mid(BufDatos.BufTeclas(bucle), 2, 1)) = 1 Then
                        Set Reg1 = Base.OpenRecordset("select * from equipos where numero_eqp=" & CStr(Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))))
                        If Reg1.EOF Then
                            myTec.Destino_NMostrador = 0 'Atencion upper
                        Else
                            myTec.Destino_NMostrador = Reg1.Fields("secc_maqui")
                        End If
                        Reg1.Close
                        Set Reg1 = Nothing
                    Else
                        myTec.Destino_NMostrador = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    End If
                    myTec.Destino_CodiTec = Val(Mid(BufDatos.BufTeclas(bucle), 5, 2))
                    myTec.PluArticulo = Val(Mid(BufDatos.BufTeclas(bucle), 7, 4))
                Else
                    myTec.PluArticulo = 0
                    myTec.Destino_NMostrador = 0
                    
                    If Val(Val(Mid(BufDatos.BufTeclas(bucle), 2, 1))) = 0 Then
                        myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    Else
                        myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2)) + 10
                    End If
                    '1.8.3
                    'If lUpper Then
                    '    MyTEC.Section = 1
                    'End If
                    ''''''
                    If Val(Mid(BufDatos.BufTeclas(bucle), 2, 1)) = 1 Then
                        Set Reg1 = Base.OpenRecordset("select * from equipos where numero_eqp=" & CStr(Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))))
                        If Reg1.EOF Then
                            myTec.Destino_NMostrador = 0 'Atencion upper
                        Else
                            myTec.Destino_NMostrador = Reg1.Fields("secc_maqui")
                        End If
                        Reg1.Close
                        Set Reg1 = Nothing
                    Else
                        myTec.Destino_NMostrador = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    End If

                    myTec.CodArticulo = Val(Mid(BufDatos.BufTeclas(bucle), 8, 6))
                    myTec.Destino_CodiTec = Val(Mid(BufDatos.BufTeclas(bucle), 5, 3))
                    'If lUpper Then
                    '    MyTEC.Destino_Tabla = Val(Mid(BufDatos.BufTeclas(bucle), 14, 2))
                    'Else
                        myTec.Destino_Tabla = Val(Mid(BufDatos.BufTeclas(bucle), 14))
                    'End If
                    'If MyTEC.Destino_Tabla <> 0 Then
                    '    MyTEC.Destino_Tabla = MyTEC.Destino_Tabla
                    'End If
                    '1.8.4--> en las teclados se empieza por la tabla 1
                    If lUpper Then myTec.Destino_Tabla = myTec.Destino_Tabla + 1
                    ''''''
                End If
                Resp = Alta_Teclas(myTec)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    If Len(BufDatos.BufTeclas(bucle)) = 10 Then
                        Print #Arch2, "6" & BufDatos.BufTeclas(bucle)
                    Else
                        Print #Arch2, "@" & BufDatos.BufTeclas(bucle)
                    End If
                    FrmBalGloDat.dato "NOK -->  + Tec. " & BufDatos.BufTeclas(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + Tec. " & BufDatos.BufTeclas(bucle)
                End If
            End If
            If left(BufDatos.BufTeclas(bucle), 1) = "1" And BGrande = 0 Then
                ' Baja
                'MyTEC.PluArticulo = 0
                'MyTEC.SecArticulo = 0
                'MyTEC.Section = Val(Mid(BufDatos.BufTeclas(Bucle), 3, 2))
                'MyTEC.Destino_NMostrador = MyTEC.Section
                'MyTEC.Destino_CodiTec = Val(Mid(BufDatos.BufTeclas(Bucle), 5, 3))
                'MyTEC.CodArticulo = Val(Mid(BufDatos.BufTeclas(Bucle), 8, 6))
                'MyTEC.Destino_Tabla = Val(Mid(BufDatos.BufTeclas(Bucle), 14, 3))
                If Len(BufDatos.BufTeclas(bucle)) = 10 Then
                    myTec.CodArticulo = 0
                    myTec.Destino_Tabla = 0
                    myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    myTec.Destino_NMostrador = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    myTec.Destino_CodiTec = Val(Mid(BufDatos.BufTeclas(bucle), 5, 2))
                    myTec.PluArticulo = Val(Mid(BufDatos.BufTeclas(bucle), 7, 4))
                Else
                    myTec.PluArticulo = 0
                    myTec.Destino_NMostrador = 0
                    If Val(Val(Mid(BufDatos.BufTeclas(bucle), 2, 1))) = 0 Then
                        myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2))
                    Else
                        myTec.Section = Val(Mid(BufDatos.BufTeclas(bucle), 3, 2)) + 10
                    End If
                    myTec.CodArticulo = Val(Mid(BufDatos.BufTeclas(bucle), 8, 6))
                    myTec.Destino_CodiTec = Val(Mid(BufDatos.BufTeclas(bucle), 5, 3))
                    If lUpper Then
                        myTec.Destino_Tabla = Val(Mid(BufDatos.BufTeclas(bucle), 14, 2))
                    Else
                        myTec.Destino_Tabla = Val(Mid(BufDatos.BufTeclas(bucle), 14, 3))
                    End If
                    If lUpper Then myTec.Destino_Tabla = myTec.Destino_Tabla + 1
                End If
                
                Resp = Baja_Teclas(myTec)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    If Len(BufDatos.BufTeclas(bucle)) = 10 Then
                        Print #Arch2, "6" & BufDatos.BufTeclas(bucle)
                    Else
                        Print #Arch2, "@" & BufDatos.BufTeclas(bucle)
                    End If
                    FrmBalGloDat.dato "NOK -->  + Tec. " & BufDatos.BufTeclas(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + Tec. " & BufDatos.BufTeclas(bucle)
                End If
            End If
        
        Next bucle

    Next BGrande
    
    Close #Arch2
    On Error Resume Next
    Base.Close
    On Error GoTo 0
End Sub

Private Sub Procesa_Secciones(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MySec As DB_mostrador
    Dim Resp As Integer
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        For bucle = 0 To BufDatos.NSecciones
            ' Alta
            If left(BufDatos.BufSecciones(bucle), 1) = "0" And BGrande = 1 Then
                MySec.NMostrador = Val(Mid(BufDatos.BufSecciones(bucle), 2, 2))
                MySec.NSeccion = MySec.NMostrador
                MySec.Tipo = 2
                MySec.Descripcion = ""
                If Len(BufDatos.BufSecciones(bucle)) > 9 Then
                    MySec.Descripcion = Mid(BufDatos.BufSecciones(bucle), 10, 21)
                    If Val(Mid(BufDatos.BufSecciones(bucle), 31, 3)) <> 0 Then
                        MySec.NMostrador = Val(Mid(BufDatos.BufSecciones(bucle), 31, 3))
                    End If
                    'If Val(Mid(BufDatos.BufSecciones(bucle), 34, 1)) = 1 Then
                        MySec.Tipo = Val(Mid(BufDatos.BufSecciones(bucle), 34, 1))
                    'End If
                    'If Val(Mid(BufDatos.BufSecciones(bucle), 34, 1)) = 2 Then
                    '    Mysec.Tipo = 2
                    'End If
                End If
                Resp = Alta_Mostrador(MySec)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "0" & BufDatos.BufSecciones(bucle)
                    FrmBalGloDat.dato "NOK -->  + Sec. " & MySec.NMostrador
                Else
                    FrmBalGloDat.dato "OK -->  + Sec. " & MySec.NMostrador
                End If
            End If
            ' Baja
            If left(BufDatos.BufSecciones(bucle), 1) = "1" And BGrande = 0 Then
                MySec.NMostrador = Val(Mid(BufDatos.BufSecciones(bucle), 2, 2))
                MySec.NSeccion = MySec.NSeccion
                Resp = Baja_Mostrador(MySec)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "0" & BufDatos.BufSecciones(bucle)
                    FrmBalGloDat.dato "NOK -->  - Sec. " & MySec.NMostrador
                Else
                    FrmBalGloDat.dato "OK -->  - Sec. " & MySec.NMostrador
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
End Sub
Private Sub Procesa_Equipos(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MyEqp As DB_Balanzas
    Dim Resp As Integer
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        For bucle = 0 To BufDatos.NEquipos
            If left(BufDatos.BufEquipos(bucle), 1) = "0" And BGrande = 1 Then
                ' Alta
                MyEqp.Descripcion = ""
                MyEqp.DireccionIP = "127.000.000.001"
                MyEqp.NMostrador = Val(Mid(BufDatos.BufEquipos(bucle), 2, 2))
                MyEqp.NBalanza = Val(Mid(BufDatos.BufEquipos(bucle), 4, 2))
                'If UsaGamaAlta Then
                    MyEqp.Tipo = 0 ' Euroscale
                'Else
                '    MyEqp.Tipo = 2 ' SC10
                'End If
                If Len(BufDatos.BufEquipos(bucle)) > 6 Then
                    Select Case Val(Mid(BufDatos.BufEquipos(bucle), 6, 2))
                        Case 0
                            MyEqp.Tipo = 0
                        Case 1
                            MyEqp.Tipo = 1
                        'Case 2
                        '    MyEqp.Tipo = 2
                    End Select
                    'If MyEqp.Tipo = 0 Or MyEqp.Tipo = 1 Then
                        If Trim(Mid(BufDatos.BufEquipos(bucle), 8, 15)) <> "" Then
                            MyEqp.DireccionIP = Mid(BufDatos.BufEquipos(bucle), 8, 15)
                        End If
                    'End If
                End If
                Resp = Alta_Balanza(MyEqp)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "4" & BufDatos.BufEquipos(bucle)
                    FrmBalGloDat.dato "NOK -->  + Eqp. " & MyEqp.NBalanza
                Else
                    FrmBalGloDat.dato "OK -->  + Eqp. " & MyEqp.NBalanza
                End If
            End If
            If left(BufDatos.BufEquipos(bucle), 1) = "1" And BGrande = 0 Then
                ' Baja
                MyEqp.NBalanza = Val(Mid(BufDatos.BufEquipos(bucle), 4, 2))
                Resp = Baja_Balanza(MyEqp)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "4" & BufDatos.BufEquipos(bucle)
                    FrmBalGloDat.dato "OK -->  - Eqp. " & MyEqp.NBalanza
                Else
                    FrmBalGloDat.dato "OK -->  - Eqp. " & MyEqp.NBalanza
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
End Sub
Private Sub Procesa_CodBar(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MyCB As DB_CodigoBarras
    Dim Resp As Integer
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        For bucle = 0 To BufDatos.NCodBar
            If left(BufDatos.BufCodBar(bucle), 1) = "0" And BGrande = 1 Then
                ' Alta
                If Mid(BufDatos.BufCodBar(bucle), 2, 1) = "0" Then
                    MyCB.Equipo = 0
                    MyCB.Mostrador = Mid(BufDatos.BufCodBar(bucle), 3, 2)
                Else
                    MyCB.Mostrador = 0
                    MyCB.Equipo = Mid(BufDatos.BufCodBar(bucle), 3, 2)
                End If
                MyCB.Datos(0) = Mid(BufDatos.BufCodBar(bucle), 5, 12)
                MyCB.Datos(1) = Mid(BufDatos.BufCodBar(bucle), 17, 12)
                MyCB.Datos(2) = Mid(BufDatos.BufCodBar(bucle), 29, 12)
                MyCB.Datos(3) = Mid(BufDatos.BufCodBar(bucle), 41, 12)
                MyCB.Datos(4) = Mid(BufDatos.BufCodBar(bucle), 53, 12)
                If Len(BufDatos.BufCodBar(bucle)) = 124 Then
                    MyCB.Datos(5) = Mid(BufDatos.BufCodBar(bucle), 65, 12)
                    MyCB.Datos(6) = Mid(BufDatos.BufCodBar(bucle), 77, 12)
                    MyCB.Datos(7) = Mid(BufDatos.BufCodBar(bucle), 89, 12)
                    MyCB.Datos(8) = Mid(BufDatos.BufCodBar(bucle), 101, 12)
                    MyCB.Datos(9) = Mid(BufDatos.BufCodBar(bucle), 113, 12)
                Else
                    MyCB.Datos(5) = Mid(BufDatos.BufCodBar(bucle), 5, 12)
                    MyCB.Datos(6) = Mid(BufDatos.BufCodBar(bucle), 17, 12)
                    MyCB.Datos(7) = Mid(BufDatos.BufCodBar(bucle), 29, 12)
                    MyCB.Datos(8) = Mid(BufDatos.BufCodBar(bucle), 41, 12)
                    MyCB.Datos(9) = Mid(BufDatos.BufCodBar(bucle), 53, 12)
                End If
                Resp = Alta_CodigoBarras(MyCB)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "7" & BufDatos.BufCodBar(bucle)
                    FrmBalGloDat.dato "NOK -->  + CB. " & BufDatos.BufCodBar(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + CB. " & BufDatos.BufCodBar(bucle)
                End If
            End If
            If left(BufDatos.BufCodBar(bucle), 1) = "1" And BGrande = 0 Then
                ' Baja
                If Mid(BufDatos.BufCodBar(bucle), 2, 1) = "0" Then
                    MyCB.Equipo = 0
                    MyCB.Mostrador = Mid(BufDatos.BufCodBar(bucle), 3, 2)
                Else
                    MyCB.Mostrador = 0
                    MyCB.Equipo = Mid(BufDatos.BufCodBar(bucle), 3, 2)
                End If
                MyCB.Datos(0) = ""
                MyCB.Datos(1) = ""
                MyCB.Datos(2) = ""
                MyCB.Datos(3) = ""
                MyCB.Datos(4) = ""
                MyCB.Datos(5) = ""
                MyCB.Datos(6) = ""
                MyCB.Datos(7) = ""
                MyCB.Datos(8) = ""
                MyCB.Datos(9) = ""
                Resp = Alta_CodigoBarras(MyCB)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "7" & BufDatos.BufCodBar(bucle)
                    FrmBalGloDat.dato "NOK -->  - CB. " & BufDatos.BufCodBar(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  - CB. " & BufDatos.BufCodBar(bucle)
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
End Sub
Private Sub Procesa_Articulos(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MyArt As DB_Articulo
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim nSuma As Integer
    Dim lGroup As Boolean
    Dim nLastShop As Integer
    'Dim lAlba As Boolean
    Dim nT As Integer
    
    If Dir(App.Path & "\alba.txt") <> "" Then
        lAlba = True
    Else
        lAlba = False
    End If
    'If = False Then
        Set Base = OpenDatabase(Base_General)
        Set Registro = Base.OpenRecordset("select * from articulo")
    'End If
    
    'Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo")
    
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        'If BGrande = 1 Then
        '    BGrande = BGrande
        'End If
        For bucle = 0 To BufDatos.NArticulos
            If (left(BufDatos.BufArticulos(bucle), 1) = "0" Or left(BufDatos.BufArticulos(bucle), 1) = "2") And BGrande = 1 Then
                MyArt.onkey = 1
                If left(BufDatos.BufArticulos(bucle), 1) = "2" Then MyArt.onkey = 0
                MyArt.lMix = False
                MyArt.nPoid = 0
                MyArt.tara = 0
                MyArt.IVA = 0
                MyArt.GrpConserv = 0
                MyArt.Preferente = 0
                MyArt.Etiqueta = 255
                MyArt.Merma = 0
                MyArt.ean13 = ""
                MyArt.Presel = 0
                MyArt.codigo = Mid(BufDatos.BufArticulos(bucle), 2, 6)
                MyArt.Mostrador = Mid(BufDatos.BufArticulos(bucle), 8, 2)
                MyArt.Plu = Mid(BufDatos.BufArticulos(bucle), 10, 4)
                'c2f 1.7.8 MyArt.subsec = Mid(BufDatos.BufArticulos(Bucle), 14, 3)
                'If lUpper = False Then
                '    MyArt.subsec = Mid(BufDatos.BufArticulos(bucle), 14, 3)
                '    'c2f 1.7.8 MyArt.familia = Mid(BufDatos.BufArticulos(Bucle), 17, 4)
                '    MyArt.familia = Mid(BufDatos.BufArticulos(bucle), 18, 4)
                'Else
                    MyArt.subsec = Mid(BufDatos.BufArticulos(bucle), 14, 3)
                    'c2f 1.7.8 MyArt.familia = Mid(BufDatos.BufArticulos(Bucle), 17, 4)
                    MyArt.familia = Mid(BufDatos.BufArticulos(bucle), 17, 4)
                
                'End If
                If UsaEuro Then
                    MyArt.precio = Val(Mid(BufDatos.BufArticulos(bucle), 21, 5)) / 100
                Else
                    MyArt.precio = Val(Mid(BufDatos.BufArticulos(bucle), 21, 5)) / (10 ^ decimales)
                End If
                If Mid(BufDatos.BufArticulos(bucle), 26, 1) = "W" Or Mid(BufDatos.BufArticulos(bucle), 26, 1) = "P" Then
                    MyArt.WGH = False
                Else
                    MyArt.WGH = True
                End If
                
                MyArt.caducidad = Mid(BufDatos.BufArticulos(bucle), 27, 3)
                
                MyArt.TipoLetra(0) = Mid(BufDatos.BufArticulos(bucle), 30, 1)
                MyArt.Descriptivos(0) = Mid(BufDatos.BufArticulos(bucle), 31, 25)
                
                
                If Len(BufDatos.BufArticulos(bucle)) > 60 And lFornes = False Then
                    If Len(BufDatos.BufArticulos(bucle)) = 62 Then
                        If lAlba Then
                            MyArt.TipoLetra(0) = 0
                            MyArt.IVA = Mid(BufDatos.BufArticulos(bucle), 30, 1)
                            nT = Right(BufDatos.BufArticulos(bucle), 2)
                            'TiendaActual = nT
                            If nT <> TiendaActual Then GoTo EsOtraTienda
                        End If
                    
                    Else
                        
                        If Len(BufDatos.BufArticulos(bucle)) = 92 Or Len(BufDatos.BufArticulos(bucle)) = 91 Then
                            MyArt.tara = Mid(BufDatos.BufArticulos(bucle), 61, 5)
                            MyArt.IVA = Mid(BufDatos.BufArticulos(bucle), 66, 1)
                            MyArt.GrpConserv = Mid(BufDatos.BufArticulos(bucle), 67, 3)
                            MyArt.Preferente = Mid(BufDatos.BufArticulos(bucle), 70, 3)
                            If TiendaActual <> 999 Then
                                MyArt.Etiqueta = Mid(BufDatos.BufArticulos(bucle), 73, 3)
                            Else
                                MyArt.Etiqueta = 999
                            End If
                            MyArt.Merma = Mid(BufDatos.BufArticulos(bucle), 76, 2)
                            'MyArt.EAN13 = Trim(Mid(BufDatos.BufArticulos(Bucle), 78, 12))
                            MyArt.ean13 = Mid(BufDatos.BufArticulos(bucle), 78, 12)
                            If Len(Trim(MyArt.ean13)) = 12 Then
                                If MyArt.ean13 = "000000000000" Then
                                    MyArt.ean13 = ""
                                Else
                                    MyArt.ean13 = Trim(MyArt.ean13) & Chr(48 + DigitoControl(Trim(MyArt.ean13)))
                                End If
                            End If
                            '''''''''''''''''''
                            MyArt.Presel = Mid(BufDatos.BufArticulos(bucle), 90, 2)
                        
                        Else
                            MyArt.tara = Mid(BufDatos.BufArticulos(bucle), 56, 5)
                            MyArt.IVA = Mid(BufDatos.BufArticulos(bucle), 61, 1)
                            MyArt.GrpConserv = Mid(BufDatos.BufArticulos(bucle), 62, 3)
                            MyArt.Preferente = Mid(BufDatos.BufArticulos(bucle), 65, 3)
                            If TiendaActual <> 999 Then
                                MyArt.Etiqueta = Mid(BufDatos.BufArticulos(bucle), 68, 3)
                            Else
                                MyArt.Etiqueta = 999
                            End If
                            MyArt.Merma = Mid(BufDatos.BufArticulos(bucle), 71, 2)
                            'MyArt.EAN13 = Trim(Mid(BufDatos.BufArticulos(Bucle), 78, 12))
                            MyArt.ean13 = Mid(BufDatos.BufArticulos(bucle), 73, 12)
                            If Len(Trim(MyArt.ean13)) = 12 Then
                                If MyArt.ean13 = "000000000000" Then
                                    MyArt.ean13 = ""
                                Else
                                    MyArt.ean13 = Trim(MyArt.ean13) & Chr(48 + DigitoControl(Trim(MyArt.ean13)))
                                End If
                            End If
                            '''''''''''''''''''
                            MyArt.Presel = Mid(BufDatos.BufArticulos(bucle), 85, 2)
                            'If lAlba Then
                            '    nT = Mid(BufDatos.BufArticulos(bucle), 92, 2)
                            '    TiendaActual = nT
                            'End If
                        End If
                    End If
                Else
                    If lFornes And Len(BufDatos.BufArticulos(bucle)) > 60 Then
                        MyArt.Preferente = Mid(BufDatos.BufArticulos(bucle), 61, 2)
                    End If
                End If
                MyArt.TRM1 = 0
                MyArt.TRM2 = 0
                MyArt.PRC1 = 0
                MyArt.PRC2 = 0
                
                'If lAlba Then
                '    'ALBA INFORMATICA
                '    If nLastShop <> nT Then
                '        Base_General = App.Path & "\t" & Format(nT, "00") & "\" & "dbase.mdb"
                '        Hydra_INI = App.Path & "\t" & Format(nT, "00") & "\" & "hydra.ini"
                '        Call LeerParametrosInicio
                '        Set Base = AbrirBase 'OpenDatabase(Base_General)
                '        Set Registro = Base.OpenRecordset("select * from articulo")
                '    '''''''''''''''''
                '    End If
                '    nLastShop = nT
                'End If
                
                '''''''''''''''
                'If lUpperNW Then
                If lUpper Then
                    If lJpg = False Then
                        MyArt.Imagen = Format(MyArt.codigo, "000000") & ".png"
                    Else
                        MyArt.Imagen = Format(MyArt.codigo, "000000") & ".jpg"
                    End If
                End If
                If lFornes Then
                    'If lJpg = False Then
                        MyArt.Imagen = CStr(MyArt.codigo) & ".png"
                    'Else
                    '    MyArt.Imagen = Format(MyArt.codigo, "000000") & ".jpg"
                    'End If
                End If
                
                Resp = Alta_Articulo_Mod(MyArt, Base, Registro)
                If Resp = 45 Then
                    Resp = Baja_Articulo(MyArt)
                    Resp = Alta_Articulo_Mod(MyArt, Base, Registro)
                End If
                
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "5" & BufDatos.BufArticulos(bucle)
                    FrmBalGloDat.dato "NOK -->  + PLU. " & BufDatos.BufArticulos(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + PLU. " & BufDatos.BufArticulos(bucle)
                    'If luppernw Then
                    'Atencionfamilia
                    'If lGroup Then
                    '    'Dim sS As Variant
                    '    'Dim ffff As Variant
                    '    'Dim pppp As Integer
                    '    'Dim sdesfam As String
                    '    'Dim simgfam As String
                    '    Dim ssSecs(1) As Long
                    '    Dim nResp As Variant
                    '    'simgfam = ""
                    '    'sS = MyArt.Mostrador
                    '    ''ffff = (sS * 10000) + ffff
                    '    ''If ffff < 100 Then
                    '    ''    ffff = (sS * 100) + ffff
                    '    ''Else
                    '    ''    ffff = (sS * 10000) + ffff
                    '    ''End If
                    '    'pppp = (MyArt.familia) Mod (100 * MyArt.Mostrador)
                    '    'nResp = Alta_Familias(MyArt.familia, "", 0, "", False, , pppp)
                    '    'If nResp = 0 Then
                    '    '    ssSecs(0) = MyArt.familia
                    '    '    nResp = Alta_Familias_Counter(MyArt.familia, "", 0, ssSecs(), False, , "", sS)
                    '    'End If
                    '    nResp = Alta_Familias(MyArt.familia, "", 0, "", False, , MyArt.familia)
                    '    If nResp = 0 Then
                    '        ssSecs(0) = MyArt.Mostrador
                    '        nResp = Alta_Familias_Counter(MyArt.familia, "", 0, ssSecs(), False, , "", MyArt.Mostrador)
                    '    End If
                    'End If
                End If
            End If
            If left(BufDatos.BufArticulos(bucle), 1) = "1" And BGrande = 0 Then
                ' Baja
                'If lAlba Then
                '    nT = Mid(BufDatos.BufArticulos(bucle), 92, 2)
                '    TiendaActual = nT
                '    'ALBA INFORMATICA
                '    If nLastShop <> nT Then
                '        Base_General = App.Path & "\t" & Format(nT, "00") & "\" & "dbase.mdb"
                '        Hydra_INI = App.Path & "\t" & Format(nT, "00") & "\" & "hydra.ini"
                '        Call LeerParametrosInicio
                '        Set Base = AbrirBase 'OpenDatabase(Base_General)
                '        Set Registro = Base.OpenRecordset("select * from articulo")
                '    '''''''''''''''''
                '    End If
                '    nLastShop = nT
                'End If
                If lAlba Then
                    nT = Right(BufDatos.BufArticulos(bucle), 2)
                    'TiendaActual = nT
                    If nT <> TiendaActual Then GoTo EsOtraTienda
                End If
                
                MyArt.codigo = Mid(BufDatos.BufArticulos(bucle), 2, 6)
                Resp = Baja_Articulo(MyArt)
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "5" & BufDatos.BufArticulos(bucle)
                    FrmBalGloDat.dato "NOK -->  - PLU. " & BufDatos.BufArticulos(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  - PLU. " & BufDatos.BufArticulos(bucle)
                End If
            End If
EsOtraTienda:
        Next bucle
    Next BGrande
    Close #Arch2
    'c2f 1.7.19
    'Base.Close
End Sub
Private Sub Procesa_EanArticulos(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nEAN - 1

        MyCod = Val(Mid(BufDatos.BufEAN(bucle), 2, 6))
        
        Registro.FindFirst "codigo=" & MyCod & " and borrado=false"

        If Registro.EOF Then
            Print #Arch2, BufDatos.BufEAN(bucle)
            FrmBalGloDat.dato "NOK -->  + EAN: " & BufDatos.BufEAN(bucle)
        Else
            With Registro
                Edit_Record Registro
                .Fields("art_cb") = Mid(BufDatos.BufEAN(bucle), 8, 12) & Chr(48 + DigitoControl(Mid(BufDatos.BufEAN(bucle), 8, 12)))
                '.Fields("tran_texto") = ""
                .Fields("tran_plu") = ""
                '.Fields("tran_plusc10") = ""
                '.Fields("tran_textosc10") = ""
                '.Fields("tran_tx1") = ""
                '.Fields("tran_tx1sc10") = ""
                .Update
                FrmBalGloDat.dato "OK -->  + EAN: " & BufDatos.BufEAN(bucle)
            End With
        End If
    Next bucle
    Close #Arch2
    Base.Close
End Sub


Private Sub Procesa_Tarifas(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    Dim nTarifa As Integer
    Dim valTarifa As Double
    
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nTarifas - 1

        MyCod = Val(Mid(BufDatos.BufTarifas(bucle), 1, 6))
        
        Registro.FindFirst "codigo=" & MyCod & " and borrado=false"

        If Registro.EOF Then
            Print #Arch2, BufDatos.BufTarifas(bucle)
            FrmBalGloDat.dato "NOK -->  + Tarifa: " & BufDatos.BufTarifas(bucle)
        Else
            nTarifa = Val(Mid(BufDatos.BufTarifas(bucle), 7, 2))

            valTarifa = Mid(BufDatos.BufTarifas(bucle), 9, 7)
            If UsaEuro Then
                valTarifa = valTarifa / 100
            Else
                valTarifa = valTarifa / (10 ^ decimales)
            End If
            
            With Registro
                Edit_Record Registro
                .Fields("tar" & CStr(nTarifa)) = valTarifa
                .Fields("tran_plu") = ""
                .Update
                FrmBalGloDat.dato "OK -->  + TARIFA: " & BufDatos.BufTarifas(bucle)
            End With
        End If
    Next bucle
    Close #Arch2
    Base.Close
End Sub

Private Sub Procesa_Imagen(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nImage - 1

        MyCod = Val(Mid(BufDatos.BufImagen(bucle), 1, 6))
        
        Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & MyCod & " and borrado=false")

        If Registro.EOF Then
            Print #Arch2, BufDatos.BufImagen(bucle)
            FrmBalGloDat.dato "NOK -->  + Imagen: " & BufDatos.BufImagen(bucle)
        Else
            'With Registro
                Registro.Edit
                Registro.Fields("imagen") = Trim(Mid(BufDatos.BufImagen(bucle), 7))
                '.Fields("tran_texto") = ""
                Registro.Fields("tran_plu") = ""
                '.Fields("tran_plusc10") = ""
                '.Fields("tran_textosc10") = ""
                '.Fields("tran_tx1") = ""
                '.Fields("tran_tx1sc10") = ""
                Registro.Update
                FrmBalGloDat.dato "OK -->  + Imagen: " & BufDatos.BufImagen(bucle)
            'End With
        End If
        Registro.Close
        Set Registro = Nothing
        
    Next bucle
    Close #Arch2
    Base.Close
End Sub

Private Sub Procesa_Etiqueta(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nEtiqueta - 1

        MyCod = Val(Mid(BufDatos.BufEtiqueta(bucle), 2, 6))
        
        Registro.FindFirst "codigo=" & MyCod & " and borrado=false"

        If Registro.EOF Then
            Print #Arch2, BufDatos.BufImagen(bucle)
            FrmBalGloDat.dato "NOK -->  + Etiqueta: " & BufDatos.BufEtiqueta(bucle)
        Else
            With Registro
                Edit_Record Registro
                .Fields("tran_ep") = Trim(Mid(BufDatos.BufEtiqueta(bucle), 8))
                '.Fields("tran_texto") = ""
                .Fields("tran_plu") = ""
                '.Fields("tran_plusc10") = ""
                '.Fields("tran_textosc10") = ""
                '.Fields("tran_tx1") = ""
                '.Fields("tran_tx1sc10") = ""
                .Update
                FrmBalGloDat.dato "OK -->  + Imagen: " & BufDatos.BufEtiqueta(bucle)
            End With
        End If
    Next bucle
    Close #Arch2
    Base.Close
End Sub

Private Sub Procesa_TEXT15(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim Registro As dao.Recordset
    Dim registro2 As dao.Recordset
    Dim nT As Integer
    Dim nPlu As Long
    Dim nSec As Integer
    Dim lAlta As Boolean
    
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.nText15 - 1

        MyCod = Val(Mid(BufDatos.BufText15(bucle), 1, 6))
        nT = Val(Mid(BufDatos.BufText15(bucle), 7, 1))
        
        'Registro.FindFirst "codigo=" & MyCod & " and borrado=false"
        Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & MyCod)
        
        If Registro.EOF Then
            Print #Arch2, BufDatos.BufText15(bucle)
            FrmBalGloDat.dato "NOK -->  + Texto15: " & BufDatos.BufText15(bucle)
        Else
            
            nPlu = Registro.Fields("plu")
            nSec = Registro.Fields("secc_maqui")
            Set registro2 = Base.OpenRecordset("select * from text15 where codigo=" & CStr(MyCod))
            If registro2.EOF Then
                registro2.AddNew
                lAlta = True
            Else
                registro2.Edit
                lAlta = False
            End If

            
            If lAlta Then
                Registro.Edit
                Registro.Fields("tran_plu") = ""
                Registro.Update
                registro2.Fields("codigo") = MyCod
                registro2.Fields("plu") = nPlu
                registro2.Fields("mostrador") = nSec
                registro2.Fields("tran_txt15") = ""
                If IsNull(Trim(Mid(BufDatos.BufText15(bucle), 8))) Or Len(Trim(Mid(BufDatos.BufText15(bucle), 8))) = 0 Then
                    registro2.Fields("text" & CStr(nT)) = ""
                Else
                    registro2.Fields("text" & CStr(nT)) = Trim(Mid(BufDatos.BufText15(bucle), 8))
                End If
            Else
                If Trim(registro2.Fields("text" & CStr(nT))) <> Trim(Mid(BufDatos.BufText15(bucle), 8)) Then
                    Registro.Edit
                    Registro.Fields("tran_plu") = ""
                    Registro.Update
                    registro2.Fields("codigo") = MyCod
                    registro2.Fields("plu") = nPlu
                    registro2.Fields("mostrador") = nSec
                    registro2.Fields("tran_txt15") = ""
                    'registro2.Fields("text" & CStr(nT)) = Trim(Mid(BufDatos.BufText15(bucle), 9))
                    If IsNull(Trim(Mid(BufDatos.BufText15(bucle), 8))) Or Len(Trim(Mid(BufDatos.BufText15(bucle), 8))) = 0 Then
                        registro2.Fields("text" & CStr(nT)) = ""
                    Else
                        registro2.Fields("text" & CStr(nT)) = Trim(Mid(BufDatos.BufText15(bucle), 8))
                    End If
                    
                    lAlta = True
                End If
            End If
            
            'registro2.Fields("text" & CStr(nT)) = Trim(Mid(BufDatos.BufText15(bucle), 9))
            
            
            If lAlta Then registro2.Update
            
            FrmBalGloDat.dato "OK -->  + Texto: " & Trim(Mid(BufDatos.BufText15(bucle), 8))

        End If
        
        Registro.Close
        Set Registro = Nothing
        
    Next bucle
    'Registro.Close
    'Set Registro = Nothing
    Close #Arch2
    Base.Close
End Sub

Private Sub Procesa_Textos(BufDatos As Tipo_Buf)
    '///////////////////////////////////////////////
    ' esta función de alta/baja directamente
    ' accediendo a la base de datos
    '///////////////////////////////////////////////
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim Resp As Integer
    Dim Base As dao.Database
    Dim MyOK As Boolean
    Dim MyCod As Long
    Dim MyBloque As Integer
    Dim Registro As dao.Recordset
    Dim Regtmp As dao.Recordset
    Dim sI As String
    Dim nReg As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo")
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For bucle = 0 To BufDatos.NTextos - 1
        ' Alta
        If left(BufDatos.BufTextos(bucle), 1) = "Z" Or left(BufDatos.BufTextos(bucle), 1) = "=" Then
            MyBloque = 0
        Else
            If left(BufDatos.BufTextos(bucle), 1) = "M" Then
                MyBloque = 2 '2.0.34
            Else
                If left(BufDatos.BufTextos(bucle), 1) = "O" Then
                    MyBloque = 3 '2.0.34
                Else
                    MyBloque = 1
                End If
            End If
        End If
        If left(BufDatos.BufTextos(bucle), 1) = "=" Then
            sI = "="
            Set Regtmp = Base.OpenRecordset("select * from articulo where secc_maqui=" & Val(Mid(BufDatos.BufTextos(bucle), 2, 2)) & " and plu=" & Val(Mid(BufDatos.BufTextos(bucle), 4, 4)))
            'Registro.FindFirst "secc_maqui=" & Val(Mid(BufDatos.BufTextos(Bucle), 2, 2)) & " and plu=" & Val(Mid(BufDatos.BufTextos(Bucle), 4, 4))
            If Regtmp.EOF Then
                MyCod = 0
            Else
                MyCod = Regtmp.Fields("codigo")
            End If
            Regtmp.Close
            Set Regtmp = Nothing
        Else
            '2.0.34 sI = "Z"
            sI = left(BufDatos.BufTextos(bucle), 1)
            ''''''''''''''''
            MyCod = Val(Mid(BufDatos.BufTextos(bucle), 2, 6)) 'c2f 1.8.2 era 1,6 (mal)
        End If
        
        BufDatos.BufTextos(bucle) = Mid(BufDatos.BufTextos(bucle), 2)
        Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & MyCod & " and borrado=false")
        'Registro.FindFirst "codigo=" & MyCod & " and borrado=false"
        'If Registro.EOF Or Registro.NoMatch Then ojito...
        If Registro.EOF Then
            '2.0.34
            'If MyBloque = 0 Then
            '    Print #Arch2, sI & BufDatos.BufTextos(Bucle)
            '    FrmBalGloDat.dato "NOK -->  + TXT1. " & BufDatos.BufTextos(Bucle)
            'Else
            '    Print #Arch2, "Y" & BufDatos.BufTextos(Bucle)
            '    FrmBalGloDat.dato "NOK -->  + TXT2. " & BufDatos.BufTextos(Bucle)
            'End If
            ''''''''
            Print #Arch2, sI & BufDatos.BufTextos(bucle)
            FrmBalGloDat.dato "NOK -->(No COD.) + TXT" & sI & ". " & BufDatos.BufTextos(bucle)
        Else
            If lUpper Then
                If MyBloque = 0 Then sTxtTmp = ""
                sTxtTmp = sTxtTmp & Trim(Mid(BufDatos.BufTextos(bucle), 7))
                If MyBloque = 3 Then
                    If Len(Trim(sTxtTmp)) > 0 Then
                        Set nReg = Base.OpenRecordset("select * from text15 where codigo=" & CStr(MyCod))
                        If nReg.EOF Then
                            nReg.AddNew
                            nReg.Fields("codigo") = Registro.Fields("codigo")
                            nReg.Fields("plu") = Registro.Fields("plu")
                            nReg.Fields("mostrador") = Registro.Fields("secc_maqui")
                        Else
                            nReg.Edit
                        End If
                        nReg.Fields("text1") = sTxtTmp
                        nReg.Fields("tran_txt15") = " "
                        nReg.Update
                        nReg.Close
                        Set nReg = Nothing
                    End If
                End If
            Else
                With Registro
                    Edit_Record Registro
                    '2.0.34
                    If MyBloque = 2 Or MyBloque = 3 Then
                        Set nReg = Base.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyCod))
                        If nReg.EOF Then
                            nReg.AddNew
                            nReg.Fields("codigo") = Registro.Fields("codigo")
                            nReg.Fields("plu") = Registro.Fields("plu")
                            nReg.Fields("mostrador") = Registro.Fields("secc_maqui")
                        Else
                            nReg.Edit
                        End If
                        If MyBloque = 2 Then
                            nReg.Fields("txt_21") = RTrim(Mid(BufDatos.BufTextos(bucle), 7, 25))
                            nReg.Fields("txt_22") = RTrim(Mid(BufDatos.BufTextos(bucle), 32, 25))
                            nReg.Fields("txt_23") = RTrim(Mid(BufDatos.BufTextos(bucle), 57, 25))
                            nReg.Fields("txt_24") = RTrim(Mid(BufDatos.BufTextos(bucle), 82, 25))
                            nReg.Fields("txt_25") = RTrim(Mid(BufDatos.BufTextos(bucle), 107, 25))
                            nReg.Fields("txt_26") = RTrim(Mid(BufDatos.BufTextos(bucle), 132, 25))
                            nReg.Fields("txt_27") = RTrim(Mid(BufDatos.BufTextos(bucle), 157, 25))
                            nReg.Fields("txt_28") = RTrim(Mid(BufDatos.BufTextos(bucle), 182, 25))
                            nReg.Fields("txt_29") = RTrim(Mid(BufDatos.BufTextos(bucle), 207, 25))
                            nReg.Fields("txt_30") = RTrim(Mid(BufDatos.BufTextos(bucle), 232, 25))
                        Else
                            nReg.Fields("txt_31") = RTrim(Mid(BufDatos.BufTextos(bucle), 7, 25))
                            nReg.Fields("txt_32") = RTrim(Mid(BufDatos.BufTextos(bucle), 32, 25))
                            nReg.Fields("txt_33") = RTrim(Mid(BufDatos.BufTextos(bucle), 57, 25))
                            nReg.Fields("txt_34") = RTrim(Mid(BufDatos.BufTextos(bucle), 82, 25))
                            nReg.Fields("txt_35") = RTrim(Mid(BufDatos.BufTextos(bucle), 107, 25))
                            nReg.Fields("txt_36") = RTrim(Mid(BufDatos.BufTextos(bucle), 132, 25))
                            nReg.Fields("txt_37") = RTrim(Mid(BufDatos.BufTextos(bucle), 157, 25))
                            nReg.Fields("txt_38") = RTrim(Mid(BufDatos.BufTextos(bucle), 182, 25))
                            nReg.Fields("txt_39") = RTrim(Mid(BufDatos.BufTextos(bucle), 207, 25))
                            nReg.Fields("txt_40") = RTrim(Mid(BufDatos.BufTextos(bucle), 232, 25))
                        End If
                        nReg.Fields("tran_txt2") = " "
                        nReg.Update
                        nReg.Close
                        Set nReg = Nothing
                        
                    Else
                        If MyBloque = 0 Then
                            If sI = "=" Then
                            
                                '.Fields("des_plu1") = Trim(Mid(BufDatos.BufTextos(Bucle), 8, 25))
                                ''.Fields("tip_let1") = Val(Mid(BufDatos.BufTextos(Bucle), 7, 1))
                                
                                '.Fields("des_plu2") = Trim(Mid(BufDatos.BufTextos(Bucle), 34, 25))
                                '.Fields("tip_let2") = Val(Mid(BufDatos.BufTextos(Bucle), 33, 1))
                                
                                '.Fields("des_plu3") = Trim(Mid(BufDatos.BufTextos(Bucle), 60, 25))
                                '.Fields("tip_let3") = Val(Mid(BufDatos.BufTextos(Bucle), 59, 1))
                                
                                '.Fields("des_plu4") = Trim(Mid(BufDatos.BufTextos(Bucle), 86, 25))
                                '.Fields("tip_let4") = Val(Mid(BufDatos.BufTextos(Bucle), 85, 1))
                                
                                '.Fields("des_plu5") = Trim(Mid(BufDatos.BufTextos(Bucle), 112, 25))
                                '.Fields("tip_let5") = Val(Mid(BufDatos.BufTextos(Bucle), 111, 1))
                                
                                '.Fields("des_plu6") = Trim(Mid(BufDatos.BufTextos(Bucle), 138, 25))
                                '.Fields("tip_let6") = Val(Mid(BufDatos.BufTextos(Bucle), 137, 1))
                                
                                '.Fields("des_plu7") = Trim(Mid(BufDatos.BufTextos(Bucle), 164, 25))
                                '.Fields("tip_let7") = Val(Mid(BufDatos.BufTextos(Bucle), 163, 1))
                                
                                '.Fields("des_plu8") = Trim(Mid(BufDatos.BufTextos(Bucle), 190, 25))
                                '.Fields("tip_let8") = Val(Mid(BufDatos.BufTextos(Bucle), 189, 1))
                                ''.Fields("des_plu9") = Trim(Mid(BufDatos.BufTextos(Bucle), 216, 25))
                                ''.Fields("des_plux") = Trim(Mid(BufDatos.BufTextos(Bucle), 233, 25))
                                
                                .Fields("des_plu1") = RTrim(Mid(BufDatos.BufTextos(bucle), 8, 25))
                                .Fields("des_plu2") = RTrim(Mid(BufDatos.BufTextos(bucle), 34, 25))
                                .Fields("tip_let2") = Val(Mid(BufDatos.BufTextos(bucle), 33, 1))
                                .Fields("des_plu3") = RTrim(Mid(BufDatos.BufTextos(bucle), 60, 25))
                                .Fields("tip_let3") = Val(Mid(BufDatos.BufTextos(bucle), 59, 1))
                                .Fields("des_plu4") = RTrim(Mid(BufDatos.BufTextos(bucle), 86, 25))
                                .Fields("tip_let4") = Val(Mid(BufDatos.BufTextos(bucle), 85, 1))
                                .Fields("des_plu5") = RTrim(Mid(BufDatos.BufTextos(bucle), 112, 25))
                                .Fields("tip_let5") = Val(Mid(BufDatos.BufTextos(bucle), 111, 1))
                                .Fields("des_plu6") = RTrim(Mid(BufDatos.BufTextos(bucle), 138, 25))
                                .Fields("tip_let6") = Val(Mid(BufDatos.BufTextos(bucle), 137, 1))
                                .Fields("des_plu7") = RTrim(Mid(BufDatos.BufTextos(bucle), 164, 25))
                                .Fields("tip_let7") = Val(Mid(BufDatos.BufTextos(bucle), 163, 1))
                                .Fields("des_plu8") = RTrim(Mid(BufDatos.BufTextos(bucle), 190, 25))
                                .Fields("tip_let8") = Val(Mid(BufDatos.BufTextos(bucle), 189, 1))
                            
                            Else
                                'c2f 2.0.19 ... se cambio de 1 a 0 por 2 a x
                                '.Fields("des_plu2") = Trim(Mid(BufDatos.BufTextos(Bucle), 7, 25))
                                '.Fields("des_plu3") = Trim(Mid(BufDatos.BufTextos(Bucle), 32, 25))
                                '.Fields("des_plu4") = Trim(Mid(BufDatos.BufTextos(Bucle), 57, 25))
                                '.Fields("des_plu5") = Trim(Mid(BufDatos.BufTextos(Bucle), 82, 25))
                                '.Fields("des_plu6") = Trim(Mid(BufDatos.BufTextos(Bucle), 107, 25))
                                '.Fields("des_plu7") = Trim(Mid(BufDatos.BufTextos(Bucle), 132, 25))
                                '.Fields("des_plu8") = Trim(Mid(BufDatos.BufTextos(Bucle), 157, 25))
                                '.Fields("des_plu9") = Trim(Mid(BufDatos.BufTextos(Bucle), 182, 25))
                                '.Fields("des_plu0") = Trim(Mid(BufDatos.BufTextos(Bucle), 207, 25))
                                '.Fields("des_plux") = Trim(Mid(BufDatos.BufTextos(Bucle), 232, 25))
                                .Fields("des_plu2") = RTrim(Mid(BufDatos.BufTextos(bucle), 7, 25))
                                .Fields("des_plu3") = RTrim(Mid(BufDatos.BufTextos(bucle), 32, 25))
                                .Fields("des_plu4") = RTrim(Mid(BufDatos.BufTextos(bucle), 57, 25))
                                .Fields("des_plu5") = RTrim(Mid(BufDatos.BufTextos(bucle), 82, 25))
                                .Fields("des_plu6") = RTrim(Mid(BufDatos.BufTextos(bucle), 107, 25))
                                .Fields("des_plu7") = RTrim(Mid(BufDatos.BufTextos(bucle), 132, 25))
                                .Fields("des_plu8") = RTrim(Mid(BufDatos.BufTextos(bucle), 157, 25))
                                .Fields("des_plu9") = RTrim(Mid(BufDatos.BufTextos(bucle), 182, 25))
                                .Fields("des_plu0") = RTrim(Mid(BufDatos.BufTextos(bucle), 207, 25))
                                .Fields("des_plux") = RTrim(Mid(BufDatos.BufTextos(bucle), 232, 25))
                            
                            End If
                            
                        Else
                            '.Fields("des_plu11") = Trim(Mid(BufDatos.BufTextos(Bucle), 7, 25))
                            '.Fields("des_plu12") = Trim(Mid(BufDatos.BufTextos(Bucle), 32, 25))
                            '.Fields("des_plu13") = Trim(Mid(BufDatos.BufTextos(Bucle), 57, 25))
                            '.Fields("des_plu14") = Trim(Mid(BufDatos.BufTextos(Bucle), 82, 25))
                            '.Fields("des_plu15") = Trim(Mid(BufDatos.BufTextos(Bucle), 107, 25))
                            '.Fields("des_plu16") = Trim(Mid(BufDatos.BufTextos(Bucle), 132, 25))
                            '.Fields("des_plu17") = Trim(Mid(BufDatos.BufTextos(Bucle), 157, 25))
                            '.Fields("des_plu18") = Trim(Mid(BufDatos.BufTextos(Bucle), 182, 25))
                            '.Fields("des_plu19") = Trim(Mid(BufDatos.BufTextos(Bucle), 207, 25))
                            '.Fields("des_plu20") = Trim(Mid(BufDatos.BufTextos(Bucle), 232, 25))
                            
                            .Fields("des_plu11") = RTrim(Mid(BufDatos.BufTextos(bucle), 7, 25))
                            .Fields("des_plu12") = RTrim(Mid(BufDatos.BufTextos(bucle), 32, 25))
                            .Fields("des_plu13") = RTrim(Mid(BufDatos.BufTextos(bucle), 57, 25))
                            .Fields("des_plu14") = RTrim(Mid(BufDatos.BufTextos(bucle), 82, 25))
                            .Fields("des_plu15") = RTrim(Mid(BufDatos.BufTextos(bucle), 107, 25))
                            .Fields("des_plu16") = RTrim(Mid(BufDatos.BufTextos(bucle), 132, 25))
                            .Fields("des_plu17") = RTrim(Mid(BufDatos.BufTextos(bucle), 157, 25))
                            .Fields("des_plu18") = RTrim(Mid(BufDatos.BufTextos(bucle), 182, 25))
                            .Fields("des_plu19") = RTrim(Mid(BufDatos.BufTextos(bucle), 207, 25))
                            .Fields("des_plu20") = RTrim(Mid(BufDatos.BufTextos(bucle), 232, 25))
                            
                            If lAgora = False Then .Fields("tran_tx1") = ""
                        End If
                        If lAgora = False Then .Fields("tran_texto") = ""
                        If lAgora = False Then .Fields("tran_plu") = ""
                        If lAgora = False Then
                            .Fields("tran_plusc10") = ""
                            .Fields("tran_textosc10") = ""
                            .Fields("tran_tx1") = ""
                            .Fields("tran_tx1sc10") = ""
                            .Update
                        End If
                    End If
                    
                    If MyBloque = 0 Then
                        FrmBalGloDat.dato "OK -->  + TXT1. " & BufDatos.BufTextos(bucle)
                    Else
                        FrmBalGloDat.dato "OK -->  + TXT2. " & BufDatos.BufTextos(bucle)
                    End If
                End With
            End If
        End If
        Registro.Close
        Set Registro = Nothing
    Next bucle
    Close #Arch2
    Base.Close
End Sub
Private Sub Procesa_Vendedores(BufDatos As Tipo_Buf)
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MyVnd As DB_Vendedor
    Dim Resp As Integer
    Dim lEsG As Boolean
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    lEsG = True
    For BGrande = 0 To 1
                
        For bucle = 0 To BufDatos.NVendedores
        
            If left(BufDatos.BufVendedores(bucle), 1) = "0" And BGrande = 1 Then
                ' Alta
                '"G" 1 (Literal) Record Type "G"
                '<R-M / U>  1 digit "0" Create-Update, "1" Remove
                'Counter    2 digit Counter (01-99)
                'Identifier 4 digits Seller’s identifier (0001-9999)
                'Key        4 digits Seller’s key into Counter (0001 - 9999)
                'Name      30 charac. Seller’s Name
                '<CR + LF> 2 CHARAC. End of register
                If Mid(BufDatos.BufVendedores(bucle), 2, 1) = "0" Then
                    If lEsG Then
                        If Len(BufDatos.BufVendedores(bucle)) = 43 Then
                            MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 2, 2)
                            MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 4, 6)
                            MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 10, 4)
                            MyVnd.estado = True
                            MyVnd.Descripcion = Mid(BufDatos.BufVendedores(bucle), 14, 30)
                        
                        Else
                            MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 2, 2)
                            MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 4, 4)
                            MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 8, 4)
                            MyVnd.estado = True
                            MyVnd.Descripcion = Mid(BufDatos.BufVendedores(bucle), 12, 30)
                        End If
                    Else
                        MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 5, 4)
                        MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 3, 2)
                        MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 15, 2)
                        If MyVnd.NVendedor < 100 Then
                            MyVnd.NVendedor = MyVnd.seccion * 100 + MyVnd.NVendedor
                        End If
                        MyVnd.estado = True
                        MyVnd.Descripcion = ""
                    End If
                    Resp = Alta_Vendedor(MyVnd)
                Else
                    Resp = 1000 ' no se gestionan vendedores a toda la tienda
                End If
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "3" & BufDatos.BufVendedores(bucle)
                    FrmBalGloDat.dato "NOK -->  + Vnd. " & BufDatos.BufVendedores(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + Vnd. " & BufDatos.BufVendedores(bucle)
                End If
            End If
            If left(BufDatos.BufVendedores(bucle), 1) = "1" And BGrande = 0 Then
                ' Baja
                'If Mid(BufDatos.BufVendedores(bucle), 2, 1) = "0" Then
                    
                        If Len(BufDatos.BufVendedores(bucle)) = 43 Then
                            MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 2, 2)
                            MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 4, 6)
                            MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 10, 4)
                        Else
                            MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 2, 2)
                            MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 4, 4)
                            MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 8, 4)
                        End If
                    'MyVnd.NVendedor = Mid(BufDatos.BufVendedores(bucle), 5, 4)
                    'MyVnd.seccion = Mid(BufDatos.BufVendedores(bucle), 3, 2)
                    'MyVnd.tecla = Mid(BufDatos.BufVendedores(bucle), 15, 2)
                    'If MyVnd.NVendedor < 100 Then
                    '    MyVnd.NVendedor = MyVnd.seccion * 100 + MyVnd.NVendedor
                    'End If
                    
                    
                    MyVnd.estado = True
                    MyVnd.Descripcion = ""
                    Resp = Baja_vendedor(MyVnd)
                'Else
                '    Resp = 1000 ' no se gestionan vendedores a toda la tienda
                'End If
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "3" & BufDatos.BufVendedores(bucle)
                    FrmBalGloDat.dato "NOK -->  - Vnd. " & BufDatos.BufVendedores(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  - Vnd. " & BufDatos.BufVendedores(bucle)
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
End Sub
Private Sub Procesa_Paises(BufDatos As Tipo_Buf)
    ' No hay función de altas/bajas en HydraTouch para
    ' países, se hace directamente en la base
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim Resp As Integer
    Dim MyPais As Integer
    Dim MyNombre As String
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim MyMod As Boolean
    Set Base = OpenDatabase(Base_General)
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        For bucle = 0 To BufDatos.NPaises
            ' Alta
            If left(BufDatos.BufPaises(bucle), 1) = "0" And BGrande = 1 Then
                MyPais = Val(Mid(BufDatos.BufPaises(bucle), 2, 3))
                MyNombre = Mid(BufDatos.BufPaises(bucle), 5)
                MyMod = False
                Set Registro = Base.OpenRecordset("select * from paises where borrado=false and codigo=" & MyPais)
                If Not Registro.EOF Then
                    Edit_Record Registro
                    If Registro.Fields("codigo") <> MyPais Or Trim(Registro.Fields("nombre")) <> Trim(MyNombre) Then
                        Registro.Fields("modificado") = True
                        MyMod = True
                    End If
                Else
                    MyMod = True
                    Registro.AddNew
                    Registro.Fields("modificado") = False
                End If
                Registro.Fields("codigo") = MyPais
                Registro.Fields("nombre") = MyNombre
                If MyMod Then
                    Registro.Fields("tran_pais") = " "
                End If
                Registro.Fields("borrado") = False
                Registro.Update
                Resp = 0
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "P" & BufDatos.BufPaises(bucle)
                    FrmBalGloDat.dato "NOK -->  + Land. " & BufDatos.BufPaises(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + Land. " & BufDatos.BufPaises(bucle)
                End If
            End If
            ' Baja
            If left(BufDatos.BufPaises(bucle), 1) = "1" And BGrande = 0 Then
                MyPais = Val(Mid(BufDatos.BufPaises(bucle), 2, 3))
                Set Registro = Base.OpenRecordset("select * from paises where borrado=false and codigo=" & MyPais)
                If Not Registro.EOF Then
                    Edit_Record Registro
                    Registro.Fields("borrado") = True
                    Registro.Fields("tran_pais") = " "
                    Resp = 0
                    Registro.Update
                Else
                    Resp = 1000
                End If
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "P" & BufDatos.BufPaises(bucle)
                    FrmBalGloDat.dato "NOK -->  - Land. " & BufDatos.BufPaises(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  - Land. " & BufDatos.BufPaises(bucle)
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
    Base.Close
End Sub
Private Sub Procesa_Fichas(BufDatos As Tipo_Buf)
    ' No hay función de altas/bajas en HydraTouch para
    ' fichas, se hace directamente en la base
    Dim Arch2 As Integer
    Dim bucle As Integer
    Dim BGrande As Integer
    Dim MyFicha As DB_Ficha
    Dim Resp As Integer
    Dim MyPais As Integer
    Dim MyNombre As String
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim MyMod As Boolean
    Dim MyOK As Boolean
    Set Base = OpenDatabase(Base_General)
    Arch2 = FreeFile()
    Open App.Path & "\globaltouchno.dat" For Append As #Arch2
    For BGrande = 0 To 1
        For bucle = 0 To BufDatos.NFichas - 1
            ' Alta
            'V0nnnIIIIIIIIIIIIIIIIIIII,c       ,ee  ,r   ,s   ,llllllllllllllllllll,pppppp,CCC    ,DDD     ,NNN       ,PPP       ,SSS       ,Fecha_Despiece,Fecha_Nacimiento,Fecha_Sacrificio,RSI_Despiece        ,RSI_Sacrificio      ,texto
            '  NumIdentificador        categori edad raza sexo lote                 peso   crianza despiece nacimiento producción sacrificio --------       --------         --------         -------------------- --------------------
            '"V"                1,1
            ' (0/1)             2,1
            'Ficha              3,3
            'Identificador      6,20
            'Categoria         26,1
            'Edad              27,2
            'Raza              29,1
            'Sexo              30,1
            'Lote              31,20
            'Peso              51,6
            'Cod.P. Crianza    57,3
            'Cod.P. Despiece   60,3
            'Cod.P. Nacimiento 63,3
            'Cod.P. Producción 66,3
            'Cod.P. Sacrificio 69,3
            'Fecha despiece    72,8
            'Fecha nacimiento  80,8
            'fecha sacrificio  88,8
            'RSI despiece      96,20
            'RSI sacrificio   116,20
            '"gap"            136,5
            'texto            141,20
            
            
            If left(BufDatos.BufFichas(bucle), 1) = "0" And BGrande = 1 Then
                MyOK = True
                'If IsNumeric(Mid(BufDatos.BufFichas(bucle), 2, 3)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 25, 1)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 26, 2)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 28, 1)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 50, 6)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 56, 3)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 59, 3)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 62, 3)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 65, 3)) And _
                'IsNumeric(Mid(BufDatos.BufFichas(bucle), 68, 3)) Then
'V0001IDENTIFICADOR FICHA1     lote ficha1               084084084   084        01/01/1902/02/19RSI DESPIECE        RSI SACRIFICIO      -----....................
                    MyFicha.codigo = Val(Mid(BufDatos.BufFichas(bucle), 2, 3))
                    MyFicha.Identificador = Mid(BufDatos.BufFichas(bucle), 5, 20)
                    MyFicha.Categoria = Val(Mid(BufDatos.BufFichas(bucle), 25, 1))
                    MyFicha.Edad = Val(Mid(BufDatos.BufFichas(bucle), 26, 2))
                    MyFicha.Raza = Val(Mid(BufDatos.BufFichas(bucle), 28, 1))
                    MyFicha.Sexo = Mid(BufDatos.BufFichas(bucle), 29, 1)
                    MyFicha.lote = Mid(BufDatos.BufFichas(bucle), 30, 20)
                    MyFicha.Peso = 0 'Mid(BufDatos.BufFichas(bucle), 50, 6)
                    MyFicha.Cod_Crianza = Val(Mid(BufDatos.BufFichas(bucle), 56, 3))
                    MyFicha.Cod_Despiece = Val(Mid(BufDatos.BufFichas(bucle), 59, 3))
                    MyFicha.Cod_nacimiento = Val(Mid(BufDatos.BufFichas(bucle), 62, 3))
                    MyFicha.Cod_Produccion = Val(Mid(BufDatos.BufFichas(bucle), 65, 3))
                    MyFicha.Cod_Sacrificio = Val(Mid(BufDatos.BufFichas(bucle), 68, 3))
                    MyFicha.Fecha_Despiece = Mid(BufDatos.BufFichas(bucle), 71, 8)
                    MyFicha.Fecha_Nacimiento = Mid(BufDatos.BufFichas(bucle), 79, 8)
                    MyFicha.Fecha_Sacrificio = Mid(BufDatos.BufFichas(bucle), 87, 8)
                    MyFicha.RSI_Despiece = Mid(BufDatos.BufFichas(bucle), 95, 20)
                    MyFicha.RSI_Sacrificio = Mid(BufDatos.BufFichas(bucle), 115, 20)
                    MyFicha.texto = Mid(BufDatos.BufFichas(bucle), 140)
                    On Error Resume Next
                        DateSerial left(MyFicha.Fecha_Nacimiento, 2), Mid(MyFicha.Fecha_Nacimiento, 4, 2), Right(MyFicha.Fecha_Nacimiento, 2)
                    If Err.Number <> 0 Then
                        MyOK = False
                    Else
                        'DateSerial left(MyFicha.Fecha_Nacimiento, 2), Mid(MyFicha.Fecha_Nacimiento, 4, 2), Right(MyFicha.Fecha_Nacimiento, 2)
                        'If Err.Number <> 0 Then
                        '    MyOK = False
                        'Else
                            DateSerial left(MyFicha.Fecha_Sacrificio, 2), Mid(MyFicha.Fecha_Sacrificio, 4, 2), Right(MyFicha.Fecha_Sacrificio, 2)
                            If Err.Number <> 0 Then MyOK = False
                        'End If
                    End If
                    On Error GoTo 0
                    'If (MyFicha.Cod_Crianza <> 0 Or MyFicha.Cod_nacimiento <> 0) And MyFicha.Cod_Produccion <> 0 Then MyOK = False
                    'If MyFicha.Cod_Crianza = 0 And MyFicha.Cod_nacimiento = 0 And MyFicha.Cod_Produccion = 0 Then MyOK = False
                'Else
                '    MyOK = False
                'End If
                If MyOK Then
                    Set Registro = Base.OpenRecordset("select * from  fichavacuno where borrado=false and codigo=" & Chr(34) & MyFicha.codigo & Chr(34))
                    If Not Registro.EOF Then
                        Edit_Record Registro
                        MyMod = True
                        Registro.Fields("modificado") = True
                    Else
                        MyMod = True
                        Registro.AddNew
                        Registro.Fields("modificado") = False
                    End If
                    With Registro
                        !codigo = MyFicha.codigo
                        !Identificador = MyFicha.Identificador
                        !codnacimiento = MyFicha.Cod_nacimiento
                        !codcrianza = MyFicha.Cod_Crianza
                        !codsacrificio = MyFicha.Cod_Sacrificio
                        !coddespiece = MyFicha.Cod_Despiece
                        !codproduccion = MyFicha.Cod_Produccion
                        !rsisacrificio = MyFicha.RSI_Sacrificio
                        !rsidespiece = MyFicha.RSI_Despiece
                        !Categoria = MyFicha.Categoria
                        !Raza = MyFicha.Raza
                        !fecha_sac = DateSerial(left(MyFicha.Fecha_Sacrificio, 2), Mid(MyFicha.Fecha_Sacrificio, 4, 2), Right(MyFicha.Fecha_Sacrificio, 2))
                        !Edad = MyFicha.Edad
                        !Sexo = 0 'MyFicha.Sexo
                        !Peso = MyFicha.Peso / 10
                        !texto = MyFicha.texto
                        !fecha_nac = DateSerial(left(MyFicha.Fecha_Nacimiento, 2), Mid(MyFicha.Fecha_Nacimiento, 4, 2), Right(MyFicha.Fecha_Nacimiento, 2))
                        !fecha_des = DateSerial(left(MyFicha.Fecha_Sacrificio, 2), Mid(MyFicha.Fecha_Sacrificio, 4, 2), Right(MyFicha.Fecha_Sacrificio, 2))
                        !lote = MyFicha.lote
                        If MyMod Then !tranficha = " "
                        !borrado = False
                        .Update
                        Resp = 0
                    End With
                Else
                    Resp = 1000
                End If
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "V" & BufDatos.BufFichas(bucle)
                    FrmBalGloDat.dato "NOK -->  + Card. " & BufDatos.BufFichas(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  + Card. " & BufDatos.BufFichas(bucle)
                End If
            End If
            ' Baja
            If left(BufDatos.BufFichas(bucle), 1) = "1" And BGrande = 0 Then
                MyFicha.codigo = Val(Mid(BufDatos.BufFichas(bucle), 2, 3))
                Set Registro = Base.OpenRecordset("select * from fichavacuno where borrado=false and codigo=" & Chr(34) & MyFicha.codigo & Chr(34))
                If Not Registro.EOF Then
                    Edit_Record Registro
                    Registro.Fields("borrado") = True
                    Registro.Fields("tranficha") = " "
                    Resp = 0
                    Registro.Update
                Else
                    Resp = 1000
                End If
                If Resp <> 0 And Resp <> 1 And Resp <> 4 Then
                    Print #Arch2, "V" & BufDatos.BufFichas(bucle)
                    FrmBalGloDat.dato "NOK -->  - Card. " & BufDatos.BufFichas(bucle)
                Else
                    FrmBalGloDat.dato "OK -->  - Card. " & BufDatos.BufFichas(bucle)
                End If
            End If
        Next bucle
    Next BGrande
    Close #Arch2
    Base.Close
End Sub

Public Sub Exporta_GlobalDat()
    ' genera un global.dat con la información de la base
    ' (incompleto)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegAux As dao.Recordset
    Dim Arch As Integer
    Dim Buf As String
    Dim sParte As String
    Load FrmBalGloDat
    FrmBalGloDat.Show
    Arch = FreeFile()
    Open App.Path & "\global.dat" For Output As #Arch
    Set Base = OpenDatabase(Base_General)
        Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    If Len(!descripcio) < 21 Then
                        Buf = "00" & Format(!secc_maqui, "00") & "000000" & !descripcio & Space(21 - Len(!descripcio)) & Format(!codi_ident, "000")
                    Else
                        sParte = Mid(!descripcio, 1, 21)
                        Buf = "00" & Format(!secc_maqui, "00") & "000000" & sParte & Format(!codi_ident, "000")
                    End If
                    Select Case !enviardatos
                        Case "GA"
                            Buf = Buf & "1"
                        Case "SC10"
                            Buf = Buf & "2"
                        Case Else
                            Buf = Buf & "0"
                    End Select
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    .Movenext
                Loop
            End If
        End With
        Set Registro = Base.OpenRecordset("select * from equipos where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    Buf = "40" & Format(!codi_ident, "00") & Format(!numero_eqp, "00")
                    Select Case !modelo
                        Case 0
                            Buf = Buf & "00" & !prog_eqp
                        Case 1
                            Buf = Buf & "01" & !prog_eqp
                        Case 100
                            Buf = Buf & "02" & Space(15)
                    End Select
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    .Movenext
                Loop
            End If
        End With
        Set Registro = Base.OpenRecordset("select * from articulo where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    Buf = "50" & Format(!codigo, "000000") & Format(!codi_ident, "00") & Format(!Plu, "0000")
                    Buf = Buf & Format(!codi_sub, "0000") & Format(!codi_fam, "000")
                    If UsaEuro Then
                        Buf = Buf & Format(!Euros * 100, "00000")
                    Else
                        Buf = Buf & Format(!precio * (10 ^ decimales), "00000")
                    End If
                    Buf = Buf & !codi_pes & Format(!caducidad, "000") & "0" & Mid(Trim(!des_plu1), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu1), 1, 25)))
                    Buf = Buf & "00000" & Format(!tara, "00000") & !tipo_iva & Format(!grupo_conserv, "000")
                    Buf = Buf & Format(!Pref, "000") & Format(!Etq, "000") & Format(!fcb, "00")
                    If Not IsNull(!art_cb) Then
                        'Buf = Buf & !art_cb & Space(12 - Len(!art_cb))
                        If Len(Trim(!art_cb)) < 12 Then
                            Buf = Buf & Trim(!art_cb) & Space(12 - Len(Trim(!art_cb)))
                        Else
                            Buf = Buf & Mid(!art_cb, 1, 12)
                        End If
                    Else
                        Buf = Buf & Space(12)
                    End If
                    Buf = Buf & Format(!ning, "00")
                    FrmBalGloDat.dato Buf
                    If Len(Buf) <> 92 Then
                        
                    End If
                    Print #Arch, Buf
                    
                    Buf = "Z" & Format(!codigo, "000000") & Mid(Trim(!des_plu2), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu2), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu3), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu3), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu4), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu4), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu5), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu5), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu6), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu6), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu7), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu7), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu8), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu8), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu9), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu9), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu0), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu0), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plux), 1, 25) & Space(25 - Len(Mid(Trim(!des_plux), 1, 25)))
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    Buf = "Y" & Format(!codigo, "000000") & Mid(Trim(!des_plu11), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu11), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu12), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu12), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu13), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu13), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu14), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu14), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu15), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu15), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu16), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu16), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu17), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu17), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu18), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu18), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu19), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu19), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu20), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu20), 1, 25)))
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    .Movenext
                Loop
            End If
        End With
        Set Registro = Base.OpenRecordset("select * from teclas where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    If IsNull(!numero_eqp) Then
                        Set RegAux = Base.OpenRecordset("select codigo,codi_ident from articulo where borrado=false and codigo=" & !codigo)
                        If Not RegAux.EOF Then
                            Buf = "@00" & Format(RegAux!codi_ident, "00")
                        Else
                            Buf = "Z"
                        End If
                    Else
                        Buf = "@01" & Format(!numero_eqp, "00")
                    End If
                    Buf = Buf & Format(!codi_tec, "000") & Format(!codigo, "000000") & Format(!Tabla, "00")
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    .Movenext
                Loop
            End If
        End With
        Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    .Movenext
                Loop
            End If
        End With
        Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    .Movenext
                Loop
            End If
        End With
        Registro.Close
    Base.Close
    Close #Arch
    Unload FrmBalGloDat
End Sub
    
Public Sub Exporta_GlobalDat_Art(Base As dao.Database)
    ' genera un global.dat con la información de la base
    ' (incompleto)
    'Dim base as dao.database
    Dim Registro As dao.Recordset
    Dim RegAux As dao.Recordset
    Dim Arch As Integer
    Dim Buf As String
    On Error GoTo errglbart
    CadenadeLog "Entrada en Exporta_globaldat_art con:" & Base.name
    Load FrmBalGloDat
    FrmBalGloDat.Show
    Arch = FreeFile()
    Open App.Path & "\global.dat" For Output As #Arch
    CadenadeLog "Creado  archivo Global.dat para escritura..."
        'Set Registro = Base.OpenRecordset("select * from articulo where ((borrado=false) and (tran_plu<>'*'))") 1.7.24
        Set Registro = Base.OpenRecordset("select * from articulo where (isnull(tran_plu) or tran_plu<>" & Chr(34) & "*" & Chr(34) & " or isnull(tran_texto) or tran_texto<>" & Chr(34) & "*" & Chr(34) & " or isnull(tran_tx1) or tran_tx1<>" & Chr(34) & "*" & Chr(34) & ")")
        
        With Registro
            CadenadeLog "Número de artículos modificados:" & CStr(Registro.Recordcount)
            
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    'Buf = "50" & Format(!codigo, "000000") & Format(!codi_ident, "00") & Format(!plu, "0000") 1.7.24
                    'se consideran las bajas...
                    If !borrado Then
                        Buf = "51"
                    Else
                        Buf = "50"
                    End If
                    Buf = Buf & Format(!codigo, "000000") & Format(!codi_ident, "00") & Format(!Plu, "0000")
                    ''''''''''''''''''''''''''''
                    
                    Buf = Buf & Format(!codi_sub, "0000") & Format(!codi_fam, "000")
                    If UsaEuro Then
                        Buf = Buf & Format(!Euros * 100, "00000")
                    Else
                        Buf = Buf & Format(!precio * (10 ^ decimales), "00000")
                    End If
                    Buf = Buf & !codi_pes & Format(!caducidad, "000") & "0" & Mid(Trim(!des_plu1), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu1), 1, 25)))
                    Buf = Buf & "00000" & Format(!tara, "00000") & !tipo_iva & Format(!grupo_conserv, "000")
                    Buf = Buf & Format(!Pref, "000") & Format(!Etq, "000") & Format(!fcb, "00")
                    If Not IsNull(!art_cb) Then
                        'Buf = Buf & !art_cb & Space(12 - Len(!art_cb))
                        If Len(Trim(!art_cb)) < 12 Then
                            Buf = Buf & Trim(!art_cb) & Space(12 - Len(Trim(!art_cb)))
                        Else
                            Buf = Buf & Mid(!art_cb, 1, 12)
                        End If
                    Else
                        Buf = Buf & Space(12)
                    End If
                    Buf = Buf & Format(!ning, "00")
                    FrmBalGloDat.dato Buf
                    If Len(Buf) <> 92 Then
                        
                    End If
                    Print #Arch, Buf
                    Buf = "Z" & Format(!codigo, "000000") & Mid(Trim(!des_plu2), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu2), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu3), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu3), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu4), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu4), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu5), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu5), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu6), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu6), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu7), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu7), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu8), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu8), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu9), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu9), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu0), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu0), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plux), 1, 25) & Space(25 - Len(Mid(Trim(!des_plux), 1, 25)))
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    Buf = "Y" & Format(!codigo, "000000") & Mid(Trim(!des_plu11), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu11), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu12), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu12), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu13), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu13), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu14), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu14), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu15), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu15), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu16), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu16), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu17), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu17), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu18), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu18), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu19), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu19), 1, 25)))
                    Buf = Buf & Mid(Trim(!des_plu20), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu20), 1, 25)))
                    FrmBalGloDat.dato Buf
                    Print #Arch, Buf
                    If TiendaActual = 999 Then
                        .Edit
                        !tran_plu = "*"
                        !tran_texto = "*"
                        !tran_tx1 = "*"
                        .Update
                    End If
                    
                    .Movenext
                Loop
            End If
        End With
        Registro.Close
    'base.Close
errglbart:
    If Err.Number <> 0 Then
        CadenadeLog "Error generando Global.dat --> " & CStr(Err.Number) & " Desc.:" & Err.Description
    End If
    On Error GoTo 0
    Close #Arch
    Unload FrmBalGloDat
End Sub
    
Public Sub Exporta_GlobalDat_Art_todos(Base As dao.Database)
    ' genera un global.dat con la información de la base
    ' (incompleto)
    'Dim base as dao.database
    Dim Registro As dao.Recordset
    Dim RegAux As dao.Recordset
    Dim Arch As Integer
    Dim Buf As String
    Load FrmBalGloDat
    FrmBalGloDat.Show
    Arch = FreeFile()
    Open App.Path & "\global.dat" For Output As #Arch
        Set Registro = Base.OpenRecordset("select * from articulo where borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    Buf = "50" & Format(!codigo, "000000") & Format(!codi_ident, "00") & Format(!Plu, "0000")
                    Buf = Buf & Format(!codi_sub, "0000") & Format(!codi_fam, "000")
                    If UsaEuro Then
                        Buf = Buf & Format(!Euros * 100, "00000")
                    Else
                        Buf = Buf & Format(!precio * (10 ^ decimales), "00000")
                    End If
                    Buf = Buf & !codi_pes & Format(!caducidad, "000") & "0" & Mid(Trim(!des_plu1), 1, 25) & Space(25 - Len(Mid(Trim(!des_plu1), 1, 25)))
                    Buf = Buf & "00000" & Format(!tara, "00000") & !tipo_iva & Format(!grupo_conserv, "000")
                    Buf = Buf & Format(!Pref, "000") & Format(!Etq, "000") & Format(!fcb, "00")
                    
                    If Not IsNull(!art_cb) Then
                        'Buf = Buf & !art_cb & Space(12 - Len(!art_cb))
                        If Len(Trim(!art_cb)) < 12 Then
                            Buf = Buf & Trim(!art_cb) & Space(12 - Len(Trim(!art_cb)))
                        Else
                            Buf = Buf & Mid(!art_cb, 1, 12)
                        End If
                    Else
                        Buf = Buf & Space(12)
                    End If
                    Buf = Buf & Format(!ning, "00")
                    FrmBalGloDat.dato Buf
                    If Len(Buf) <> 92 Then
                        
                    End If
                    Print #Arch, Buf
                    
                    .Movenext
                Loop
            End If
        End With
        Registro.Close
    'base.Close
    Close #Arch
    Unload FrmBalGloDat
End Sub
    

Public Sub Alba_EnviaTiquetsPendientes(Euroscale_o_SC10 As Boolean, Optional MiFecha As Date)
Dim db As Database
Dim rst As Recordset
Dim rstAux As Recordset
' Cabecera
Dim B1 As String
Dim B2 As String
Dim Codigo_mos As String
Dim Codigo_bal As String
Dim Codigo_ven As String
Dim Cod_Bar As String
Dim Lineas As String
Dim precio As String
Dim Tipo As String
Dim Fecha As String
Dim hora As String
Dim LineaAnulada As String
'********
' Linea
Dim Num_lin As String
Dim Cod_Plu As String
Dim Cod_Cod As String
Dim Precio_Uni As String
Dim Pes_Uni As String
Dim Total As String
Dim Signo As String
Dim Tipo_Linea As String
Dim Anulado As String
'************
Dim MiEnvio As String
Dim nintentos As Integer
Dim bucle As Integer
Dim RegIVA As Recordset
Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim NombreTiquet As String
Dim BufferLineas() As String
Dim BufferLineasAD1() As String
Dim MiCuentaLineas As Integer
Dim BucleLineas As Integer
Dim HayCabecera As Boolean
Dim PrecioSinIVA As Double
Dim PrecioConIVA As Double
Dim BufferIVA As String
Dim BufferIVA2 As String
Dim MiIva(5) As Tipo_Grupo_IVA
Dim BucleIVA As Integer
Dim NombreF As Tipo_NombreTiquet
Dim Fichero2 As Integer
Dim nContAs As Integer
Dim sVm As String
Dim sVmS As String
Dim sCadAd As String
Dim scadAD1 As String
Dim sCadfP As String
Dim sTiqGen As String
Dim sPes_Uni As String
Dim sFactura As String
Dim sParte As String

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim sSQL As String
CadenadeLog "Traspaso Tiques Pendientes>"
    If IsDate(MiFecha) And CStr(MiFecha) <> "0:00:00" Then
        sTiqGen = "tqgen" & Format(Day(MiFecha), "00")
        If Month(MiFecha) < 10 Then
            sTiqGen = sTiqGen & CStr(Month(MiFecha))
        Else
            sTiqGen = sTiqGen & Chr(Asc("A") + Month(MiFecha) - 10)
        End If
    Else
    
        sTiqGen = "tqgen" & Format(Day(Now), "00")
        If Month(Now) < 10 Then
            sTiqGen = sTiqGen & CStr(Month(Now))
        Else
            sTiqGen = sTiqGen & Chr(Asc("A") + Month(Now) - 10)
        End If
    
    End If
    
    'Fichero1 = FreeFile()
    If Euroscale_o_SC10 Then
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    Else
        MiTabla1 = "cabeceraSC10"
        MiTabla2 = "ticketsSC10"
    End If
    sParte = ""
    If IsDate(MiFecha) And CStr(MiFecha) <> "0:00:00" Then
        sParte = "#" & Format(MiFecha, "dd/mm/yyyy") & "#"
    End If
    Set db = AbrirBase()
    '*******************
    ' carga tipos de IVA
    '*******************
    'Set RegIVA = db.OpenRecordset("SELECT * FROM iva")
    'For Bucle = 0 To 4
    '    RegIVA.FindFirst "tipo=" & (Bucle + 1)
    '    If RegIVA.EOF Or RegIVA.NoMatch Then
    '        MiIva(Bucle).porcentaje = 0
    '    Else
    '        MiIva(Bucle).porcentaje = RegIVA!Valor
    '    End If
    'Next Bucle
    '*******************
    ' /carga tipos de IVA
    '*******************
    If sParte = "" Then
    sSQL = "SELECT * " & _
           "FROM " & MiTabla1 & " " & _
           "WHERE reserv3=" & Chr(34) & "NO" & Chr(34) & " or reserv3=" & Chr(34) & "--" & Chr(34) & " " & _
           "ORDER BY val(nume)"
    Else
    db.Execute "update " & MiTabla1 & " set reserv3='NO' where d_fecha=" & sParte
    sSQL = "SELECT * " & _
           "FROM " & MiTabla1 & " " & _
           "WHERE d_fecha=" & sParte & _
           "ORDER BY val(nume)"
    
    End If
    
    Set rst = db.OpenRecordset(sSQL)
    If Not rst.EOF Then
        rst.MoveFirst
        ' si no está enviada la cabecera, se envía
        Do Until rst.EOF
        With rst
            If !reserv3 = "NO" Or !reserv3 = "--" Then
                HayCabecera = True
                Codigo_mos = Format(Val(!seccion), "00")
                Codigo_bal = Format(Val(!Balanza), "00")
                Codigo_ven = Format(Val(!vended), "0000")
                If Not IsNull(!factura) Then
                    sFactura = !factura
                Else
                    sFactura = Space(25)
                End If
                If Len(sFactura) < 25 Then
                    sFactura = sFactura & Space(25 - Len(sFactura))
                End If
                
                If !codbar <> "" Then
                    Cod_Bar = !codbar
                Else
                    Cod_Bar = Space(13)
                End If
                Lineas = Format(Val(!numlin), "000")
                If !Importe >= 0 Then
                    precio = !Importe
                    precio = Quitar_Coma(precio, 8)
                Else
                    precio = (-1) * !Importe
                    precio = "-" & Quitar_Coma(precio, 7)
                End If
                ' Tipo V -> venta red
                ' Tipo L -> venta local
                ' Tipo A -> venta anulado red
                ' Tipo B -> venta anulado local
                '************
                ' de que tipo
                ' 0 --> Venta
                ' 1 --> Super
                ' 2 --> Envasado
                ' 3 --> Autoservicio
                ' 4 --> Venta Local
                ' 5 --> Super Local
                ' 6 --> Envasado Local
                ' 7 --> Autoservicio Local
                ' 9 --> anulados....
                '*************************
                If (!STPETIC And 64) = 0 Then
                    
                    If (!STPETIC And 1) = 0 Then
                        Tipo = "V"
                    Else
                        Tipo = "L"
                    End If
                    
                    
                Else
                    
                    If (!STPETIC And 1) = 0 Then
                        Tipo = "A"
                    Else
                        Tipo = "B"
                    End If
                    
                End If
                NombreF.Balanza = Val(!Balanza)
                Select Case !TYPTIC
                    Case 1
                        NombreF.Modo = 1
                    Case 2
                        NombreF.Modo = 2
                    Case 3
                        NombreF.Modo = 3
                        
                        If Tipo = "A" Or Tipo = "B" Then
                            Tipo = "F"
                        Else
                            Tipo = "E"
                        End If
                        
                    Case 4
                        NombreF.Modo = 4
                    Case 5
                        NombreF.Modo = 1
                    Case 6
                        NombreF.Modo = 2
                    Case 7
                        NombreF.Modo = 3
                    Case 8
                        NombreF.Modo = 4
                End Select
                NombreF.cliente = Val(!codcli)
                NombreF.vendedor = Val(!vended)
                NombreF.seccion = Val(!seccion)
                NombreF.tiquet = Val(!nume)
                Fecha = Format(!Fecha, "ddmmyyyy")
                hora = Format(!hora, "hhmm")
                ' Enviar la cabecera
                MiEnvio = "H" & _
                Codigo_mos & Codigo_bal & Codigo_ven & Cod_Bar & _
                Lineas & precio & Tipo & Fecha & hora
                ' Envia la cadena
                sCadAd = ""
                'If Dir(App.Path & "\nscli.txt") <> "" Then
                '    If IsNull(!numlote) Then
                '        sCadAd = Space(20)
                '    Else
                '        sCadAd = Trim(!numlote) & Space(20 - Len(Trim(!numlote)))
                '    End If
                '    sCadAd = sCadAd & Format(Val(!codcli), "000000")
                'Else
                    sCadAd = ""
                'End If
                sCadfP = ""
                'If Dir(App.Path & "\nsfpa.txt") <> "" Then
                '    If IsNull(!tipoPago) Then
                '        sCadfP = "0"
                '    Else
                '        sCadfP = CStr(!tipoPago)
                '    End If
                'Else
                '    sCadfP = ""
                'End If
                'If Dir(App.Path & "\nsfpa.txt") <> "" Then 'begines...
                    If IsNull(!tipoPago) Then
                        sCadfP = "00"
                    Else
                        sCadfP = Format(!tipoPago, "00")
                    End If
                'Else
                '    sCadfP = ""
                'End If
                
                ReDim BufferLineas(1)
                BufferLineas(0) = MiEnvio & sCadAd & "0" '& Format(TiendaActual, "00")
                MiCuentaLineas = 1
                NombreTiquet = ObtenNombre(NombreF)
                nContAs = 0
                sVm = "0"
                If InStr(1, NombreTiquet, "*") <> 0 Then
                    For MiCuentaLineas = 1 To Len(NombreTiquet)
                        If Mid(NombreTiquet, MiCuentaLineas, 1) = "*" Then
                            nContAs = nContAs + 1
                            
                        End If
                    Next MiCuentaLineas
                Else
                    sVm = Codigo_bal
                    sVmS = Codigo_mos
                End If
                If nContAs <> 0 Then
                    NombreTiquet = Mid(NombreTiquet, nContAs + 1)
                    'If Dir(App.Path & "\transpluasc.exe") <> "" Then
                    '    sVm = Mid(Cod_Bar, nContAs + 1, 2)
                    'Else
                        sVm = Mid(Cod_Bar, nContAs + 1, nContAs)
                    'End If
                    sVmS = Mid(Cod_Bar, 3, 2)
                    NombreTiquet = sVm & NombreTiquet
                    'ATENCIÓN HIPERUSERA:
                    ' SE PONE COMO SECCIÓN LOS 2 DÍGITOS SIGUIENTES EN EL CÓDIGO DE BARRAS A LA
                    ' bALANZA rELATIVA
                    'If Dir(App.Path & "\transpluasc.exe") <> "" Or Dir(App.Path & "\_transpluasc.exe") <> "" Then
                    '    NombreTiquet = Mid(NombreTiquet, 1, 1) & Mid(Cod_Bar, 3, 2) & Mid(NombreTiquet, 4)
                    'End If
                    '''''''''''''''''''''
                End If
                '//Inco ver si se activa cabecera If Dir(App.Path & "\sncli.txt") = "" Then
                    MiCuentaLineas = 0
                '//End If
                'NombreTiquet = Codigo_bal & Format(!nume, "0000") & "." & _
                'Left(fecha, 2) & Hex$(Val(Mid(fecha, 3, 2)))
                NombreTiquet = NombreTiquet & "." & left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
                rst.Edit
                .Fields("reserv3") = "--"
                .Update
            Else
                HayCabecera = False
            End If
        End With
        If sParte = "" Then
        Set rstAux = db.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
        Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & rst!nume & Chr(34) & _
                 " and numbal=" & rst!Balanza & " and date=" & Chr(34) & rst!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & rst!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & rst!TYPTIC & Chr(34) & " order by val(numlin), lincan") 'c2f 1.6.12 - 1
        Else
        Set rstAux = db.OpenRecordset("select * from " & MiTabla2 & _
                 " where nume=" & Chr(34) & rst!nume & Chr(34) & _
                 " and numbal=" & rst!Balanza & " and date=" & Chr(34) & rst!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & rst!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & rst!TYPTIC & Chr(34) & " order by val(numlin), lincan") 'c2f 1.6.12 - 1
        
        End If
                 
        With rstAux
            If Not .EOF Then
                .MoveFirst
                 'For BucleIVA = 0 To 4
                 '   MiIva(BucleIVA).Base = 0
                 '   MiIva(BucleIVA).Importe = 0
                 'Next BucleIVA
                 Do Until .EOF
                    With rstAux
                        Num_lin = Format(!numlin, "000")
                        Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        Cod_Cod = Format(!Code, "000000")
                        'If !price >= 0 Then
                        '    Precio_Uni = !price
                        '    Precio_Uni = Quitar_Coma(Precio_Uni, 8)
                        'Else
                        '    Precio_Uni = (-1) * !price
                        '    Precio_Uni = (-1) * Quitar_Coma(Precio_Uni, 7)
                        'End If
                        'If !IVA >= 0 And !IVA < 5 Then
                        '    MiIva(!IVA).Importe = MiIva(!IVA).Importe + !amount
                        '    MiIva(!IVA).Base = MiIva(!IVA).Base + PRsinIVA(!amount, MiIva(!IVA).porcentaje)
                        'End If
                        If !units = 0 Then
                            sPes_Uni = Format(Abs(!Weight), "000.000")
                            'If !Weight >= 0 Then
                            '    Pes_Uni = !Weight
                            '    Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            'Else
                            '    Pes_Uni = (-1) * !Weight
                            '    Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            'End If
                            Tipo_Linea = "P"
                        Else
                            sPes_Uni = Format(Abs(!units), "000.000")
                            'If !units > 0 Then
                            '    Pes_Uni = !units
                            '    Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            'Else
                            '    Pes_Uni = (-1) * !units
                            '   Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            'End If
                            Tipo_Linea = "U"
                        End If
                        If !amount < 0 Then
                            'Total = (-1) * !amount
                            'Total = "-" & Quitar_Coma(Total, 7)
                            Signo = "-"
                        Else
                            'Total = !amount
                            'Total = Quitar_Coma(Total, 8)
                            Signo = "+"
                        End If
                        If !lincan = 0 Then
                            LineaAnulada = "N"
                        Else
                            LineaAnulada = "S"
                        End If
                        '*MiEnvio = "L" & Codigo_mos & _
                        '*Cod_Bar & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Total & Signo & _
                        '*Tipo_Linea & LineaAnulada & Cod_Cod & Format(Val(!IVA), "0")
                        ' Envia la cadena
                        'Format(rst!nume, "0000") & Num_lin & Codigo_mos & Format(Val(sVm), "00") & Format(Val(!Vendor), "0000")
                        scadAD1 = ""
                        'If Dir(App.Path & "\transpluasc.exe") <> "" Or Dir(App.Path & "\_transpluasc.exe") <> "" Then
                        'MiEnvio = Format(rst!nume, "0000") & Num_lin & sVmS & "0" & sVm & Format(Val(!Vendor), "0000") & _
                        '        Cod_Bar & Cod_Cod & Cod_Plu & Format(Abs(rst!Importe), "00000.00") & Format(Abs(!Price), "00000.00") & _
                        '        sPes_Uni & Format(Abs(!amount), "00000.00") & Signo & Tipo_Linea & LineaAnulada & Mid(CStr(!IVA), 1, 1) & _
                        '        Tipo & Format(!D_FECHA, "ddmmyyyy") & Format(!D_HORA, "hhmm")
                       '
                        'Else
                            'ver MaIllan
                            MiEnvio = Format(rst!nume, "00000") & Num_lin & sVmS & sVm & Format(Val(!Vendor), "0000") & _
                                    Cod_Bar & Cod_Cod & Cod_Plu & Format(Abs(rst!Importe), "00000.00") & Format(Abs(!Price), "00000.00") & _
                                    sPes_Uni & Format(Abs(!amount), "00000.00") & Signo & Tipo_Linea & LineaAnulada & Mid(CStr(!IVA), 1, 1) & _
                                    Tipo & Format(!D_FECHA, "ddmmyyyy") & Format(!D_HORA, "hhmm") & Format(TiendaActual, "00")
                            If Dir(App.Path & "\fichalote.txt") <> "" Then
                                scadAD1 = Format(Val(!reserv1), "0000")
                            End If
                        'End If
                        MiEnvio = sin_Coma(MiEnvio)
                        ReDim Preserve BufferLineas(MiCuentaLineas + 1)
                        ReDim Preserve BufferLineasAD1(MiCuentaLineas + 1)
                        BufferLineas(MiCuentaLineas) = MiEnvio
                        BufferLineasAD1(MiCuentaLineas) = scadAD1
                        MiCuentaLineas = MiCuentaLineas + 1
                        rstAux.Edit
                        .Fields("reserv3") = "SI"
                        .Update
                        .Movenext
                    End With
                 Loop
                 
                 If HayCabecera Then
                     'Fichero1 = FreeFile()
                     On Error GoTo 0
                     'Open App.Path & "\tiquet.buffer" For Output As #Fichero1
                     'If lFichGen = True Then
                        Fichero2 = FreeFile()
                        Open App.Path & "\" & sTiqGen For Append As #Fichero2
                     'End If
                     For BucleLineas = 0 To MiCuentaLineas - 1
                        'Print #Fichero1, BufferLineas(BucleLineas) & sCadAd & sCadfP & BufferLineasAD1(BucleLineas)
                        'If lFichGen = True Then
                            Print #Fichero2, BufferLineas(BucleLineas) & sCadAd & "0" & BufferLineasAD1(BucleLineas) & sCadfP & sFactura
                        'End If
                     Next BucleLineas
                     'Close #Fichero1
                     'If lFichGen = True Then
                        Close #Fichero2
                     'End If
                     'frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905) & " " & Cod_Bar
                     'FileCopy App.Path & "\tiquet.buffer", App.Path & "\" & NombreTiquet
                     'FileCopy App.Path & "\tiquet.buffer", var.PathExpor & "\" & NombreTiquet
                 
                 End If
                 
            Else
                rst.Edit
                rst.Fields("reserv3") = "SI"
                rst.Update
            End If
        End With
        rst.Movenext
        Loop
    End If
    db.Close
    Exit Sub
ErrorPath:
    MsgBox "Error al acceder a " & App.Path
End Sub
    
Public Function sin_Coma(ByVal MiCadena As String) As String

    ' *****************
    ' Esta función sustituye
    ' el caracter "," (coma) por
    ' el caracter "." (punto)
    Dim bucle As Integer
    Dim Buffer As String
    Buffer = ""
    For bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, bucle, 1) = "," Then
            Buffer = Buffer & "."
        Else
            Buffer = Buffer & Mid(MiCadena, bucle, 1)
        End If
    Next bucle
    sin_Coma = Buffer
End Function


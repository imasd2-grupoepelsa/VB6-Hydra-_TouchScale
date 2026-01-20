Attribute VB_Name = "SQLSERVER"
Option Explicit

Public MyConnObj As New ADODB.Connection
Public MyRecSet As New ADODB.Recordset
Public strStream As New ADODB.Stream
Dim MyMinSec As Integer

'Items
'-----
'Id   INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(127) NOT NULL
'ShortName Not VarChar(64)
'TexT Not VarChar(1024)
'Text1 Not VarChar(110)
'Text2 Not VarChar(110)
'Text3 Not VarChar(110)
'Text4 Not VarChar(110)
'Text5 Not VarChar(110)
'Family INTEGER
'Price DECIMAL(6,2) DEFAULT 0
'Discount Not DECIMAL(4, 2)
'SaleForm SMALLINT DEFAULT 1
'PLU INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'VAT SMALLINT DEFAULT 0
'EAN13 VarChar(13)
'Expiring SMALLINT DEFAULT 0
'BestBefore SMALLINT DEFAULT 0
'Tare DECIMAL(5,3) DEFAULT 0
'Label Not VarChar(64)
'Icon Not VarChar(127)
'QLevel Not VarChar(16)
'Unique Index(Code), Index(Name(19))'
' ShortName --> Descriptivo 1 == Name (Hasta 64 de longitud)
' Text1     --> Descriptivos 2 a 5
' Text2     --> Descriptivos 6 a 9
' Text3     --> Descriptivos 10 a 13
' Text4     --> Descriptivos 14 a 17
' Text5     --> Descriptivos 18 a 21
' Text      --> Se guarda en 2040 troceado de 80 en 80
' Discount  --> Se guarda en TIP_LETX
' Label     --> Se guarda en Tran_EP
' Icon y QLevel --> Se guardan en Imagen --> Icon &  " _+_ " & QLevel --> NO SE GUARDA QLevel
'

'PLUKeys
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Section INTEGER 'grupo en la Balanza
'Board INTEGER
'PLUKey INTEGER
'Code INTEGER
'Unique Index(Section, Board, PLUKey)'
'
'QLevels
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code VARCHAR(16) UNIQUE NOT NULL
'Icon Not VarChar(64)
'Description Not VarChar(127)''
'
'Vendors
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL'
'
'VendorKeys
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Counter INTEGER
'VendorKey INTEGER
'Vendor INTEGER UNIQUE
'Unique Index(Counter, VendorKey)'
'
'HTickets
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NetStat CHAR(1) DEFAULT ' '
'Captured CHAR(1) DEFAULT ' '
'tckType CHAR(1)
'status CHAR(1)
'Number INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'Scale INTEGER DEFAULT 0
'Vendor INTEGER
'DatTim DateTime
'ReloadFrom INTEGER DEFAULT 0
'Customer INTEGER DEFAULT 0
'tckDiscount DECIMAL(4,2) DEFAULT 0
'Unique Index(Vendor, id)'
'
'LTickets
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'IdHTicket INTEGER
'lineType SMALLINT
'SaleForm SMALLINT
'Item INTEGER
'Weight DECIMAL(5, 3)
'Tare DECIMAL(5,3) DEFAULT 0
'TTare CHAR(1) DEFAULT ' '
'lineDiscount DECIMAL(4, 2)
'Price DECIMAL(6, 2)
'Amount DECIMAL(8, 2)
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Index (IdHTicket)
'
'PayTickets
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'IdHTicket INTEGER
'PayMethod INTEGER
'Delivered DECIMAL(8,2) DEFAULT 0
'Index (IdHTicket)'
'
'PayMethods
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Name VARCHAR(64) UNIQUE NOT NULL'
'
'PacketLabels
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Number INTEGER DEFAULT 0
'Counter INTEGER DEFAULT 0
'Scale INTEGER DEFAULT 0
'Vendor INTEGER
'lineType SMALLINT
'Item INTEGER
'SaleForm SMALLINT
'Weight DECIMAL(5, 3)
'Tare DECIMAL(5,3) DEFAULT 0
'TTare CHAR(1) DEFAULT ' '
'Price DECIMAL(6, 2)
'lineDiscount DECIMAL(4, 2)
'Amount DECIMAL(8, 2)
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Index (IdHTicket)
'
'LocalVendorPacket
'-----------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Vendor INTEGER UNIQUE
'Auto BOOLEAN
'ClearStart BOOLEAN
'Repetitions SMALLINT DEFAULT 0
'Copies SMALLINT DEFAULT 0
'EAN13 CHAR(1) DEFAULT '1'
'BovineCard SMALLINT(3)
'EarTag VarChar(30)
'BovineLot VarChar(20)
'Customer INTEGER DEFAULT 0
'PresetON BOOLEAN DEFAULT FALSE
'Item INTEGER DEFAULT 0
'SaleForm SMALLINT
'Price DECIMAL(6, 2)
'Units INTEGER UNSIGNED NOT NULL DEFAULT 0
'Tare DECIMAL(5, 3)
'Discount DECIMAL(4, 2)'
'
'LocalPacketPresets
'------------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Vendor INTEGER
'Level SMALLINT
'Type SMALLINT
'Value DECIMAL(6, 3)
'Number INTEGER DEFAULT 1
'ItemTotal INTEGER DEFAULT 0
'LowerLabels INTEGER UNSIGNED NOT NULL DEFAULT 0
'TotalUnits INTEGER UNSIGNED NOT NULL DEFAULT 0
'TotalWeight DECIMAL(12,3) UNSIGNED NOT NULL DEFAULT 0
'TotalAmount DECIMAL(12,2) UNSIGNED NOT NULL DEFAULT 0
'Unique Index(Vendor, Level)
'
'SServiceLabels
'--------------
'LIKE PacketLabels'
'
'NumTickets
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) DEFAULT 'T'
'Identifier INTEGER
'Number INTEGER
'UNIQUE INDEX (Identifier,Type)'
'
'EraseTickets
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) UNIQUE DEFAULT 'T'
'DatTim DateTime'
'
'NumberingTickets
'----------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Type CHAR(1) UNIQUE DEFAULT 'T'
'Numbering SMALLINT DEFAULT 0
'ResetNumber BOOLEAN DEFAULT FALSE
'SaveMode CHAR(1) DEFAULT ' ''
'
'Counters
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL
'Description Not VarChar(127)'
'
'Scales
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'MAC VARCHAR(17) UNIQUE NOT NULL
'IP Not VarChar(16)
'Number INTEGER
'Name VARCHAR(64) NOT NULL
'Type CHAR(1)'
'
'Families
'--------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(54) NOT NULL
'IsBovine BOOLEAN DEFAULT FALSE
'Unique Index(Code)
'Label Not VarChar(64)
'Index (Name(19))'
' IsBovine  --> se guarda en Posicion
' Label     --> se guarda en Imagen

'
'Bovine
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NumCard SMALLINT(3) UNIQUE NOT NULL
'EarTag Not VarChar(30)
'BirthLand SMALLINT
'BirthDate Date
'BreedingLand SMALLINT
'SlaughterLand SMALLINT
'SlaughterReg VarChar(30)
'SlaughterDate Date
'ButcherLand SMALLINT
'ButcherReg VarChar(30)
'Lot VarChar(20)
'Index EarTag'
'
'Customers
'---------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(255) NOT NULL
'Address VarChar(255)
'City VarChar(255)
'WEB VarChar(255)
'Phone VarChar(255)
'TexT VarChar(255)
'Unique Index(Code)
'Index (Name(19))'
'
'Trademark
'---------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'CO_Name Not VarChar(80)
'Description Not VarChar(80)
'Business Not VarChar(80)
'Slogan Not VarChar(80)
'Headquarter Not VarChar(80)
'CO_WEB Not VarChar(80)
'Shop_Name Not VarChar(80)
'Address Not VarChar(80)
'City Not VarChar(80)
'Phone Not VarChar(80)
'TIN Not VarChar(80)
'Shop_WEB Not VarChar(80)
'Text1 Not VarChar(80)
'Text2 Not VarChar(80)
'Text3 Not VarChar(80)
'Text4 Not VarChar(80)
'Text5 Not VarChar(80)
'Text6 Not VarChar(80)'
'
'status
'------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Version SMALLINT DEFAULT 0'
'
'BarcodeEAN13
'------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'NumberType CHAR(1)
'Number SMALLINT DEFAULT 0
'NetworkSales Not CHAR(12)
'NetworkLineSLB Not CHAR(12)
'NetworkTotalSLB Not CHAR(12)
'NetworkSS Not CHAR(12)
'NetworkPacket1 Not CHAR(12)
'NetworkPacket2 Not CHAR(12)
'LocalSales Not CHAR(12)
'LocalLineSLB Not CHAR(12)
'LocalTotalSLB Not CHAR(12)
'LocalSS Not CHAR(12)
'LocalPacket1 Not CHAR(12)
'LocalPacket2 Not CHAR(12)
'LocalTotal Not CHAR(12)
'Unique Index(NumberType, Number)'
'
'Users
'-----
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'User VARCHAR(16) UNIQUE NOT NULL DEFAULT ''
'password Not VarChar(31)
'Name VARCHAR(127) NOT NULL
'LevelLab BOOLEAN DEFAULT FALSE
'LevelAdmin BOOLEAN DEFAULT FALSE
'LevelConfig BOOLEAN DEFAULT FALSE
'LevelProg BOOLEAN DEFAULT FALSE
'LevelVendor BOOLEAN DEFAULT FALSE'
'
'-- Local tables
'LocalHTickets
'-------------
'LIKE HTickets'
'
'LocalLTickets
'-------------
'LIKE LTickets
'
'LocalPayTickets
'---------------
'LIKE PayTickets
'
'LocalPacketLabels
'-----------------
'LIKE PacketLabels
'
'LocalSServiceLabels
'-------------------
'LIKE SServiceLabels
'
'LocalNumTickets
'---------------
'LIKE NumTickets
'
'LocalEraseTickets
'-----------------
'LIKE EraseTickets
''
'
'LocalStatus
'-----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Version SMALLINT DEFAULT 0
'Language Not VarChar(31)
'ItemsUpdated BOOLEAN DEFAULT FALSE
'EraseTickets BOOLEAN DEFAULT FALSE
'RemoveScale BOOLEAN DEFAULT FALSE
'
'LocalCountries
'--------------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(127) NOT NULL
'Unique Index(Code)
'Index (Name(19))

'Versión 4 de Touch-Scale no tiene Tabla Valores IVA's
'Public Sub Absorver_Iva()
'
'    Unload FrmIVA
'
'    CancelHumano = False
'    canCelar = 0
'    frmMonitor.cmdmonitorstop.Enabled = True
'
'    MostrarMonitor
'
'    frmControl.Switch_mode 0
'    If EstadoBal = 0 Then 'Protección mínima...
'        frmMonitor.MostrarInfo "Lectura Ivas..."
'        If canCelar = 0 And CancelHumano = False Then Call LeeIVAS
'    End If
'    'cerrar conexión
'    On Error Resume Next
'    MyConnObj.Close
'    Set MyConnObj = Nothing
'    MyRecSet.Close
'    Set MyRecSet = Nothing
'    On Error GoTo 0
'    EstadoBal = 1000
'    frmMonitor.cmdmonitorstop.Enabled = False
'    HabilitarSalir
'    frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
'    If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"'
'
'End Sub

Public Sub Absorver_Estructura()
                
    'Borrar datos previos existentes...
    'Revisar actual Hydra para ver netID
    CancelHumano = False
    canCelar = 0
    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = True
    frmMonitor.Label1.Caption = ""
    
    MostrarMonitor
    
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then 'Protección mínima...
        frmMonitor.MostrarInfo CargaCadena(329) & " ..."  '"Secciones..."
        Call LeeSecciones
        frmMonitor.MostrarInfo CargaCadena(303) & " ..."  '"Equipos..."
        If canCelar = 0 And CancelHumano = False Then Call LeeEquipos
        frmMonitor.MostrarInfo CargaCadena(67) & " ..." '"Familias..."
        If canCelar = 0 And CancelHumano = False Then Call LeeFamilias
        frmMonitor.MostrarInfo CargaCadena(13) & " ..." '"Vendedores..."
        If canCelar = 0 And CancelHumano = False Then Call LeeVendedores
        frmMonitor.MostrarInfo CargaCadena(214) & " ..." '"Ivas..."
        If canCelar = 0 And CancelHumano = False Then Call LeeIVAS
        frmMonitor.MostrarInfo CargaCadena(1084) & " ..." '"Articulos..."
        If canCelar = 0 And CancelHumano = False Then Call LeeArticulos
        'TouchScale ... prueba captura paises...
        frmMonitor.MostrarInfo CargaCadena(619) & " ..." '"Fichas de vacuno..."
        If canCelar = 0 And CancelHumano = False Then Call capturaPaises
        If canCelar = 0 And CancelHumano = False Then Call capturaFichas
        frmMonitor.MostrarInfo CargaCadena(23) & " ..." '"Códigos de barras..."
        If canCelar = 0 And CancelHumano = False Then Call capturaCB
        frmMonitor.MostrarInfo CargaCadena(1056) & " ..." '"Clientes..."
        If canCelar = 0 And CancelHumano = False Then Call capturaClientes
        ''''''''''''''''''''''''''''''''''''''''
        
        Call Corregir_Nulos
        'frmMonitor.MostrarInfo "Lectura Datos Trazabilidad..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTraza '(*)
        'frmMonitor.MostrarInfo "Lectura de Tipos de Etiquetas..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTipoEtq '(*)
        'frmMonitor.MostrarInfo "Lectura de Totales..."
        'If canCelar = 0 And CancelHumano = False Then Call LeeTickets
        
    End If
    'cerrar conexión
    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0
    EstadoBal = 1000
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.Command1.Enabled = True
    HabilitarSalir
    frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
    If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
End Sub

Public Function CodiIdent_FS(ByVal FS As Long) As Long
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim nFs As Long
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    Set myRec = Mybase.OpenRecordset("select * from seccion where secc_maqui=" & CStr(FS))
    If Not myRec.EOF Then
        nFs = myRec.Fields("codi_ident")
    End If
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    If Err.Number <> 0 Then
        CodiIdent_FS = 0
    Else
        CodiIdent_FS = nFs
    End If
    On Error GoTo 0
End Function

'Items
'-----
'DROP TABLE IF EXISTS `pcscale`.`items`;
'CREATE TABLE  `pcscale`.`items` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(127) NOT NULL,
'  `ShortName` varchar(64) NOT NULL,
'  `Text` varchar(1024) NOT NULL,
'  `Text1` varchar(110) NOT NULL,
'  `Text2` varchar(110) NOT NULL,
'  `Text3` varchar(110) NOT NULL,
'  `Text4` varchar(110) NOT NULL,
'  `Text5` varchar(110) NOT NULL,
'  `Family` int(11) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT '0.00',
'  `Discount` decimal(4,2) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `PLU` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `VAT` decimal(1,0) unsigned DEFAULT '0',
'  `EAN13` varchar(13) DEFAULT NULL,
'  `Expiring` smallint(6) DEFAULT '0',
'  `BestBefore` smallint(6) DEFAULT '0',
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `Label` varchar(64) NOT NULL,
'  `Icon` varchar(127) NOT NULL,
'  `QLevel` 8varchar(16) NOT NULL,
'  `Mixed` smallint(6) DEFAULT '0',
'  `MixedWeight` decimal(5,3) DEFAULT '0.000',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
'Repartir los descriptivos... Para posible compatibilidad...
' ShortName --> Descriptivo 1 == Name (Hasta 64 de longitud)
' name      -->
' Text1     -->
' Text2     -->
' Text3     -->
' Text4     -->
' Text5     --> Descriptivos 1 al 20
' Text      --> Se guarda en 2040 troceado de 80 en 80
Public Sub LeeArticulos()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset 'tabla articulo
    'Dim MyRecIva As DAO.Recordset
    Dim Myrec2040 As DAO.Recordset
    Dim MyRecTmp As DAO.Recordset
    Dim MyRecAux As New ADODB.Recordset
    
    Dim nFam As Long
    Dim nSec As Long
    Dim nPlu As Long
    Dim nPrc As Double
    'Dim n As Long
    Dim sDesc As String
    Dim nLen As Long
    Dim ncont As Double
    Dim sParte As String
    Dim a As Variant
    Dim nContEvent As Integer
    
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.items where code<>0 order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                
                nFam = MyRecSet.Fields("family")
                If nFam = 0 Then nFam = 1
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(nFam))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("codi_fam") = 1
                    myRec.Fields("codi_ident") = 0
                    myRec.Fields("secc_maqui") = 0
                    myRec.Fields("posicion") = 0
                    myRec.Fields("txt_fam") = "DPT. " & CStr(nFam)
                    myRec.Fields("imagen") = ""
                    myRec.Update
                End If
                myRec.Close
                Set myRec = Nothing
                
                Set myRec = Mybase.OpenRecordset("select * from articulo where codigo=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                
                nSec = MyRecSet.Fields("counter")
                nPlu = MyRecSet.Fields("plu")
                myRec.Fields("codi_ident") = CodiIdent_FS(nSec)
                myRec.Fields("secc_maqui") = nSec
                myRec.Fields("codi_sub") = 0
                myRec.Fields("codi_fam") = nFam
                myRec.Fields("plu") = nPlu
                myRec.Fields("codigo") = MyRecSet.Fields("code")
                If MyRecSet.Fields("saleform") = 1 Then
                    myRec.Fields("codi_pes") = "W"
                Else
                    myRec.Fields("codi_pes") = "U"
                End If
                
                
                myRec.Fields("tipo_iva") = MyRecSet.Fields("vat")
                
                nPrc = MyRecSet.Fields("price")
                myRec.Fields("euros") = nPrc
                
                myRec.Fields("prc3") = Format(Now, "yymmddhhmm")
                myRec.Fields("tara") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("tara_envasado") = 1000 * MyRecSet.Fields("tare")
                myRec.Fields("caducidad") = MyRecSet.Fields("Expiring")
                myRec.Fields("pref") = MyRecSet.Fields("BestBefore")
                'Atención! Touch Scale no dispone de:
                '-Peso Aproximado
                '-Merma
                '-Undades de Preselección
                'MyRec.Fields("poid") = 0 'MyRecSet.Fields("pesoaproximado")
                myRec.Fields("fcb") = 0 'MyRecSet.Fields("merma")
                myRec.Fields("ning") = 0 'MyRecSet.Fields("unidadespreseleccion")
                'Atención! La etiqueta es por nombre... Se ha de crear un fichero
                ' en el PC con los nombres de etiquetas que existen en la Balanza
                ' para que se pueda asociar mediante un índice... O crear un campo
                ' en Articulo para contener el nombre...
                ' DE MOMENTO SE RECOGE NOMBRE DE ETIQUETA EN TRAN_EP (max 60 caracteres)
                myRec.Fields("etq") = 0 'ETQ se hace 0 hasta solucionar
                If IsNull(MyRecSet.Fields("label")) Then
                    myRec.Fields("tran_ep") = ""
                    myRec.Fields("etq") = 0
                Else
                    myRec.Fields("tran_ep") = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("label"))))
                    Select Case ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("label")))
                    Case "1.xml"
                        myRec.Fields("etq") = 1
                    Case "2.xml"
                        myRec.Fields("etq") = 2
                    Case "3.xml"
                        myRec.Fields("etq") = 3
                    Case "4.xml"
                        myRec.Fields("etq") = 4
                    Case "5.xml"
                        myRec.Fields("etq") = 5
                    Case "6.xml"
                        myRec.Fields("etq") = 6
                    Case "7.xml"
                        myRec.Fields("etq") = 7
                    Case "8.xml"
                        myRec.Fields("etq") = 8
                    Case "9.xml"
                        myRec.Fields("etq") = 9
                    End Select
                    
                End If
                '1.0.2 ShortName no se captura...
                myRec.Fields("des_plu1") = Mid(Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))), 1, 25) 'era shortname
                '''''''''''''''''''''''''''''''''
                'Atención! no se dispone de Imagen ni Color
                myRec.Fields("art_cb") = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("ean13"))))
                myRec.Fields("grupo_cons") = 0 'Grupo de conservación
                myRec.Fields("imagen") = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("icon")))) '& " _+_ " & _
                                         'Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("qlevel"))))
                myRec.Fields("ning") = MyRecSet.Fields("discount")
                
                If MyRecSet.Fields("mixed") <> 0 Then
                    myRec.Fields("lmix") = True
                Else
                    myRec.Fields("lmix") = False
                End If
                myRec.Fields("poid") = 1000 * (MyRecSet.Fields("mixedweight"))
                
'''''''Ingredientes a tabla 2040
'                If Not IsNull(Trim(MyRecSet.Fields("text"))) Then
'                    Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("code")))
'                    If myRec.EOF Then
'                        Myrec2040.AddNew
'                    Else
'                        Myrec2040.Edit
'                    End If
'                    Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
'                    Myrec2040.Fields("plu") = nPlu
'                    Myrec2040.Fields("mostrador") = nSec
'                    Myrec2040.Fields("tran_txt2") = "-"
'                    sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text"))))
'                    nLen = Len(sDesc)
'                    ncont = nLen / 80
'
'                    If ncont <> Int(nLen / 80) Then ncont = Int(ncont) + 1
'
'                    For nLen = 1 To ncont
'
'                        Select Case nLen
'                        Case 1
'                            Myrec2040.Fields("txt_21") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 2
'                            Myrec2040.Fields("txt_22") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 3
'                            Myrec2040.Fields("txt_23") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 4
'                            Myrec2040.Fields("txt_24") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 5
'                            Myrec2040.Fields("txt_25") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 6
'                            Myrec2040.Fields("txt_26") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 7
'                            Myrec2040.Fields("txt_27") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 8
'                            Myrec2040.Fields("txt_28") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 9
'                            Myrec2040.Fields("txt_29") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 10
'                            Myrec2040.Fields("txt_30") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 11
'                            Myrec2040.Fields("txt_31") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 12
'                            Myrec2040.Fields("txt_32") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 13
'                            Myrec2040.Fields("txt_33") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 14
'                            Myrec2040.Fields("txt_34") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 15
'                            Myrec2040.Fields("txt_35") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 16
'                            Myrec2040.Fields("txt_36") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 17
'                            Myrec2040.Fields("txt_37") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 18
'                            Myrec2040.Fields("txt_38") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 19
'                            Myrec2040.Fields("txt_39") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        Case 20
'                            Myrec2040.Fields("txt_40") = Mid(sDesc, ((nLen - 1) * 80) + 1, 80)
'                        End Select
'                    Next nLen
'                    Myrec2040.Update
'                    Myrec2040.Close
'                    Set Myrec2040 = Nothing
'                End If

'                'sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text1")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text2")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text3")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text4")))) & _
'                '        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text5"))))
                
                sDesc = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text1")))) & _
                        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text2")))) & _
                        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text3")))) & _
                        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text4")))) & _
                        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text5")))) & _
                        Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("text"))))
                
If MyRecSet.Fields("code") = 98790 Then
    nLen = nLen
End If

                nLen = Len(sDesc)
                ncont = nLen / 25
                
                If ncont <> Int(nLen / 25) Then ncont = Int(ncont) + 1
                    
                For nLen = 1 To ncont
                    Select Case nLen
                    'Select Case MyRecAux.Fields("numlinea")
                        Case 1
                            myRec.Fields("des_plu2") = Trim(Mid(sDesc, 1, 25))
                        Case 2
                            myRec.Fields("des_plu3") = Trim(Mid(sDesc, 26, 25))
                        Case 3
                            myRec.Fields("des_plu4") = Trim(Mid(sDesc, 51, 25))
                        Case 4
                            myRec.Fields("des_plu5") = Trim(Mid(sDesc, 76, 25))
                        Case 5
                            myRec.Fields("des_plu6") = Trim(Mid(sDesc, 101, 25))
                        Case 6
                            myRec.Fields("des_plu7") = Trim(Mid(sDesc, 126, 25))
                        Case 7
                            myRec.Fields("des_plu8") = Trim(Mid(sDesc, 151, 25))
                        Case 8
                            myRec.Fields("des_plu9") = Trim(Mid(sDesc, 176, 25))
                        Case 9
                            myRec.Fields("des_plu0") = Trim(Mid(sDesc, 201, 25))
                        Case 10
                            myRec.Fields("des_plux") = Trim(Mid(sDesc, 226, 25))
                        Case 11
                            myRec.Fields("des_plu11") = Trim(Mid(sDesc, 251, 25))
                        Case 12
                            myRec.Fields("des_plu12") = Trim(Mid(sDesc, 276, 25))
                        Case 13
                            myRec.Fields("des_plu13") = Trim(Mid(sDesc, 301, 25))
                        Case 14
                            myRec.Fields("des_plu14") = Trim(Mid(sDesc, 326, 25))
                        Case 15
                            myRec.Fields("des_plu15") = Trim(Mid(sDesc, 351, 25))
                        Case 16
                            myRec.Fields("des_plu16") = Trim(Mid(sDesc, 376, 25))
                        Case 17
                            myRec.Fields("des_plu17") = Trim(Mid(sDesc, 401, 25))
                        Case 18
                            myRec.Fields("des_plu18") = Trim(Mid(sDesc, 426, 25))
                        Case 19
                            myRec.Fields("des_plu19") = Trim(Mid(sDesc, 451, 25))
                        Case 20
                            myRec.Fields("des_plu20") = Trim(Mid(sDesc, 476, 25))
                        Case Else
                            If nLen > 20 And nLen < 41 Then
                                Set Myrec2040 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecSet.Fields("code")))
                                If myRec.EOF Then
                                    Myrec2040.AddNew
                                Else
                                    Myrec2040.Edit
                                End If
                                Myrec2040.Fields("codigo") = CStr(MyRecSet.Fields("code"))
                                Myrec2040.Fields("plu") = nPlu
                                Myrec2040.Fields("mostrador") = nSec
                                Myrec2040.Fields("tran_txt2") = "-"
                                Myrec2040.Fields("txt_" & CStr(nLen)) = Trim(Mid(sDesc, 476 + ((nLen - 20) * 25), 25))
                                Myrec2040.Update
                                Myrec2040.Close
                                Set Myrec2040 = Nothing
                            End If
                    End Select
            
                Next nLen
                    

                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                'Call frmControl.Balanza_ComOK
                nContEvent = nContEvent + 1
                If nContEvent >= 10 Then
                    nContEvent = 0
                    Do_Events
                End If
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

'Families
'--------
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(54) NOT NULL,
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub LeeFamilias()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.families order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast

            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                'atención! V4 Touch Scale las Familias son Genéricas nos se asignan a sección.
                myRec.Fields("codi_fam") = MyRecSet.Fields("code")
                myRec.Fields("codi_ident") = 0 'MyRecSet.Fields("idseccion")
                myRec.Fields("secc_maqui") = 0 'MyRecSet.Fields("idseccion")
                myRec.Fields("posicion") = MyRecSet.Fields("isbovine") 'MyRecSet.Fields("posicion")
                myRec.Fields("txt_fam") = Mid(Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))), 1, 35)
                myRec.Fields("imagen") = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("label"))))
                'atención! V4 Touch Scale en pcscale.families se almacena si es familia Vacuno y etiqueta
                '          a asignar a los artículos de la Familia. Se deberá recorrer la tabla de artículos
                '          para que los que coincidan con la Familia asignarles la Etiqueta. Por otra parte
                '          en Hydra se indica una sola familia de Vacuno en Configuración y se asigna número
                '          de etiqueta especial, habrá que asignar nombre en lugar de número.
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
        
        Else 'se crea la familia 1
        
            Set myRec = Mybase.OpenRecordset("select * from fam_code where codi_fam=1")
            If myRec.EOF Then
                myRec.AddNew
                myRec.Fields("codi_fam") = 1
                myRec.Fields("codi_ident") = 0
                myRec.Fields("secc_maqui") = 0
                myRec.Fields("posicion") = 0
                myRec.Fields("txt_fam") = "DPT. 1"
                myRec.Fields("imagen") = ""
                myRec.Update
            End If
    
        End If
    End If
    ''''''''''''''''''''''''''''''''
    
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    On Error GoTo 0
    
End Sub

'Vendors
'-------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Code INTEGER UNIQUE
'Name VARCHAR(64) NOT NULL'
'
'VendorKeys
'----------
'Id INTEGER PRIMARY KEY AUTO_INCREMENT
'Counter INTEGER
'VendorKey INTEGER
'Vendor INTEGER UNIQUE
'Unique Index(Counter, VendorKey)'
'atención! V4 Touch Scale un vendedor no puede estar en varias Secciones
Public Sub LeeVendedores()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "Select * from pcscale.vendorkeys order by vendor", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast
            frmMonitor.barra1.Value = 0

            frmMonitor.barra1.Max = MyRecSet.Recordcount
            
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("vendor")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                
                myRec.Fields("ident_vend") = MyRecSet.Fields("vendor")
                myRec.Fields("codi_ident") = MyRecSet.Fields("counter")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("counter")
                myRec.Fields("tec_vend") = MyRecSet.Fields("vendorkey")
                myRec.Fields("tran_vend") = "-"
                
                myRec.Update
                
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
            MyRecSet.Close
            Set MyRecSet = Nothing
            myRec.Close
            Set myRec = Nothing
            
            MyRecSet.Open "Select * from pcscale.vendors order by code", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                On Error GoTo 0
                EstadoBal = 1000
                canCelar = 1
                Call frmControl.Balanza_ComError(0)
            Else
'*******
                If Not MyRecSet.EOF Then
                    MyRecSet.MoveLast
                    frmMonitor.barra1.Value = 0
                    frmMonitor.barra1.Max = MyRecSet.Recordcount
                    
                    MyRecSet.MoveFirst
                    'n = 1
                    Do While Not MyRecSet.EOF
                        Set myRec = Mybase.OpenRecordset("select * from vendedor where ident_vend=" & CStr(MyRecSet.Fields("code")))
                        
                        If Not myRec.EOF Then
                            myRec.MoveFirst
                            Do While Not myRec.EOF
                                myRec.Edit
                    
                                myRec.Fields("codi_vend") = MyRecSet.Fields("code")
                                myRec.Fields("nombre") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 30)
                                myRec.Fields("prog_vend") = "*"
                                 
                                myRec.Update
                                myRec.Movenext
                            Loop
                             
                        End If
                        
                        MyRecSet.Movenext
                        
                        Call MonitorBarra
                        Call frmControl.Balanza_ComOK
                        
                        Do_Events
                        
                        If canCelar = 1 Or CancelHumano = True Then GoTo salir
                    
                    Loop
                End If
'*******
            End If
            
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    On Error GoTo 0
    
    
End Sub

'Counters
'--------
'DROP TABLE IF EXISTS `pcscale`.`counters`;
'CREATE TABLE  `pcscale`.`counters` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Description` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
Public Sub LeeSecciones()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim Cm As String
    
    
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select * from pcscale.counters order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
   
   
        If Not MyRecSet.EOF Then
            MyRecSet.MoveLast


            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecSet.Recordcount
        
            MyRecSet.MoveFirst
            Do While Not MyRecSet.EOF
                If MyMinSec = 0 Then
                    MyMinSec = MyRecSet.Fields("code")
                End If
                Set myRec = Mybase.OpenRecordset("select * from seccion where codi_ident=" & CStr(MyRecSet.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                myRec.Fields("com") = 0
                myRec.Fields("sec_ip") = ""
                myRec.Fields("codi_ident") = MyRecSet.Fields("code")
                myRec.Fields("secc_maqui") = MyRecSet.Fields("code")
                myRec.Fields("nombre") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 25)
                myRec.Fields("descripcio") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("description"))), 1, 25)
                'atención! V4 touch scale se dejará com GA y se verá si cambiar...
                myRec.Fields("enviardatos") = "GA"
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
   
   
        End If
    End If
    ''''''''''''''''''''''''''''''''
salir:
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    On Error GoTo 0
End Sub

'Scales
'------
'DROP TABLE IF EXISTS `pcscale`.`scales`;
'CREATE TABLE  `pcscale`.`scales` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `MAC` varchar(17) NOT NULL,
'  `IP` varchar(16) NOT NULL,
'  `Number` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Type` char(1) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `MAC` (`MAC`)
') ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
'Tabla scales:
'.- Type: ‘S’, ‘N’, ‘A’
'    Balanza servidora (‘S’), esclava (‘N’), balanza no “conectada” (‘A’)
Public Sub LeeEquipos()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim MyRecSec As DAO.Recordset
    Dim sData As String
    Dim Cm As String
    
    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select * from pcscale.scales order by number", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
   
   
        If Not MyRecSet.EOF Then
            frmMonitor.barra1.Value = 0
            MyRecSet.MoveLast

            
            frmMonitor.barra1.Max = MyRecSet.Recordcount
            MyRecSet.MoveFirst
            'n = 1
            Do While Not MyRecSet.EOF
                Set myRec = Mybase.OpenRecordset("select * from equipos where numero_eqp=" & CStr(MyRecSet.Fields("number")))
                If myRec.EOF Then
                    myRec.AddNew
                Else
                    myRec.Edit
                End If
                myRec.Fields("numero_eqp") = MyRecSet.Fields("number")
                sData = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("ip"))))
                If GeneraTextoIP(sData) Then
                    myRec.Fields("prog_eqp") = sData
                End If
                'If MyRecSet.Fields("tipoconexion") = 1 Then
                '    MyRec.Fields("modelo") = 0
                'Else
                    myRec.Fields("modelo") = MyRecSet.Fields("type")
                'End If
                myRec.Fields("descripcio") = Mid(Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))), 1, 21)
                myRec.Fields("codi_ident") = MyMinSec
                myRec.Fields("secc_maqui") = MyMinSec
                
                myRec.Update
                
                'n = n + 1
                MyRecSet.Movenext
                
                Call MonitorBarra
                Call frmControl.Balanza_ComOK
                                
                Do_Events
                
                If canCelar = 1 Or CancelHumano = True Then GoTo salir
                
            Loop
            
   
        End If
        
        
    End If
salir:
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    On Error GoTo 0
End Sub

'DROP TABLE IF EXISTS `pcscale`.`vats`;
'CREATE TABLE  `pcscale`.`vats` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` decimal(1,0) unsigned NOT NULL,
'  `Name` varchar(16) NOT NULL,
'  `Percent` decimal(4,2) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
Public Sub LeeIVAS()
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset 'tabla articulo
    Dim MyRecAux As New ADODB.Recordset
    'Dim nFam As Long
    'Dim nSec As Long
    'Dim nPlu As Long
    'Dim n As Long

    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    'n = 0
    MyRecAux.CursorLocation = adUseClient
    MyRecAux.Open "Select * from pcscale.vats order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        EstadoBal = 1000
        canCelar = 1
        Call frmControl.Balanza_ComError(0)
    Else
        If Not MyRecAux.EOF Then
            MyRecAux.MoveLast
            frmMonitor.barra1.Value = 0
            frmMonitor.barra1.Max = MyRecAux.Recordcount
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
             'MyRecSet.Open "Select * from dbo.tipoimpuesto where idtipoimpuesto=" & CStr(MyRecAux.Fields("idtipoimpuesto1")), MyConnObj, adOpenKeyset, adLockOptimistic
             'If Not MyRecSet.EOF Then
                Set myRec = Mybase.OpenRecordset("select * from iva where tipo=" & CStr(MyRecAux.Fields("code")))
                If myRec.EOF Then
                    myRec.AddNew
                    myRec.Fields("tipo") = MyRecAux.Fields("code")
                   'n = n + 1
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("code")
                   myRec.Fields("descripcion") = ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("name")))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                   'If n > 4 Then
                   '   GoTo Salir
                   'End If
                Else
                '    n = n + 1
                   myRec.Edit
                   myRec.Fields("valor") = MyRecAux.Fields("percent")
                   myRec.Fields("tran_iva") = "-"
                   myRec.Fields("interno") = MyRecAux.Fields("code")
                   myRec.Fields("descripcion") = ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("name")))
                   myRec.Update
                   myRec.Close
                   Set myRec = Nothing
                
                End If
             'End If
             MyRecAux.Movenext
             'MyRecSet.Close
             'Set MyRecSet = Nothing
             Call MonitorBarra
             Call frmControl.Balanza_ComOK

             Do_Events

             If canCelar = 1 Or CancelHumano = True Then GoTo salir

            Loop

        End If

    End If
    ''''''''''''''''''''''''''''''''

salir:
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecSet.Close
    Set MyRecSet = Nothing
    myRec.Close
    Set myRec = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
    End If
    Mybase.Close
    Set Mybase = Nothing
FIN:
    On Error GoTo 0

End Sub

'DROP TABLE IF EXISTS `pcscale`.`localcountries`;
'CREATE TABLE  `pcscale`.`localcountries` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Alpha2` char(2) NOT NULL,
'  `Name` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Alpha2` (`Alpha2`),
'  UNIQUE KEY `Code` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=245 DEFAULT CHARSET=utf8;
Public Sub capturaPaises()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
    Set Mybase = DAO.OpenDatabase(Base_General)

    MyRecSet.Open "Select * from pcscale.localcountries order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Mybase.Execute ("delete from paises")
        Set MyRecVar = Mybase.OpenRecordset("select * from paises")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            MyRecVar.AddNew
            MyRecVar.Fields("codigo") = MyRecSet.Fields("code")
            MyRecVar.Fields("nombre") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 100) 'Format(MyRecSet.Fields("code"), "000") & "-" & Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("alpha2"))), 1, 2) & "-" & "   -" & Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 89)
            MyRecVar.Fields("valor") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name"))), 1, 100)
            MyRecVar.Fields("tran_pais") = " "
            MyRecVar.Fields("borrado") = False
            MyRecVar.Fields("modificado") = True
            MyRecVar.Fields("abrv_3") = "   "
            On Error Resume Next
            MyRecVar.Fields("abrv_2") = Mid(ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("alpha2"))), 1, 2)
            On Error GoTo 0
            MyRecVar.Fields("tipo") = 1
            MyRecVar.Update
            MyRecSet.Movenext
        Loop
        MyRecVar.Close
        Set MyRecVar = Nothing
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'DROP TABLE IF EXISTS `pcscale`.`bovine`;
'CREATE TABLE  `pcscale`.`bovine` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumCard` smallint(3) NOT NULL,
'  `EarTag` varchar(30) NOT NULL, --> Crotal
'  `BirthLand` smallint(6) DEFAULT NULL, --> Pais Nacimiento
'  `BirthDate` date DEFAULT NULL, --> Fecha nacimiento
'  `BreedingLand` smallint(6) DEFAULT NULL, --> Pais Crianza
'  `SlaughterLand` smallint(6) DEFAULT NULL,--> Pais Sacrificio
'  `SlaughterReg` varchar(30) DEFAULT NULL, --> Registro Sacrificio
'  `SlaughterDate` date DEFAULT NULL, --> Fecha Sacrificio
'  `ButcherLand` smallint(6) DEFAULT NULL, --> Pais Despiece
'  `ButcherReg` varchar(30) DEFAULT NULL, --> Registro despiece
'  `Lot` varchar(20) DEFAULT NULL, --> Lote
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumCard` (`NumCard`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub capturaFichas()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.Open "Select * from pcscale.bovine order by numcard", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Set MyRecVar = Mybase.OpenRecordset("select * from fichavacuno where codigo=" & Chr(34) & CStr(MyRecSet.Fields("numcard")) & Chr(34))
            If MyRecVar.EOF Then
                MyRecVar.AddNew
                'MyRecVar.Fields("codigo") = MyRecSet.Fields("numcard")
                'MyRecVar.Fields("identificador") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("eartag")))
                'MyRecVar.Fields("codnacimiento") = MyRecSet.Fields("birthland")
                'MyRecVar.Fields("fecha_nac") = Format(MyRecSet.Fields("birthdate"), "dd/mm/yyyy")
                'MyRecVar.Fields("codcrianza") = MyRecSet.Fields("breedingland")
                'MyRecVar.Fields("codsacrificio") = MyRecSet.Fields("slaughterland")
                'MyRecVar.Fields("rsisacrificio") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("SlaughterReg")))
                'MyRecVar.Fields("fecha_sac") = Format(MyRecSet.Fields("SlaughterDate"), "dd/mm/yyyy")
                'MyRecVar.Fields("coddespiece") = MyRecSet.Fields("ButcherLand")
                'MyRecVar.Fields("rsidespiece") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("ButcherReg")))
                'MyRecVar.Fields("lote") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("lot")))
            Else
                MyRecVar.Edit
            End If
                MyRecVar.Fields("codigo") = MyRecSet.Fields("numcard")
                MyRecVar.Fields("identificador") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("eartag")))
                MyRecVar.Fields("codnacimiento") = MyRecSet.Fields("birthland")
                MyRecVar.Fields("fecha_nac") = Format(MyRecSet.Fields("birthdate"), "dd/mm/yyyy")
                MyRecVar.Fields("codcrianza") = MyRecSet.Fields("breedingland")
                MyRecVar.Fields("codsacrificio") = MyRecSet.Fields("slaughterland")
                MyRecVar.Fields("rsisacrificio") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("SlaughterReg")))
                MyRecVar.Fields("fecha_sac") = Format(MyRecSet.Fields("SlaughterDate"), "dd/mm/yyyy")
                MyRecVar.Fields("coddespiece") = MyRecSet.Fields("ButcherLand")
                MyRecVar.Fields("rsidespiece") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("ButcherReg")))
                If IsNull(MyRecSet.Fields("lot")) Then
                    MyRecVar.Fields("lote") = ""
                Else
                    MyRecVar.Fields("lote") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("lot")))
                End If
            'End If
            
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

Public Function consultaSeccMaquiBalanza(ByVal nBal) As Integer
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
    Set Mybase = DAO.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select secc_maqui from equipos where numero_eqp=" & CStr(nBal))
    If MyRecVar.EOF Then
        consultaSeccMaquiBalanza = 1
    Else
        consultaSeccMaquiBalanza = MyRecVar.Fields("secc_maqui")
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
End Function
'DROP TABLE IF EXISTS `pcscale`.`barcodeean13`;
'CREATE TABLE  `pcscale`.`barcodeean13` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumberType` char(1) DEFAULT NULL,
'  `Number` smallint(6) DEFAULT '0',
'  `NetworkSales` char(12) NOT NULL,
'  `NetworkLineSLB` char(12) NOT NULL,
'  `NetworkTotalSLB` char(12) NOT NULL,
'  `NetworkSS` char(12) NOT NULL,
'  `NetworkPacket1` char(12) NOT NULL,
'  `NetworkPacket2` char(12) NOT NULL,
'  `LocalSales` char(12) NOT NULL,
'  `LocalLineSLB` char(12) NOT NULL,
'  `LocalTotalSLB` char(12) NOT NULL,
'  `LocalSS` char(12) NOT NULL,
'  `LocalPacket1` char(12) NOT NULL,
'  `LocalPacket2` char(12) NOT NULL,
'  `LocalTotal` char(12) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumberType` (`NumberType`,`Number`)
') ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
'Typo "G" se desprecia, se cogen "S" (Equipo) y "C" (Mostrador)
Public Sub capturaCB()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim nCodiIdent As Integer
Dim nSeccMaqui As Integer
Dim nC As Integer
Dim Cm As String

    Set Mybase = DAO.OpenDatabase(Base_General)
    ''''''''''''''''''''''''''
    Cm = "delete from codbar"
    Mybase.Execute Cm
    ''''''''''''''''''''''''''
    MyRecSet.Open "Select * from pcscale.barcodeean13 where numbertype<>'G'", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from codbar")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            If MyRecSet.Fields("numbertype") = "S" Then
                nSeccMaqui = consultaSeccMaquiBalanza(MyRecSet.Fields("number"))
                nCodiIdent = nSeccMaqui
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where numero_eqp=" & CStr(MyRecSet.Fields("number")))
            Else
                Set MyRecVar = Mybase.OpenRecordset("select * from codbar where secc_maqui=" & CStr(MyRecSet.Fields("number")) & " and isnull(numero_eqp)")
                nSeccMaqui = MyRecSet.Fields("number")
                nCodiIdent = nSeccMaqui
            End If
            If MyRecVar.EOF Then
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.AddNew
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("numbertype") = "S" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("networksales")))
                        MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkLineSLB")))
                        MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkTotalSLB")))
                        MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkSS")))
                        MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkPacket1")))
                    Else
                        MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localsales")))
                        MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localLineSLB")))
                        MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localTotalSLB")))
                        MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localSS")))
                        MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localPacket1")))
                    End If
                    MyRecVar.Update
                Next nC
            Else
                For nC = 0 To 1 '0=red, 1=local
                    MyRecVar.Edit
                    MyRecVar.Fields("codi_ident") = nCodiIdent
                    MyRecVar.Fields("secc_maqui") = nSeccMaqui
                    If MyRecSet.Fields("numbertype") = "S" Then
                        MyRecVar.Fields("numero_eqp") = MyRecSet.Fields("number")
                    End If
                    MyRecVar.Fields("secc_eqp") = nC
                    If nC = 0 Then
                        MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("networksales")))
                        MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkLineSLB")))
                        MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkTotalSLB")))
                        MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkSS")))
                        MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("NetworkPacket1")))
                    Else
                        MyRecVar.Fields("CB_VNT") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localsales")))
                        MyRecVar.Fields("CB_SUP") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localLineSLB")))
                        MyRecVar.Fields("CB_MIX") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localTotalSLB")))
                        MyRecVar.Fields("CB_EV1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localSS")))
                        MyRecVar.Fields("CB_EV2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("localPacket1")))
                    End If
                    MyRecVar.Update
                Next nC
    
            End If
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub

'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(255) NOT NULL,
'  `Address` varchar(255) DEFAULT NULL,
'  `City` varchar(255) DEFAULT NULL,
'  `WEB` varchar(255) DEFAULT NULL,
'  `Phone` varchar(255) DEFAULT NULL,
'  `Text` varchar(255) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub capturaClientes()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
    Set Mybase = DAO.OpenDatabase(Base_General)
    MyRecSet.Open "Select * from pcscale.customers order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
    
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Set MyRecVar = Mybase.OpenRecordset("select * from cliente where cod_cli=" & CStr(MyRecSet.Fields("code")))
            If MyRecVar.EOF Then
                MyRecVar.AddNew
                MyRecVar.Fields("cod_cli") = MyRecSet.Fields("code")
                MyRecVar.Fields("nom_cli") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))
                MyRecVar.Fields("data1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("address")))
                MyRecVar.Fields("data2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("city")))
                MyRecVar.Fields("data3") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("web")))
                MyRecVar.Fields("data4") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("phone")))
            Else
                MyRecVar.Edit
                MyRecVar.Fields("cod_cli") = MyRecSet.Fields("code")
                MyRecVar.Fields("nom_cli") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("name")))
                MyRecVar.Fields("data1") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("address")))
                MyRecVar.Fields("data2") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("city")))
                MyRecVar.Fields("data3") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("web")))
                MyRecVar.Fields("data4") = ConvertUtf8BytesToString(DecodeBase64(MyRecSet.Fields("phone")))
            End If
            
            MyRecVar.Update
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecSet.Movenext
        Loop
    
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing
    Mybase.Close

End Sub



'//////////
'Envios Touch Scale
'\\\\\\\\\\
''''''''''''''''''''''
'ENVIO DE ARTÍCULOS...
''''''''''''''''''''''
'Se envían:
'   IVAS
'   Familias Altas/bajas
'   Artículos Altas/bajas
'   Teclas (Siempre) Altas/bajas
'Atención PLU MIXTO
'DROP TABLE IF EXISTS `pcscale`.`items`;
'CREATE TABLE  `pcscale`.`items` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(127) NOT NULL,
'  `ShortName` varchar(64) NOT NULL,
'  `Text` varchar(1024) NOT NULL,
'  `Text1` varchar(110) NOT NULL,
'  `Text2` varchar(110) NOT NULL,
'  `Text3` varchar(110) NOT NULL,
'  `Text4` varchar(110) NOT NULL,
'  `Text5` varchar(110) NOT NULL,
'  `Family` int(11) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT '0.00',
'  `Discount` decimal(4,2) NOT NULL,
'  `SaleForm` smallint(6) DEFAULT '1',
'  `PLU` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `VAT` decimal(1,0) unsigned DEFAULT '0',
'  `EAN13` varchar(13) DEFAULT NULL,
'  `Expiring` smallint(6) DEFAULT '0',
'  `BestBefore` smallint(6) DEFAULT '0',
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `Label` varchar(64) NOT NULL,
'  `Icon` varchar(127) NOT NULL,
'  `QLevel` varchar(16) NOT NULL,
'  `Mixed` smallint(6) DEFAULT '0',
'  `MixedWeight` decimal(5,3) DEFAULT '0.000',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;
Public Function MUE_EnviaArticuloNW(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean, Optional MiCodigo As Long) As Boolean
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecTmp As DAO.Recordset
Dim myRecTmp1 As DAO.Recordset
Dim MyRecFam As DAO.Recordset
Dim MyRecSec As DAO.Recordset
Dim nD As Double
Dim MyRecAux As New ADODB.Recordset
Dim MyRecAux1 As New ADODB.Recordset
Dim MyRecFamSec As New ADODB.Recordset
Dim Cm As String
Dim nF As Integer
Dim nL As Integer
Dim lEsAlta As Boolean
Dim nC(99) As Long
Dim ncont As Integer
Dim Resp As Long
Dim nPrice As Double
Dim nContPrecios As Integer
Dim nIdFinal As Double
Dim sColorTmp As String
Dim PosiRel As Long
Dim lEnvImage As Boolean
Dim nContEvnt As Integer
Dim s(5) As String

CadenadeLog "Entrada en Envio de Articulos... Corregir nulos..."
Call Corregir_Nulos
'''''''''''''''''

    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINARTICULO
    End If
    On Error GoTo 0
    On Error GoTo FINARTICULO
    
'''''''PROGRAMACIÓN IVA's
    Set myRecTmp1 = Mybase.OpenRecordset("select * from iva")
    If Not myRecTmp1.EOF Then
        myRecTmp1.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = myRecTmp1.Recordcount
        myRecTmp1.MoveFirst
        Do While Not myRecTmp1.EOF
            MyRecAux.Open "select * from pcscale.vats where code=" & CStr(myRecTmp1.Fields("tipo")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            End If
            
            MyRecAux.Fields("code") = myRecTmp1.Fields("tipo")
            MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes(myRecTmp1.Fields("descripcion")))
            MyRecAux.Fields("percent") = myRecTmp1.Fields("valor")
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            myRecTmp1.Movenext
        
            Call MonitorBarra
        
            'Do_Events '--> ver
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 20 Then
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
        
        Loop
    
    End If
    myRecTmp1.Close
    Set myRecTmp1 = Nothing
'''''''''''''''''''''''''''''''


    'Programación Familias...
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(54) NOT NULL,
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`),
'  UNIQUE KEY `Code_2` (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
            'baja de familias...
            Set MyRecFam = Mybase.OpenRecordset("select * from fam_code where posicion=525")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
            If Not MyRecFam.EOF Then
                
                MyRecFam.MoveFirst
                
                Do While Not MyRecFam.EOF
                
                    Cm = "delete from pcscale.families where code=" & CStr(MyRecFam.Fields("codi_fam"))
                    MyConnObj.Execute Cm
                    Cm = "delete from pcscale.items where family=" & CStr(MyRecFam.Fields("codi_fam"))
                    MyConnObj.Execute Cm

                    MyRecFam.Delete
                    
                    If Not MyRecFam.EOF Then MyRecFam.Movenext
                Loop
                
            End If
            MyRecFam.Close
            Set MyRecFam = Nothing
            ''''''''''''''''''''
            Set MyRecFam = Mybase.OpenRecordset("select * from fam_code order by codi_fam")  '& " and secc_maqui=" & CStr(MyRecVar.Fields("secc_maqui")) & ")")
            If Not MyRecFam.EOF Then
                
                MyRecFam.MoveFirst
                
                Do While Not MyRecFam.EOF
                    
                    lEsAlta = False
                    
                    On Error Resume Next
                    MyRecAux.Open "select * from pcscale.families where code=" & CStr(MyRecFam.Fields("codi_fam")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Err.Number <> 0 Then
                        Call frmControl.Balanza_ComError(0)
                        GoTo FINARTICULO
                    End If
                    On Error GoTo FINARTICULO
                    If MyRecAux.EOF Then
                        lEsAlta = True
                    End If
                    
                    If lEsAlta Then
                        MyRecAux.AddNew
                    End If
                    MyRecAux.Fields("code") = MyRecFam.Fields("codi_fam")
                    MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes(MyRecFam.Fields("txt_fam")))
                    MyRecAux.Fields("label") = Trim(ConvertStringToUtf8Bytes(MyRecFam.Fields("imagen")))
                    If MyRecFam.Fields("posicion") = 1 Then
                        MyRecAux.Fields("isbovine") = 1
                    Else
                        MyRecAux.Fields("isbovine") = 0
                    End If
                    'MyRecAux.UpdateBatch adAffectAll
                    MyRecAux.Update
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                    MyRecFam.Movenext
                Loop
            
            End If
            'MyRecAux.Fields("descripcion") = MyRecFam.Fields("txt_fam")
            MyRecFam.Close
            Set MyRecFam = Nothing
            lEsAlta = False
            
    
    '''''''''''''''''''''''''
    Set MyRecVar = Mybase.OpenRecordset("select * from articulo where borrado=true order by codigo")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo "Baja Artículo:" & CStr(MyRecVar.Fields("codigo"))
            
            Cm = "delete from pcscale.items where code=" & CStr(MyRecVar.Fields("codigo"))
            MyConnObj.Execute Cm
            Cm = "delete from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo"))
            Mybase.Execute Cm
            
            Do_Events
            
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    
    
    
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo where tran_plu<>'*' order by codigo")
        lEnvImage = False
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from articulo order by codigo")
        lEnvImage = True
    End If
    If Not MyRecVar.EOF Then
        MyRecVar.MoveLast
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
            
        nContEvnt = 0
        frmMonitor.MostrarInfo "Envio Artículos...:" & CStr(frmMonitor.barra1.Max)
        CadenadeLog "Artículos a enviar...:" & CStr(frmMonitor.barra1.Max)
        Do While Not MyRecVar.EOF
            
            lEsAlta = False
            On Error Resume Next
           
            MyRecAux.Open "select * from pcscale.items where code=" & MyRecVar.Fields("codigo"), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            End If
            
            MyRecAux.Fields("code") = MyRecVar.Fields("codigo")
            MyRecAux.Fields("shortname") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu1")))
            MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu1")))
            
            ''Cm = ""
            ''Set myRecTmp1 = MyBase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
            ''If Not myRecTmp1.EOF Then
            ''    For ncont = 21 To 40
            ''        Cm = Cm & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
            ''    Next ncont
            ''End If
            ''myRecTmp1.Close
            ''Set myRecTmp1 = Nothing
            ''If Trim(Cm) <> "" Then
            ''    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(Mid(Cm, 1, 1024))
            ''Else
            ''    If lEsAlta Then MyRecAux.Fields("text") = ""
            ''End If
            ''5.4.1 If lEsAlta Then MyRecAux.Fields("text") = ""
            '''''''
            'Cm = ""
            'For ncont = 3 To 20
            '    Select Case ncont
            '    'Case 1
            '    '    Cm = Cm & Trim(MyRecVar.Fields("des_plu2"))
            '    'Case 2
            '    '    Cm = Cm & Trim(MyRecVar.Fields("des_plu3")) & " "
            '    Case 3
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu4")) & " "
            '    Case 4
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu5")) & " " & Chr(10)
            '    Case 5
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu6")) & " "
            '    Case 6
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu7")) & " " & Chr(10)
            '    Case 7
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu8")) & " "
            '    Case 8
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu9")) & " " & Chr(10)
            '    Case 9
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu0")) & " "
            '    Case 10
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plux")) & " " & Chr(10)
            '    Case 11
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu11")) & " "
            '    Case 12
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu12")) & " " & Chr(10)
            '    Case 13
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu13")) & " "
            '    Case 14
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu14")) & " " & Chr(10)
            '    Case 15
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu15")) & " "
            '    Case 16
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu16")) & " " & Chr(10)
            '    Case 17
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu17")) & " "
            '    Case 18
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu18")) & " " & Chr(10)
            '    Case 19
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu19")) & " "
            '    Case 20
            '        Cm = Cm & Trim(MyRecVar.Fields("des_plu20"))
           '
            '    End Select
            
            'Next ncont
            ''5.4.1
            ''s(1) = Mid(Cm, 1, 100)
            ''s(2) = Mid(Cm, 101, 100)
            ''s(3) = Mid(Cm, 201, 100)
            ''s(4) = Mid(Cm, 301, 100)
            ''s(5) = Mid(Cm, 401)
            'If Trim(MyRecVar.Fields("des_plu2")) <> "" Or Trim(MyRecVar.Fields("des_plu3")) <> "" Then
            '    MyRecAux.Fields("text1") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu2"))) & " " & ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("des_plu3")))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text1") = ""
            'End If
            ''If Trim(s(2)) <> "" Then
            '    MyRecAux.Fields("text2") = "" 'ConvertStringToUtf8Bytes(Trim(s(2)))
            ''Else
            ''    If lEsAlta Then MyRecAux.Fields("text2") = ""
            ''End If
            ''If Trim(s(3)) <> "" Then
            '    MyRecAux.Fields("text3") = "" 'ConvertStringToUtf8Bytes(Trim(s(3)))
            ''Else
            ''    If lEsAlta Then MyRecAux.Fields("text3") = ""
            ''End If
            ''If Trim(s(4)) <> "" Then
            '    MyRecAux.Fields("text4") = "" 'ConvertStringToUtf8Bytes(Trim(s(4)))
            ''Else
            ''    If lEsAlta Then MyRecAux.Fields("text4") = ""
            ''End If
            ''If Trim(s(5)) <> "" Then
            '    MyRecAux.Fields("text5") = "" 'ConvertStringToUtf8Bytes(Trim(s(5)))
            ''Else
            ''    If lEsAlta Then MyRecAux.Fields("text5") = ""
            ''end If
            'If Trim(Cm) <> "" Then
            '    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(Cm)
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text") = ""
            'End If
            
            
            Cm = ""
            For ncont = 1 To 20
                Select Case ncont
                Case 1
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu2"))
                Case 2
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu3"))
                Case 3
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu4"))
                Case 4
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu5"))
                Case 5
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu6"))
                Case 6
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu7"))
                Case 7
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu8"))
                Case 8
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu9"))
                Case 9
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu0"))
                Case 10
                    Cm = Cm & Trim(MyRecVar.Fields("des_plux"))
                Case 11
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu11"))
                Case 12
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu12"))
                Case 13
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu13"))
                Case 14
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu14"))
                Case 15
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu15"))
                Case 16
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu16"))
                Case 17
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu17"))
                Case 18
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu18"))
                Case 19
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu19"))
                Case 20
                    Cm = Cm & Trim(MyRecVar.Fields("des_plu20"))
           
                End Select
            Next ncont
            
            Set myRecTmp1 = Mybase.OpenRecordset("select * from lintxt2040 where codigo=" & CStr(MyRecVar.Fields("codigo")))
            If Not myRecTmp1.EOF Then
                For ncont = 21 To 40
                    Cm = Cm & Trim(myRecTmp1.Fields("txt_" & CStr(ncont)))
                Next ncont
            End If
            myRecTmp1.Close
            Set myRecTmp1 = Nothing
            
            If Trim(Cm) <> "" Then
                Cm = ConvertStringToUtf8Bytes(Cm)
                MyRecAux.Fields("text") = Mid(Cm, 1, 1024)
            Else
                If lEsAlta Then MyRecAux.Fields("text") = ""
            End If
            If lEsAlta Then MyRecAux.Fields("text1") = ""
            If lEsAlta Then MyRecAux.Fields("text2") = ""
            If lEsAlta Then MyRecAux.Fields("text3") = ""
            If lEsAlta Then MyRecAux.Fields("text4") = ""
            If lEsAlta Then MyRecAux.Fields("text5") = ""
            
            MyRecAux.Fields("family") = MyRecVar.Fields("codi_fam")
            
            MyRecAux.Fields("price") = MyRecVar.Fields("euros")
            If MyRecVar.Fields("codi_pes") = "W" Then
                MyRecAux.Fields("saleform") = 1
            Else
                MyRecAux.Fields("saleform") = 0
            End If
            MyRecAux.Fields("plu") = MyRecVar.Fields("plu")
            MyRecAux.Fields("counter") = MyRecVar.Fields("secc_maqui")
            MyRecAux.Fields("vat") = MyRecVar.Fields("tipo_iva")
            'If lEsAlta Then
                MyRecAux.Fields("discount") = MyRecVar.Fields("ning")
                If lEsAlta Then MyRecAux.Fields("qlevel") = ""
            'End If
            ''''''''''''
            If Trim(MyRecVar.Fields("art_cb")) <> "" Then
                MyRecAux.Fields("ean13") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("art_cb")))
            Else
                MyRecAux.Fields("ean13") = ""
            End If
            MyRecAux.Fields("expiring") = MyRecVar.Fields("caducidad")
            MyRecAux.Fields("bestbefore") = MyRecVar.Fields("pref")
            MyRecAux.Fields("tare") = MyRecVar.Fields("tara") / 1000
            If Trim(MyRecVar.Fields("tran_ep")) <> "" Then
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("tran_ep")))
            Else
                
                Select Case MyRecVar.Fields("etq")
                Case 0
                    MyRecAux.Fields("label") = ""
                Case 255
                    MyRecAux.Fields("label") = ""
                Case Else
                    MyRecAux.Fields("label") = ConvertStringToUtf8Bytes(CStr(MyRecVar.Fields("etq")) & ".xml")
                End Select
                
            End If
            '" _+_ "
            s(1) = separaFich(Trim(MyRecVar.Fields("imagen")))
            If Trim(s(1)) <> "" Then
                MyRecAux.Fields("icon") = Trim(ConvertStringToUtf8Bytes(s(1)))
            Else
                If lEsAlta Then MyRecAux.Fields("icon") = ""
            End If
            If MyRecVar.Fields("lmix") Then
                MyRecAux.Fields("mixed") = 1
            Else
                MyRecAux.Fields("mixed") = 0
            End If
            MyRecAux.Fields("mixedweight") = MyRecVar.Fields("poid") / 1000
            
            MyRecAux.Update
            
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_plu") = "*"
            MyRecVar.Update
            
            MyRecVar.Movenext
            
            Call MonitorBarra

            'Call frmControl.Balanza_ComOK
            
            'Do_Events '--> ver
            nContEvnt = nContEvnt + 1
            If nContEvnt >= 20 Then
                nContEvnt = 0
                DoEvents
                Sleep (1)
            End If
            '''''''''''''''''''
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirArt
            
        Loop
        
    End If
    
''''''''''''TECLAS...
'DROP TABLE IF EXISTS `pcscale`.`plukeys`;
'CREATE TABLE  `pcscale`.`plukeys` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Section` int(11) DEFAULT NULL,
'  `Board` int(11) DEFAULT NULL,
'  `PLUKey` int(11) DEFAULT NULL,
'  `Code` int(11) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Section` (`Section`,`Board`,`PLUKey`),
'  KEY `Code` (`Code`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
    frmMonitor.MostrarInfo CargaCadena(234)  '"ENVÍO DE TECLAS RÁPIDAS"
    Dim nEqp As Integer
        
    Set MyRecSec = Mybase.OpenRecordset("select * from seccion order by codi_ident")
    If Not MyRecSec.EOF Then
        MyRecSec.MoveFirst
        Do While Not MyRecSec.EOF
            
            Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and (borrado=true or codigo=0) and codi_ident=" & CStr(MyRecSec.Fields("codi_ident")))
            If Not myRecTmp1.EOF Then
        
                myRecTmp1.MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = myRecTmp1.Recordcount
                myRecTmp1.MoveFirst
                Do While Not myRecTmp1.EOF
                    
                    '4.6.0
                    'If IsNull(myRecTmp1.Fields("numero_eqp")) Then
                    '    nEqp = 0
                    'Else
                        nEqp = Val(myRecTmp1.Fields("tran_tecsc10")) - 1 'guardar en texto GRUPO (1 a 9)
                    'End If
                    '''''''
                    
                    Cm = "delete from pcscale.plukeys where section=" & CStr(nEqp + 1) & " and board=" & CStr(myRecTmp1.Fields("tabla")) & " and plukey=" & CStr(myRecTmp1.Fields("codi_tec"))
                    MyConnObj.Execute Cm
                    
                    'myRecTmp1.Delete
                    If Not myRecTmp1.EOF Then myRecTmp1.Movenext
                
                    Call MonitorBarra
                
                    'Do_Events '--> ver
                    nContEvnt = nContEvnt + 1
                    If nContEvnt >= 20 Then
                        nContEvnt = 0
                        DoEvents
                        Sleep (1)
                    End If
                    '''''''''''''''''''
                
                Loop
            
            End If
            myRecTmp1.Close
            Set myRecTmp1 = Nothing
            If Modificaciones Then
                Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=" & CStr(MyRecSec.Fields("codi_ident")) & " and ((tran_tec<>'*') or isnull(tran_tec))")
            Else
                Set myRecTmp1 = Mybase.OpenRecordset("select * from teclas where prog_tec='TOUCH' and codigo<>0 and codi_ident=" & CStr(MyRecSec.Fields("codi_ident"))) ' where codigo=" & CStr(MyRecVar.Fields("codigo")))
            End If
            If Not myRecTmp1.EOF Then
                myRecTmp1.MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = myRecTmp1.Recordcount
                myRecTmp1.MoveFirst
                Do While Not myRecTmp1.EOF
                    'If IsNull(myRecTmp1.Fields("numero_eqp")) Then
                    '    nEqp = 0
                    'Else
                        nEqp = myRecTmp1.Fields("tran_tecsc10") - 1
                    'End If
                    MyRecAux.Open "select * from pcscale.plukeys where section=" & CStr(nEqp + 1) & " and board=" & CStr(myRecTmp1.Fields("tabla")) & " and plukey=" & CStr(myRecTmp1.Fields("codi_tec")), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Err.Number <> 0 Then
                        'Call frmControl.Balanza_ComError(0)
                        GoTo FINARTICULO
                    End If
                    On Error GoTo FINARTICULO
                    lEsAlta = False
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                        lEsAlta = True
                    End If
                    
                    MyRecAux.Fields("section") = nEqp + 1
                    MyRecAux.Fields("board") = myRecTmp1.Fields("tabla")
                    MyRecAux.Fields("plukey") = myRecTmp1.Fields("codi_tec")
                    MyRecAux.Fields("code") = myRecTmp1.Fields("codigo")
                    
                    MyRecAux.Update
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                    myRecTmp1.Edit
                    'If myRecTmp1.Fields("tran_tec") = "E" Then
                    '    myRecTmp1.Fields("tran_tec") = "A"
                    'Else
                    '    myRecTmp1.Fields("tran_tec") = "T"
                    'End If
                    myRecTmp1.Fields("tran_tec") = "*"
                    myRecTmp1.Update
                    myRecTmp1.Movenext
                
                    Call MonitorBarra
                
                    'Do_Events '--> ver
                    nContEvnt = nContEvnt + 1
                    If nContEvnt >= 20 Then
                        nContEvnt = 0
                        DoEvents
                        Sleep (1)
                    End If
                    '''''''''''''''''''
                
                Loop
            
            End If
            myRecTmp1.Close
            Set myRecTmp1 = Nothing

            MyRecSec.Movenext
        Loop
    End If
'''''''''''''''''''''''''''''''
    

SalirArt:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    
CadenadeLog "Fin envio Artículos."
    
    MUE_EnviaArticuloNW = True
    Exit Function
FINARTICULO:
    CadenadeLog "Error en MUE_EnviaArticuloNW:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecAux.Close
    Set MyRecAux = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
    MUE_EnviaArticuloNW = False
End Function

'Adaptar TouchScale Barna
'DROP TABLE IF EXISTS `pcscale`.`bovine`;
'CREATE TABLE  `pcscale`.`bovine` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `NumCard` smallint(3) NOT NULL,
'  `EarTag` varchar(30) NOT NULL,
'  `BirthLand` smallint(6) DEFAULT NULL,
'  `BirthDate` date DEFAULT NULL,
'  `BreedingLand` smallint(6) DEFAULT NULL,
'  `SlaughterLand` smallint(6) DEFAULT NULL,
'  `SlaughterReg` varchar(30) DEFAULT NULL,
'  `SlaughterDate` date DEFAULT NULL,
'  `ButcherLand` smallint(6) DEFAULT NULL,
'  `ButcherReg` varchar(30) DEFAULT NULL,
'  `Lot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `NumCard` (`NumCard`)
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
Public Sub MUE_EnviarFichasVacuno(ByVal NSistema As Integer, ByVal Modificaciones As Boolean, ByVal Marcar As Boolean, Optional MyCod As Long)
'Dim Crianza As Long
'Dim Nacimiento As Long
'Dim Produccion As Long
'Dim sCrianza As String
'Dim sNacimiento As String
'Dim sProduccion As Long
'Dim Sacrificio As Long
'Dim sSacrificio As String
'Dim Despiece As Long
'Dim sDespiece As Long
'Dim sSexo As String
'Dim sRaza As String
'Dim sCategoria As String
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecTmp As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nC1 As Integer
Dim Cm As String
'Dim nLast As Double

    If MyCod = 0 Then frmMonitor.MostrarInfo CargaCadena(619) '"Envío de Fichas Vacuno"
    
    Set Mybase = DAO.OpenDatabase(Base_General)
    
        Do_Events
        If MyCod = 0 Then
            If Modificaciones Then
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno where  (tranficha<>'*' OR tranficha is null)")
            Else
                Set MyRecVar = Mybase.OpenRecordset _
                ("select * from fichavacuno ")
            End If
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from fichavacuno where borrado=false and codigo=" & Chr(34) & MyCod & Chr(34))
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    If !borrado = False Then
                        'Crianza = !codcrianza
                        'On Error Resume Next
                        'sCrianza = Trim(!valorpcrianza)
                        'Nacimiento = !codnacimiento
                        'sNacimiento = Trim(!valorpnacimiento)
                        'Produccion = !codproduccion
                        'sProduccion = Trim(!valorpproduccion)
                        'Sacrificio = !codpsacrificio
                        'sSacrificio = Trim(!valorpsacrificio)
                        'Despiece = !codpdespiece
                        'sDespiece = Trim(!valorpdespiece)
                        'On Error GoTo 0
                        'If !codcrianza = 0 And !codnacimiento = 0 Then
                        '    Crianza = Produccion
                        '    sCrianza = sProduccion
                        '    Nacimiento = Produccion
                        '    sNacimiento = sProduccion
                        'End If
                        'Valores paises, categoria, raza, sexo
                        
                        ''''''''''''''''''''''''''''''''''''''
                        MyRecDes.Open "Select * from pcscale.bovine where numcard=" & CStr(!codigo), MyConnObj, adOpenKeyset, adLockOptimistic
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            If MyRecDes.EOF Then

                                MyRecDes.AddNew
                                MyRecDes.Fields("numcard") = !codigo
                                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(!Identificador))
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(!rsisacrificio))
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(!rsidespiece))
                                
                                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(!lote))
                                    
                            
                            Else
                                   
                                MyRecDes.Fields("numcard") = !codigo
                                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(!Identificador))
                                MyRecDes.Fields("birthland") = !codnacimiento
                                MyRecDes.Fields("birthdate") = Format(!fecha_nac, "yyyy-mm-dd")
                                MyRecDes.Fields("breedingland") = !codcrianza
                                MyRecDes.Fields("slaughterland") = !codsacrificio
                                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(!rsisacrificio))
                                MyRecDes.Fields("SlaughterDate") = Format(!fecha_sac, "yyyy-mm-dd")
                                MyRecDes.Fields("ButcherLand") = !coddespiece
                                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(!rsidespiece))
                                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(!lote))
                                   
                               
                            End If
                            
                            MyRecDes.Update
                            MyRecDes.Close
                            Set MyRecDes = Nothing
                            
                            frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo, "000") & " --> OK"
                            'Edit_Record Registro
                            .Edit
                            .Fields("tranficha") = "*"
                            .Update
    
                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        Cm = "delete from pcscale.bovine where numcard=" & CStr(!codigo)
                        MyConnObj.Execute Cm
                        ''''''''''''''''''''
                        frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tranficha") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirFicha:
                    
                Loop
            End If
        End With
    
SalirFicha:
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing

salir:

End Sub


Public Sub MUE_EnviaCLientes(NSistema As Integer, Modificaciones As Boolean, Marcar As Boolean)
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
'Dim MyRecTmp As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim Cm As String

    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCLIENTES
    End If
    On Error GoTo 0
    On Error GoTo FINCLIENTES
    Set MyRecVar = Mybase.OpenRecordset("select * from cliente where borrado=true")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo "Baja Cliente:" & CStr(MyRecVar.Fields("cod_cli"))
            Cm = "delete pcscale.customers where code=" & CStr(MyRecVar.Fields("cod_cli"))
            MyConnObj.Execute Cm
            
            MyRecVar.Delete
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente where tran_cli<>'*'")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from cliente")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            
            MyRecDes.Open "Select * from pcscale.customers where code=" & CStr(MyRecVar.Fields("cod_cli")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINCLIENTES
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            
            frmMonitor.MostrarInfo "Envio Cliente:" & CStr(MyRecVar.Fields("cod_cli"))
            
            MyRecDes.Fields("code") = MyRecVar.Fields("cod_cli")
            MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nom_cli")))
            MyRecDes.Fields("address") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("data1")))
            MyRecDes.Fields("city") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data2")))
            MyRecDes.Fields("web") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data3")))
            MyRecDes.Fields("phone") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("Data4")))
            MyRecDes.Fields("text") = ""
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_cli") = "*"
            MyRecVar.Update
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCliente
            
        Loop
    
    
    End If
SalirCliente:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCLIENTES:
    CadenadeLog "Error en MUE_EnviaClientes:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
End Sub

Public Function Ii(ByVal sS As String) As String
Dim nPos As Integer
Dim sR As String
    sR = sS
    nPos = InStr(1, sR, "I")
    If nPos <> 0 Then
        Do While InStr(1, sR, "I") <> 0
            nPos = InStr(1, sR, "I")
            sR = Mid(sR, 1, nPos - 1) & "i" & Mid(sR, nPos + 1)
        Loop
    End If
    Ii = sR
End Function

'Adaptar TouchScale Barna
Public Sub MUE_EnviaCodigosBarras(NSistema As Integer, Modificaciones As Boolean, Optional Marcar As Boolean)
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecTmp As DAO.Recordset
'Dim MyRecDes As New ADODB.Recordset
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String

    frmMonitor.MostrarInfo CargaCadena(233)
    
    Set Mybase = DAO.OpenDatabase(Base_General)
    
        Do_Events
        If Modificaciones Then
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar where (tran_bar<>'*' OR tran_bar is null)")
        Else
            Set MyRecVar = Mybase.OpenRecordset _
            ("select * from codbar")
        End If
        
        With MyRecVar
            If Not .EOF Then
                .MoveLast
                frmMonitor.barra1.Value = 0
                frmMonitor.barra1.Max = MyRecVar.Recordcount
                .MoveFirst
                Do Until .EOF
                    ' *******************
                    ' Envío de datos
                    ' *******************
                    
                    If !borrado = False Then
                        If !numero_eqp <> 0 Then
                            ' Equipo
                            CadenadeLog "Consulta Equipo " & CStr(!numero_eqp)
                            MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(!numero_eqp), MyConnObj, adOpenKeyset, adLockOptimistic
                        Else
                            ' Seccion
                            CadenadeLog "Consulta Seccion " & CStr(!secc_maqui)
                            MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(!secc_maqui), MyConnObj, adOpenKeyset, adLockOptimistic
                        End If
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            EstadoBal = 1000
                            canCelar = 1
                            Call frmControl.Balanza_ComError(0)
                            GoTo salir
                        Else
                            
                            If MyRecAux.EOF Then
                                MyRecAux.AddNew
                                MyRecAux.Fields("localsales") = ""
                                MyRecAux.Fields("localLineSLB") = ""
                                MyRecAux.Fields("localTotalSLB") = ""
                                MyRecAux.Fields("localSS") = ""
                                MyRecAux.Fields("localPacket1") = ""
                                MyRecAux.Fields("localPacket2") = ""
                                MyRecAux.Fields("localtotal") = ""
                                MyRecAux.Fields("networksales") = ""
                                MyRecAux.Fields("NetworkLineSLB") = ""
                                MyRecAux.Fields("NetworkTotalSLB") = ""
                                MyRecAux.Fields("NetworkSS") = ""
                                MyRecAux.Fields("NetworkPacket1") = ""
                                MyRecAux.Fields("NetworkPacket2") = ""
                                CadenadeLog "Nuevo item"
                            End If
                            
                            If IsNull(!numero_eqp) Then
                                ' Seccion
                                MyRecAux.Fields("number") = !secc_maqui
                                MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("C")
                            Else
                                If !numero_eqp <> 0 Then
                                    ' Balanza
                                    MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("S")
                                    MyRecAux.Fields("number") = !numero_eqp
                                Else
                                    ' Seccion
                                    MyRecAux.Fields("number") = !secc_maqui
                                    MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("C")
                                End If
                            End If
                            If !secc_eqp = 0 Then
                                ' Red
                                MyRecAux.Fields("networksales") = ConvertStringToUtf8Bytes(Trim(!CB_VNT))
                                MyRecAux.Fields("NetworkLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_SUP)))
                                MyRecAux.Fields("NetworkTotalSLB") = ConvertStringToUtf8Bytes((Trim(!CB_MIX))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_MIX)))
                                MyRecAux.Fields("NetworkSS") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                MyRecAux.Fields("NetworkPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                MyRecAux.Fields("NetworkPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV2)))
                            
                            Else
                                ' Local
                                MyRecAux.Fields("localsales") = ConvertStringToUtf8Bytes(Trim(!CB_VNT))
                                MyRecAux.Fields("localLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_SUP)))
                                MyRecAux.Fields("localTotalSLB") = ConvertStringToUtf8Bytes((Trim(!CB_MIX))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_MIX)))
                                MyRecAux.Fields("localSS") = ConvertStringToUtf8Bytes(Trim(Ii(!CB_EV1)))
                                MyRecAux.Fields("localPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                                MyRecAux.Fields("localPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV2)))
                                MyRecAux.Fields("localtotal") = ConvertStringToUtf8Bytes((Trim(!CB_MIX))) 'ConvertStringToUtf8Bytes(Ii(Trim(!CB_EV1)))
                            End If
                               
                            
                            MyRecAux.Update
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            If IsNull(!numero_eqp) Then
                                frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                            Else
                                If !numero_eqp Then
                                    frmMonitor.MostrarDato CargaCadena(23) & " EQP.:" & Format(!numero_eqp, "000") & " --> OK"
                                Else
                                    frmMonitor.MostrarDato CargaCadena(23) & " SEC.:" & Format(!secc_maqui, "000") & " --> OK"
                                End If
                            End If
                            .Edit
                            .Fields("tran_bar") = "*"
                            .Update

                        End If
                        
                    Else
                            
                        'Borrar en SQLSERVER
                        If !numero_eqp <> 0 Then
                            Cm = "delete from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(!numero_eqp)
                        Else
                            Cm = "delete from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(!secc_maqui)
                        End If
                        
                        MyConnObj.Execute Cm
                        ''''''''''''''''''''
                        'frmMonitor.MostrarDato CargaCadena(666) & " " & Format(!codigo + 1, "000") & " --> OK"
                        
                        'Delete_Record Registro
                        If Not .EOF Then .Delete
    
                    End If
                    
                    Call MonitorBarra
                    Call frmControl.Balanza_ComOK
                    
                    Do_Events
                    
                    MyRecVar.Edit
                    MyRecVar.Fields("tran_bar") = "*"
                    MyRecVar.Update
                    
                    If Not .EOF Then .Movenext
                    
                
                    If canCelar = 1 Or CancelHumano = True Then GoTo SalirBarras
                    
                Loop
            End If
        End With
    
SalirBarras:
    MyRecVar.Close
    Set MyRecVar = Nothing
salir:
    Mybase.Close
    Set Mybase = Nothing

End Sub


Public Sub MUE_EnviaVendedor(NSistema As Integer, Modificaciones As Boolean)
    '///////////////////////////////////////
    '// Envío de vendedores (TouchScale)    //
    '///////////////////////////////////////
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
Dim Cm As String
Dim lEsAlta As Boolean
Dim sColorTmp As String
Dim nColor As Integer

    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINVENDEDOR
    End If
    On Error GoTo 0
    On Error GoTo FINVENDEDOR
    
    Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where borrado=true or tran_vend<>'*'")
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            'frmMonitor.MostrarInfo "Baja Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
            CadenadeLog "Baja Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
            Cm = "delete from pcscale.vendorkeys where counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend"))
            MyConnObj.Execute Cm
            Cm = "delete from pcscale.vendors where code=" & CStr(MyRecVar.Fields("ident_vend"))
            MyConnObj.Execute Cm
     
            If MyRecVar.Fields("borrado") = True Then
                MyRecVar.Delete
            End If
            If Not MyRecVar.EOF Then MyRecVar.Movenext
        Loop
    
    End If
    MyRecVar.Close
    Set MyRecVar = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Modificaciones Then
        Set MyRecVar = Mybase.OpenRecordset("select * from vendedor where tran_vend<>'*'")
    Else
        Set MyRecVar = Mybase.OpenRecordset("select * from vendedor")
    End If
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            frmMonitor.MostrarInfo "Envio Vendedor:" & CStr(MyRecVar.Fields("ident_vend"))
'DROP TABLE IF EXISTS `pcscale`.`vendorkeys`;
'CREATE TABLE  `pcscale`.`vendorkeys` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Counter` int(11) DEFAULT NULL,
'  `VendorKey` int(11) DEFAULT NULL,
'  `Vendor` int(11) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Vendor` (`Vendor`),
'  UNIQUE KEY `Counter` (`Counter`,`VendorKey`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
            'MyRecDes.Open "Select * from pcscale.vendorkeys where vendor=" & CStr(MyRecVar.Fields("ident_vend")) & " and counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
            MyRecDes.Open "Select * from pcscale.vendorkeys where counter=" & CStr(MyRecVar.Fields("secc_maqui")) & " and vendorkey=" & CStr(MyRecVar.Fields("tec_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINVENDEDOR
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
                lEsAlta = True
            Else
                lEsAlta = False
            End If

            MyRecDes.Fields("vendor") = MyRecVar.Fields("ident_vend")
            MyRecDes.Fields("counter") = MyRecVar.Fields("secc_maqui")
            MyRecDes.Fields("vendorkey") = MyRecVar.Fields("tec_vend")
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
                
            On Error Resume Next
            MyRecDes.Open "select * from pcscale.vendors where code=" & CStr(MyRecVar.Fields("ident_vend")), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Call frmControl.Balanza_ComError(0)
                GoTo FINVENDEDOR
            End If
            On Error GoTo FINVENDEDOR
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            MyRecDes.Fields("code") = MyRecVar.Fields("ident_vend")
            MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nombre")))

            MyRecDes.Update
            
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            MyRecVar.Edit
            MyRecVar.Fields("tran_vend") = "*"
            MyRecVar.Fields("prog_vend") = "*"
            MyRecVar.Update
            MyRecVar.Movenext

            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirVend

        Loop
    
    
    End If
SalirVend:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINVENDEDOR:
    CadenadeLog "Error en MUE_EnviaVendedor:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
       
End Sub

Public Function Capa1A_GA_Erase_Totals(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para TouchScale  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim Cm As String
    Dim nF As Integer
    Dim nC As Long
    Dim ncont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
   
On Error Resume Next
            MyRecSet.Close
            Set MyRecSet = Nothing
On Error GoTo 0

    CadenadeLog "Capa1A_GA_Erase_totals"
    'On Error Resume Next
    
    'CadenadeLog "Identificación Balanza..."
    'MyRecAux.Open "select * from dbo.balanza", MyConnObj, adOpenKeyset, adLockOptimistic
    'If Err.Number <> 0 Then
    '    Resp = 1000
    '    GoTo FIN_Capa1A_GA_Erase_Totals
    'End If
    'On Error GoTo FIN_Capa1A_GA_Erase_Totals
    'If Not MyRecAux.EOF Then
    '    'ejemplo:FB9F5AA7-3AB9-46A4-8EDE-F1B57957C5B5
    '    'CadenadeLog "Captura idbalanza"
    '    IdEqp = MyRecAux.Fields("idbalanza")
    'Else
    '    IdEqp = "00000000-0000-0000-0000-" & Format(Now(), "ddmmyyhhmmss") '"00000000-0000-0000-0000-000000000000"
    'End If
    'MyRecAux.Close
    'Set MyRecAux = Nothing
    ''CadenadeLog IdEqp
    '
    ''On Error Resume Next
    
    On Error GoTo FIN_Capa1A_GA_Erase_Totals
    
    'CadenadeLog "Selección por tipo documento..."
    'MyRecAux1.Open "select A1.TipoDocumento tipdoc, sum(A2.unidades) uni, sum(A2.peso) pes, SUM(a2.importeconimpuestos) imp, COUNT(distinct(a2.IdTransaction)) ope from dbo.Documento A1, dbo.LineaDocumento A2 where (A1.IsFinalizado = 1 and A2.IdDocumento = A1.IdDocumento) group by A1.tipodocumento", MyConnObj, adOpenKeyset, adLockOptimistic
    'On Error GoTo FIN_Capa1A_GA_Erase_Totals
    'If Not MyRecAux1.EOF Then
    '    MyRecAux1.MoveFirst
    '    Do While Not MyRecAux1.EOF
    '
    '        'MyRecSet.Open "select * from dbo.acumuladoventas", MyConnObj, adOpenKeyset, adLockOptimistic
    '        'MyRecSet.AddNew
    '        Cm = "insert into dbo.acumuladoventas values ('" & _
    '             IdEqp & "'," & CStr(MyRecAux1.Fields("tipdoc")) & ",convert(date,'" & Format(Now(), "yyyy-mm-dd") & " 00:00:00.000" & "')," & Quitar_Coma(CStr(MyRecAux1.Fields("imp")), Len(CStr(MyRecAux1.Fields("imp")))) & "," & _
    '             CStr(MyRecAux1.Fields("ope")) & "," & CStr(MyRecAux1.Fields("uni")) & "," & Quitar_Coma(CStr(MyRecAux1.Fields("pes")), Len(CStr(MyRecAux1.Fields("pes")))) & ")"
    '        MyConnObj.Execute Cm
    '        '''''''''''''''''''''''''''''''
    '        ''MyRecSet.Fields("idbalanza") = 0
    '        'CadenadeLog "añadido registro en Acumuladoventas."
    '        'On Error Resume Next
    '        'MyRecSet.Fields("idbalanza") = IdEqp
    '        'On Error GoTo 0
    '        'On Error GoTo FIN_Capa1A_GA_Erase_Totals
    '
    '        'MyRecSet.Fields("tipodocumento") = MyRecAux1.Fields("tipdoc")
    '        'MyRecSet.Fields("fecha") = CDate(Now()) 'Format(Now(), "yyyy-dd-mm") & " 00:00:00.000"
    '        'MyRecSet.Fields("importe") = MyRecAux1.Fields("imp")
    '        'MyRecSet.Fields("operaciones") = MyRecAux1.Fields("ope")
    '        'MyRecSet.Fields("unidades") = MyRecAux1.Fields("uni")
    '        'MyRecSet.Fields("peso") = MyRecAux1.Fields("pes")
    '        'MyRecSet.Update
    '
    '        'MyRecSet.Close
    '        'Set MyRecSet = Nothing
    '
    '        MyRecAux1.Movenext
    '    Loop
    'End If
    ''MyRecSet.Close
    ''Set MyRecSet = Nothing
    'MyRecAux1.Close
    'Set MyRecAux1 = Nothing
    '
    
'*******
                    

'*******
    
    
    CadenadeLog "Iniciar Numeración si Procede..."
    
    If No_Reinit = 0 Then 'Iniciar numeración tiques...
        Cm = "truncate pcscale.localnumtickets"
        MyConnObj.Execute Cm
        Cm = "truncate pcscale.numtickets"
        MyConnObj.Execute Cm
        Cm = "truncate pcscale.erasetickets"
        MyConnObj.Execute Cm
    End If
    CadenadeLog "Borrado Documentos y Lineas Documento..."
    Cm = "truncate pcscale.paytickets"
    MyConnObj.Execute Cm
    Cm = "truncate pcscale.HTickets"
    MyConnObj.Execute Cm
    Cm = "truncate pcscale.LTickets"
    MyConnObj.Execute Cm
    Cm = "truncate pcscale.sservicelabels"
    MyConnObj.Execute Cm
    Cm = "truncate pcscale.packetlabels"
    MyConnObj.Execute Cm
    
    'Añadididas 1.0.4
    Cm = "truncate pcscale.localerasetickets"
    MyConnObj.Execute Cm

    Cm = "truncate pcscale.localhtickets"
    MyConnObj.Execute Cm


    Cm = "truncate pcscale.localltickets"
    MyConnObj.Execute Cm


    Cm = "truncate pcscale.localhtickets"
    MyConnObj.Execute Cm


    Cm = "truncate pcscale.localpacketlabels"
    MyConnObj.Execute Cm


    Cm = "truncate pcscale.localpacketpresets"
    MyConnObj.Execute Cm


    Cm = "truncate pcscale.localpaytickets"
    MyConnObj.Execute Cm

    Cm = "truncate pcscale.localsservicelabels"
    MyConnObj.Execute Cm
    ''''''''''''''''''''''''''''''''
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_GA_Erase_Totals = Resp
End Function

''''¡¡¡¡¡¡¡¡¡¡¡¡¡ Solo utilizar en final día!!!!!!!!!!!!!!!!!!!'''''''
'   Se borran datos de la Balanza
'   Se deben introducir tiques en dbasetouch.mdb con typtic=4
'   En principio la información de cabecera no es válida...
'   Se pensaba en utilizar las tablas de TouchScale.MDB pero se va a traspasar directamente a dbasetouch.mdb
Public Function Tiques(ByVal lLocal As Boolean) As Integer
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset

Dim MyBaseORG1 As DAO.Database
Dim MyRsORG1 As DAO.Recordset
Dim MyrsART As DAO.Recordset
Dim MyRsVEND As DAO.Recordset

Dim MyBaseBAL As DAO.Database
Dim MyRsBAL As DAO.Recordset
Dim siP As String
Dim MyRsDes As New ADODB.Recordset
Dim myRsKEYS As New ADODB.Recordset
Dim myRsTMP As New ADODB.Recordset
Dim Cm As String
Dim nF As Integer
Dim ncont As Integer
Dim sS As String
Dim nImporte As Long
Dim sHour As String
Dim sFecha As String
'Dim nF As Integer

'    lComunica = True
'nF = FreeFile()
'Open App.Path & "\LOGTOUCH.TXT" For Output As #nF
'Print #nF, "Recogida Tiques Touch-Scale..."
    
    sHour = Format(Time, "hh:mm:ss")
    sFecha = Format(Date, "dd/mm/yy")
    
    'Set MyBaseBAL = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    'Set MyRsBAL = MyBaseBAL.OpenRecordset("select * from balanzas where ip<>'127-0-0-1' order by scalenr")
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbasetouch.mdb")
    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where borrado=false")
    ''Call AbrirTotales
    ''If lTotalesAbierta Then
    ''    nF = nF
    ''End If
    ''Cm = "select * from articulo where borrado=false"
    ''MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic


    'If Not MyRsBAL.EOF Then
        'MyRsBAL.MoveFirst
        'Do While Not MyRsBAL.EOF
        
            'On Error GoTo ProblemaPing
        
            'sIp = toIpAdd(Trim(MyRsBAL.Fields("ip")))
            
            
            'If ElPing(sIp) = 1 Then
    
            '    Call AbrirMySQL(sIp)
                
            '    If lMySqlAbierta Then
                    For ncont = 0 To 1 '0=local, 1=Red
                        If ncont = 0 Then
'Print #nF, "Recogida Tiques local Touch-Scale..."
'DROP TABLE IF EXISTS `pcscale`.`htickets`;
'CREATE TABLE  `pcscale`.`htickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdNetHticket` int(11) DEFAULT '0',
'  `NetStat` char(1) DEFAULT '',
'  `Captured` char(1) DEFAULT '',
'  `tckType` char(1) DEFAULT NULL,
'  `Status` char(1) DEFAULT NULL,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `DatTim` datetime DEFAULT NULL,
'  `ReloadFrom` int(11) DEFAULT '0',
'  `Customer` int(11) DEFAULT '0',
'  `tckDiscount` decimal(4,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Vendor` (`Vendor`,`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                            
'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
'CREATE TABLE  `pcscale`.`ltickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                            Cm = "select * from pcscale.LocalLTickets"
                            MyRsDes.Open Cm, MyConnObj, adOpenDynamic, adLockOptimistic
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF

                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG!seccion = CStr(8)
                                    MyRsORG!Balanza = MyRsBAL!scalenr
                                    MyRsORG!vended = CStr(6099)
                                    MyRsORG!numlin = CStr(1)
                                    MyRsORG!lincan = CStr(0)
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG!Importe = MyRsDes!price
                                        nImporte = (MyRsDes!price) * 100
                                    Else
                                        MyRsORG!Importe = Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
                                        nImporte = 100 * (Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2))
                                    End If
                                    MyRsORG!hora = sHour
                                    MyRsORG!onoff = "Local"
                                    MyRsORG!STPETIC = CStr(16)
                                    MyRsORG!codcli = CStr(0)
                                    MyRsORG!Fecha = sFecha
                                    MyRsORG!TYPTIC = CStr(4)
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
                                    sS = sS & CStr(DigitoControl(sS))
                                    MyRsORG!codbar = sS
                                    MyRsORG!reserv1 = "No Enviado"
                                    MyRsORG!reserv2 = "*"
                                    MyRsORG!reserv3 = "SI"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sFecha
                                    MyRsORG1!Hour = sHour
                                    MyRsORG1!numbal = MyRsBAL!scalenr
                                    MyRsORG1!vendor = CStr(6099)
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG1!numlin = CStr(1)
                                    MyRsORG1!secc = CStr(8)
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    
                                    If MyrsART.EOF Then
                                        MyRsORG1!Plu = CStr(0)
                                        MyRsORG1!Item = ""
                                    Else
                                        MyRsORG1!Plu = CStr(MyrsART!Plu)
                                        MyRsORG1!Item = MyrsART!des_plu1
                                    End If
                                    MyRsORG1!price = MyRsDes!price
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG1!amount = MyRsDes!price
                                        MyRsORG1!weigth = 0
                                        MyRsORG1!units = 1
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
                                        MyRsORG1.Fields("Weight") = MyRsDes.Fields("Weight")
                                        MyRsORG1!units = 0
                                    End If
                                    MyRsORG1!STPETIC = CStr(16)
                                    MyRsORG1!lincan = CStr(0)
                                    MyRsORG1!IVA = 0
                                    MyRsORG1!TYPTIC = CStr(4)
                                    
                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sFecha, "dd/mm/yy")
                                            !codi_ident = 188
                                            !ident_vend = 6099
                                            !secc_maqui = 8
                                            !Modo = 4
                                            !tiquets = 0
                                            !operacion = 0
                                            !cancela = 0
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !Peso = 0
                                            !unidad = 0
                                            ' ****
                                            !pesetas = 0
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
                                            !Supertot = 0
                                            !Total = 0
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
                                            !operacion = !operacion + 1
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!price)
                                                !unidad = !unidad + 1
                                                !Total = !Total + (MyRsDes!price)
                                                !Supertot = !Total + MyRsDes!price
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
                                                !Peso = !Peso + (MyRsDes.Fields("Weight"))
                                                !Total = !Total + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
                                                !Supertot = !Total + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    
                                    MyRsDes.Movenext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                        
                            MyRsVEND.Clone
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        Else
                            
'Print #nF, "Recogida Tiques Red Touch-Scale..."
                            
                            Cm = "select * from pcscale.LTickets"
                            MyRsDes.Open Cm, MyConnObj, adOpenDynamic, adLockOptimistic
'Print #nF, "Paso 001"
                            
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
'Print #nF, "Paso 002"
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF
                                    
                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
'Print #nF, "Paso 003"
                                    MyRsORG!seccion = CStr(8)
'Print #nF, "Paso 004"
                                    MyRsORG!Balanza = 1 'Ver número Balanza
'Print #nF, "Paso 005"
                                    MyRsORG!vended = CStr(6099)
'Print #nF, "Paso 006"
                                    MyRsORG!numlin = CStr(1)
'Print #nF, "Paso 007"
                                    MyRsORG!lincan = CStr(0)
'Print #nF, "Paso 008"
                                    If MyRsDes!saleform = 0 Then
'Print #nF, "Paso 009"
                                        
                                        MyRsORG!Importe = MyRsDes!price
'Print #nF, "Paso 0010"
                                        nImporte = (MyRsDes!price) * 100
'Print #nF, "Paso 0011"
                                    Else
                                        MyRsORG!Importe = Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0012"
                                        nImporte = 100 * (Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2))
'Print #nF, "Paso 0013"
                                    End If
                                    MyRsORG!hora = sHour
'Print #nF, "Paso 0014"
                                    MyRsORG!onoff = "Red"
'Print #nF, "Paso 0015"
                                    MyRsORG!STPETIC = CStr(16)
'Print #nF, "Paso 0016"
                                    MyRsORG!codcli = CStr(0)
'Print #nF, "Paso 0017"
                                    MyRsORG!Fecha = sFecha
'Print #nF, "Paso 0018"
                                    MyRsORG!TYPTIC = CStr(4)
'Print #nF, "Paso 0019"
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
'Print #nF, "Paso 0020"
                                    sS = sS & CStr(DigitoControl(sS))
'Print #nF, "Paso 0021"
                                    MyRsORG!codbar = sS
'Print #nF, "Paso 0022"
                                    MyRsORG!reserv1 = "No Enviado"
'Print #nF, "Paso 0023"
                                    MyRsORG!reserv2 = "*"
'Print #nF, "Paso 0024"
                                    MyRsORG!reserv3 = "SI"
'Print #nF, "Paso 0025"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sFecha
'Print #nF, "Paso 0026"
                                    MyRsORG1!Hour = sHour
'Print #nF, "Paso 0027"
                                    MyRsORG1!numbal = 1 'MyRsBAL!scalenr
'Print #nF, "Paso 0028"
                                    MyRsORG1!vendor = CStr(6099)
'Print #nF, "Paso 0029"
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
'Print #nF, "Paso 0030"
                                    MyRsORG1!numlin = CStr(1)
'Print #nF, "Paso 0031"
                                    MyRsORG1!secc = CStr(8)
'Print #nF, "Paso 0032"
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
'Print #nF, "Paso 0033"
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
'Print #nF, "Paso 0034"
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    If MyrsART.EOF Then
                                        MyRsORG1!Plu = CStr(0)
'Print #nF, "Paso 0035"
                                        MyRsORG1!Item = ""
'Print #nF, "Paso 0036"
                                    Else
                                        MyRsORG1!Plu = CStr(MyrsART!Plu)
'Print #nF, "Paso 0037"
                                        MyRsORG1!Item = MyrsART!des_plu1
'Print #nF, "Paso 0038"
                                    End If
                                    MyRsORG1!price = MyRsDes!price
'Print #nF, "Paso 0039"
                                    If MyRsDes!saleform = 0 Then
'Print #nF, "Paso 0040"
                                        MyRsORG1!amount = MyRsDes!price
'Print #nF, "Paso 0041"
                                        MyRsORG1.Fields("weight") = 0
'Print #nF, "Paso 0042"
                                        MyRsORG1!units = 1
'Print #nF, "Paso 0043"
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0044"
                                        MyRsORG1.Fields("Weight") = MyRsDes.Fields("Weight")
'Print #nF, "Paso 0045"
                                        MyRsORG1!units = 0
'Print #nF, "Paso 0046"
                                    End If

                                    MyRsORG1!STPETIC = CStr(16)
'Print #nF, "Paso 0046"
                                    MyRsORG1!lincan = CStr(0)
'Print #nF, "Paso 0047"
                                    MyRsORG1!IVA = 0
'Print #nF, "Paso 0048"
                                    MyRsORG1!TYPTIC = CStr(4)
'Print #nF, "Paso 0049"
                                    

                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
'Print #nF, "Paso 0050"
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sFecha, "dd/mm/yy")
'Print #nF, "Paso 0051"
                                            !codi_ident = 188
'Print #nF, "Paso 0052"
                                            !ident_vend = 6099
'Print #nF, "Paso 0053"
                                            !secc_maqui = 8
'Print #nF, "Paso 0054"
                                            !Modo = 4
'Print #nF, "Paso 0055"
                                            !tiquets = 0
'Print #nF, "Paso 0056"
                                            !operacion = 0
'Print #nF, "Paso 0057"
                                            !cancela = 0
'Print #nF, "Paso 0058"
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
'Print #nF, "Paso 0059"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !Peso = 0
'Print #nF, "Paso 0060"
                                            !unidad = 0
'Print #nF, "Paso 0061"
                                            ' ****
                                            !pesetas = 0
'Print #nF, "Paso 0062"
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
'Print #nF, "Paso 0063"
                                            !Supertot = 0
'Print #nF, "Paso 0064"
                                            !Total = 0
'Print #nF, "Paso 0065"
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
'Print #nF, "Paso 0066"
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
'Print #nF, "Paso 0067"
                                            !operacion = !operacion + 1
'Print #nF, "Paso 0068"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
'Print #nF, "Paso 0069"
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!price)
'Print #nF, "Paso 0070"
                                                !unidad = !unidad + 1
'Print #nF, "Paso 0071"
                                                !Total = !Total + (MyRsDes!price)
'Print #nF, "Paso 0072"
                                                !Supertot = !Total + MyRsDes!price
'Print #nF, "Paso 0073"
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0074"
                                                !Peso = !Peso + (MyRsDes.Fields("Weight"))
'Print #nF, "Paso 0075"
                                                !Total = !Total + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0076"
                                                !Supertot = !Total + Round((MyRsDes!price) * (MyRsDes.Fields("Weight")), 2)
'Print #nF, "Paso 0077"
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    MyRsDes.Movenext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                            MyRsVEND.Close
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        End If
                    
                    Next ncont
                            
                    'Borrado de los datos de Totales en Balanza
'Print #nF, "Borrado Tiques Touch-Scale..."
                    On Error Resume Next
                    Cm = "truncate pcscale.paytickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.HTickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.LTickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.sservicelabels"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.packetslabels"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    
                    'Cm = "delete from pcscale.LocalHTickets"
                    'myconnobj.Execute Cm
                    'Cm = "delete from pcscale.LocalLTickets"
                    'myconnobj.Execute Cm
                    'Cm = "delete from pcscale.localnumtickets"
                    'myconnobj.Execute Cm
                    
                    'Intento reiniciar numeración
                    'Cm = "update pcscale.status set resetticketnumber=1 where id=1"
                    'myconnobj.Execute Cm
                    'Cm = "update pcscale.erasetickets set dattim=now()"
                    'myconnobj.Execute Cm
                    On Error Resume Next
                    Cm = "truncate pcscale.localnumtickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.numtickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    On Error Resume Next
                    Cm = "truncate pcscale.erasetickets"
                    MyConnObj.Execute Cm
                    On Error GoTo 0
                    
                    ''''''''''''''''''''''''''''''

            '    End If
            
            'End If
        
            'If lMySqlAbierta Then
            '    myconnobj.Close
            '    Set DbMysqli = Nothing
            '    lMySqlAbierta = False
            'End If
            '
            'MyRsBAL.Movenext
        
        'Loop
        
    'End If

    On Error Resume Next
    
    'MyRsBAL.Close
    'Set MyRsBAL = Nothing
    
    MyBaseBAL.Close
    Set MyBaseBAL = Nothing
    
    MyrsART.Close
    Set MyrsART = Nothing

    MyBaseORG.Close
    Set MyBaseORG = Nothing
    'If lTotalesAbierta Then
    '    DbTotales.Close
    '    Set DbTotales = Nothing
    'End If
    
'    On Error GoTo 0
'ProblemaPing:
    'If Err.Number <> 0 Then
        'Print #nF, "Error:" & CStr(Err.Number) & " Desc.:" & Err.Description
        'On Error GoTo 0
        'On Error Resume Next
        'If lMySqlAbierta Then
        '    myconnobj.Close
        '    Set DbMysqli = Nothing
        'End If
        'On Error GoTo 0
        
    'End If
    'On Error Resume Next
    'Close #nF
    On Error GoTo 0
    'lComunica = False
    
End Function


'Print #nF, "Recogida Tiques local Touch-Scale..."
'DROP TABLE IF EXISTS `pcscale`.`htickets`;
'CREATE TABLE  `pcscale`.`htickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdNetHticket` int(11) DEFAULT '0',
'  `NetStat` char(1) DEFAULT '',
'  `Captured` char(1) DEFAULT '',
'  `tckType` char(1) DEFAULT NULL,
'  `Status` char(1) DEFAULT NULL,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `DatTim` datetime DEFAULT NULL,
'  `ReloadFrom` int(11) DEFAULT '0',
'  `Customer` int(11) DEFAULT '0',
'  `tckDiscount` decimal(4,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Vendor` (`Vendor`,`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                            
'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
'CREATE TABLE  `pcscale`.`ltickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
'DROP TABLE IF EXISTS `pcscale`.`paytickets`;

'CREATE TABLE  `pcscale`.`paytickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `PayMethod` int(11) DEFAULT NULL,
'  `Delivered` decimal(8,2) DEFAULT '0.00',
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`paymethods`;
'CREATE TABLE  `pcscale`.`paymethods` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Name` varchar(64) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Name` (`Name`)
') ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`packetlabels`;
'CREATE TABLE  `pcscale`.`packetlabels` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `Price` decimal(6,2) DEFAULT NULL,
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  `Customer` int(11) DEFAULT '0',
'  `DatTim` datetime DEFAULT NULL,
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB AUTO_INCREMENT=70 DEFAULT CHARSET=utf8;

'DROP TABLE IF EXISTS `pcscale`.`sservicelabels`;
'CREATE TABLE  `pcscale`.`sservicelabels` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Number` int(11) DEFAULT '0',
'  `Counter` int(11) DEFAULT '0',
'  `Scale` int(11) DEFAULT '0',
'  `Vendor` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `Price` decimal(6,2) DEFAULT NULL,
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  `Customer` int(11) DEFAULT '0',
'  `DatTim` datetime DEFAULT NULL,
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;
'Cabeceras de tique:
'.- NetStat: ‘N’, ‘L’.
'    Tique en red (‘N’) o local (‘L’).
'    Los tiques de red copiados en local se guardan en la tabla LocalHTicktes con NetStat=’N’.
'.- TckType: ‘N’, ‘O’
'    Tique “Normal” o  “Order” (encargo).
'    Los tiques encargo sólo se marcan como tales, no se realiza ninguna gestión especial.
'.- Status: ‘O’, ‘P’, ‘C’, ’D’, ‘R’, ‘N’, ‘X’
'    Estado del tique: abierto (‘O’), en impresión (‘P’), cerrado (‘C’), cancelado (‘D’),
'    recargado en tique abierto (‘R’),  recargado en tique cerrado (‘N’), cancelado manualmente (‘X’)
'Líneas de tique:
'.- LineType: ‘1’, ‘0’
'    Línea “normal” (‘1’) o cancelada (‘0’).
'.- TTare: ‘ ‘ , ‘W’, ‘M’, ‘I’
'    Tipo de tara aplicado en la operación. Tara acumulativa/pesada (‘W‘), manual (‘M’),
'    programada en el artículo (‘I’), sin tara (‘ ‘)


Public Sub LeeTickets()
    
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim Cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim ncont As Long
    Dim dLImporte As Double
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    '=============
    'Exit Sub
    '=============
    
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    On Error GoTo FIN
    Set Mybase = DAO.OpenDatabase(Base_General)
    
    'CadenadeLog "1"
    
    'HydraTouch Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    '*************************
    'CadenadeLog "select * from pcscale.htickets where captured='' and status='C' order by id"
    
    'MyRecSetTmp.CursorLocation = adUseClient
    'MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and status='C' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    'MyRecSetTmp.CursorLocation = adUseServer
    MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and status='C' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    
    'CadenadeLog "Estado Recordset:" & CStr(MyRecSetTmp.State)
    'CadenadeLog "Número de Registros:" & CStr(MyRecSetTmp.Recordcount)
    'CadenadeLog "2"
    'MyRecSetTmp.MoveFirst
    If Not MyRecSetTmp.EOF Then
        'frmMonitor.barra1.Value = 0
        'MyRecSetTmp.MoveLast
        'frmMonitor.barra1.Max = MyRecSetTmp.Recordcount
        'CadenadeLog "2.5"
        'CadenadeLog "Número de Registros:" & CStr(MyRecSetTmp.Recordcount)
        'MyRecSetTmp.MoveFirst
        'CadenadeLog "2.75"
        Do While Not MyRecSetTmp.EOF
            
            'CadenadeLog "3"
            
            Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
            Fhora = Mid(Ffecha, 12, 8)
            Ffecha = Mid(Ffecha, 1, 10)
            
            'CadenadeLog "HydraTouch:" & Ffecha & " :" & Fhora
            
            IdDoc = MyRecSetTmp.Fields("id")
            ''''medio de pago
            '1=efectivo
            '2=t.crédito
            '3=cheque
            Cabecera.tipoPago = 1
            Cabecera.cantidadEntre = 0
            MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            
            'CadenadeLog "4"
            
            If Not MyRecAux.EOF Then
                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            '''''''''''''''''
            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'CadenadeLog "5"
            
            ReDim Cabecera.Lines(nNumLin)
            
            MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
            dLImporte = MyRecAux.Fields("importe")
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'recoger datos
            Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
            Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
            Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
            Cabecera.NCliente = MyRecSetTmp.Fields("customer")
            Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.hora = Fhora
            Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
            Cabecera.ImporteTotal = dLImporte 'myrecsettmp.Fields("totalneto")
            Cabecera.NLineas = nNumLin
            Cabecera.LineaInicial = 1
            Cabecera.LineaFinal = nNumLin
            Cabecera.LineasMensaje = nNumLin
            
            'cabecera.codigo ???
            
            ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
            '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
            If MyRecSetTmp.Fields("tcktype") = "N" Then
                Cabecera.CadenaEstadoTiquet = "Local"
            Else
                Cabecera.CadenaEstadoTiquet = "Red"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            dLImporte = Round(MyRecAux.Fields("importe"), 2)
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'CadenadeLog "6"
            
            Cabecera.ImporteBruto = dLImporte 'myrecsettmp.Fields("totalbruto")
            Cabecera.Impuestos = 0 'myrecsettmp.Fields("totalimpuestos")
            Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal  '.Fields("totaldescuento")
            Cabecera.documentoabono = 0 'myrecsettmp.Fields("iddocumentoabono")
            
            frmMonitor.MostrarDato "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
            CadenadeLog "HydraTouch:" & "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
            
            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            
            'CadenadeLog "7"
            
            Cabecera.LineasCanceladas = nNumLin
            
            'versión Julio(0=(Mixto y Super y TPV), 1=Venta, 2=Envasado, 3=Autoservicio)
            'HydraTouch Tipos de Tiquet (typtic)
            ' 1 --> Rollo
            ' 2 --> Super
            ' 3 --> Envasado
            ' 4 --> Autoservicio
            ' 5 --> Rollo Local
            ' 6 --> Super Local
            ' 7 --> Envasado Local
            ' 8 --> Autoservicio Local
            '*************************
            ''''''''''''''''''''''''''''
            '1.0.0 (0=Factura pasa a 7 red, 8 local
            '       1=tique pasa a 1 red, 4 local
            '       2=Envasado pasa a 2 red, 5 local
            '       3=AutoServicio pasa a 3 red, 6 local
            Select Case MyRecSetTmp.Fields("status")
            Case "C" 'tique venta ??? --> ¿Cómo identifica los diferentes modos?
                If Cabecera.CadenaEstadoTiquet = "Local" Then
                        Cabecera.tipoTiquet = 4
                Else
                        Cabecera.tipoTiquet = 1
                End If
            'Case "O"
            '    If Cabecera.CadenaEstadoTiquet = "Local" Then
            '        Cabecera.tipoTiquet = 5
            '    Else
            '        Cabecera.tipoTiquet = 2
            '    End If
            'Case Else
            '    If Cabecera.CadenaEstadoTiquet = "Local" Then
            '        Cabecera.tipoTiquet = 6
            '    Else
            '        Cabecera.tipoTiquet = 3
            '    End If
            End Select
            
            
            'también se vé afectado si es cancelado o no pero de mamento cerrado...
            Cabecera.EstadoTiquet = 16
            '********************
            ' este dato se introduce
            ' tras pasar por actualiza_tabla_cabecera
            'Incluir_en_Base As Boolean
            'Cabecera.tipoPago = myrecsettmp.Fields("idmoneda")
            'Cabecera.cantidadEntre = myrecsettmp.Fields("totalneto") + myrecsettmp.Fields("importecambio")
            
            Cabecera.Cambio = Cabecera.ImporteTotal - Cabecera.cantidadEntre  ' myrecsettmp.Fields("importecambio")
            Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")
            'lote As String
            'lEncargo As Boolean


            'LINEAS

            nNumLin = 0
            MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            
            'CadenadeLog "8"
            
            If Not MyRecAux.EOF Then
                
                'CadenadeLog "8.1"
                
                MyRecAux.MoveFirst
                
                'CadenadeLog "8.2"
                
                Do While Not MyRecAux.EOF
                    
                    'CadenadeLog "8.3"
'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
'CREATE TABLE  `pcscale`.`ltickets` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `IdHTicket` int(11) DEFAULT NULL,
'  `lineType` smallint(6) DEFAULT NULL,
'  `SaleForm` smallint(6) DEFAULT NULL,
'  `Item` int(11) DEFAULT NULL,
'  `Weight` decimal(5,3) DEFAULT NULL,
'  `Tare` decimal(5,3) DEFAULT '0.000',
'  `TTare` char(1) DEFAULT '',
'  `lineDiscount` decimal(4,2) DEFAULT NULL,
'  `Price` decimal(6,2) DEFAULT NULL,
'  `Amount` decimal(8,2) DEFAULT NULL,
'  `BovineCard` smallint(3) DEFAULT NULL,
'  `EarTag` varchar(30) DEFAULT NULL,
'  `BovineLot` varchar(20) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  KEY `IdHTicket` (`IdHTicket`)
') ENGINE=InnoDB DEFAULT CHARSET=latin1;
                    nNumLin = nNumLin + 1
                    If MyRecAux.Fields("saleform") = 1 Then
                        Cabecera.Lines(nNumLin).EsPesado = True
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                    Else
                        Cabecera.Lines(nNumLin).EsPesado = False
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                    End If
                    
                    'CadenadeLog "8.4"
                    
                    Cabecera.Lines(nNumLin).FactorPeso = 1
                    Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                    
                    'CadenadeLog "8.5"
                    
                    Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                    Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                    'ver qué se hace IVA As Integer
                    'ver qué se hace con TipoLinea As Byte
                    
                    'CadenadeLog "8.6"
                    
                    Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("linediscount") 'porcentaje
                    Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("amount") 'MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                    Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2) 'MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                    
                    'CadenadeLog "8.7"
                    
                    'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                    If MyRecAux.Fields("linetype") = 0 Then
                        Cabecera.Lines(nNumLin).LineaCancelada = True
                    Else
                        Cabecera.Lines(nNumLin).LineaCancelada = False
                    End If
                    'If MyRecAux.Fields("precio") <> MyRecAux.Fields("preciodefecto") Then
                    '    Cabecera.Lines(nNumLin).SehaModificado = True
                    'Else
                        Cabecera.Lines(nNumLin).SehaModificado = False
                    'End If
                    '¿? ModificadoTemporal As Boolean
                    
                    'CadenadeLog "8.8"
                    'c2f 1.1.1 aclarar ficha Vacuno...
                    'If IsNull(MyRecAux.Fields("eartag")) Or MyRecAux.Fields("eartag") = "" Then
                        Cabecera.Lines(nNumLin).FichaVacuno = 0
                    'Else
                    '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("eartag")
                    'End If
                    'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                    'Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    'CadenadeLog "8.9"
                    
                    MyRecAux.Movenext
                Loop
                MyRecAux.Close
                Set MyRecAux = Nothing
                Actualiza_Tabla_Cabecera Mybase, Cabecera
                Do_Events
                If Cabecera.Incluir_en_Base = True Then
                    Actualiza_Tabla_Tickets Mybase, Cabecera
                    Do_Events
                    Actualiza_Tabla_Gtarti Mybase, Cabecera
                    Do_Events
                    Actualiza_Tabla_GtVend Mybase, Cabecera
                    Do_Events
                    actualiza_tabla_gtHora Mybase, Cabecera
                    Do_Events
                    Actualiza_Tabla_Gtsecs Mybase, Cabecera
                    Do_Events
                    actualiza_tabla_gtInfVen Mybase, Cabecera
                    Do_Events
                    Actualiza_Tabla_Log Mybase, Cabecera
                End If
            End If



            '''''''''''''''
            'marcar...
            'cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
            Cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
            MyConnObj.Execute Cm
            ''''''''''

            MyRecSetTmp.Update

            MyRecSetTmp.Movenext
        
            Call MonitorBarra
            Call frmControl.Balanza_ComOK
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirLeeTique
            Do_Events
            
        Loop
    End If

SalirLeeTique:
''''''''

    On Error Resume Next
    'myRec.Close
    'Set myRec = Nothing
    MyRecSetTmp.Close
    Set MyRecSetTmp = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
FIN:
    If Err.Number <> 0 Then
        CadenadeLog "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
    End If
    On Error GoTo 0

End Sub

'DROP TABLE IF EXISTS `pcscale`.`counters`;
'CREATE TABLE  `pcscale`.`counters` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` int(11) DEFAULT NULL,
'  `Name` varchar(64) NOT NULL,
'  `Description` varchar(127) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaDesSec()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nResp As Integer

'    nResp = ElPing(sAddIpTouch)
'    If nResp = 1 Then

    On Error GoTo FINDESSEC
    Set Mybase = DAO.OpenDatabase(Base_General)
    Set MyRecVar = Mybase.OpenRecordset("select * from seccion where borrado=false")
    On Error GoTo FINDESSEC
    If Not MyRecVar.EOF Then
        frmMonitor.barra1.Value = 0
        frmMonitor.barra1.Max = MyRecVar.Recordcount
        MyRecVar.MoveFirst
        Do While Not MyRecVar.EOF
            MyRecDes.Open "Select * from pcscale.counters where code=" & CStr(MyRecVar.Fields("secc_maqui")), MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            frmMonitor.MostrarInfo "Envio Touch Sección:" & CStr(MyRecVar.Fields("secc_maqui"))
            MyRecDes.Fields("code") = MyRecVar.Fields("secc_maqui")
            If Not IsNull(MyRecVar.Fields("nombre")) Then
                MyRecDes.Fields("name") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("nombre")))
            Else
                MyRecDes.Fields("name") = ConvertStringToUtf8Bytes("Sección: " & CStr(MyRecVar.Fields("secc_maqui")))
            End If
            If Not IsNull(MyRecVar.Fields("descripcio")) Then
                MyRecDes.Fields("description") = ConvertStringToUtf8Bytes(Trim(MyRecVar.Fields("descripcio")))
            Else
                MyRecDes.Fields("description") = ConvertStringToUtf8Bytes("---Sección " & CStr(MyRecVar.Fields("secc_maqui")) & "---")
            End If
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            MyRecVar.Movenext
            
            Call MonitorBarra
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirDesSec
            
        Loop
    
    
    End If
SalirDesSec:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINDESSEC:
    CadenadeLog "Error en MUE_EnviaDesSec:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
'    Else
'        canCelar = 1
'    End If
End Sub

'La idea de almacenamiento de datos se ha cambiado... Aquí fugura la Original.
' Como queda ahora sería la misma para Adaptador de protocolo, Alcampo y Sabeco...
'DROP TABLE IF EXISTS `pcscale`.`trademark`;
'CREATE TABLE  `pcscale`.`trademark` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `CO_Name` varchar(80) NOT NULL,          --> cab 1
'  `Description` varchar(80) NOT NULL,      --> cab 2
'  `Business` varchar(80) NOT NULL,         --> cab 3
'  `Slogan` varchar(80) NOT NULL,           --> cab 4
'  `Headquarter` varchar(80) NOT NULL,      --> cab 5
'  `CO_WEB` varchar(80) NOT NULL,           --> ley 1
'  `Shop_Name` varchar(80) NOT NULL,        --> ley 2
'  `Address` varchar(80) NOT NULL,          --> ley 3
'  `City` varchar(80) NOT NULL,             --> ley 4
'  `Phone` varchar(80) NOT NULL,            --> ley 5
'-------Solo Cabecera y Leyenda
'  `TIN` varchar(80) NOT NULL,              --> TxtF 1
'  `Shop_WEB` varchar(80) NOT NULL,         --> txtF 2
'  `Text1` varchar(80) NOT NULL,            --> txtF 3
'  `Text2` varchar(80) NOT NULL,            --> txtF 4
'  `Text3` varchar(80) NOT NULL,            --> txtF 5
'  `Text4` varchar(80) NOT NULL,            --> txtF 6
'  `Text5` varchar(80) NOT NULL,            --> txtF 7
'  `Text6` varchar(80) NOT NULL,            --> txtF 8
'''''''''''''''''''''''''''''''''
'  PRIMARY KEY (`Id`)
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
Public Sub MUE_EnviaCabLey()
Dim Mybase As DAO.Database
Dim MyRecVar As DAO.Recordset
Dim MyRecDes As New ADODB.Recordset
'Dim nResp As Integer
'    nResp = ElPing(sAddIpTouch)
'    If nResp = 1 Then

    On Error Resume Next
    Set Mybase = DAO.OpenDatabase(Base_General)
    If Err.Number <> 0 Then
        GoTo FINCABLEY
    End If
    On Error GoTo 0
    
    Set MyRecVar = Mybase.OpenRecordset("select * from trade")
    If Not MyRecVar.EOF Then
        'frmmonitor.barra1.Value = 0
        'frmmonitor.barra1.Max = MyRecVar.Recordcount
        'MyRecVar.MoveFirst
        'Do While Not MyRecVar.EOF
            
            MyRecDes.Open "Select * from pcscale.trademark", MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                'Call frmControl.Balanza_ComError(0)
                GoTo FINCABLEY
            End If
            If MyRecDes.EOF Then
                MyRecDes.AddNew
            End If
            
            frmMonitor.MostrarInfo "Envio Textos Empresa:"
            MyRecDes.Fields("co_name") = ConvertStringToUtf8Bytes(MyRecVar.Fields("company"))
            MyRecDes.Fields("shop_name") = ConvertStringToUtf8Bytes(MyRecVar.Fields("name"))
            MyRecDes.Fields("address") = ConvertStringToUtf8Bytes(MyRecVar.Fields("address"))
            MyRecDes.Fields("headquarter") = ConvertStringToUtf8Bytes(MyRecVar.Fields("quarter"))
            MyRecDes.Fields("co_web") = ConvertStringToUtf8Bytes(MyRecVar.Fields("web"))
            MyRecDes.Fields("shop_web") = ConvertStringToUtf8Bytes(MyRecVar.Fields("shopweb"))
            MyRecDes.Fields("description") = ConvertStringToUtf8Bytes(MyRecVar.Fields("description"))
            MyRecDes.Fields("slogan") = ConvertStringToUtf8Bytes(MyRecVar.Fields("slogan"))
            MyRecDes.Fields("city") = "" 'ConvertStringToUtf8Bytes(MyRecVar.Fields("leyenda04"))
            MyRecDes.Fields("phone") = "" 'ConvertStringToUtf8Bytes(MyRecVar.Fields("leyenda05"))
            
            MyRecVar.Close
            Set MyRecVar = Nothing
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
            
            'MyRecVar.Movenext
            
            'Call MonitorBarra
            'Call frmControl.Balanza_ComOK
            
            Do_Events
            
            If canCelar = 1 Or CancelHumano = True Then GoTo SalirCabLey
            
        'Loop
    
    
    End If
SalirCabLey:
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    Mybase.Close
    Set Mybase = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0
    Exit Sub
FINCABLEY:
    CadenadeLog "Error en MUE_EnviaCabLey:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error Resume Next
    MyRecVar.Close
    Set MyRecVar = Nothing
    MyRecDes.Close
    Set MyRecDes = Nothing
    Mybase.Close
    Set Mybase = Nothing
    On Error GoTo 0
    canCelar = 1
'    Else
'        canCelar = 1
'    End If
End Sub

'***********************************
'////////////////////////////
'TouchScale --> BASE MAESTRA
Public Sub AN_BaseMaestra(Modificaciones As Boolean, Optional ByVal lC As Boolean, Optional ByVal nEnv As Integer)
    Dim Bucle As Integer
    Dim Base As DAO.Database
    Dim SeBorra As Boolean
    Dim Beco As Long
    Dim Registro As DAO.Recordset
    Dim nEnviados As Integer
    
    On Error GoTo FIN

    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = False
    'If Not Taquion Then '1.7.2
        frmMonitor.Command1.Enabled = False
        frmMonitor.cmdmonitorstop.Enabled = True
    'End If
    frmMonitor.Label1 = ""
    
    'nEnviados = 0
    
    lCogeTiquet = False
    canCelar = 0
    CancelHumano = False
    
    MostrarMonitor
    
    MUE_LanzaBaseMaestra Modificaciones, nEnv
    nEnviados = 2

    '*****************************
    ' borrado de datos eliminados
    'SE REALIZARÁN EN ENVIO...
    '*****************************
    'Set Base = OpenDatabase(Base_General) '1.7.2 OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from articulo where borrado=true and tran_plu='*' and tran_pluSC10='*' and tran_en='*' and tran_el='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If (Not .EOF) Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from publi where borrado=true and tran_pub='*' and tran_pubSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from cabley where borrado=true and tran_cab='*' and tran_cabSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from codbar where borrado=true and tran_bar='*' and tran_barSC10='*' and tran_en='*' and tran_el='*' and tran_ep='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from vendedor where borrado=true and tran_vend='*' and tran_vendSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
    'Set Registro = Base.OpenRecordset _
    '("select * from teclas where borrado=true")
    ''2.0.30 ("select * from teclas where borrado=true and tran_tec='*' and tran_tecSC10='*'")
    'With Registro
    '    If Not .EOF Then
    '        .MoveFirst
    '        Do Until .EOF
    '            If Not .EOF Then Delete_Record Registro
    '            If Not .EOF Then .Movenext
    '        Loop
    '    End If
    'End With
FIN:
On Error GoTo 0

    On Error Resume Next
    'Registro.Close
    '1.7.2 CerrarBase Base
    Set Base = Nothing
    'Workspaces(0).close
    
    
    'If Not Taquion Then
        frmMonitor.Label1.Caption = CargaCadena(445)  '"Comunicación Finalizada"
        frmMonitor.cmdmonitorstop.Enabled = False
    'End If
    On Error GoTo 0
    HabilitarSalir
    lCogeTiquet = True
End Sub

'TouchScale -->MUE_LanzaBaseMaestra
'REVISADA VER COMO SE REALIZA FINALMENTE VACUNO...
Public Function MUE_LanzaBaseMaestra(Modificaciones As Boolean, Optional ByVal nEnv As Integer) As Boolean
'************************************************************************************
' envío de base maestra, realizando un ciclo por todas los sistemas existentes
'************************************************************************************
    Dim Bucle As Integer
    Dim Marcar As Boolean
    Dim ErrorPrevio As Boolean
    
    ErrorPrevio = False
    Marcar = False
    
    MUE_Sistemas
    
    For Bucle = 0 To SisEur.NSistemas - 1
        If MultiEuroscale Then
            frmControl.Switch_mode 700 + Bucle
        Else
            frmControl.Switch_mode 0
        End If
        If Bucle = SisEur.NSistemas - 1 Then
            If Not ErrorPrevio Then Marcar = True
        End If
        If CancelHumano Then Exit For
        canCelar = 0
        'touchScale Barna...
        'Incluir envio de Secciones (No Existe)
        'Incluir envio de Balanzas (No Existe)
        'Incluir envio de IVAS (No Existe)
        
        'Buscar "adaptar touchscale barna"
        'MUE_EnviaDesSec
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaDesSec
        'MUE_EnviaCabLey
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaCabLey
        ''''''''''''''''''''''''''''''''''
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 1) Then MUE_EnviaVendedor Bucle, Modificaciones
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 2) Then MUE_EnviaCLientes Bucle, Modificaciones, Marcar
        'se envian en artículos... If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 3) Then MUE_EnviaIVAS Bucle, Modificaciones, Marcar
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 4) Then MUE_EnviaCodigosBarras Bucle, Modificaciones
        ''If UsaFichasVacuno Then 'ACLARAR...
            'Los paises se capturan... If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 5) Then MUE_EnviarPaises Bucle, Modificaciones, Marcar
            If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 5) Then MUE_EnviarFichasVacuno Bucle, Modificaciones, Marcar
        ''End If
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 Then Call capturaPaises
        If canCelar = 0 And Not CancelHumano And EstadoBal = 0 And (nEnv = 0 Or nEnv = 7) Then MUE_EnviaArticuloNW Bucle, Modificaciones, Marcar
    
        'cerrar conexión
        On Error Resume Next
        MyConnObj.Close
        Set MyConnObj = Nothing
        On Error GoTo 0
        EstadoBal = 1000
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
        frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
        If CancelHumano Then frmMonitor.Label1 = "Comunicación suspendida por Usuario"
    
    Next Bucle

    
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'Adaptar TouchScale Barna
Public Sub BorraSrv()
Dim Cm As String
Dim MyRecDes As New ADODB.Recordset
On Error Resume Next
    
    CadenadeLog "BORRASRV..."
    
    On Error Resume Next
    CadenadeLog "pcscale.Barcodeean13"
    Cm = "delete from pcscale.Barcodeean13"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.Bovine"
    Cm = "delete from pcscale.Bovine"
    MyConnObj.Execute Cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "pcscale.customers"
    Cm = "delete from pcscale.customers"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.families"
    Cm = "delete from pcscale.families"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.htickets"
    Cm = "delete from pcscale.htickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.items"
    Cm = "delete from pcscale.items"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localerasetickets"
    Cm = "delete from pcscale.localerasetickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
      
    On Error Resume Next
    CadenadeLog "pcscale.localhtickets"
    Cm = "delete from pcscale.localhtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localtickets"
    Cm = "delete from pcscale.localtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpacketlabels"
    Cm = "delete from pcscale.localpacketlabels"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpacketpresets"
    Cm = "delete from pcscale.localpacketpresets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpaytickets"
    Cm = "delete from pcscale.localpaytickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localplukeyschanges"
    Cm = "delete from pcscale.localplukeyschanges"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localsservicelabels"
    Cm = "delete from pcscale.localsservicelabels"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localvendorchanges"
    Cm = "delete from pcscale.localvendorchanges"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.localvendorpacket"
    Cm = "delete from pcscale.localvendorpacket"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.ltickets"
    Cm = "delete from pcscale.ltickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    'c2f 09-08-2012
    'If ReiniciarNum <> 0 Then
        On Error Resume Next
        CadenadeLog "pcscale.erasetickets"
        Cm = "truncate pcscale.erasetickets"
        MyConnObj.Execute Cm
        On Error GoTo 0
        
        On Error Resume Next
        CadenadeLog "pcscale.localnumtickets"
        Cm = "truncate pcscale.localnumtickets"
        MyConnObj.Execute Cm
        On Error GoTo 0

        On Error Resume Next
        CadenadeLog "pcscale.numtickets"
        Cm = "truncate pcscale.numtickets"
        MyConnObj.Execute Cm
        On Error GoTo 0
    'End If
    ''''''''''''''''''
    
    'c2f 09-08-2012 esta tabla es de configuración no alterar...
    'On Error Resume Next
    'CadenadeLog "pcscale.numberingtickets"
    'cm = "delete from pcscale.numberingtickets"
    'MyConnObj.Execute cm
    'On Error GoTo 0
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    CadenadeLog "pcscale.packetlabels"
    Cm = "delete from pcscale.packetlabels"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.paymethods"
    Cm = "delete from pcscale.paymethods"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.paytickets"
    Cm = "delete from pcscale.paytickets"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.plukeys"
    Cm = "delete from pcscale.plukeys"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.qlevels"
    Cm = "delete from pcscale.qlevels"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.sservicelabels"
    Cm = "delete from pcscale.sservicelabels"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.trademark"
    Cm = "delete from pcscale.trademark"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.vendorkeys"
    Cm = "delete from pcscale.vendorkeys"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.vendors"
    Cm = "delete from pcscale.vendors"
    MyConnObj.Execute Cm
    On Error GoTo 0

    Cm = "REPLACE INTO Trademark VALUES(1,'','','','','','','','','','','','','','','','','','');"
    MyConnObj.Execute Cm
    
    'c2f 09-08-2012 --> Ver con corrección Balanza
    ' En principio dejo los código de barras necesarios Prueba ECI para el equipo 15
    ' ¡¡¡ QUITAR MÁS ADELANTE Y VER QUÉ CÓDIGOS DE BARRAS SE DEJAN POR DEFECTO...!!!
    'cm = "REPLACE INTO BarcodeEAN13 VALUES(1,'S',15, '2502215TTTTT', '26CCCCCiiiii', '2502215TTTTT', '26CCCCCiiiii', '26CCCCCiiiii', '25CCCCCiiiii', '2YTTTTTIIIII', '20VVViiiiiii', '10VVVIIIIIII', '2YTTTTTiiiii', '2CCCCCCiiiii', '2CCCCCCiiiii',  '2CCCCCCIIIII');"
    'MyConnObj.Execute cm
    '''''''
    
    
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
    End If

On Error GoTo 0

End Sub

Public Function ElPing(ElHost As String) As Integer
    '///////////////////////////////////////////////
    '//  Lanza ping a la dirección que se indique //
    '//  llamando al OCX de ping                  //
    '///////////////////////////////////////////////
    Dim Contador As Integer
    Dim Echo As ICMP_ECHO_REPLY
    Dim Resp As Long
    Resp = ICMP_Ping(ElHost, "alive", Echo)
    If Resp = 0 Then
        ElPing = 1
    Else
        ElPing = 0
    End If
End Function


Attribute VB_Name = "TouchScale"
Option Explicit
''''''''''''''''''''''''

'Recupera un código en función de Sección/PLU
Public Function GetCodeSC(ByVal nS As Integer, ByVal nP As Long) As Long
Dim MyRecAux As New ADODB.Recordset
    CadenadeLog "Localizar Code: Sección:" & CStr(nS) & " PLU:" & CStr(nP)
    MyRecAux.Open "select * from pcscale.items where counter=" & CStr(nS) & " and plu=" & CStr(nP), MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecAux.EOF Then
        GetCodeSC = MyRecAux.Fields("code")
    Else
        GetCodeSC = 0
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
End Function
'Envio de teclas directas:
'Lmode              ' 0 para programar a toda una sección, 1 para programar una balanza
'lTarget            ' número de sección o balanza a programar
'lTable             ' Tabla de teclas directas (de 0 a 9)
'lIdentifierType    ' 0, se va a indicar un PLU, 1, se va a indicar un código
'Identifier         ' PLU o Código, si vale cero, se borra la tecla.
'KeyNumber          ' número de tecla a programar.
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
Public Function Capa1A_SendKey_Touch(Lmode As Long, Target As Long, lTable As Long, lIdentifierType As Long, lIdentifier As Long, KeyNumber As Long) As Long
    '///////////////////////////////////////
    '// Envío de Teclas (TouchScale)      //
    '///////////////////////////////////////
    Dim MyOK As Boolean
    Dim MyBuf As String
    Dim Resp As Long
    Dim MiStr As String
    Dim Bucle As Integer
    Dim MyRecAux As New ADODB.Recordset
    Dim Cm As String
    Dim lCodeSC As Long
    Dim lLTarget As Long
    '''''''''''''''
    
'Envio de teclas directas:
'Lmode              ' 0 para programar a toda una sección, 1 para programar una balanza
'lTarget            ' número de sección o balanza a programar
'lTable             ' Tabla de teclas directas (de 0 a 9)
'lIdentifierType    ' 0, se va a indicar un PLU, 1, se va a indicar un código
'Identifier         ' PLU o Código, si vale cero, se borra la tecla.
'KeyNumber          ' número de tecla a programar.
    
    If Lmode = 1 Then lLTarget = Target + 10
    
    If Ocupado = False Then
        MyOK = True
        If (KeyNumber < 0) Or (KeyNumber > 999) Then MyOK = False
        If MyOK Then
            
            If lIdentifier = 0 Then 'baja
                Cm = "delete from pcscale.plukeys where section=" & CStr(lLTarget) & " and board=" & CStr(lTable) & " and plukey=" & CStr(KeyNumber)
                MyConnObj.Execute Cm
            Else
                If KeyNumber = 0 Then
                    Cm = "delete from pcscale.plukeys where code=" & CStr(lIdentifier)
                    MyConnObj.Execute Cm
                Else
                    On Error Resume Next
                    MyRecAux.Open "select * from pcscale.plukeys where section=" & CStr(lLTarget) & " and board=" & CStr(lTable) & " and plukey=" & CStr(KeyNumber), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                        GoTo FINSENDKEY
                    End If
                    On Error GoTo FINSENDKEY
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                    End If
                    
                    'If Lmode = 1 Then
                    If Target <> 0 Then
                        MyRecAux.Fields("section") = lLTarget
                    Else
                        MyRecAux.Fields("section") = 1
                    End If
                    MyRecAux.Fields("board") = lTable
                    MyRecAux.Fields("plukey") = KeyNumber
                    
                    'poner código de artículo se obtiene de la tabla Items con GetCodeSC
                    If lIdentifierType = 0 Then
                        lCodeSC = GetCodeSC(lLTarget, lIdentifier)
                        If lCodeSC = 0 Then
                            Resp = 1002
                            GoTo FINSENDKEY
                        End If
                        ''''''''''''''''
                        MyRecAux.Fields("code") = lCodeSC
                    Else
                        MyRecAux.Fields("code") = lIdentifier
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    MyRecAux.Update
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                End If
            End If
        Else
            Resp = 1002
        End If
    Else
        Resp = 1003
    End If
FINSENDKEY:
    If Err.Number <> 0 Then
        On Error GoTo 0
        Resp = 1000
    End If
    PonDato "Capa1A_SendKey_touch", Resp
    Capa1A_SendKey_Touch = Resp
End Function

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
'Nom_Cli : Name, max.255 caracteres ' Nombre del Cliente
'Data1 : Address max.255 caracteres ' Datos adicionales del cliente
'Data2 : City max.255 caracteres ' Datos adicionales del cliente
'Data3 : Web max.255 caracteres ' Datos adicionales del cliente
'Data4 : Phone max.255 caracteres ' Datos adicionales del cliente
'Erase_Client : Flag de borrado, si está a cero, se considera que se trata de una Alta, y si vale uno, un borrado, en cuyo caso los datos "Nom_Cli", "Data1", "Data2", "Data3" y "Data4" no son tenidos en cuenta.
Public Function Capa1A_SendClient_Touch(Client_Code As Long, Nom_Cli As String, Data1 As String, Data2 As String, Data3 As String, Data4 As String, Erase_Client As Long) As Long
    '///////////////////////////////////////
    '// Envío de Teclas (TouchScale)      //
    '///////////////////////////////////////
    Dim MyOK As Boolean
    Dim MyBuf As String
    Dim Resp As Long
    Dim MiStr As String
    Dim Bucle As Integer
    Dim MyRecAux As New ADODB.Recordset
    Dim Cm As String
    
    If Ocupado = False Then
        MyOK = True
        If Len(Nom_Cli) > 255 Or Len(Data1) > 255 Or Len(Data2) > 255 Or Len(Data3) > 255 Or Len(Data4) > 255 Then MyOK = False
        If Erase_Client <> 0 And Erase_Client <> 1 Then MyOK = False
        If MyOK Then
            
            If Erase_Client = 1 Then 'baja
                
                Cm = "delete from pcscale.customers where code=" & CStr(Client_Code)
                MyConnObj.Execute Cm
                
            Else
                On Error Resume Next
                MyRecAux.Open "select * from pcscale.customers where code=" & CStr(Client_Code), MyConnObj, adOpenKeyset, adLockOptimistic
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    GoTo FINSENDCLIENT
                End If
                On Error GoTo FINSENDCLIENT
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                End If
                
                MyRecAux.Fields("code") = Client_Code
                MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(Nom_Cli))
                MyRecAux.Fields("address") = ConvertStringToUtf8Bytes(Trim(Data1))
                MyRecAux.Fields("city") = ConvertStringToUtf8Bytes(Trim(Data2))
                MyRecAux.Fields("web") = ConvertStringToUtf8Bytes(Trim(Data3))
                MyRecAux.Fields("phone") = ConvertStringToUtf8Bytes(Trim(Data4))
                
                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
            End If
        Else
            Resp = 1002
        End If
    Else
        Resp = 1003
    End If
FINSENDCLIENT:
    If Err.Number <> 0 Then
        On Error GoTo 0
        Resp = 1000
    End If
    
    PonDato "Capa1A_SendClient_touch", Resp
    Capa1A_SendClient_Touch = Resp
End Function

'DROP TABLE IF EXISTS `pcscale`.`vats`;
'CREATE TABLE  `pcscale`.`vats` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `Code` decimal(1,0) unsigned NOT NULL,
'  `Name` varchar(16) NOT NULL,
'  `Percent` decimal(4,2) DEFAULT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
Public Function capa1A_sendVat_Touch(vat As Epel_Vat) As Long
'**************
Dim res As Long
Dim MyBuf As String
Dim MyOK As Boolean
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String
Dim nC As Integer
Dim nD As Double
'******************
    If Ocupado = False Then
        MyOK = True
        If vat.type1 < 0 Or vat.type1 > 9999 Then MyOK = False
        If vat.type2 < 0 Or vat.type2 > 9999 Then MyOK = False
        If vat.type3 < 0 Or vat.type3 > 9999 Then MyOK = False
        If vat.type4 < 0 Or vat.type4 > 9999 Then MyOK = False
        If vat.type5 < 0 Or vat.type5 > 9999 Then MyOK = False
        If MyOK Then
            On Error GoTo FINVAT
            Cm = "delete from pcscale.vats"
            MyConnObj.Execute Cm
            MyRecAux.Open "select * from pcscale.vats", MyConnObj, adOpenKeyset, adLockOptimistic
            For nC = 1 To 5
                MyRecAux.AddNew
                Select Case nC
                Case 1
                MyRecAux.Fields("code") = 0
                nD = vat.type1 / 100
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes("IVA " & Format(nD, "0.00")))
                MyRecAux.Fields("percent") = nD
                
                Case 2
                MyRecAux.Fields("code") = 1
                nD = vat.type2 / 100
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes("IVA " & Format(nD, "0.00")))
                MyRecAux.Fields("percent") = nD
                
                Case 3
                MyRecAux.Fields("code") = 2
                nD = vat.type3 / 100
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes("IVA " & Format(nD, "0.00")))
                MyRecAux.Fields("percent") = nD
                
                Case 4
                MyRecAux.Fields("code") = 3
                nD = vat.type4 / 100
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes("IVA " & Format(nD, "0.00")))
                MyRecAux.Fields("percent") = nD
                
                Case 5
                MyRecAux.Fields("code") = 4
                nD = vat.type5 / 100
                MyRecAux.Fields("name") = Trim(ConvertStringToUtf8Bytes("IVA " & Format(nD, "0.00")))
                MyRecAux.Fields("percent") = nD
                
                End Select
                MyRecAux.Update
            Next nC
            
            MyRecAux.Close
            Set MyRecAux = Nothing
        Else
            res = 1002
        End If
    Else
        res = 1003
    End If
FINVAT:
    If Err.Number <> 0 Then
        CadenadeLog "Err. en capa1A_sendvat_touch:" & Err.Description & " Code:" & CStr(Err.Number)
        res = 1000
    End If
    On Error GoTo 0
    capa1A_sendVat_Touch = res
End Function

'TouchScale
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
Public Function Capa1B_SendCard_Touch(card As Epel_Card) As Long
Dim MyRecDes As New ADODB.Recordset
Dim Cm As String
Dim Resp As Long

    ' *******************
    ' Envío de datos
    ' *******************
    If card.lBaja = False Then
        MyRecDes.Open "Select * from pcscale.bovine where numcard=" & CStr(card.Number), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            GoTo salir
        Else
            If MyRecDes.EOF Then

                MyRecDes.AddNew
                MyRecDes.Fields("numcard") = card.Number
                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(card.Identifier))
                MyRecDes.Fields("birthland") = card.BirthCode
                MyRecDes.Fields("birthdate") = Format(card.BirthDate, "yyyy-mm-dd")
                MyRecDes.Fields("breedingland") = card.BreedingCode
                MyRecDes.Fields("slaughterland") = card.SlaughterCode
                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(card.SlaughterReg))
                MyRecDes.Fields("SlaughterDate") = Format(card.SlaughterDate, "yyyy-mm-dd")
                MyRecDes.Fields("ButcherLand") = card.ButcheringCode
                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(card.ButcheringReg))
                
                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(card.SetNumber))
                    
            
            Else
                   
                MyRecDes.Fields("numcard") = card.Number
                MyRecDes.Fields("eartag") = ConvertStringToUtf8Bytes(Trim(card.Identifier))
                MyRecDes.Fields("birthland") = card.BirthCode
                MyRecDes.Fields("birthdate") = Format(card.BirthDate, "yyyy-mm-dd")
                MyRecDes.Fields("breedingland") = card.BreedingCode
                MyRecDes.Fields("slaughterland") = card.SlaughterCode
                MyRecDes.Fields("SlaughterReg") = ConvertStringToUtf8Bytes(Trim(card.SlaughterReg))
                MyRecDes.Fields("SlaughterDate") = Format(card.SlaughterDate, "yyyy-mm-dd")
                MyRecDes.Fields("ButcherLand") = card.ButcheringCode
                MyRecDes.Fields("ButcherReg") = ConvertStringToUtf8Bytes(Trim(card.ButcheringReg))
                
                MyRecDes.Fields("lot") = ConvertStringToUtf8Bytes(Trim(card.SetNumber))
               
            End If
            
            MyRecDes.Update
            MyRecDes.Close
            Set MyRecDes = Nothing
            
        End If
        
    Else
        'Borrar
        Cm = "delete from pcscale.bovine where numcard=" & CStr(card.Number)
        MyConnObj.Execute Cm
        ''''''''''''''''''''
    End If
                    
    
salir:
    Resp = 0
    If Err.Number <> 0 Then
        Resp = 1000
        On Error GoTo 0
    End If
    Capa1B_SendCard_Touch = Resp
End Function

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

Public Function Capa1A_SendItem_Touch(Item As Epel_Item) As Long
Dim nD As Double
Dim MyRecAux As New ADODB.Recordset
Dim lEsAlta As Boolean
Dim Resp As Long

    
            On Error Resume Next
            MyRecAux.Open "select * from pcscale.items where code=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                GoTo FINARTICULO
            End If
            On Error GoTo FINARTICULO
            lEsAlta = False
            If MyRecAux.EOF Then
                MyRecAux.AddNew
                lEsAlta = True
            End If
            
            MyRecAux.Fields("code") = Item.Code
            MyRecAux.Fields("shortname") = ConvertStringToUtf8Bytes(Trim(Item.Text))
            MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(Item.Text))
            If lEsAlta Then MyRecAux.Fields("text") = ""
            If lEsAlta Then MyRecAux.Fields("text1") = ""
            If lEsAlta Then MyRecAux.Fields("text2") = ""
            If lEsAlta Then MyRecAux.Fields("text3") = ""
            If lEsAlta Then MyRecAux.Fields("text4") = ""
            If lEsAlta Then MyRecAux.Fields("text5") = ""
            MyRecAux.Fields("family") = Item.Family
            MyRecAux.Fields("price") = Item.Price / 100
            MyRecAux.Fields("saleform") = Item.weight
            MyRecAux.Fields("plu") = Item.Plu
            MyRecAux.Fields("counter") = Item.Sec
            MyRecAux.Fields("vat") = 0
            MyRecAux.Fields("discount") = 0
            If lEsAlta Then MyRecAux.Fields("qlevel") = ""
            MyRecAux.Fields("ean13") = ""
            MyRecAux.Fields("expiring") = Item.Caducity
            MyRecAux.Fields("bestbefore") = Item.Pref
            MyRecAux.Fields("tare") = Item.Tare / 1000
            If Item.Etq = 255 Then Item.Etq = 0
            Select Case Item.Etq
            Case 1
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("1.xml")
            Case 2
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("2.xml")
            Case 3
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("3.xml")
            Case 4
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("4.xml")
            Case 5
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("5.xml")
            Case 6
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("6.xml")
            Case 7
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("7.xml")
            Case 8
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("8.xml")
            Case 9
                MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("9.xml")
            
            Case Else
                MyRecAux.Fields("label") = ""
            End Select
                        
            If lEsAlta Then MyRecAux.Fields("icon") = Item.sImagen
            If lEsAlta Then MyRecAux.Fields("mixed") = 0
            If lEsAlta Then MyRecAux.Fields("mixedweight") = 0
            
            MyRecAux.Update
            
            MyRecAux.Close
            Set MyRecAux = Nothing

FINARTICULO:
    Resp = 0
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_SendItem_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1A_SendItem_Touch = Resp
End Function


' Envio de descriptivos
Public Function Capa1B_SendDesItem_Touch(fItem As Epel_FullItem) As Long
Dim MyRecAux As New ADODB.Recordset
Dim Resp As Long
Dim ncont As Integer
Dim Cm As String
Dim s(5) As String

    Resp = 0
    
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.items where code=" & CStr(fItem.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FinDesF
    End If
    On Error GoTo FinDesF

    If MyRecAux.EOF Then
        Resp = 1004
        GoTo FinDesF
    End If

    Cm = ""
    ' Descriptivos 1 a 20
    For ncont = 0 To 19
        Cm = Cm & Trim(fItem.Desc(ncont))
    Next ncont
    ' Descriptivos 21 a 40
    For ncont = 0 To 19
        Cm = Cm & Trim(fItem.Desc2040(ncont))
    Next ncont
        
    MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(Mid(Cm, 1, 1024))
        
    On Error Resume Next
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    On Error GoTo FinDesF

FinDesF:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_SendDesItem_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_SendDesItem_Touch = Resp
End Function

' Envío articulo completo (TouchScale)
Public Function Capa1B_SendFullItem_Touch(FullItem As Epel_FullItem) As Long
Dim nD As Double
Dim MyRecAux As New ADODB.Recordset
Dim lEsAlta As Boolean
Dim Resp As Long
Dim ncont As Integer
Dim Cm As String
Dim s(5) As String
   
    Resp = 0
   
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.items where code=" & CStr(FullItem.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FinArticuloF
    End If
    On Error GoTo FinArticuloF
    lEsAlta = False
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        lEsAlta = True
    End If
    
    MyRecAux.Fields("code") = FullItem.Code
    MyRecAux.Fields("shortname") = ConvertStringToUtf8Bytes(Trim(FullItem.Text))
    MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(FullItem.Text))
    ' Descriptivo Texto Largo
    Cm = ""
    ' Descriptivos 1 a 20
    For ncont = 0 To 19
        Cm = Cm & Trim(FullItem.Desc(ncont))
    Next ncont
    ' Descriptivos 21 a 40
    For ncont = 0 To 19
        Cm = Cm & Trim(FullItem.Desc2040(ncont))
    Next ncont
    If Trim(Cm) <> "" Then
        MyRecAux.Fields("text") = ConvertStringToUtf8Bytes(Mid(Cm, 1, 1024))
    Else
        If lEsAlta Then MyRecAux.Fields("text") = ""
    End If
    If lEsAlta Then MyRecAux.Fields("text1") = ""
    If lEsAlta Then MyRecAux.Fields("text2") = ""
    If lEsAlta Then MyRecAux.Fields("text3") = ""
    If lEsAlta Then MyRecAux.Fields("text4") = ""
    If lEsAlta Then MyRecAux.Fields("text5") = ""
    
    MyRecAux.Fields("family") = FullItem.Family
    MyRecAux.Fields("price") = FullItem.Price / 100
    MyRecAux.Fields("saleform") = FullItem.weight
    MyRecAux.Fields("plu") = FullItem.Plu
    MyRecAux.Fields("counter") = FullItem.Sec
    MyRecAux.Fields("vat") = FullItem.vat
    MyRecAux.Fields("discount") = 0
    If lEsAlta Then MyRecAux.Fields("qlevel") = ""
    MyRecAux.Fields("ean13") = ConvertStringToUtf8Bytes(FullItem.EAN13)
    MyRecAux.Fields("expiring") = FullItem.Caducity
    MyRecAux.Fields("bestbefore") = FullItem.Pref
    MyRecAux.Fields("tare") = FullItem.Tare / 1000
    
    If FullItem.Etq = 255 Then FullItem.Etq = 0
    Select Case FullItem.Etq
    Case 1
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("1.xml")
    Case 2
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("2.xml")
    Case 3
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("3.xml")
    Case 4
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("4.xml")
    Case 5
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("5.xml")
    Case 6
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("6.xml")
    Case 7
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("7.xml")
    Case 8
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("8.xml")
    Case 9
        MyRecAux.Fields("label") = ConvertStringToUtf8Bytes("9.xml")
    
    Case Else
        MyRecAux.Fields("label") = ""
    End Select
    
    If lEsAlta Then MyRecAux.Fields("icon") = FullItem.sImagen
    
    If FullItem.lMix = True Then
        If lEsAlta Then MyRecAux.Fields("mixed") = 1
    Else
        If lEsAlta Then MyRecAux.Fields("mixed") = 0
    End If
    If lEsAlta Then MyRecAux.Fields("mixedweight") = FullItem.nPeso / 1000
    
    MyRecAux.Update
    
    MyRecAux.Close
    Set MyRecAux = Nothing

FinArticuloF:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_SendFullItem_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_SendFullItem_Touch = Resp
End Function



Public Function Capa1B_ConsNumTqtEci(ByRef mNumTQT As Long, ByRef nBalMar) As Long
Dim Resp As Long
    On Error Resume Next
    Resp = 0
    
    MyRecSet.Open "select * from pcscale.htickets where ((captured='' or isnull(captured)) and status='C') order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
        mNumTQT = MyRecSet.Fields("id")
        nBalMar = MyRecSet.Fields("scale")

    Else
        mNumTQT = 0
        nBalMar = 0

    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing

FinConsultaF:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_ConsNumTqtEci:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_ConsNumTqtEci = Resp
End Function

Public Function Capa1B_ConsNumTqtEci_Local(ByRef mNumTQT As Long, ByRef nBalMar) As Long
Dim Resp As Long
    On Error Resume Next
    Resp = 0
    
    MyRecSet.Open "select * from pcscale.localhtickets where ((captured='' or isnull(captured)) and status='C') order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
        mNumTQT = MyRecSet.Fields("id")
        nBalMar = MyRecSet.Fields("scale")

    Else
        mNumTQT = 0
        nBalMar = 0

    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing

FinConsultaFLocal:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_ConsNumTqtEci_Local:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_ConsNumTqtEci_Local = Resp
End Function


Public Function Capa1B_ConsEqpCon(ByRef sCadEquipos As String) As Long
Dim Resp As Long
Dim mNumEqp As Integer
    On Error Resume Next
    Resp = 0
    sCadEquipos = "00000000000000000000" & "00000000000000000000" & "00000000000000000000" & "00000000000000000000" & "00000000000000000000" & "00000000000000000000"
    
    MyRecSet.Open "select * from pcscale.scales where (number<61 and number>0) order by number", MyConnObj, adOpenKeyset, adLockOptimistic
    
    If Not MyRecSet.EOF Then
        MyRecSet.MoveFirst
        
        Do While Not MyRecSet.EOF
            mNumEqp = 0
            mNumEqp = MyRecSet.Fields("number")
            
            If mNumEqp > 0 And mNumEqp < 61 Then
                Select Case mNumEqp
                    Case 1
                        sCadEquipos = "01" & Mid(sCadEquipos, 3, 58)
                    
                    Case 60
                        sCadEquipos = Mid(sCadEquipos, 1, 58) & "60"
                    
                    Case Else
                        sCadEquipos = Mid(sCadEquipos, 1, (mNumEqp - 1) * 2) & Format(mNumEqp, "00") & Mid(sCadEquipos, (mNumEqp * 2) + 1, 60 - (mNumEqp * 2))
                    
                End Select
            End If
            
            MyRecSet.Movenext
        Loop
    
        MyRecSet.Close
        Set MyRecSet = Nothing
    End If
End Function

Public Function Capa1B_MarNumTqtEci(ByVal mNumTQT As Long, ByVal nBalMar As Long) As Long
Dim Cm As String
Dim Resp As Long
    On Error Resume Next
    Resp = 0
    Cm = "update pcscale.htickets set captured='*' where (id=" & CStr(mNumTQT) & " and scale=" & CStr(nBalMar) & ")"
    CadenadeLog "marcado:" & Cm
    MyConnObj.Execute Cm

FinMarcaF:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_MarNumTqtEci:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_MarNumTqtEci = Resp
End Function

Public Function Capa1B_MarNumTqtEci_Local(ByVal mNumTQT As Long, ByVal nBalMar As Long) As Long
Dim Cm As String
Dim Resp As Long
    On Error Resume Next
    Resp = 0
    Cm = "update pcscale.localhtickets set captured='*' where (id=" & CStr(mNumTQT) & " and scale=" & CStr(nBalMar) & ")"
    CadenadeLog "marcado Local:" & Cm
    MyConnObj.Execute Cm

FinMarcaFLocal:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_MarNumTqtEci_Local:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_MarNumTqtEci_Local = Resp
End Function

' Consulta de articulo completo
Public Function Capa1B_ConsFullItem_Touch(ByVal mMetodo As Integer, ByVal mCodigo As Long, ByVal mProtDesc As Integer, ByRef mFullItem As Epel_FullItem) As Long
Dim Resp As Long
Dim ncont As Integer
Dim Cm As String
Dim Cierra As Integer
Dim MyRecAux As New ADODB.Recordset
    Resp = 0
    Cierra = 0
    On Error Resume Next
    Select Case mMetodo
        Case 1  ' Igual
            MyRecAux.Open "select * from pcscale.items where code=" & CStr(mCodigo), MyConnObj, adOpenKeyset, adLockOptimistic
        
        Case 3  ' Mayor o Igual
            MyRecAux.Open "select * from pcscale.items where code>=" & CStr(mCodigo) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
        
        Case Else
            Resp = 1003
            GoTo SalFunCons
    End Select
    If Err.Number <> 0 Then
        GoTo FinArticuloC
    End If
    On Error GoTo FinArticuloC
    Cierra = 1
    If MyRecAux.EOF Then
        Resp = 1003
        GoTo SalFunCons
    End If
    
    mFullItem.Code = MyRecAux.Fields("code")
    mFullItem.Text = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("name"))))
    If Len(mFullItem.Text) > 25 Then mFullItem.Text = Mid(mFullItem.Text, 1, 25)
    mFullItem.Family = MyRecAux.Fields("family")
    mFullItem.Price = MyRecAux.Fields("price") * 100
    mFullItem.weight = MyRecAux.Fields("saleform")
    mFullItem.vat = MyRecAux.Fields("vat")
    mFullItem.Plu = MyRecAux.Fields("plu")
    mFullItem.Sec = MyRecAux.Fields("counter")
    mFullItem.Caducity = MyRecAux.Fields("expiring")
    mFullItem.Tare = MyRecAux.Fields("tare") * 1000
    mFullItem.Pref = MyRecAux.Fields("bestbefore")
    mFullItem.EAN13 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("ean13"))))
    mFullItem.sImagen = MyRecAux.Fields("icon")
    If MyRecAux.Fields("mixed") = 1 Then
        mFullItem.lMix = True
    Else
        mFullItem.lMix = False
    End If
    mFullItem.nPeso = MyRecAux.Fields("mixedweight") * 1000
    Select Case MyRecAux.Fields("label")
        Case "1.xml"
            mFullItem.Etq = 1
        Case "1.xml"
            mFullItem.Etq = 2
        Case "3.xml"
            mFullItem.Etq = 3
        Case "4.xml"
            mFullItem.Etq = 4
        Case "5.xml"
            mFullItem.Etq = 5
        Case "6.xml"
            mFullItem.Etq = 6
        Case "7.xml"
            mFullItem.Etq = 7
        Case "8.xml"
            mFullItem.Etq = 8
        Case "9.xml"
            mFullItem.Etq = 9
        
        Case Else
            mFullItem.Etq = 0
    End Select
    mFullItem.Loss = 0
    mFullItem.Presel = 0
    mFullItem.subsec = 0
    mFullItem.Font = 0
    ncont = 0
    For ncont = 0 To 9
        mFullItem.FDesc(ncont) = 0
    Next
    
    ' Descriptivos
    Cm = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("text"))))
    ' 1 a 20
    For ncont = 1 To 19
        If Len(Cm) > 25 Then
            mFullItem.Desc(ncont) = Mid(Cm, 1, 25)
            Cm = Mid(Cm, 26, Len(Cm) - 25)
        Else
            mFullItem.Desc(ncont) = Cm
            Cm = ""
        End If
    Next
    ' 21 a 40
    For ncont = 0 To 19
        If Len(Cm) > 25 Then
            mFullItem.Desc2040(ncont) = Mid(Cm, 1, 25)
            Cm = Mid(Cm, 26, Len(Cm) - 25)
        Else
            mFullItem.Desc2040(ncont) = Cm
            Cm = ""
        End If
    Next
        

FinArticuloC:
    If Cierra = 1 Then
        MyRecAux.Close
        Set MyRecAux = Nothing
        Cierra = 0
    End If
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_ConsFullItem_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
SalFunCons:
    If Cierra = 1 Then
        MyRecAux.Close
        Set MyRecAux = Nothing
        Cierra = 0
    End If
    On Error GoTo 0
    Capa1B_ConsFullItem_Touch = Resp
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

' Peticion de tique (TouchScale)
Public Function Capa1B_QueryTicket_Touch(ByVal lType As Long, Cabecera As Epel_CabTick, ByVal bNoMarcar As Integer) As Long
    Dim RespTqt As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim ncont As Long
    Dim dLImporte As Double
    Dim NoHayTqt As Boolean
    Dim fTipoPago As Long
    NoHayTqt = False
    RespTqt = 0
    On Error Resume Next
   
    'HydraTouch Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    MyRecSet.Open "select * from pcscale.htickets where ((captured='' or isnull(captured)) and status='C') order by id", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Ffecha = Format(MyRecSet.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
    CadenadeLog Ffecha
        Fhora = Mid(Ffecha, 12, 8)
        Ffecha = Mid(Ffecha, 1, 10)
        IdDoc = MyRecSet.Fields("id")
        ''''medio de pago
        '1=efectivo
        '2=t.crédito
        '3=cheque
        fTipoPago = 1
        Cabecera.cantidadEntre = 0
        MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            fTipoPago = MyRecAux.Fields("paymethod")
            Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
        End If
        Select Case fTipoPago
            Case 1
                Cabecera.tipoPago = 0
            Case 2
                Cabecera.tipoPago = 2
            Case 3
                Cabecera.tipoPago = 1
            Case Else
                Cabecera.tipoPago = 4
        End Select
        MyRecAux.Close
        Set MyRecAux = Nothing
        '''''''''''''''''
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        ReDim Cabecera.Lines(nNumLin)
            
        MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype = 1", MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = MyRecAux.Fields("importe")
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        'recoger datos
        Cabecera.mIdDoc = IdDoc
        Cabecera.NTicket = MyRecSet.Fields("number")
        Cabecera.Vendor = MyRecSet.Fields("vendor")
        Cabecera.Section = MyRecSet.Fields("counter")
        Cabecera.Client = MyRecSet.Fields("customer")
        Cabecera.date = Format(Ffecha, "dd/mm/yyyy")
        Cabecera.Hour = Format(Fhora, "hh:mm:ss")
        Cabecera.Machine = MyRecSet.Fields("scale")
        Cabecera.amount = dLImporte
        Cabecera.nLines = nNumLin
        
        Cabecera.EAN13 = MyRecSet.Fields("ean13")
            
        ' Ticket normal o encargo
        If MyRecSet.Fields("tcktype") = "N" Then
            Cabecera.lEncargo = False
        Else
            Cabecera.lEncargo = True
        End If
        
        ' Ticket local o red
        If MyRecSet.Fields("NetStat") = "N" Then
            Cabecera.type = 1
        Else
            Cabecera.type = 0
        End If
            
        MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = Round(MyRecAux.Fields("importe"), 2)
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        Cabecera.P1 = dLImporte 'MyRecSet.Fields("totalbruto")
        Cabecera.P2 = Cabecera.P1 - Cabecera.amount  '.Fields("totaldescuento")
                       
        Cabecera.discount = MyRecSet.Fields("tckDiscount") ' Porcentaje
        
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        Cabecera.nLinCanceled = nNumLin
            
        'LINEAS
        nNumLin = 0
        MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
                nNumLin = nNumLin + 1
                If MyRecAux.Fields("saleform") = 1 Then
                    Cabecera.Lines(nNumLin).Weighted = True
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                Else
                    Cabecera.Lines(nNumLin).Weighted = False
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                End If
                Cabecera.Lines(nNumLin).Item_Price = MyRecAux.Fields("price")
                    
                Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("item")
                Cabecera.Lines(nNumLin).discount = MyRecAux.Fields("linediscount") 'porcentaje
                Cabecera.Lines(nNumLin).amount = MyRecAux.Fields("amount")
                If Cabecera.Lines(nNumLin).amount > 0 Then
                    Cabecera.Lines(nNumLin).positive = 1
                Else
                    Cabecera.Lines(nNumLin).positive = 0
                End If
                If MyRecAux.Fields("linetype") = 0 Then
                    Cabecera.Lines(nNumLin).Cancelled = True
                Else
                    Cabecera.Lines(nNumLin).Cancelled = False
                End If
                    
                Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("eartag")
                Cabecera.Lines(nNumLin).Number = nNumLin
                
                ' Modo de linea --> 1 Por codigo, 0 por Seccion - PLU
                Cabecera.Lines(nNumLin).ModeLine = 2  ' El 2 indica que se tome el del configurador
                
                MyRecAux.Movenext
            Loop
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If

        '''''''''''''''
        'marcar...
        If bNoMarcar = 0 Then
            Cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
            MyConnObj.Execute Cm
        End If
        ''''''''''

        Do_Events
    Else
        NoHayTqt = True
    End If

SalirLeeTique:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
FIN:
    On Error GoTo 0
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryTicket_Touch. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespTqt = 1004
    Else
        If NoHayTqt = True Then
            RespTqt = 4
        Else
            RespTqt = 0
        End If
    End If
    Capa1B_QueryTicket_Touch = RespTqt
End Function

' obtener PLU en función Código de artículo...
Public Function myPLU(ByVal nC As Long) As Long
Dim MyRecAux As New ADODB.Recordset
    CadenadeLog "Localizar plu del code=" & CStr(nC)
    MyRecAux.Open "select * from pcscale.items where code=" & CStr(nC), MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecAux.EOF Then
        myPLU = MyRecAux.Fields("plu")
    Else
        myPLU = 0
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
End Function
' Peticion de tique (TouchScale)
Public Function Capa1B_QueryNumTicket_Touch(ByVal lIdDoc As Long, Cabecera As Epel_CabTick, ByVal bNoMarcar As Integer) As Long
    Dim RespTqt As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim ncont As Long
    Dim dLImporte As Double
    Dim NoHayTqt As Boolean
    Dim fTipoPago As Long
    NoHayTqt = False
    RespTqt = 0
    On Error Resume Next
   
    'HydraTouch Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    MyRecSet.Open "select * from pcscale.htickets where id=" & lIdDoc, MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Ffecha = Format(MyRecSet.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
    CadenadeLog Ffecha
        Fhora = Mid(Ffecha, 12, 8)
        Ffecha = Mid(Ffecha, 1, 10)
        IdDoc = MyRecSet.Fields("id")
        ''''medio de pago
        '1=efectivo
        '2=t.crédito
        '3=cheque
        fTipoPago = 1
        Cabecera.cantidadEntre = 0
        MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            fTipoPago = MyRecAux.Fields("paymethod")
            Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
        End If
        Select Case fTipoPago
            Case 1
                Cabecera.tipoPago = 0
            Case 2
                Cabecera.tipoPago = 2
            Case 3
                Cabecera.tipoPago = 1
            Case Else
                Cabecera.tipoPago = 4
        End Select
        MyRecAux.Close
        Set MyRecAux = Nothing
        
        '''''''''''''''''
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        ReDim Cabecera.Lines(nNumLin)
            
        MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype = 1", MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = MyRecAux.Fields("importe")
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        'recoger datos
        Cabecera.mIdDoc = IdDoc
        Cabecera.NTicket = MyRecSet.Fields("number")
        Cabecera.Vendor = MyRecSet.Fields("vendor")
        Cabecera.Section = MyRecSet.Fields("counter")
        Cabecera.Client = MyRecSet.Fields("customer")
        Cabecera.date = Format(Ffecha, "dd/mm/yyyy")
        Cabecera.Hour = Format(Fhora, "hh:mm:ss")
        Cabecera.Machine = MyRecSet.Fields("scale")
        Cabecera.amount = dLImporte
        Cabecera.nLines = nNumLin
        
        Cabecera.EAN13 = MyRecSet.Fields("ean13")
            
        ' Ticket normal o encargo
        If MyRecSet.Fields("tcktype") = "N" Then
            Cabecera.lEncargo = False
        Else
            Cabecera.lEncargo = True
        End If
        
        ' Ticket local o red
        If MyRecSet.Fields("NetStat") = "N" Then
            Cabecera.type = 1
        Else
            Cabecera.type = 0
        End If
            
        MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = Round(MyRecAux.Fields("importe"), 2)
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        Cabecera.P1 = dLImporte 'MyRecSet.Fields("totalbruto")
        Cabecera.P2 = Cabecera.P1 - Cabecera.amount  '.Fields("totaldescuento")
                       
        Cabecera.discount = MyRecSet.Fields("tckDiscount") ' Porcentaje
        
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        Cabecera.nLinCanceled = nNumLin
            
        'LINEAS
        nNumLin = 0
        MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
                nNumLin = nNumLin + 1
                If MyRecAux.Fields("saleform") = 1 Then
                    Cabecera.Lines(nNumLin).Weighted = True
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                Else
                    Cabecera.Lines(nNumLin).Weighted = False
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                End If
                Cabecera.Lines(nNumLin).Item_Price = MyRecAux.Fields("price")
                    
                If Dir(App.Path & "\byplu") <> "" Then
                    If MyRecAux.Fields("item") = 0 Then
                        CadenadeLog "PLU Genérico... (Código=0), se deja PLU=0."
                        Cabecera.Lines(nNumLin).item_code = 0
                    Else
                        'CadenadeLog "Conversión a PLU del código:" & CStr(MyRecAux.Fields("item")) & " --> " & CStr(myPLU(MyRecAux.Fields("item")))
                        Cabecera.Lines(nNumLin).item_code = myPLU(MyRecAux.Fields("item"))
                    End If
                Else
                    Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("item")
                End If
                '''''''''''''''
                
                Cabecera.Lines(nNumLin).discount = MyRecAux.Fields("linediscount") 'porcentaje
                Cabecera.Lines(nNumLin).amount = MyRecAux.Fields("amount")  'MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                If Cabecera.Lines(nNumLin).amount > 0 Then
                    Cabecera.Lines(nNumLin).positive = 1
                Else
                    Cabecera.Lines(nNumLin).positive = 0
                End If
                If MyRecAux.Fields("linetype") = 0 Then
                    Cabecera.Lines(nNumLin).Cancelled = True
                Else
                    Cabecera.Lines(nNumLin).Cancelled = False
                End If
                    
                Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("eartag")
                Cabecera.Lines(nNumLin).Number = nNumLin
                
                ' Modo de linea --> 1 Por codigo, 0 por Seccion - PLU
                Cabecera.Lines(nNumLin).ModeLine = 2  ' El 2 indica que se tome el del configurador
                
                MyRecAux.Movenext
            Loop
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If

        '''''''''''''''
        'marcar...
        If bNoMarcar = 0 Then
            Cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
            MyConnObj.Execute Cm
        End If
        ''''''''''

        Do_Events
    Else
        NoHayTqt = True
    End If

SalirLeeTique:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
FIN:
    On Error GoTo 0
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryNumTicket_Touch. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespTqt = 1004
    Else
        If NoHayTqt = True Then
            RespTqt = 4
        Else
            RespTqt = 0
        End If
    End If
    Capa1B_QueryNumTicket_Touch = RespTqt
End Function

' Peticion de tique (TouchScale)
Public Function Capa1B_QueryNumTicket_Touch_Local(ByVal lIdDoc As Long, Cabecera As Epel_CabTick, ByVal bNoMarcar As Integer) As Long
    Dim RespTqt As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim Cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim ncont As Long
    Dim dLImporte As Double
    Dim NoHayTqt As Boolean
    Dim fTipoPago As Long
    NoHayTqt = False
    RespTqt = 0
    On Error Resume Next
   
    'HydraTouch Tipos de Tiquet (typtic)
    ' 1 --> Rollo
    ' 2 --> Super
    ' 3 --> Envasado
    ' 4 --> Autoservicio
    ' 5 --> Rollo Local
    ' 6 --> Super Local
    ' 7 --> Envasado Local
    ' 8 --> Autoservicio Local
    MyRecSet.Open "select * from pcscale.localhtickets where id=" & lIdDoc, MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        Ffecha = Format(MyRecSet.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
    CadenadeLog Ffecha
        Fhora = Mid(Ffecha, 12, 8)
        Ffecha = Mid(Ffecha, 1, 10)
        IdDoc = MyRecSet.Fields("id")
        ''''medio de pago
        '1=efectivo
        '2=t.crédito
        '3=cheque
        fTipoPago = 1
        Cabecera.cantidadEntre = 0
        MyRecAux.Open "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            fTipoPago = MyRecAux.Fields("paymethod")
            Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
        End If
        Select Case fTipoPago
            Case 1
                Cabecera.tipoPago = 0
            Case 2
                Cabecera.tipoPago = 2
            Case 3
                Cabecera.tipoPago = 1
            Case Else
                Cabecera.tipoPago = 4
        End Select
        MyRecAux.Close
        Set MyRecAux = Nothing
        
        '''''''''''''''''
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        ReDim Cabecera.Lines(nNumLin)
            
        MyRecAux.Open "select sum(amount) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc) & " and linetype = 1", MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = MyRecAux.Fields("importe")
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        'recoger datos
        Cabecera.mIdDoc = IdDoc
        Cabecera.NTicket = MyRecSet.Fields("number")
        Cabecera.Vendor = MyRecSet.Fields("vendor")
        Cabecera.Section = MyRecSet.Fields("counter")
        Cabecera.Client = MyRecSet.Fields("customer")
        Cabecera.date = Format(Ffecha, "dd/mm/yyyy")
        Cabecera.Hour = Format(Fhora, "hh:mm:ss")
        Cabecera.Machine = MyRecSet.Fields("scale")
        Cabecera.amount = dLImporte
        Cabecera.nLines = nNumLin
        
        Cabecera.EAN13 = MyRecSet.Fields("ean13")
            
        ' Ticket normal o encargo
        If MyRecSet.Fields("tcktype") = "N" Then
            Cabecera.lEncargo = False
        Else
            Cabecera.lEncargo = True
        End If
        
        ' Ticket local o red
        If MyRecSet.Fields("NetStat") = "N" Then
            Cabecera.type = 1
        Else
            Cabecera.type = 0
        End If
            
        MyRecAux.Open "select sum(price*weight) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        dLImporte = Round(MyRecAux.Fields("importe"), 2)
        MyRecAux.Close
        Set MyRecAux = Nothing
            
        Cabecera.P1 = dLImporte 'MyRecSet.Fields("totalbruto")
        Cabecera.P2 = Cabecera.P1 - Cabecera.amount  '.Fields("totaldescuento")
                       
        Cabecera.discount = MyRecSet.Fields("tckDiscount") ' Porcentaje
        
        MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
        nNumLin = MyRecAux.Fields("lineas")
        MyRecAux.Close
        Set MyRecAux = Nothing
        Cabecera.nLinCanceled = nNumLin
            
        'LINEAS
        nNumLin = 0
        MyRecAux.Open "select * from pcscale.localltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            MyRecAux.MoveFirst
            Do While Not MyRecAux.EOF
                nNumLin = nNumLin + 1
                If MyRecAux.Fields("saleform") = 1 Then
                    Cabecera.Lines(nNumLin).Weighted = True
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                Else
                    Cabecera.Lines(nNumLin).Weighted = False
                    Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("weight")
                End If
                Cabecera.Lines(nNumLin).Item_Price = MyRecAux.Fields("price")
                    
                If Dir(App.Path & "\byplu") <> "" Then
                    If MyRecAux.Fields("item") = 0 Then
                        CadenadeLog "PLU Genérico... (Código=0), se deja PLU=0."
                        Cabecera.Lines(nNumLin).item_code = 0
                    Else
                        'CadenadeLog "Conversión a PLU del código:" & CStr(MyRecAux.Fields("item")) & " --> " & CStr(myPLU(MyRecAux.Fields("item")))
                        Cabecera.Lines(nNumLin).item_code = myPLU(MyRecAux.Fields("item"))
                    End If
                Else
                    Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("item")
                End If
                '''''''''''''''
                
                Cabecera.Lines(nNumLin).discount = MyRecAux.Fields("linediscount") 'porcentaje
                Cabecera.Lines(nNumLin).amount = MyRecAux.Fields("amount")  'MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                If Cabecera.Lines(nNumLin).amount > 0 Then
                    Cabecera.Lines(nNumLin).positive = 1
                Else
                    Cabecera.Lines(nNumLin).positive = 0
                End If
                If MyRecAux.Fields("linetype") = 0 Then
                    Cabecera.Lines(nNumLin).Cancelled = True
                Else
                    Cabecera.Lines(nNumLin).Cancelled = False
                End If
                    
                Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("eartag")
                Cabecera.Lines(nNumLin).Number = nNumLin
                
                ' Modo de linea --> 1 Por codigo, 0 por Seccion - PLU
                Cabecera.Lines(nNumLin).ModeLine = 2  ' El 2 indica que se tome el del configurador
                
                MyRecAux.Movenext
            Loop
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If

        '''''''''''''''
        'marcar...
        If bNoMarcar = 0 Then
            Cm = "update pcscale.localhtickets set captured='*' where id=" & CStr(IdDoc)
            MyConnObj.Execute Cm
        End If
        ''''''''''

        Do_Events
    Else
        NoHayTqt = True
    End If

SalirLeeTique:

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
FIN:
    On Error GoTo 0
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryNumTicket_Touch_Local. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespTqt = 1004
    Else
        If NoHayTqt = True Then
            RespTqt = 4
        Else
            RespTqt = 0
        End If
    End If
    Capa1B_QueryNumTicket_Touch_Local = RespTqt
End Function

Public Function Capa1B_MarcarTqt_Touch(ByVal IdDoc As Long) As Long
Dim Cm As String
On Error Resume Next
    Cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
    MyConnObj.Execute Cm
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        Capa1B_MarcarTqt_Touch = 1003
    Else
        Capa1B_MarcarTqt_Touch = 0
    End If
End Function

' Borrado de articulos (TouchScale)
Public Function Capa1B_Kill_All_Item_Touch() As Long
Dim Cm As String
Dim RespKAI As Long
    On Error Resume Next

    Cm = "delete from pcscale.items"
    MyConnObj.Execute Cm

    'On Error GoTo 0
    
    If Err.Number <> 0 Then
        CadenadeLog "Error en Kill_All_Item_Touch. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespKAI = 1010
    Else
        RespKAI = 0
    End If
    Capa1B_Kill_All_Item_Touch = RespKAI
End Function

' Borrado de vendedores (TouchScale)
Public Function Capa1B_Kill_All_Vendor_Touch() As Long
Dim Cm As String
Dim RespKAV As Long
    On Error Resume Next

    Cm = "delete from pcscale.vendors"
    MyConnObj.Execute Cm

    'On Error GoTo 0
    
    If Err.Number <> 0 Then
        CadenadeLog "Error en Kill_All_Vendor_Touch. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespKAV = 1010
    Else
        RespKAV = 0
    End If
    Capa1B_Kill_All_Vendor_Touch = RespKAV
End Function

' Borrado teclas de PLU (TouchScale)
Public Function Capa1B_Kill_All_PluKeys_Touch() As Long
Dim Cm As String
Dim RespKAP As Long
    On Error Resume Next

    Cm = "delete from pcscale.plukeys"
    MyConnObj.Execute Cm

    'On Error GoTo 0
    
    If Err.Number <> 0 Then
        CadenadeLog "Error en Kill_All_PluKeys_Touch. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        RespKAP = 1010
    Else
        RespKAP = 0
    End If
    Capa1B_Kill_All_PluKeys_Touch = RespKAP
End Function


'DROP TABLE IF EXISTS `pcscale`.`trademark`;
'CREATE TABLE `pcscale`.`trademark` (
'`Id` int(11) NOT NULL AUTO_INCREMENT,
'`CO_Name` varchar(80) NOT NULL,
'`Description` varchar(80) NOT NULL,
'`Business` varchar(80) NOT NULL,
'`Slogan` varchar(80) NOT NULL,
'`Headquarter` varchar(80) NOT NULL,
'`CO_WEB` varchar(80) NOT NULL,
'`Shop_Name` varchar(80) NOT NULL,
'`Address` varchar(80) NOT NULL,
'`City` varchar(80) NOT NULL,
'`Phone` varchar(80) NOT NULL,
'`TIN` varchar(80) NOT NULL,
'`Shop_WEB` varchar(80) NOT NULL,
'`Text1` varchar(80) NOT NULL,
'`Text2` varchar(80) NOT NULL,
'`Text3` varchar(80) NOT NULL,
'`Text4` varchar(80) NOT NULL,
'`Text5` varchar(80) NOT NULL,
'`Text6` varchar(80) NOT NULL,
'PRIMARY KEY (`Id`)
') ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;
'DROP TABLE IF EXISTS `pcscale`.`counters`;
'CREATE TABLE `pcscale`.`counters` (
'`Id` int(11) NOT NULL AUTO_INCREMENT,
'`Code` int(11) DEFAULT NULL,
'`Name` varchar(64) NOT NULL,
'`Description` varchar(127) NOT NULL,
'PRIMARY KEY (`Id`),
'UNIQUE KEY `Code` (`Code`)
') ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;

' Envío cabeceras y leyendas (TouchScale)
Public Function Capa1A_SendHeadingLegend_Touch(HeadingLegend As Epel_HeadingLegend) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim sMiText As String

' Mode = 0 --> Cabeceras, 1 Leyendas, 2 Textos Libres
' Number (Solo Textos libres) --> 0 (1 a 5), 1 (6 a 10)
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.trademark", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FINCABLEY
    End If
    On Error GoTo FINCABLEY
    
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("co_name") = ""
        MyRecAux.Fields("description") = ""
        MyRecAux.Fields("business") = ""
        MyRecAux.Fields("slogan") = ""
        MyRecAux.Fields("headquarter") = ""
        MyRecAux.Fields("co_web") = ""
        MyRecAux.Fields("shop_name") = ""
        MyRecAux.Fields("address") = ""
        MyRecAux.Fields("city") = ""
        MyRecAux.Fields("phone") = ""
        MyRecAux.Fields("tin") = ""
        MyRecAux.Fields("shop_web") = ""
        MyRecAux.Fields("text1") = ""
        MyRecAux.Fields("text2") = ""
        MyRecAux.Fields("text3") = ""
        MyRecAux.Fields("text4") = ""
        MyRecAux.Fields("text5") = ""
        MyRecAux.Fields("text6") = ""
        CadenadeLog "Nuevo item"
    End If
    
    MyRecAux.Fields("co_name") = ConvertStringToUtf8Bytes(HeadingLegend.Heading3)
    MyRecAux.Fields("shop_name") = ConvertStringToUtf8Bytes(HeadingLegend.Heading4)
    MyRecAux.Fields("address") = ConvertStringToUtf8Bytes(HeadingLegend.Heading5)
    MyRecAux.Fields("headquarter") = ConvertStringToUtf8Bytes(HeadingLegend.Legend1)
    MyRecAux.Fields("co_web") = ConvertStringToUtf8Bytes(HeadingLegend.Legend2)
    MyRecAux.Fields("shop_web") = ConvertStringToUtf8Bytes(HeadingLegend.Legend3)
    MyRecAux.Fields("description") = ConvertStringToUtf8Bytes(HeadingLegend.Legend4)
    MyRecAux.Fields("slogan") = ConvertStringToUtf8Bytes(HeadingLegend.Legend5)
    MyRecAux.Fields("city") = ""
    MyRecAux.Fields("phone") = ""
            
    MyRecAux.Update
    
    MyRecAux.Close
    Set MyRecAux = Nothing

    On Error Resume Next
    MyRecAux.Open "select * from pcscale.counters where code =" & CStr(HeadingLegend.Number), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FINCABLEY
    End If
    On Error GoTo FINCABLEY
    
    If MyRecAux.EOF Then
        Resp = 1004
    Else
        MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(HeadingLegend.Heading1))
        MyRecAux.Fields("description") = ConvertStringToUtf8Bytes(Trim(HeadingLegend.Heading2))
        Resp = 0
        MyRecAux.Update
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    
FINCABLEY:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_SendHeadingLegend_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1A_SendHeadingLegend_Touch = Resp
End Function

' Borrado de vendedores para TouchScale
Public Function Capa1A_SendDelVendor_Touch(ByVal IDV As Long) As Long
Dim Resp As Long
Dim Cm As String
    Resp = 0
    
    ' Borrado de vendedor
    On Error Resume Next
    If IDV <> 0 Then
        Cm = "delete from pcscale.vendors where code=" & CStr(IDV)
    Else
        ' Borrado de todos los vendedores
        Cm = "delete from pcscale.vendors"
    End If
    MyConnObj.Execute Cm
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinDelVnd
    End If
    On Error GoTo FinDelVnd
                 
FinDelVnd:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_SendDelVendor_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1A_SendDelVendor_Touch = Resp
End Function

' Envío de vendedores (TouchScale)
Public Function Capa1A_SendVendor_Touch(IDV As Long, sName As String, nSection As Long, nKey As Long, sImgVnd As String, Status As Long, EraseVendor As Long) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String
    Resp = 0
    If Status = 1 Then
        ' Programacion de Tecla de vendedor
        On Error Resume Next
        Cm = "delete from pcscale.vendorkeys where counter=" & CStr(nSection) & " and vendorkey=" & CStr(nKey)
        MyConnObj.Execute Cm
        If Err.Number <> 0 Then
            Resp = 1003
            GoTo FinVnd
        End If
        On Error GoTo FinVnd
                 
        On Error Resume Next
        Cm = "delete from pcscale.vendorkeys where vendor=" & CStr(IDV)
        MyConnObj.Execute Cm
        If Err.Number <> 0 Then
            Resp = 1003
            GoTo FinVnd
        End If
        On Error GoTo FinVnd
                 
        On Error Resume Next
        CadenadeLog "Programacion de tecla (" & nKey & ") de vendedor " & IDV & " de la seccion " & nSection
        MyRecAux.Open "Select * from pcscale.vendorkeys where vendor=" & CStr(IDV), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1003
            GoTo FinVnd
        End If
        On Error GoTo FinVnd
    
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        End If
        
        MyRecAux.Fields("vendor") = IDV
        MyRecAux.Fields("vendorkey") = nKey
        MyRecAux.Fields("counter") = nSection
        
        
        MyRecAux.Update
        MyRecAux.Close
        Set MyRecAux = Nothing
    Else
        If Status = 0 Then
            ' Programacion de Vendedor
            On Error Resume Next
            
            CadenadeLog "Programacion de Vendedor " & IDV & " - " & Trim(sName)
            MyRecAux.Open "Select * from pcscale.vendors where code=" & CStr(IDV), MyConnObj, adOpenKeyset, adLockOptimistic
            
            If Err.Number <> 0 Then
                Resp = 1003
                GoTo FinVnd
            End If
            On Error GoTo FinVnd
                       
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            
            MyRecAux.Fields("code") = IDV
            MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(sName))
            
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
        Else
            Resp = 1003
            GoTo FinVnd
        End If
    End If
FinVnd:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_SendVendor_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1A_SendVendor_Touch = Resp
End Function

' Envío de paises (TouchScale)
Public Function Capa1B_SendCountry_Touch(ByVal CountryCode As Long, ByVal CountryName As String, ByVal EraseCountry As Long) As Long
Dim Resp As Long
Dim Cm As String
Dim MyRecAux As New ADODB.Recordset
    Resp = 0
    
    If EraseCountry = 1 Then
        ' Borrado de pais
        On Error Resume Next
        CadenadeLog "Borrado de pais: " & CStr(CountryCode) & " - " & Trim(CountryName)
        Cm = "delete from pcscale.localcountries where code=" & CStr(CountryCode)
        MyConnObj.Execute Cm
        If Err.Number <> 0 Then
            Resp = 1003
            GoTo FinCountry
        End If
        On Error GoTo FinCountry
    Else
        ' Programacion de Pais
        On Error Resume Next
        
        CadenadeLog "Programacion de Pais " & CStr(CountryCode) & " - " & Trim(CountryName)
        MyRecAux.Open "Select * from pcscale.localcountries where code=" & CStr(CountryCode), MyConnObj, adOpenKeyset, adLockOptimistic
        
        If Err.Number <> 0 Then
            Resp = 1003
            GoTo FinCountry
        End If
        On Error GoTo FinCountry
                   
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        End If
        
        MyRecAux.Fields("code") = CountryCode
        MyRecAux.Fields("alpha2") = ConvertStringToUtf8Bytes(Trim(Mid(CountryName, 1, 2)))
        MyRecAux.Fields("name") = ConvertStringToUtf8Bytes(Trim(CountryName))
        
        MyRecAux.Update
        MyRecAux.Close
        Set MyRecAux = Nothing
    
    End If
FinCountry:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_SendCountry_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1B_SendCountry_Touch = Resp
End Function

' Eliminar Articulo (TouchScale)
Public Function Capa1B_EliArt_Touch(ByVal Mode As Long, ByVal Code As Long) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String
    Resp = 0
    On Error Resume Next
    If Mode = 0 Then    ' Por codigo
        Cm = "delete from pcscale.items where code=" & CStr(Code)
    Else
        Cm = "delete from pcscale.items where plu=" & CStr(Code) & " and counter=" & CStr(Mode)
    End If
    MyConnObj.Execute Cm
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinArt
    End If
    On Error GoTo FinArt
FinArt:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_EliArt_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1B_EliArt_Touch = Resp
End Function

' Borrar ficha vacuno (TouchScale)
Public Function Capa1B_Erase_Card_Touch(ByVal cCode As Long) As Long
Dim Resp As Long
Dim Cm As String
Dim MyRecAux As New ADODB.Recordset
    Resp = 0
    
    On Error Resume Next
    Cm = "delete from pcscale.localcountries where code=" & CStr(cCode)
    MyConnObj.Execute Cm
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinCard
    End If
    On Error GoTo FinCard
FinCard:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_Erase_Card_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1B_Erase_Card_Touch = Resp
End Function

' Peticion de Cod. Barras (TouchScale)
Public Function Capa1B_Query_BarCode_Touch(ByRef Bcode As Epel_BarCode, ByVal Lmode As Long) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
    Resp = 0
    On Error Resume Next
    If Lmode = 1 Then
        ' Consulta por equipo
        MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(Bcode.Number), MyConnObj, adOpenKeyset, adLockOptimistic
    Else
        ' Consulta por seccion
        MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(Bcode.Number), MyConnObj, adOpenKeyset, adLockOptimistic
    End If
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinCBarras
    End If
    On Error GoTo FinCBarras
    
    If MyRecAux.EOF Then
        Resp = 1010
        GoTo FinCBarras
    End If

    Bcode.Number = MyRecAux.Fields("number")
    ' Red
    Bcode.Net_Sale = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("networksales"))))
    Bcode.Net_Super = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("NetworkLineSLB"))))
    Bcode.Net_Mix = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("NetworkTotalSLB"))))
    Bcode.Net_Pack1 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("NetworkPacket1"))))
    Bcode.Net_Pack2 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("NetworkPacket2"))))
    ' Local
    Bcode.Local_Sale = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localsales"))))
    Bcode.Local_Super = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localLineSLB"))))
    Bcode.Local_Mix = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localTotalSLB"))))
    Bcode.Local_Pack1 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localPacket1"))))
    Bcode.Local_Pack2 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localPacket2"))))
    Bcode.Local_Total1 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("localtotal"))))

    MyRecAux.Close
    Set MyRecAux = Nothing

FinCBarras:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_Query_BarCode_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1B_Query_BarCode_Touch = Resp
End Function

' conversión de I a i
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

' Envío de codigos de barras (TouchScale)
Public Function Capa1A_SendBarCode_Touch(BarCode As Epel_BarCode) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset

    On Error Resume Next
    If BarCode.Mode = 1 Then
        ' Equipo
        CadenadeLog "Consulta Equipo " & CStr(BarCode.Number)
        MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='S' and number=" & CStr(BarCode.Number), MyConnObj, adOpenKeyset, adLockOptimistic
    Else
        ' Seccion
        CadenadeLog "Consulta Seccion " & CStr(BarCode.Number)
        MyRecAux.Open "Select * from pcscale.barcodeean13 where numbertype='C' and number=" & CStr(BarCode.Number), MyConnObj, adOpenKeyset, adLockOptimistic
    End If
    
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinCBarras
    End If
    On Error GoTo FinCBarras

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
       
    If BarCode.Mode = 1 Then
        ' Balanza
        MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("S")
    Else
        ' Seccion
        MyRecAux.Fields("numbertype") = ConvertStringToUtf8Bytes("C")
    End If
    
    
    If BarCode.Red = 1 Then
        ' Red
        MyRecAux.Fields("number") = BarCode.Number
        MyRecAux.Fields("networksales") = ConvertStringToUtf8Bytes(Trim(BarCode.Net_Sale))
        MyRecAux.Fields("NetworkLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
        MyRecAux.Fields("NetworkTotalSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Sale)))
        MyRecAux.Fields("NetworkSS") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
        MyRecAux.Fields("NetworkPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
        MyRecAux.Fields("NetworkPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack2)))
        CadenadeLog "Red:" & BarCode.Net_Sale & "|" & BarCode.Net_Super & "|" & BarCode.Net_Mix & "|" & BarCode.Net_Pack1 & "|" & BarCode.Net_Pack2
    Else
        If BarCode.Red = 0 Then
            ' Local
            MyRecAux.Fields("number") = BarCode.Number
            MyRecAux.Fields("localsales") = ConvertStringToUtf8Bytes(Trim(BarCode.Local_Sale))
            MyRecAux.Fields("localLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack1)))
            MyRecAux.Fields("localTotalSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Sale)))
            MyRecAux.Fields("localSS") = ConvertStringToUtf8Bytes(Trim(Ii(BarCode.Local_Pack1)))
            MyRecAux.Fields("localPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack1)))
            MyRecAux.Fields("localPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack2)))
            MyRecAux.Fields("localtotal") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Total1)))
            CadenadeLog "Local:" & BarCode.Local_Sale & "|" & BarCode.Local_Super & "|" & BarCode.Local_Mix & "|" & BarCode.Local_Pack1 & "|" & BarCode.Local_Pack2 & "|" & BarCode.Local_Total1
        Else
            ' Se programa Red y Local
            MyRecAux.Fields("number") = BarCode.Number
            MyRecAux.Fields("networksales") = ConvertStringToUtf8Bytes(Trim(BarCode.Net_Sale))
            MyRecAux.Fields("NetworkLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
            MyRecAux.Fields("NetworkTotalSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Sale)))
            MyRecAux.Fields("NetworkSS") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
            MyRecAux.Fields("NetworkPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack1)))
            MyRecAux.Fields("NetworkPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Net_Pack2)))
            MyRecAux.Fields("localsales") = ConvertStringToUtf8Bytes(Trim(BarCode.Local_Sale))
            MyRecAux.Fields("localLineSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack1)))
            MyRecAux.Fields("localTotalSLB") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Sale)))
            MyRecAux.Fields("localSS") = ConvertStringToUtf8Bytes(Trim(Ii(BarCode.Local_Pack1)))
            MyRecAux.Fields("localPacket1") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack1)))
            MyRecAux.Fields("localPacket2") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Pack2)))
            MyRecAux.Fields("localtotal") = ConvertStringToUtf8Bytes(Ii(Trim(BarCode.Local_Total1)))
            CadenadeLog "Local:" & BarCode.Local_Sale & "|" & BarCode.Local_Super & "|" & BarCode.Local_Mix & "|" & BarCode.Local_Pack1 & "|" & BarCode.Local_Pack2 & "|" & BarCode.Local_Total1 & _
                " --- Red:" & BarCode.Net_Sale & "|" & BarCode.Net_Super & "|" & BarCode.Net_Mix & "|" & BarCode.Net_Pack1 & "|" & BarCode.Net_Pack2
        End If
    End If
    
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing

FinCBarras:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_SendBarCode_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    
    Capa1A_SendBarCode_Touch = Resp
End Function

' Peticion de articulo (TouchScale)
Public Function Capa1B_Query_Item_Touch(ByRef FullItem As Epel_FullItem, ByVal Lmode As Long) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String
Dim ncont As Long
    Resp = 0
    On Error Resume Next
    MyRecAux.Open "select * from pcscale.items where code=" & CStr(Lmode), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FinArticuloF
    End If
    On Error GoTo FinArticuloF
    
    If MyRecAux.EOF Then
        Resp = 1010
        GoTo FinArticuloF
    End If

    FullItem.Code = MyRecAux.Fields("code")
    FullItem.Text = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("name"))))
    FullItem.Family = MyRecAux.Fields("family")
    FullItem.Price = MyRecAux.Fields("price") * 100
    FullItem.weight = MyRecAux.Fields("saleform")
    FullItem.Plu = MyRecAux.Fields("plu")
    FullItem.Sec = MyRecAux.Fields("counter")
    FullItem.vat = MyRecAux.Fields("vat")
    FullItem.EAN13 = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("ean13"))))
    FullItem.Caducity = MyRecAux.Fields("expiring")
    FullItem.Pref = MyRecAux.Fields("bestbefore")
    FullItem.Tare = MyRecAux.Fields("tare") * 1000
    
    Select Case MyRecAux.Fields("label")
        Case "1.xml"
            FullItem.Etq = 1
        Case "2.xml"
            FullItem.Etq = 2
        Case "3.xml"
            FullItem.Etq = 3
        Case "4.xml"
            FullItem.Etq = 4
        Case "5.xml"
            FullItem.Etq = 5
        Case "6.xml"
            FullItem.Etq = 6
        Case "7.xml"
            FullItem.Etq = 7
        Case "8.xml"
            FullItem.Etq = 8
        Case "9.xml"
            FullItem.Etq = 9
        
        Case Else
            FullItem.Etq = 0
    End Select
    
    If MyRecAux.Fields("mixed") = 1 Then
        FullItem.lMix = True
    Else
        FullItem.lMix = False
    End If
    FullItem.nPeso = MyRecAux.Fields("mixedweight") * 1000
    
    ' Descriptivo Texto Largo
    Cm = Trim(ConvertUtf8BytesToString(DecodeBase64(MyRecAux.Fields("text"))))
    For ncont = 0 To 19
        If Len(Cm) > 25 Then
            FullItem.Desc2040(ncont) = Mid(Cm, 1, 25)
            Cm = Mid(Cm, 26, Len(Cm) - 25)
        Else
            If Len(Cm) = 0 Then
                FullItem.Desc2040(ncont) = ""
            Else
                FullItem.Desc2040(ncont) = Cm
                Cm = ""
            End If
        End If
    Next ncont
    
    MyRecAux.Close
    Set MyRecAux = Nothing

FinArticuloF:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1B_Query_Item_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1B_Query_Item_Touch = Resp
End Function

' Borrado de totales (TouchScale)
Public Function Capa1A_GA_Erase_Totals_Touch(ByVal No_Reinit As Long) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim Cm As String
    Resp = 0
    
    On Error Resume Next
    Cm = "delete from pcscale.htickets"
    If No_Reinit = 1 Then
        ' Borra solo los tickets capturados
        Cm = Cm & " where captured=*"
    End If
    MyConnObj.Execute Cm
    If Err.Number <> 0 Then
        Resp = 1003
        GoTo FinErase
    End If
    On Error GoTo FinErase
FinErase:
    If Err.Number <> 0 Then
        Resp = 1000
        CadenadeLog "Error en Capa1A_GA_Erase_Totals_Touch:" & CStr(Err.Number) & " Descr.:" & Err.Description
    End If
    On Error GoTo 0
    Capa1A_GA_Erase_Totals_Touch = Resp
End Function

Public Function BorraSrv() As Long
Dim Resp As Long
Dim Cm As String
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
    CadenadeLog "pcscale.counters"
    Cm = "delete from pcscale.counters"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.customers"
    Cm = "delete from pcscale.customers"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.erasetickets"
    Cm = "delete from pcscale.erasetickets"
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
    CadenadeLog "pcscale.localnumtickets"
    Cm = "delete from pcscale.localnumtickets"
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
    
    On Error Resume Next
    CadenadeLog "pcscale.numberingtickets"
    Cm = "delete from pcscale.numberingtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.numtickets"
    Cm = "delete from pcscale.numtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0

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
    CadenadeLog "pcscale.scales"
    Cm = "delete from pcscale.scales"
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
    
   
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    BorraSrv = Resp
End Function

' Borrado General
Public Function EraseGeneral_Touch() As Long
Dim Resp As Long
Dim Cm As String
Dim ReiniciarNum As Integer

On Error Resume Next
    
    ReiniciarNum = 0
    If Dir(App.Path & "\reiniciarnum") <> "" Then
        ReiniciarNum = 69
    End If
    
    CadenadeLog "Borrado General..."
    
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
    
    If ReiniciarNum <> 0 Then
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
    End If
    ''''''''''''''''''
    
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
    
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseGeneral_Touch = Resp
End Function
        
' Codigos de barras
Public Function EraseEAN_Touch() As Long
Dim Cm As String
Dim Resp As Long
On Error Resume Next
    
    CadenadeLog "Borrado Codigos de Barras..."
    
    On Error Resume Next
    CadenadeLog "pcscale.Barcodeean13"
    Cm = "delete from pcscale.Barcodeean13"
    MyConnObj.Execute Cm
    On Error GoTo 0
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseEAN_Touch = Resp
End Function
        
' Vendedores
Public Function EraseVendors_Touch() As Long
Dim Resp As Long
Dim Cm As String
    CadenadeLog "Borrado Vendedores..."
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
    CadenadeLog "pcscale.vendorkeys"
    Cm = "delete from pcscale.vendorkeys"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.vendors"
    Cm = "delete from pcscale.vendors"
    MyConnObj.Execute Cm
    On Error GoTo 0

    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseVendors_Touch = Resp
End Function
        
' Teclas
Public Function EraseKeys_Touch() As Long
Dim Resp As Long
Dim Cm As String
    CadenadeLog "Borrado Teclas..."

    On Error Resume Next
    CadenadeLog "pcscale.localplukeyschanges"
    Cm = "delete from pcscale.localplukeyschanges"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.plukeys"
    Cm = "delete from pcscale.plukeys"
    MyConnObj.Execute Cm
    On Error GoTo 0

    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseKeys_Touch = Resp
End Function
        
' Clientes
Public Function EraseClients_Touch() As Long
Dim Resp As Long
Dim Cm As String
    CadenadeLog "Borrado Clientes..."

    On Error Resume Next
    CadenadeLog "pcscale.customers"
    Cm = "delete from pcscale.customers"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseClients_Touch = Resp
End Function

' Borrado Totales
Public Function EraseTotals_Touch(ByVal ReiniciarNum As Integer) As Long
Dim Resp As Long
Dim Cm As String
    CadenadeLog "Borrado de Totales..."
    
    On Error Resume Next
    CadenadeLog "pcscale.htickets"
    Cm = "truncate pcscale.htickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.ltickets"
    Cm = "truncate pcscale.ltickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localerasetickets"
    Cm = "truncate pcscale.localerasetickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
      
    On Error Resume Next
    CadenadeLog "pcscale.localhtickets"
    Cm = "truncate pcscale.localhtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localtickets"
    Cm = "truncate pcscale.localtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
        
    On Error Resume Next
    CadenadeLog "pcscale.localpacketlabels"
    Cm = "truncate pcscale.localpacketlabels"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpacketpresets"
    Cm = "truncate pcscale.localpacketpresets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localpaytickets"
    Cm = "truncate pcscale.localpaytickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.localsservicelabels"
    Cm = "truncate pcscale.localsservicelabels"
    MyConnObj.Execute Cm
    On Error GoTo 0
       
    On Error Resume Next
    CadenadeLog "pcscale.numberingtickets"
    Cm = "truncate pcscale.numberingtickets"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.packetlabels"
    Cm = "truncate pcscale.packetlabels"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.paymethods"
    Cm = "truncate pcscale.paymethods"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.paytickets"
    Cm = "truncate pcscale.paytickets"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    On Error Resume Next
    CadenadeLog "pcscale.sservicelabels"
    Cm = "truncate pcscale.sservicelabels"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    If ReiniciarNum <> 0 Then
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
    End If
    
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0

    EraseTotals_Touch = Resp
End Function
        
' Borrado Articulos
Public Function ErasePlus_Touch() As Long
Dim Resp As Long
Dim Cm As String

    CadenadeLog "Borrado Plus..."
    On Error Resume Next
    CadenadeLog "pcscale.items"
    Cm = "delete from pcscale.items"
    MyConnObj.Execute Cm
    On Error GoTo 0

    On Error Resume Next
    CadenadeLog "pcscale.qlevels"
    Cm = "delete from pcscale.qlevels"
    MyConnObj.Execute Cm
    On Error GoTo 0
    
    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    ErasePlus_Touch = Resp
End Function
        
' Textos
Public Function EraseTexts_Touch() As Long
Dim Resp As Long
Dim Cm As String
    CadenadeLog "Borrado Textos..."

    On Error Resume Next
    CadenadeLog "pcscale.trademark"
    Cm = "delete from pcscale.trademark"
    MyConnObj.Execute Cm
    On Error GoTo 0

    Cm = "REPLACE INTO Trademark VALUES(1,'','','','','','','','','','','','','','','','','','');"
    MyConnObj.Execute Cm

    Resp = 0
    If Err.Number <> 0 Then
        CadenadeLog "KILL --> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        Resp = 1000
    End If

On Error GoTo 0
    EraseTexts_Touch = Resp
End Function


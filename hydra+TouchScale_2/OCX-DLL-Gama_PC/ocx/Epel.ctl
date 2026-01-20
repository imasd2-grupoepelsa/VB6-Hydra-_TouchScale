VERSION 5.00
Begin VB.UserControl Epel 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   Picture         =   "Epel.ctx":0000
   ScaleHeight     =   660
   ScaleWidth      =   660
End
Attribute VB_Name = "Epel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**************
' Estructuras
'************
Public WithEvents RecibeEvento As CEventoBas
Attribute RecibeEvento.VB_VarHelpID = -1
Public Event TicketReceived(ByVal NTicket As Long, ByVal nSection As Long, ByVal NScale As Long)
'***********
Public Event ItemReceived(ByVal TotalID As Long, ByVal lData As Long)
'***********
Public Event ComError(ByVal Error_Code As Long)
Public Event ComOK()
'*******************

'///////////////
'// Artículos //
'///////////////
Public Type Epel_Item
    Code As Long
    Plu As Long
    Sec As Long
    subsec As Long
    Price As Double
    Family As Long
    weight As Long
    Caducity As Long
    Font As Long
    Text As String
    Tare As Long
    Pref As Long
    Etq As Long
    sImagen As String
End Type
'///////////////////////
'// Artículo completo //
'///////////////////////
Public Type Epel_FullItem
    Code As Long
    Plu As Long
    Sec As Long
    subsec As Long
    Family As Long
    Price As Long
    weight As Long
    Caducity As Long
    Font As Long
    Text As String
    Tare As Long
    Pref As Long
    Etq As Long
    vat As Long
    EAN13 As String
    Desc(20) As String
    FDesc(10) As Long
    Loss As Long
    Presel As Long
    Desc2040(20) As String
    lMix As Boolean
    nPeso As Long
    sImagen As String
    ProtDesc As String
End Type

Public Type Epel_tecla
    Lmode As Long
    lTarget As Long
    lTable As Long
    lIdentifierType As Long
    lIdentifier As Long
    KeyNumber As Long
End Type

Public Type Epel_Cliente
    Client_Code As Long
    Nom_Cli As String
    Data1 As String
    Data2 As String
    Data3 As String
    Data4 As String
    Erase_Client As Long
End Type

Public Type Epel_BarCode
    '///////////////////////////////
    '// Códigos de Barras         //
    '///////////////////////////////
    Mode As Long
    Number As Long
    Red As Integer
    Net_Sale As String
    Net_Super As String
    Net_Mix As String
    Net_Pack1 As String
    Net_Pack2 As String
    Net_Total1 As String
    Net_Total2 As String
    Local_Sale As String
    Local_Super As String
    Local_Mix As String
    Local_Pack1 As String
    Local_Pack2 As String
    Local_Total1 As String
    Local_Total2 As String
End Type

Public Type Epel_HeadingLegend
    '///////////////////////////////
    '// Cabeceras y Leyendas      //
    '///////////////////////////////
    Mode As Long
    Number As Long
    Heading1 As String
    Heading2 As String
    Heading3 As String
    Heading4 As String
    Heading5 As String
    Legend1 As String
    Legend2 As String
    Legend3 As String
    Legend4 As String
    Legend5 As String
End Type

Public Type Epel_Card
    Number As Long
    Identifier As String
    BirthCode As Long
    BreedingCode As Long
    SlaughterCode As Long
    ButcheringCode As Long
    ProductionCode As Long
    SlaughterReg As String
    ButcheringReg As String
    Category As Long
    Race As Long
    SlaughterDate As String
    Age As Long
    Sex As String
    FreeText As String
    BirthDate As String
    ButcheringDate As String
    SetNumber As String
    weight As Long
    lBaja As Boolean
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type Epel_LinTick
    '///////////////////////////////
    '// Líneas de Ticket          //
    '///////////////////////////////
    amount As Double
    Number As Long
    item_code As Long
    Item_Price As Double
    Item_Pes As Double
    positive As Long
    Weighted As Long
    Cancelled As Long
    discount As Double
    BovineCard As Long
    ModeLine As Integer
End Type

Public Type Epel_CabTick
    '///////////////////////////////
    '// Cabeceras de Ticket       //
    '///////////////////////////////
    mIdDoc As Long
    type As Long '0--> local - 1 red
    NTicket As Long
    Section As Long
    Vendor As Long
    Client As Long
    Machine As Long
    nLines As Long
    positive As Long
    amount As Double
    Hour As String
    date As String
    Lines() As Epel_LinTick
    Cancelled As Long
    P1 As Double ' importe bruto  // ---> cas.n-> para gama baja almacena el puntero a la linea
    P2 As Double ' descuento      //              inicial de esa cabecera
    Task As Long ' 0--> Normal, 1 --> Tiquet de encargo
    tipoPago As Long
    cantidadEntre As Double
    lote As String
    lEncargo As Boolean
    nLinCanceled As Integer
    discount As Double
    EAN13 As String
End Type

Public Type Epel_Vat
    '/////////////////////////////////////////////
    '// Tipos de IVA (Euroscale / Gama Baja)    //
    '/////////////////////////////////////////////
    type1 As Double
    type2 As Double
    type3 As Double
    type4 As Double
    type5 As Double
End Type

Public Type Epel_LinBovine
    IdNumDocumento As Long
    IdNumLinea As Long
    scale_num As Integer
    scale_step As Integer
    scale_work As Integer
    item_code As Long
    item_cad As Long
    item_fam As Long
    item_weighted As Long
    item_bovine As String
    line_price As Double
    line_weight As Double
    line_amount As Double
    line_tare As Double
    line_abono As Long
    item_descripcion As String
End Type

'=========================
'Consulta estado Conexión
'=========================
' Necesaria para ScalePos y TouchScale
Public Function Get_Status() As Long
    Get_Status = Estado
End Function

'=====================================
'*************************************
' Envio de datos por protocolo       *
'*************************************
'=====================================
' Necesaria para ScalePos y TouchScale
Public Function Send_Item(Item As Epel_Item) As Long
'***************
Dim Resp As Long
'***************
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Artículo"
        Resp = Capa1A_SendItem_Touch(Item)
    Else
        Resp = 1004
    End If
    Send_Item = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Send_FullItem(FullItem As Epel_FullItem) As Long
    '/////////////////////////////////////////////
    '// envío de Artículos                      //
    '/////////////////////////////////////////////
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Artículo Completo"
        Resp = Capa1B_SendFullItem_Touch(FullItem)
    Else
        Resp = 1004
    End If
    Send_FullItem = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para TouchScale
Public Function Send_DesItem(FullItem As Epel_FullItem) As Long
    '/////////////////////////////////////////////
    '// envío de descriptivos                   //
    '/////////////////////////////////////////////
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de descriptivos"
        Resp = Capa1B_SendDesItem_Touch(FullItem)
    Else
        Resp = 1004
    End If
    Send_DesItem = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

'Envio de teclas directas:
'Lmode              ' 0 para programar a toda una sección, 1 para programar una balanza
'lTarget            ' número de sección o balanza a programar
'lTable             ' Tabla de teclas directas (de 0 a 9)
'lIdentifierType    ' 0, se va a indicar un PLU, 1, se va a indicar un código
'Identifier         ' PLU o Código, si vale cero, se borra la tecla.
'KeyNumber          ' número de tecla a programar.
'No ScalePos, Si TouchScale, ScalePos responde OK y no hace nada
Public Function Send_Key(MyTecla As Epel_tecla) As Long
    '/////////////////////////////////////////////
    '// envío de Teclas Directas                //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Send_key-->" & " Mode: " & CStr(MyTecla.Lmode) & " Target:" & CStr(MyTecla.lTarget) & " Table:" & CStr(MyTecla.lTable) & " Id.TYpe:" & CStr(MyTecla.lIdentifierType) & " Identifier:" & CStr(MyTecla.lIdentifier) & " Key:" & CStr(MyTecla.KeyNumber)
        Resp = Capa1A_SendKey_Touch(MyTecla.Lmode, MyTecla.lTarget, MyTecla.lTable, MyTecla.lIdentifierType, MyTecla.lIdentifier, MyTecla.KeyNumber)
    Else
        Resp = 1004
    End If

    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Send_Key = Resp
End Function

' Textos
Public Function EraseTexts() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseTexts_Touch
    Else
        Resp = 1004
    End If
    EraseTexts = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        

' Teclas
Public Function EraseKeys() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseKeys_Touch
    Else
        Resp = 1004
    End If
    EraseKeys = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Kill_All_PluKeys() As Long
Dim res As Long
    If (Estado = 1) Or (Estado = 2) Or (Estado = 3) Then
        CadenadeLog "Recibida orden de borrado de Teclas"
        res = Capa1B_Kill_All_PluKeys_Touch
    Else
        res = 1004
    End If
    Kill_All_PluKeys = res
    If res >= 1000 Then
        Call iRaise(1, res)
    Else
        Call iRaise(0, 0)
    End If
End Function

' TouchScale
Public Function Cons_FullItem(ByVal mMetodo As Integer, ByVal mCodigo As Long, ByVal mProtDesc As Integer, ByRef mFullItem As Epel_FullItem) As Long
    '/////////////////////////////////////////////
    '// Consulta de Artículos                      //
    '/////////////////////////////////////////////
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de consulta de Artículo Completo"
        Resp = Capa1B_ConsFullItem_Touch(mMetodo, mCodigo, mProtDesc, mFullItem)
    Else
        Resp = 1004
    End If
    Cons_FullItem = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Query_Item(ByRef Item As Epel_FullItem, ByVal Lmode As Long) As Long
Dim res As Long
Dim res1 As Long

    If (Estado = 1) Or (Estado = 2) Or (Estado = 3) Then
        CadenadeLog "Recibida orden de consulta de artículos"
        res = Capa1B_Query_Item_Touch(Item, Lmode)
    Else
        res = 1004
    End If
    Query_Item = res
    If res >= 1000 Then
        Call iRaise(1, res)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Erase_Item(ByVal Mode As Long, ByVal Item_ID As Long) As Long
    '/////////////////////////////////////////////
    '// envío de Artículos                      //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Recibida orden de borrado de artículos"
        Resp = Capa1B_EliArt_Touch(Mode, Item_ID)
    Else
        Resp = 1004
    End If
    Erase_Item = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
' Necesaria para ScalePos y TouchScale
Public Function Kill_All_Item() As Long
Dim res As Long
    If (Estado = 1) Or (Estado = 2) Or (Estado = 3) Then
        CadenadeLog "Recibida orden de borrado de todos los artículos"
        res = Capa1B_Kill_All_Item_Touch
    Else
        res = 1004
    End If
    Kill_All_Item = res
    If res >= 1000 Then
        Call iRaise(1, res)
    Else
        Call iRaise(0, 0)
    End If
        
End Function
        
' Borrado Articulos
Public Function ErasePlus() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = ErasePlus_Touch
    Else
        Resp = 1004
    End If
    ErasePlus = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        

' Necesaria para ScalePos y TouchScale
Public Function Send_Vat(VatTypes As Epel_Vat) As Long
    Dim Resp As Long
    CadenadeLog "Función Send_Vat ..."
    CadenadeLog "Parámetros: "
    CadenadeLog "Tipo 0: " & CStr(VatTypes.type1)
    CadenadeLog "Tipo 1: " & CStr(VatTypes.type2)
    CadenadeLog "Tipo 2: " & CStr(VatTypes.type3)
    CadenadeLog "Tipo 3: " & CStr(VatTypes.type4)
    CadenadeLog "Tipo 4: " & CStr(VatTypes.type5)

    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Tipos de IVA"
        Resp = capa1A_sendVat_Touch(VatTypes)
    Else
        Resp = 1004
    End If
    Send_Vat = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Send_BarCode(BarCode As Epel_BarCode) As Long
    '/////////////////////////////////////////////
    '// envío de códigos de Barras              //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Código de Barras"
        Resp = Capa1A_SendBarCode_Touch(BarCode)
    Else
        Resp = 1004
    End If
    Send_BarCode = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Codigos de barras
Public Function EraseEAN() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseEAN_Touch
    Else
        Resp = 1004
    End If
    EraseEAN = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        

' Necesaria para ScalePos y TouchScale
Public Function Query_BarCode(ByRef Bcode As Epel_BarCode, ByVal Lmode As Long) As Long
    ' lmode -->
    ' 0 Todos
    ' 1 Red
    ' 2 Local
    Dim res As Long
    If (Estado = 1) Or (Estado = 2) Or (Estado = 3) Then
        CadenadeLog "Recibida orden de consulta de códigos de barras"
        res = Capa1B_Query_BarCode_Touch(Bcode, Lmode)
    Else
        res = 1004
    End If
    Query_BarCode = res
    If res >= 1000 Then
        Call iRaise(1, res)
    Else
        Call iRaise(0, 0)
    End If
    
End Function

' Necesaria para ScalePos y TouchScale
Public Function Send_HeadingLegend(HeadingLegend As Epel_HeadingLegend) As Long
    '/////////////////////////////////////////////
    '// envío de cabeceras y leyendas           //
    '/////////////////////////////////////////////
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Cabeceras/Leyendas"
        Resp = Capa1A_SendHeadingLegend_Touch(HeadingLegend)
    Else
        Resp = 1004
    End If
    Send_HeadingLegend = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Send_Vendor(ByVal IDV As Long, ByVal sName As String, ByVal nSection As Long, ByVal nKey As Long, ByVal Status As Long, ByVal EraseVendor As Long) As Long
    '/////////////////////////////////////////////
    '// envío de Vendedores                     //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        CadenadeLog "Recibida orden de Envío de Vendedor"
        Resp = Capa1A_SendVendor_Touch(IDV, sName, nSection, nKey, CStr(IDV), Status, EraseVendor)
    Else
        Resp = 1004
    End If
    Send_Vendor = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para TouchScale
Public Function Send_DelVendor(ByVal IDV As Long) As Long
    '/////////////////////////////////////////////
    '// envío de Vendedores                     //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        CadenadeLog "Recibida orden de Envío de Vendedor"
        Resp = Capa1A_SendDelVendor_Touch(IDV)
    Else
        Resp = 1004
    End If
    Send_DelVendor = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
' Necesaria para ScalePos y TouchScale
Public Function Kill_All_Vendor() As Long
Dim res As Long
    If (Estado = 1) Or (Estado = 2) Or (Estado = 3) Then
        CadenadeLog "Recibida orden de borrado de todos los Vendedores"
        res = Capa1B_Kill_All_Vendor_Touch
    Else
        res = 1004
    End If
    Kill_All_Vendor = res
    If res >= 1000 Then
        Call iRaise(1, res)
    Else
        Call iRaise(0, 0)
    End If
End Function
' Vendedores
Public Function EraseVendors() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseVendors_Touch
    Else
        Resp = 1004
    End If
    EraseVendors = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function


'Client_Code : valor 0 a 99999 ' Código del Cliente
'Nom_Cli : Texto, 25 caracteres ' Nombre del Cliente
'Data1 : Texto 25 caracteres ' Datos adicionales del cliente
'Data2 : Texto 25 caracteres ' Datos adicionales del cliente
'Data3 : Texto 25 caracteres ' Datos adicionales del cliente
'Data4 : Texto 25 caracteres ' Datos adicionales del cliente
'Para TouchScale
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
' Necesaria para ScalePos y TouchScale
Public Function Send_Client(MyCliente As Epel_Cliente) As Long
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Recibida orden de envío de Cliente"
        Resp = Capa1A_SendClient_Touch(MyCliente.Client_Code, MyCliente.Nom_Cli, MyCliente.Data1, MyCliente.Data2, MyCliente.Data3, MyCliente.Data4, MyCliente.Erase_Client)
    Else
        Resp = 1004
    End If
    Send_Client = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Clientes
Public Function EraseClients() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseClients_Touch
    Else
        Resp = 1004
    End If
    EraseClients = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        


' Necesaria para ScalePos y TouchScale
Public Function Send_Country(ByVal CountryCode As Long, ByVal CountryName As String, ByVal EraseCountry As Long) As Long
    '/////////////////////////////////////////////
    '// envío/borrado de países                 //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        CadenadeLog "Recibida orden de envío de país"
        Resp = Capa1B_SendCountry_Touch(CountryCode, CountryName, EraseCountry)
    Else
        Resp = 1004
    End If
    Send_Country = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' Necesaria para ScalePos y TouchScale
Public Function Send_Card(card As Epel_Card) As Long
    '///////////////////////////////////////
    '// Solicitud de Pedido (Euroscale)   //
    '///////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Then
        CadenadeLog "Recibida orden de Envío de ficha"
        Resp = Capa1B_SendCard_Touch(card)
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Send_Card = Resp
End Function

' Necesaria para ScalePos y TouchScale
Public Function Erase_Card(ByVal cCode As Long) As Long
    '/////////////////////////////////////////////
    '// borrado de Fichas de vacuno             //
    '/////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Recibida orden de borrado de ficha de vacuno"
        Resp = Capa1B_Erase_Card_Touch(cCode)
    Else
        Resp = 1004
    End If
    Erase_Card = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function


' TouchScale
Public Function ConsEqpCon(ByRef MisEquipos As String) As Long
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de consulta de equipos conectados"
        Resp = Capa1B_ConsEqpCon(MisEquipos)
    Else
        Resp = 1004
    End If
    ConsEqpCon = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' TouchScale
Public Function ConsNumTqtEci(ByRef MiNumTQT As Long, ByRef nBalMar) As Long
    Dim Resp As Long
    If Estado = 1 Then
        CadenadeLog "Recibida orden de consulta de tique"
        Resp = Capa1B_ConsNumTqtEci(MiNumTQT, nBalMar)
    Else
        Resp = 1004
    End If
    ConsNumTqtEci = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' TouchScale
'C2f 10-09-2012 Local
Public Function ConsNumTqtEci_Local(ByRef MiNumTQT As Long, ByRef nBalMar) As Long
    Dim Resp As Long
    If Estado = 1 Then
        Resp = Capa1B_ConsNumTqtEci_Local(MiNumTQT, nBalMar)
    Else
        Resp = 1004
    End If
    ConsNumTqtEci_Local = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' TouchScale
Public Function MarNumTqtEci(ByVal MiNumTQT As Long, ByVal nBalMar As Long) As Long
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de marcar tique"
        Resp = Capa1B_MarNumTqtEci(MiNumTQT, nBalMar)
    Else
        Resp = 1004
    End If
    MarNumTqtEci = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function

' TouchScale
'c2f 10-09-2012 Local
Public Function MarNumTqtEci_Local(ByVal MiNumTQT As Long, ByVal nBalMar As Long) As Long
    Dim Resp As Long
    
    If Estado = 1 Then
        CadenadeLog "Recibida orden de marcar tique Local"
        Resp = Capa1B_MarNumTqtEci_Local(MiNumTQT, nBalMar)
    Else
        Resp = 1004
    End If
    MarNumTqtEci_Local = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function



' Necesaria para ScalePos y TouchScale
Public Function Query_Ticket(ByVal lType As Long, Cabecera As Epel_CabTick, Optional ByVal bNoMarcar As Integer) As Long
    '///////////////////////////////////////////////////
    '// Pedir un ticket                               //
    '// (lo devuelve en una estructura Epel_CabTick)  //
    '///////////////////////////////////////////////////
    Dim Bucle As Integer
    Dim Resp As Long
    Dim bNoMarca As Integer
    If bNoMarcar = 0 Or IsNull(bNoMarcar) Then
        bNoMarca = 0
    Else
        bNoMarca = 1
    End If
        
    If Estado = 1 Or Estado = 2 Then
        CadenadeLog "Recibida petición de Ticket no pedido"
        For Bucle = 1 To 3 ' 3 Reintentos
            Resp = Capa1B_QueryTicket_Touch(lType, Cabecera, bNoMarca)
            If Resp = 0 Or Resp >= 1000 Then Exit For
        Next Bucle
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_Ticket = Resp
End Function

' Necesaria para ScalePos y TouchScale
Public Function Query_NumTicket(ByVal lIdTqt As Long, ByRef Cabecera As Epel_CabTick, Optional ByVal bNoMarcar As Integer) As Long
    '///////////////////////////////////////////////////
    '// Pedir un ticket                               //
    '// (lo devuelve en una estructura Epel_CabTick)  //
    '///////////////////////////////////////////////////
    Dim Bucle As Integer
    Dim Resp As Long
    Dim bNoMarca As Integer
    If bNoMarcar = 0 Or IsNull(bNoMarcar) Then
        bNoMarca = 0
    Else
        bNoMarca = 1
    End If
        
    If Estado = 1 Or Estado = 2 Then
        CadenadeLog "Recibida petición de Ticket no pedido por numero"
        For Bucle = 1 To 3 ' 3 Reintentos
            Resp = Capa1B_QueryNumTicket_Touch(lIdTqt, Cabecera, bNoMarca)
            If Resp = 0 Or Resp >= 1000 Then Exit For
        Next Bucle
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_NumTicket = Resp
End Function

' Necesaria para ScalePos tiques Local
'c2f 10-09-2012
Public Function Query_NumTicket_Local(ByVal lIdTqt As Long, ByRef Cabecera As Epel_CabTick, Optional ByVal bNoMarcar As Integer) As Long
    '///////////////////////////////////////////////////
    '// Pedir un ticket                               //
    '// (lo devuelve en una estructura Epel_CabTick)  //
    '///////////////////////////////////////////////////
    Dim Bucle As Integer
    Dim Resp As Long
    Dim bNoMarca As Integer
    If bNoMarcar = 0 Or IsNull(bNoMarcar) Then
        bNoMarca = 0
    Else
        bNoMarca = 1
    End If
        
    If Estado = 1 Or Estado = 2 Then
        CadenadeLog "Recibida petición de Ticket Local no pedido por numero"
        For Bucle = 1 To 3 ' 3 Reintentos
            Resp = Capa1B_QueryNumTicket_Touch_Local(lIdTqt, Cabecera, bNoMarca)
            If Resp = 0 Or Resp >= 1000 Then Exit For
        Next Bucle
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_NumTicket_Local = Resp
End Function

Public Function MarcarTqt(ByVal lId As Long) As Long
    MarcarTqt = Capa1B_MarcarTqt_Touch(lId)
End Function


'/////////////////////////////////////////////
'// Orden de Borrado de Totales             //
'/////////////////////////////////////////////
' Necesaria para ScalePos y TouchScale
Public Function Erase_Totals(ByVal No_Reinit As Long) As Long
Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        CadenadeLog "Recibida orden de Borrado de Totales "
            Resp = Capa1A_GA_Erase_Totals_Touch(No_Reinit)
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Erase_Totals = Resp
End Function


Public Function Configure(ByVal MachineID As Long, ByVal TypeConf As Long, ByVal sConfigure As String) As Long
'//////////////////////////////////
'// Configura las comunicaciones //
'//////////////////////////////////
' Necesaria para ScalePos y TouchScale
    Dim Resp As Long
    CadenadeLog "Recibida orden de configuración.Máquina : " & MachineID & ".Tipo conexión : " & TypeConf & ".Cadena Conexión : " & sConfigure
    Resp = Capa0_Configure(MachineID, TypeConf, sConfigure)
    If Resp >= 1000 Then Call iRaise(1, Resp)
    Configure = Resp
End Function

' Necesaria para ScalePos y TouchScale
Public Sub Reset()
    '////////////////////////
    '// Cierra los puertos //
    '// Dejando al OCX en  //
    '// el estado inicial  //
    '////////////////////////
    CadenadeLog "Recibida Orden de RESET"
    Capa0_Reset
End Sub

'*****************************
'*** consultas ***
'*****************************




' Borrado General
Public Function EraseGeneral() As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseGeneral_Touch
    Else
        Resp = 1004
    End If
    EraseGeneral = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        
' Borrado Totales
Public Function EraseTotals(ByVal ReinitNum As Integer) As Long
Dim Resp As Long
    If Estado = 1 Then
        Resp = EraseTotals_Touch(ReinitNum)
    Else
        Resp = 1004
    End If
    EraseTotals = Resp
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
End Function
        

Private Sub RecibeEvento_BasComError(ByVal Error_Code As Long)
    Call iRaise(1, Error_Code)
End Sub

Private Sub RecibeEvento_BasComOK()
    Call iRaise(0, 0)
End Sub

Private Sub RecibeEvento_BasItemReceived(ByVal TotalID As Long, ByVal lData As Long)
    Call iRaiseItem(TotalID, lData)
End Sub

Public Function iRaise(imErr As Integer, imResp As Long)
    If imErr = 1 Then
        RaiseEvent ComError(imResp)
    Else
        RaiseEvent ComOK
    End If
End Function

Public Function iRaiseItem(ByVal TotalID As Long, ByVal lData As Long)
    RaiseEvent ItemReceived(TotalID, lData)
End Function

'**************

Public Function ALT_TERMINATE()
    Call UserControl_Terminate
End Function

Private Sub UserControl_Initialize()
    Set RecibeEvento = New CEventoBas
    Set EnviaEvento = RecibeEvento
    PathLog = App.Path & "\logepelcom.txt"
    Capa0_Reset
    CadenadeLog "OCX Cargado"
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
        
    Capa0_Reset
       
    On Error Resume Next
    
    Set RecibeEvento = Nothing
    
    On Error Resume Next
    
    CadenadeLog "OCX Descargado"
End Sub

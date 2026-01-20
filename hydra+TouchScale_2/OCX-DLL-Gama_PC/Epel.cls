VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Epel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/////////////////////////////////////////////////////////////////////////////////////
'// Control ActiveX de comunicaciones. Este control es de uso exclusivamente        //
'// interno. No debe distribuirse a terceros. En su lugar, distribuir               //
'// la librería EPELSA.DLL que se encuentra en la subcarpeta "DLL" de               //
'// este código fuente. El distribuir el control directamente, llevará              //
'// a problemas con todos los programas que no esten diseñados en Visual Basic.     //
'// Por otra parte este OCX está diseñado para trabajar con Hydra Basic, y          //
'// puede crear conflictos con la librería epelsa.dll si es distribuida a terceros  //
'/////////////////////////////////////////////////////////////////////////////////////
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
    ' versión D
    Loss As Long
    Presel As Long
    Desc2040(20) As String  'cas.v->343
    'c2f versión 1.7.4
    lMix As Boolean
    nPeso As Long
    sImagen As String
    ProtDesc As String
End Type

Public Type Epel_FullItemXtra
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
    TRM1 As Long
    PRC1 As Long
    TRM2 As Long
    PRC2 As Long
    Imagen As String
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
    FHeading1 As Long
    FHeading2 As Long
    FHeading3 As Long
    FHeading4 As Long
    FHeading5 As Long
    FLegend1 As Long
    FLegend2 As Long
    FLegend3 As Long
    FLegend4 As Long
    FLegend5 As Long
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
'3.1.43 Estructura para añadir Código de Barras,sec,plu a la línea de tique
Public Type Internal_Extra_Line_Pos
    idTique As Long
    idLinea As Long
    idArticulo As Long
    idSeccion As Integer
    idReferencia As Long
    idBarras As String
    porcentaje As Long
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
    '**********
    ' euroscale
    '**********
    discount As Double
    BovineCard As Long
    ModeLine As Integer
End Type

Public Type Epel_CabTick
    '///////////////////////////////
    '// Cabeceras de Ticket       //
    '///////////////////////////////
    mIdDoc As Long
    type As Long '*** 0=> venta
    NTicket As Long
    Section As Long
    Vendor As Long
    Client As Long
    Machine As Long '*** numMaq
    nLines As Long '***numMax lin
    positive As Long
    amount As Double
    Hour As String
    date As String
    Lines() As Epel_LinTick
    '**********************
    ' euroscale
    '**********************
    Cancelled As Long
    P1 As Double ' importe bruto  // ---> cas.n-> para gama baja almacena el puntero a la linea
    P2 As Double ' descuento      //              inicial de esa cabecera
    Task As Long ' 0--> Normal, 1 --> Tiquet de encargo
    '///cas
    tipoPago As Long
    cantidadEntre As Double
    '///cas
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
'//////////////////
'//Datos EMPRESA //
'//////////////////
'IdEmpresa, IdFiscal, Nombre, Direccion, CodigoPostal, Poblacion, Provincia, Telefono, Email, Logo
Public Type Epel_Empresa
    IdFiscal As String '35
    Nombre  As String '50
    Direccion As String '50
    CodigoPostal As String '15
    Poblacion As String '50
    Provincia As String '50
    Telefono  As String '35
    Email As String '50
End Type
'///////////////////
'// Datos Familia //
'///////////////////
Public Type Epel_Family
    Family As Long 'código de familia
    Text As String 'Descripción
    Sec As Long 'Sección
    Position As Long 'Posición
    Color As Long 'Valor Apaleta
    Imagen As String 'Imagen
End Type
'////////////////////
'// Datos Clientes //
'////////////////////
'dbo.Cliente
'       IdCliente --> Cod_Cli
'       Codigo    --> Cod_Cli
'       Nombre    --> Nom_Cli
'       NIF       --> Data1
'       Direccion    --> Data2
'       CodigoPostal --> Data3
'       Poblacion
'       Provincia
'       Telefono --> Data4
'       Email
Public Type Epel_Clientes
    IdCliente As Long
    codigo As Long
    Nombre As String
    NIF As String
    Direccion As String
    CodigoPostal As String
    Poblacion As String
    Provincia As String
    Telefono As String
    Email As String
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
Public Function ALT_SendVat(ByVal VatTypes As String) As Long
    Dim MiVAt As Epel_Vat
    Dim Resp As Long
    If Len(VatTypes) <> 20 Then
        Resp = 9001
    Else
        CadenadeLog "Función ALT_SendVat..."
        CadenadeLog "String Parámetro: " & VatTypes
        CadenadeLog "Tipo 0: " & Left(VatTypes, 4)
        CadenadeLog "Tipo 1: " & Mid(VatTypes, 5, 4)
        CadenadeLog "Tipo 2: " & Mid(VatTypes, 9, 4)
        CadenadeLog "Tipo 3: " & Mid(VatTypes, 13, 4)
        CadenadeLog "Tipo 4: " & Right(VatTypes, 4)
        MiVAt.type1 = Val(Left(VatTypes, 4))
        MiVAt.type2 = Val(Mid(VatTypes, 5, 4))
        MiVAt.type3 = Val(Mid(VatTypes, 9, 4))
        MiVAt.type4 = Val(Mid(VatTypes, 13, 4))
        MiVAt.type5 = Val(Right(VatTypes, 4))
        Resp = Send_Vat(MiVAt)
    End If
    ALT_SendVat = Resp
End Function

'/////////////////////////////////////////////
'// Pedir todos los tiquets y dejarlos en   //
'// un fichero                              //
'/////////////////////////////////////////////
' Necesaria para ScalePos y TouchScale
Public Function Query_All_Tickets(ByVal lType As Long, ByVal Mode As Long, ByVal sFilePath As String) As Long
    Dim Resp As Long
    If Mode <> 0 And Mode <> 1 Then
        Resp = 1002
    Else
        If Estado = 1 Or Estado = 2 Then
            CadenadeLog "Recibida orden de volcado de Totales a fichero : " & sFilePath
            If IsNumeric(Mid(sFilePath, 1, 8)) Then
                sFilePath = Mid(sFilePath, 9)
                myLastDoc = Val(Mid(sFilePath, 1, 8))
            Else
                myLastDoc = 0
            End If
            Resp = F_Query_All_Tickets_Touch(lType, Mode, sFilePath, False)
        Else
            Resp = 1004
        End If
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_All_Tickets = Resp
End Function

'/////////////////////////////////////////////
'// Pedir todos los tiquets y dejarlos en   //
'// un fichero & Forma de pago.             //
'/////////////////////////////////////////////
Public Function Query_All_Tickets_P(ByVal lType As Long, ByVal Mode As Long, ByVal sFilePath As String) As Long
    Dim Resp As Long
    If Mode <> 0 And Mode <> 1 Then
        Resp = 1002
    Else
        If Estado = 1 Or Estado = 2 Then
            CadenadeLog "Recibida orden de volcado de Tique a fichero (& forma de pago): " & sFilePath
            If IsNumeric(Mid(sFilePath, 1, 8)) Then
                sFilePath = Mid(sFilePath, 9)
                myLastDoc = Val(Mid(sFilePath, 1, 8))
            Else
                myLastDoc = 0
            End If
            Resp = F_Query_All_Tickets_Touch(lType, Mode, sFilePath, True)
        Else
            Resp = 1004
        End If
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_All_Tickets_P = Resp
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

' Necesaria para ScalePos y TouchScale
Public Function Query_One_Ticket(ByVal lType As Long, ByVal Lmode As Long, ByVal lMatch As Long, ByVal lNumber As Long, Cabecera As Epel_CabTick) As Long
    '///////////////////////////////////////////////////
    '// Pedir un ticket                               //
    '// (lo devuelve en una estructura Epel_CabTick)  //
    '///////////////////////////////////////////////////
    Dim Bucle As Integer
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        If Estado = 3 Then
            Resp = 1010
        Else
            CadenadeLog "Recibida orden de Petición de ticket concreto : " & lNumber & " Modo : " & Lmode
            For Bucle = 1 To 3
                If lType > 3 And Lmode <> 2 Then
                    Resp = 1010
                Else
                    Resp = Capa1B_QueryOneTicket_Touch(lType, Lmode, lMatch, lNumber, Cabecera)
                End If
                If Resp = 0 Or Resp >= 1000 Then Exit For
            Next Bucle
        End If
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_One_Ticket = Resp
End Function

' Necesaria para ScalePos y TouchScale
Public Function Query_One_Ticket_Mark(ByVal lType As Long, ByVal Lmode As Long, ByVal lMatch As Long, ByVal lNumber As Long, Cabecera As Epel_CabTick) As Long
    '///////////////////////////////////////////////////
    '// Pedir un ticket marcándolo                   //
    '// (lo devuelve en una estructura Epel_CabTick)  //
    '///////////////////////////////////////////////////
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        Resp = Capa1B_QueryOneTicketMark_Touch(lType, Lmode, lMatch, lNumber, Cabecera)
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_One_Ticket_Mark = Resp
End Function

' Necesaria para ScalePos y TouchScale
'Public Type Epel_LinBovine
'    scale_num As Integer
'    scale_step As Integer
'    scale_work As Integer
'    item_code As Long
'    item_cad As Long
'    item_fam As Long
'    item_weighted As Long
'    item_bovine As String
'    line_price As Double
'    line_weight As Double
'    line_amount As Double
'    line_tare As Double
'End Type
' Consultamos una linea de vacuno --> Se marca para no volver a pedir
Public Function Query_Line_Bovine(ByVal tFam As Integer, ByVal tBal As Integer, MyLine As Epel_LinBovine) As Long
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        Resp = QueryLineBovine_Touch(tFam, tBal, MyLine)
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Query_Line_Bovine = Resp
End Function

' Programamos lote de linea de vacuno
Public Function Prog_Line_Bovine(MyLine As Epel_LinBovine) As Long
    Dim Resp As Long
    If Estado = 1 Or Estado = 2 Or Estado = 3 Then
        Resp = ProgLineBovine_Touch(MyLine)
    Else
        Resp = 1004
    End If
    If Resp >= 1000 Then
        Call iRaise(1, Resp)
    Else
        Call iRaise(0, 0)
    End If
    Prog_Line_Bovine = Resp
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

''SCALEPOS
'"COD","Código",1,6,0
'"MOS","Mostrador",7,8,0
'"PLU","Plu",9,12,0
'"TXT","Texto",13,37,0
'"PRC","Precio",38,44,0
'"CAD","Caducidad",45,47,0
'"CBA","Código de barras",48,60,0
'"ETQ","Etiqueta",61,63,0
'"FAM","Familia",64,66,0
'"IVA","Tipo de IVA",67,67,0
'"MER","Merma en tanto por ciento",68,72,0
'"PRE","Consumo preferente",73,75,0
'"NPR","Nº de Paquete Preseleccionado",76,80,0
'"TAR","Tara",81,85,0
'"TXT","Texto",86,110,1
'"TXT","Texto",111,135,2
'"TXT","Texto",136,160,3
'"TXT","Texto",161,185,4
'"TXT","Texto",186,210,5
'"TXT","Texto",211,235,6
'"TXT","Texto",236,260,7
'"TXT","Texto",261,285,8
'"TXT","Texto",286,310,9
'"TXT","Texto",311,335,10
'"TXT","Texto",336,360,11
'"TXT","Texto",361,385,12
'"TXT","Texto",386,410,13
'"TXT","Texto",411,435,14
'"TXT","Texto",436,460,15
'"TXT","Texto",461,485,16
'"TXT","Texto",486,510,17
'"TXT","Texto",511,535,18
'"TXT","Texto",536,560,19
'"TXT","Texto",561,585,20
'"WGH","Pesado/No pesado",586,586,0
'"TRM1","Tramo 1",587,591,0
'"PRC1","Precio del Tramo 1",592,598,0
'"TRM2","Tramo 2",599,603,0
'"PRC2","Precio del Tramo 2",604,610,0
'"IMAGE","Path + fichero Imagen",611,660,0
' Necesaria para ScalePos
Public Function ALT_SendARTCON(ByVal FullItem As String) As Long
    Dim miart As Epel_FullItem
    Dim Resp As Long
    Dim Bucle As Integer
    If Estado = 1 Then
        miart.subsec = 0
        miart.Font = 0
        For Bucle = 0 To 9
            miart.FDesc(Bucle) = 0
        Next Bucle
        
        miart.Code = Val(Mid(FullItem, 1, 6))
        miart.Sec = Val(Mid(FullItem, 7, 2))
        miart.Plu = Val(Mid(FullItem, 9, 4))
        miart.Text = Trim(Mid(FullItem, 13, 25))
        miart.Price = Val(Mid(FullItem, 38, 7))
        miart.Caducity = Val(Mid(FullItem, 45, 3))
        miart.EAN13 = Trim(Mid(FullItem, 48, 13))
        miart.Etq = Val(Mid(FullItem, 61, 3))
        miart.Family = Val(Mid(FullItem, 64, 3))
        miart.vat = Val(Mid(FullItem, 67, 1))
        miart.Loss = Val(Mid(FullItem, 68, 5))
        miart.Pref = Val(Mid(FullItem, 73, 3))
        miart.Presel = Val(Mid(FullItem, 76, 5))
        miart.Tare = Val(Mid(FullItem, 81, 5))
        For Bucle = 0 To 19
            miart.Desc(Bucle) = Trim(Mid(FullItem, 86 + (Bucle * 25), 25))
        Next Bucle
        If Mid(FullItem, 586, 1) = "W" Or Mid(FullItem, 586, 1) = "0" Then
            miart.weight = 1
        Else
            miart.weight = 0
        End If
        miartXtra.TRM1 = Val(Mid(FullItem, 587, 5))
        miartXtra.PRC1 = Val(Mid(FullItem, 592, 7))
        miartXtra.TRM2 = Val(Mid(FullItem, 599, 5))
        miartXtra.PRC2 = Val(Mid(FullItem, 604, 7))
        miartXtra.Imagen = Mid(FullItem, 611)
        Resp = Send_FullItem(miart)
    Else
        Resp = 1004
    End If
    ALT_SendARTCON = Resp
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

Private Sub RecibeEvento_BasTicketReceived(ByVal NTicket As Long, ByVal nSection As Long, ByVal NScale As Long, Cancel As Long)
    Call iRaiseTqt(NTicket, nSection, NScale)
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

Public Function iRaiseTqt(ByVal NTicket As Long, ByVal nSection As Long, ByVal NScale As Long)
    RaiseEvent TicketReceived(NTicket, nSection, NScale)
End Function
'**************
'***********
'***********
Private Sub Class_Initialize()
   
    Set RecibeEvento = New CEventoBas
    Set EnviaEvento = RecibeEvento
    PathLog = App.Path & "\logepelcom.txt"
    Capa0_Reset
    CadenadeLog "OCX Cargado"
End Sub

'******
'******
Private Sub Class_Terminate()
    
    On Error Resume Next
      
    Capa0_Reset
       
    On Error Resume Next
    
    Set RecibeEvento = Nothing
    
    On Error Resume Next
    
    CadenadeLog "OCX Descargado"
    
End Sub


Public Function ALT_TERMINATE()
    Call Class_Terminate
End Function


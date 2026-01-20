Attribute VB_Name = "ScalePos"
Option Explicit
'**************
Private Do_Cancel As Long
Public sImgVnd As String
Public miartXtra As Epel_FullItemXtra
Public Extra_lines() As Internal_Extra_Line_Pos
Public myLastDoc As Long
Private Type Tipo_Linea
    EsPesado As Boolean
    FactorPeso As Integer
    precio As Double
    cantidad As Double
    NumeroLinea As Byte
    CodigoArticulo As Long
    IVA As Integer
    TipoLinea As Byte
    DescuentoLinea As Double
    ImporteLinea As Double
    LineaCancelada As Boolean
    SehaModificado As Boolean
    ModificadoTemporal As Boolean
    FichaVacuno As Long
    ImporteBruto As Double
    descriptivo As String
    DescuentoPorcentaje As Double
End Type
Private Type tipo_cabecera
     Ntiquet As Long
     NVendedor As Long
     NMostrador As Long
     NCliente As Long
     Fecha As Variant
     hora As Variant
     NBalanza As Long
     ImporteTotal As Double
     NLineas As Long
     LineaInicial As Long
     LineaFinal As Long
     codigo As Long
     CadenaEstadoTiquet As String
     ImporteBruto As Double
     Impuestos As Double
     descuento As Double
     LineasCanceladas As Long
     tipoTiquet As Long
     EstadoTiquet As Long
     LineasMensaje As Long
     '********************
     ' este dato se introduce
     ' tras pasar por actualiza_tabla_cabecera
     Incluir_en_Base As Boolean
     tipoPago As Long
     cantidadEntre As Double
     Cambio As Double
     lote As String
     lEncargo As Boolean
     documentoabono As Long
     Lines() As Tipo_Linea
End Type

'**************
'Colores
'sColor = "#FFE0E0E0"
'sColor = "#FFC1C1C1"
'sColor = "#FFA2A2A2"
'sColor = "#FF6d6f81"
'sColor = "#FF838383"
'sColor = "#FF646464"
'sColor = "#FF454545"
'sColor = "#FF000000"
'sColor = "#FFF0FF00"
'sColor = "#FFc7de28"
'sColor = "#FFFFDA28"
'sColor = "#FFbba125"
'sColor = "#FFFFA500"
'sColor = "#FFFF7E34"
'sColor = "#FFFF7400"
'sColor = "#FF974500"
'sColor = "#FFC2FF49"
'sColor = "#FF4EFB22"
'sColor = "#FF16C700"
'sColor = "#FF00B761"
'sColor = "#FF00C7B4"
'sColor = "#FF00BAFF"
'sColor = "#FF00C7B4"
'sColor = "#FF242da3"
'sColor = "#FFC088FF"
'sColor = "#FFE082FF"
'sColor = "#FFD555FF"
'sColor = "#FFC518FF"
'sColor = "#FFFF00BE"
'sColor = "#FFFF005F"
'sColor = "#FFE20000"
'sColor = "#FFb70000"
Public Function APaleta(ByVal nColor As Integer) As String
Dim sColor As String
    Select Case nColor
    Case 0
        sColor = "#FFE0E0E0"
    Case 1
        sColor = "#FFC1C1C1"
    Case 2
        sColor = "#FFA2A2A2"
    Case 3
        sColor = "#FF6D6F81"
    Case 4
        sColor = "#FF838383"
    Case 5
        sColor = "#FF646464"
    Case 6
        sColor = "#FF454545"
    Case 7
        sColor = "#FF000000"
    Case 8
        sColor = "#FFF0FF00"
    Case 9
        sColor = "#FFC7DE28"
    Case 10
        sColor = "#FFFFDA28"
    Case 11
        sColor = "#FFBBA125"
    Case 12
        sColor = "#FFFFA500"
    Case 13
        sColor = "#FFFF7E34"
    Case 14
        sColor = "#FFFF7400"
    Case 15
        sColor = "#FF974500"
    Case 16
        sColor = "#FFC2FF49"
    Case 17
        sColor = "#FF4EFB22"
    Case 18
        sColor = "#FF16C700"
    Case 19
        sColor = "#FF00B761"
    Case 20
        sColor = "#FF00C7B4"
    Case 21
        sColor = "#FF00BAFF"
    Case 22
        sColor = "#FF00C7B4"
    Case 23
        sColor = "#FF242DA3"
    Case 24
        sColor = "#FFC088FF"
    Case 25
        sColor = "#FFE082FF"
    Case 26
        sColor = "#FFD555FF"
    Case 27
        sColor = "#FFC518FF"
    Case 28
        sColor = "#FFFF00BE"
    Case 29
        sColor = "#FFFF005F"
    Case 30
        sColor = "#FFE20000"
    Case 31
        sColor = "#FFB70000"
    Case Else
        sColor = "#FFE0E0E0"
    End Select
    APaleta = sColor
End Function

Public Function sAjusta(ByVal sParte As String) As String
Dim nL As Integer
    For nL = 1 To Len(sParte)
        If Mid(sParte, nL, 1) = "-" Then
            If Mid(sParte, nL + 1, 1) = "0" Then
                sParte = Mid(sParte, 1, nL) & Mid(sParte, nL + 2)
            End If
        End If
    Next nL
    sAjusta = sParte
End Function
Public Function Ajusta25(ByVal sS As String) As String
Dim nL As Integer
Dim sL As String
nL = Len(Trim(sS))
If nL <> 25 Then
    sL = sS & Space(25 - nL)
Else
    sL = sS
End If

Ajusta25 = sL
End Function
Public Function Quitar_Coma_S(ByVal MiCadena As String) As String
    ' *****************
    ' Esta función sustituye
    ' el caracter "," (coma) por
    ' el caracter "." (punto)
    Dim Bucle As Integer
    Dim Buffer As String
    For Bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, Bucle, 1) = "," Then
            Buffer = Buffer & "."
        Else
            Buffer = Buffer & Mid(MiCadena, Bucle, 1)
        End If
    Next Bucle
    Quitar_Coma_S = Buffer
End Function

Public Function Quitar_Coma(ByVal MiCadena As String, ByVal Longitud As Integer) As String
    ' *****************
    ' Esta función sustituye
    ' el caracter "," (coma) por
    ' el caracter "." (punto)
    Dim C1 As Integer
    Dim HayPunto As Boolean
    Dim Bucle As Integer
    Dim Buffer As String
    If Len(MiCadena) > Longitud Then
        Quitar_Coma = Space(Longitud)
        Exit Function
    End If
    For Bucle = 1 To Len(MiCadena)
        If Mid(MiCadena, Bucle, 1) = "," Or Mid(MiCadena, Bucle, 1) = "." Then
            HayPunto = True
            Buffer = Buffer & "."
        Else
            Buffer = Buffer & Mid(MiCadena, Bucle, 1)
        End If
    Next Bucle
    C1 = 0
    For Bucle = Len(Buffer) To 1 Step -1
        If Mid(Buffer, Bucle, 1) = "." Then
            HayPunto = True
            Exit For
        Else
            C1 = C1 + 1
        End If
    Next Bucle
    If HayPunto = False Then
        'If UsaEuro Or decimales = 2 Then
            Buffer = Buffer & ".00"
        'Else
        '    Select Case decimales
        '        Case 1
        '            Buffer = Buffer & ".0"
        '        Case 3
        '            Buffer = Buffer & ".000"
        '    End Select
        'End If
    Else
        Do While C1 < 2
            Buffer = Buffer & "0"
            C1 = C1 + 1
        Loop
    End If
    Do While Len(Buffer) < Longitud
        Buffer = "0" & Buffer
    Loop
    Quitar_Coma = Buffer
End Function
'**************

Public Function Capa1B_SendFamily(Family As Epel_Family)
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim lEsAlta As Boolean
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim nPrice As Double
    'datos familia:
    'Family as long 'código de familia
    'Text as string 'Descripción
    'Sec as long 'Sección
    'Position as long 'Posición
    'Color as long 'Valor Apaleta
    'Imagen as string 'Imagen
    
    lEsAlta = False
    
    On Error Resume Next
    'PonDato "Familia...", 0
    MyRecAux.Open "select * from dbo.familia where idfamilia=" & CStr(Family.Family), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFamily
    End If
    On Error GoTo FIN_Capa1B_SendFamily
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        lEsAlta = True
    End If
    MyRecAux.Fields("idfamilia") = Family.Family

    If Dir(App.Path & "\images\families\" & CStr(Family.Family)) <> "" Then
        On Error GoTo noImgFam1
        If FileLen(App.Path & "\images\families\" & CStr(Family.Family)) > 0 Then
            'Imagen
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\images\families\" & CStr(Family.Family)
            MyRecAux.Fields("imagen").Value = strStream.Read
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
        End If
noImgFam1:
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
    Else
        If Dir(App.Path & "\fmlimg.jpg") <> "" Then
            On Error GoTo noImgFam2
            If FileLen(App.Path & "\fmlimg.jpg") > 0 Then
                
                'Imagen defecto
                strStream.type = adTypeBinary
                strStream.Open
                strStream.LoadFromFile App.Path & "\fmlimg.jpg"
                MyRecAux.Fields("imagen").Value = strStream.Read
                '''''''''''''''
                strStream.Flush
                strStream.Close
                Set strStream = Nothing
                ''''''''
                MyRecAux.Fields("color") = "#FFC2FF49" '"#FFCCCCCC"

            End If
noImgFam2:
            If Err.Number <> 0 Then
                On Error GoTo 0
            End If
        Else
            If Not IsNumeric(Family.Imagen) Then
                If Dir(Family.Imagen) <> "" Then
                    If FileLen(Family.Imagen) > 0 Then
                        'Imagen
                        strStream.type = adTypeBinary
                        strStream.Open
                        strStream.LoadFromFile Trim(Family.Imagen)
                        MyRecAux.Fields("imagen").Value = strStream.Read
                        '''''''''''''''
                        strStream.Flush
                        strStream.Close
                        Set strStream = Nothing
                        ''''''''
                    End If
                
                End If
            Else
                MyRecAux.Fields("color") = APaleta(Family.Color)
            End If
            
        End If

    End If
    MyRecAux.Fields("descripcion") = Trim(Family.Text)
    If lEsAlta Then MyRecAux.Fields("codigo") = Family.Family
    MyRecAux.Update
    'si es alta...
    If lEsAlta Then
        cm = "update dbo.familia set idglobal=NEWID() where idfamilia=" & CStr(Family.Family)
        MyConnObj.Execute cm
        ''''''''''''''
    End If
    lEsAlta = False
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    On Error Resume Next
    'PonDato "FamiliasSeccion...", 0
    MyRecAux.Open "select * from dbo.familiasseccion where idfamilia=" & CStr(Family.Family), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFamily
    End If
    On Error GoTo FIN_Capa1B_SendFamily
    If MyRecAux.EOF Then
        MyRecAux.AddNew
    End If
    MyRecAux.Fields("posicion") = Family.Position
    MyRecAux.Fields("idfamilia") = Family.Family
    MyRecAux.Fields("idseccion") = Family.Sec

    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    Resp = 0
    
FIN_Capa1B_SendFamily:
    If Err.Number <> 0 Then
        CadenadeLog "Error en SendFamily. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    
    PonDato " Cod: " & Family.Family & " Sec: " & Family.Sec & " Desc: " & Family.Text, Resp
    Capa1B_SendFamily = Resp

End Function

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
Public Function Capa1A_SendCliente(miCliente As Epel_Clientes) As Long
Dim MyRecDes As New ADODB.Recordset
Dim Resp As Long
Dim lEsAlta As Boolean
Dim cm As String
    
    On Error GoTo FINCLIENTES
    If Trim(miCliente.NIF) = "" And Trim(miCliente.Nombre) = "" Then
        cm = "delete dbo.cliente where idcliente=" & CStr(miCliente.IdCliente)
        MyConnObj.Execute cm
        Resp = 0
        GoTo FINCLIENTES
    End If
    MyRecDes.Open "Select * from dbo.cliente where idcliente=" & CStr(miCliente.IdCliente), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FINCLIENTES
    End If
    If MyRecDes.EOF Then
        MyRecDes.AddNew
    End If
    MyRecDes.Fields("idcliente") = miCliente.IdCliente
    MyRecDes.Fields("codigo") = miCliente.codigo
    MyRecDes.Fields("nombre") = miCliente.Nombre
    MyRecDes.Fields("nif") = miCliente.NIF
    MyRecDes.Fields("direccion") = miCliente.Direccion
    MyRecDes.Fields("codigopostal") = miCliente.CodigoPostal
    MyRecDes.Fields("telefono") = miCliente.Telefono
    MyRecDes.Fields("poblacion") = miCliente.Poblacion
    MyRecDes.Fields("provincia") = miCliente.Provincia
    MyRecDes.Fields("email") = miCliente.Email
    
    MyRecDes.Update
    MyRecDes.Close
    Set MyRecDes = Nothing
FINCLIENTES:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_SendCliente. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1000
    Else
        CadenadeLog "Envio Cliente:" & CStr(miCliente.IdCliente) & " -->OK"
    End If
    Capa1A_SendCliente = Resp
    
End Function

Public Function Capa1A_SendBarPos(MyCod As Epel_BarCode) As Long
'dbo.ConfigCodigoBarras
'       IdCodigoBarras
'       Descripcion
'       Tipo
'       Mascara
Dim MyRecDes As New ADODB.Recordset
Dim cm As String
Dim lEsAlta As Boolean
Dim nC As Integer
Dim Resp As Long
            
            On Error GoTo FINCODBAR
            
            For nC = 1 To 5
            Select Case nC
            Case 1
                MyRecDes.Open "Select * from dbo.configcodigobarras where descripcion='CB_VNT'", MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecDes.EOF Then
                    lEsAlta = True
                    'MyRecDes.AddNew
                End If
                '       IdCodigoBarras
                '       Descripcion
                '       Tipo
                '       Mascara
                If lEsAlta Then
                    lEsAlta = False
                    cm = "insert into dbo.configcodigobarras (idcodigobarras, descripcion, tipo, mascara) values (NEWID(),'CB_VNT','EAN-13','" & MyCod.Net_Sale & "?" & "')"
                    MyConnObj.Execute cm
                    ''''''''''''''
                Else
                    cm = "update dbo.configcodigobarras set tipo='EAN-13',mascara='" & MyCod.Net_Sale & "?'" & " where descripcion='CB_VNT'"
                    MyConnObj.Execute cm
                End If
                MyRecDes.Close
                Set MyRecDes = Nothing
            Case 2
                MyRecDes.Open "Select * from dbo.configcodigobarras where descripcion='CB_SUP'", MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecDes.EOF Then
                    lEsAlta = True
                    'MyRecDes.AddNew
                End If
                '       IdCodigoBarras
                '       Descripcion
                '       Tipo
                '       Mascara
                If lEsAlta Then
                    lEsAlta = False
                    cm = "insert into dbo.configcodigobarras (idcodigobarras, descripcion, tipo, mascara) values (NEWID(),'CB_SUP','EAN-13','" & MyCod.Net_Super & "?" & "')"
                    MyConnObj.Execute cm
                Else
                    cm = "update dbo.configcodigobarras set tipo='EAN-13',mascara='" & MyCod.Net_Super & "?'" & " where descripcion='CB_SUP'"
                    MyConnObj.Execute cm
                End If
                MyRecDes.Close
                Set MyRecDes = Nothing
            Case 3
                MyRecDes.Open "Select * from dbo.configcodigobarras where descripcion='CB_MIX'", MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecDes.EOF Then
                    lEsAlta = True
                    'MyRecDes.AddNew
                End If
                '       IdCodigoBarras
                '       Descripcion
                '       Tipo
                '       Mascara
                If lEsAlta Then
                    lEsAlta = False
                    cm = "insert into dbo.configcodigobarras (idcodigobarras, descripcion, tipo, mascara) values (NEWID(),'CB_MIX','EAN-13','" & MyCod.Net_Mix & "?" & "')"
                    MyConnObj.Execute cm
                Else
                    cm = "update dbo.configcodigobarras set tipo='EAN-13',mascara='" & MyCod.Net_Mix & "?'" & " where descripcion='CB_MIX'"
                    MyConnObj.Execute cm
                End If
                MyRecDes.Close
                Set MyRecDes = Nothing
            Case 4
                MyRecDes.Open "Select * from dbo.configcodigobarras where descripcion='CB_EV1'", MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecDes.EOF Then
                    lEsAlta = True
                    'MyRecDes.AddNew
                End If
                '       IdCodigoBarras
                '       Descripcion
                '       Tipo
                '       Mascara
                If lEsAlta Then
                    lEsAlta = False
                    cm = "insert into dbo.configcodigobarras (idcodigobarras, descripcion, tipo, mascara) values (NEWID(),'CB_EV1','EAN-13','" & MyCod.Net_Pack1 & "?" & "')"
                    MyConnObj.Execute cm
                Else
                    cm = "update dbo.configcodigobarras set tipo='EAN-13',mascara='" & MyCod.Net_Pack1 & "?'" & " where descripcion='CB_EV1'"
                    MyConnObj.Execute cm
                End If
                MyRecDes.Close
                Set MyRecDes = Nothing
            Case 5
                MyRecDes.Open "Select * from dbo.configcodigobarras where descripcion='CB_EV2'", MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecDes.EOF Then
                    lEsAlta = True
                    'MyRecDes.AddNew
                End If
                '       IdCodigoBarras
                '       Descripcion
                '       Tipo
                '       Mascara
                If lEsAlta Then
                    lEsAlta = False
                    cm = "insert into dbo.configcodigobarras (idcodigobarras, descripcion, tipo, mascara) values (NEWID(),'CB_EV2','EAN-13','" & MyCod.Net_Pack2 & "?" & "')"
                    MyConnObj.Execute cm
                Else
                    cm = "update dbo.configcodigobarras set tipo='EAN-13',mascara='" & MyCod.Net_Pack2 & "?'" & " where descripcion='CB_EV2'"
                    MyConnObj.Execute cm
                End If
                'MyRecDes.Close
                'Set MyRecDes = Nothing
            End Select
            Next nC
    
            Resp = 0

FINCODBAR:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_SendBarPos:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    On Error Resume Next
    MyRecDes.Close
    Set MyRecDes = Nothing
    On Error GoTo 0


    Capa1A_SendBarPos = Resp

End Function

'Empresa.
'IdEmpresa, IdFiscal, Nombre, Direccion, CodigoPostal, Poblacion, Provincia, Telefono, Email, Logo
Public Function Capa1A_SendEmpresa(Empresa As Epel_Empresa) As Long
Dim cm As String
    cm = "update dbo.empresa set idfiscal='" & Empresa.IdFiscal & "',nombre='" & Empresa.Nombre & "',direccion='" & Empresa.Direccion & "'," & _
         "codigopostal='" & Empresa.CodigoPostal & "',poblacion='" & Empresa.Poblacion & "',provincia='" & Empresa.Provincia & "'," & _
         "telefono='" & Empresa.Telefono & "',email='" & Empresa.Email & "' where idempresa=1"
    MyConnObj.Execute cm

End Function

'***********************************
'Formas de Pago....
'1   Efectivo
'2   Tarjeta Crédito
'3   Tarjeta Débito
'4   EBT Cash
'5   Gift Card
'6   Cheque
'7   Pendiente
'14  Ticket Restaurant
'15  Chèque -Vacances
'16  Cheque Gourmet
'17  Restaurant Pass
'19  Cupón
'20  Others
'21  EBT FoodStamp
'22  Visa
'23  MasterCard
'24  American Express
'25  Traveller 's Check
'28  Multibanco
'30  Visa Electron
'31  Diners Club
'32  Big Pass
'33  Sodexho
'34  Pin
'35  Chipknip
'36  Cestaticket
'37  EC -Maestro
'38  Quick
'39  BPI -Express
'40  Gift Check
'41  Ripley
'42  Saga
'43  Otras tarjetas
'44  Cheque Comercial
'45  Cheque de Gobiernos
'46  CLAVE
'TIQUE.DAT (Cada registro contiene una Línea de Documento)
'Posición Posición
'Inicio   Fin     Contenido               Valores
'1        5       Número de Tiquet        00000 a 99999
'6        8       Número de Línea         001 a 999
'9        16      Fecha                   DD/MM/YY
'17       21      Hora                    HH:MM
'22       24      Sección Máquina         001 a 999
'25       28      Identificación Vendedor 0001 a 9999
'29       34      Código Artículo         000001 a 999999
'35       41      Precio Venta            Sin decimales 7 dig.
'42       42      Pesado / No pesado      0=Pesado, 1=No pesado
'43       49      Peso de la Venta        Sin decimales 7 dig.
'50       56      Unidades                Sin decimales 7 dig.
'57       63      Importe                 Sin decimales 7 dig.
'64       64      Tipo de Documento       (*)
'65       65      Linea Cancelada o no    0=No Cancelada
'                                         1=Cancelada    (***)
'66       90      Descriptivo artículo    25 Caracteres
'91       115     LOTE                    NºLote 25 Caracteres
'116      120     Código Cliente          00001 a 99999
'121      127     B.Importe               Importe Base.
'128      129     Forma de pago           (**)
'CR+LF                                    Fin Registro
'(*) Tipos de Documento:
'1 = Tique
'2 = Etiqueta Envasado
'3 = Etiqueta AutoServicio
'4 = Tique Local
'5 = Etiqueta Envasado Local
'6 = Etiqueta AutoServicio Local
'7 = Factura (Tique Cobrado)
'8 = Factura Local.
'Nota: Si la forma de pago es "Pendiente (07)", aunque se trate de un Tique Cobrado, el Tipo de Documento quedará como Tique. (01 ó 04).
'(**) Formas de Pago
'01 = Efectivo       02 = Tarjeta Crédito      03 = Tarjeta Débito
'04 = EBT Cash       05 = Gift Card            06 = Cheque
'07 = Pendiente      14 = Ticket Restaurant    15 = Chèque -Vacances
'16 = Cheque Gourmet 17 = Restaurant Pass      19 = Cupón
'20 = Others         21 = EBT FoodStamp        22 = Visa
'23 = MasterCard     24 = American Express     25 = Traveller 's Check
'28 = Multibanco     30 = Visa Electron        31 = Diners Club
'32 = Big Pass       33 = Sodexho              34 = Pin
'35 = Chipknip       36 = Cestaticket          37 = EC -Maestro
'38 = Quick          39 = BPI -Express         40 = Gift Check
'41 = Ripley         42 = Saga                 43 = Otras tarjetas
'44 = Cheque Comercial  45 = Cheque de Gobiernos  46 = CLAVE
'(***) Cada Linea Cancelada Genera 2 registros. (Linea original y Linea con indicador de linea cancelada con los signos negativos en los campos correspondientes).
Public Function LeeTickets(ByVal nEPL As Long) As Long
    Dim myConnTqt As New ADODB.Connection
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRec As New ADODB.Recordset
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim nCont As Long
    Dim sFich As String
    Dim nFich As Integer
    Dim sParte As String
    Dim nResp As Long
    Dim ssParte As String
    Dim sParteCab As String
    Dim nnFich As Integer
    
    On Error GoTo SalirLeeTique
    
    'Exp_Path = App.Path & "\exporta"
    If (Dir(App.Path & "\exporta", vbDirectory) = "") Then MkDir App.Path & "\exporta"
    
    'On Error Resume Next
    'Set MyBase = DAO.OpenDatabase(Base_General)
    ''''''''''''
    'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000"
    ''''''''''''
    'Para conectar mediante IP y Port.
    '''''''''''''''
    '¡ATENCIÓN A LA CONEXIÓN...
    ' Por defecto en la Balanza el puerto para conectar será: 1031
    '--> SQLSERVER Configuration Manager -->
    '    Configuración de red de SQL Server -->
    '    Protocolos de SQLEXPRESS -->
    '    TCP/IP -->Propiedades-->Direcciones IP-->IPALL-->Puertos Dinámicos TCP.
    '''''''''''''''
    'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=ScalePos;Data Source=127.0.0.1,6000" 'el puerto es normalmente 1031.
    'If Err.Number <> 0 Then
    '    On Error GoTo 0
    '    MsgBox "No es posible la Conexión..."
    '    GoTo FIN
    'End If
            
    'desmarcar tiques... PRUEBA
    ''Ffecha = Now()
    ''cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:00.000" & "'"
    ''MyConnObj.Execute cm
    'cm = "update dbo.documento set iscapturado=0"
    'MyConnObj.Execute cm
    '''''''''''''''''''''''''''
    
    
    '(ICG) versión Junio tipo documento (0=Venta, 2=Envasado, 3=autoServicio)
    '      versión Julio                (0=(Mixto y Super y TPV), 1=Venta, 2=Envasado, 3=Autoservicio)
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
    
    'En primera versión se considera:
    '   ICG           HydraTouch
    '  TotalNeto      Importe
    '  TotalBruto     Importe_bruto
    '  ImporteCambio  Cambio
    '                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
    'On Error Resume Next
    'MyConnObj.Close
    'Set MyConnObj = Nothing
    'On Error GoTo 0
    'On Error Resume Next
    'MyConnTqt.Close
    'Set MyConnTqt = Nothing
    'On Error GoTo 0
    If nEPL <> 0 Then
        myConnTqt.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=EPL;Data Source=" & sAddIp
    End If
    'Para marcar tique en fecha.
    MyRecSet.Open "select * from dbo.documento where ((SUBSTRING(convert(varchar(23),fecha,121),12,8)='00:00:00') and isfinalizado=1 and numero<>0) order by tipodocumento,numero", MyConnObj, adOpenKeyset, adLockOptimistic
    
    'MyRecSet.Open "select * from dbo.documento where ((isnull(iscapturado,0)=0 or iscapturado=0) and (isfinalizado=1) and (numero<>0)) order by tipodocumento,numero", MyConnObj, adOpenKeyset, adLockOptimistic
    'MyRecSet.Open "select * from dbo.documento where ((isnull(iscapturado,0)=0 or iscapturado=0) and (isfinalizado=1)) order by tipodocumento,numero", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        'frmMonitor.barra1.Value = 0
        'frmMonitor.barra1.Max = MyRecSet.RecordCount
        MyRecSet.MoveFirst
        Do While Not MyRecSet.EOF
            Ffecha = MyRecSet.Fields("fecha")
            Fhora = Format(MyRecSet.Fields("hora"), "0000")
            Fhora = Left(Fhora, 2) & ":" & Right(Fhora, 2) & ":00"
            IdDoc = MyRecSet.Fields("iddocumento")
            ''''medio de pago
            MyRecAux.Open "select * from dbo.mediopagoDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                Cabecera.tipoPago = MyRecAux.Fields("idmediopago")
            Else
                Cabecera.tipoPago = 0
            End If
            MyRecAux.Close
            Set MyRecAux = Nothing
            '''''''''''''''''
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            'recoger datos
            Cabecera.Ntiquet = MyRecSet.Fields("numero")
            Cabecera.NVendedor = MyRecSet.Fields("idusuario")
            Cabecera.NMostrador = MyRecSet.Fields("idseccion")
            Cabecera.NCliente = MyRecSet.Fields("idcontacto")
            Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.hora = Fhora
            Cabecera.NBalanza = MyRecSet.Fields("numeroterminal")
            Cabecera.ImporteTotal = MyRecSet.Fields("totalneto")
            Cabecera.NLineas = nNumLin
            Cabecera.LineaInicial = 1
            Cabecera.LineaFinal = nNumLin
            Cabecera.LineasMensaje = nNumLin
            
            'cabecera.codigo ???
            ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
            '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
            If MyRecSet.Fields("islocal") Then
                Cabecera.CadenaEstadoTiquet = "Local"
            Else
                Cabecera.CadenaEstadoTiquet = "Net  "
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Cabecera.ImporteBruto = MyRecSet.Fields("totalbruto")
            Cabecera.Impuestos = MyRecSet.Fields("totalimpuestos")
            Cabecera.descuento = MyRecSet.Fields("totaldescuento")
            Cabecera.documentoabono = MyRecSet.Fields("iddocumentoabono")
            
            'frmMonitor.MostrarDato "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
            
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where (Idlineaabono = -1 and iddocumento=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.LineasCanceladas = nNumLin
            
            'Adapta TipoDocumento(ICG) a TypTic HydraTouch --> versión Junio
            'Select Case MyRecSet.Fields("tipodocumento")
            'Case 0
            '    If MyRecSet.Fields("islocal") = 1 Then
            '        Cabecera.tipoTiquet = 5
            '    Else
            '        Cabecera.tipoTiquet = 1
            '    End If
            'Case 2
            '    If MyRecSet.Fields("islocal") = 1 Then
            '        Cabecera.tipoTiquet = 7
            '    Else
            '        Cabecera.tipoTiquet = 3
            '    End If
            'Case 3
            '    If MyRecSet.Fields("islocal") = 1 Then
            '        Cabecera.tipoTiquet = 8
            '    Else
            '        Cabecera.tipoTiquet = 4
            '    End If
            'End Select
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
            Select Case MyRecSet.Fields("tipodocumento")
            Case 0
                If Cabecera.tipoPago = 7 Then 'si es pendiente no se considera cobrado sino tique...
                    If MyRecSet.Fields("islocal") = 1 Then
                        Cabecera.tipoTiquet = 4
                    Else
                        Cabecera.tipoTiquet = 1
                    End If
                Else
                    If MyRecSet.Fields("islocal") = 1 Then
                        Cabecera.tipoTiquet = 8
                    Else
                        Cabecera.tipoTiquet = 7
                    End If
                End If
            Case 1
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 4
                Else
                    Cabecera.tipoTiquet = 1
                End If
            Case 2
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 5
                Else
                    Cabecera.tipoTiquet = 2
                End If
            Case 3
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.tipoTiquet = 6
                Else
                    Cabecera.tipoTiquet = 3
                End If
            End Select
            
            
            'también se vé afectado si es cancelado o no pero de mamento cerrado...
            Cabecera.EstadoTiquet = 16
            '********************
            ' este dato se introduce
            ' tras pasar por actualiza_tabla_cabecera
            'Incluir_en_Base As Boolean
            'Cabecera.tipoPago = MyRecSet.Fields("idmoneda")
            Cabecera.cantidadEntre = MyRecSet.Fields("totalneto") + MyRecSet.Fields("importecambio")
            Cabecera.Cambio = MyRecSet.Fields("importecambio")
            'lote As String
            'lEncargo As Boolean


            'Guardar TIQUE si Cliente
            If nEPL <> 0 Then
                If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) And Cabecera.NCliente <> 0 Then
                    cm = "insert into dbo.documentos values (" & CStr(MyRecSet.Fields("iddocumento")) & "," & _
                                                                CStr(Cabecera.tipoTiquet) & "," & _
                                                                CStr(MyRecSet.Fields("numero")) & "," & _
                                                                CStr(MyRecSet.Fields("numeroterminal")) & "," & _
                                                                CStr(MyRecSet.Fields("idseccion")) & "," & _
                                                                CStr(MyRecSet.Fields("idcontacto")) & "," & _
                                                                "CURRENT_TIMESTAMP)"
                                                                '"'" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "')"
                    myConnTqt.Execute cm
                End If
            End If
            '''''''''''''''''''''''''
            
            'generar fichero en "\exporta"
            'Cabecera.Ntiquet = MyRecSet.Fields("numero")
            'Cabecera.NVendedor = MyRecSet.Fields("idusuario")
            'Cabecera.NMostrador = MyRecSet.Fields("idseccion")
            'Cabecera.NCliente = MyRecSet.Fields("idcontacto")
            'Nombre Fichero: ddmmyyTTTTTTVVVVYYYBBBXXXXXX.DAT
            '               ddmmyy = Fecha
            '               TTTTTT = Número Tique
            '               VVVV = Vendedor
            '               YYY = Mostrador
            '               BBB = Balanza
            '               XXXXXX = Cliente.
            sFich = App.Path & "\exporta\" & Format(Ffecha, "ddmmyy") & _
                  Format(Cabecera.Ntiquet, "000000") & Format(Cabecera.NVendedor, "0000") & _
                  Format(Cabecera.NMostrador, "000") & Format(Cabecera.NBalanza, "000") & Format(Cabecera.NCliente, "000000") & ".dat"
            nFich = FreeFile()
            Open sFich For Output As #nFich
            sParte = Format(Cabecera.Ntiquet, "000000") & Format(Cabecera.NVendedor, "0000") & _
            Format(Cabecera.NMostrador, "000") & Format(Cabecera.NCliente, "000000") & _
            Format(Ffecha, "ddmmyy") & Cabecera.hora & Format(Cabecera.NBalanza, "000") & _
            Format(Cabecera.ImporteTotal, "000000.000") & Format(Cabecera.NLineas, "000") & _
            Cabecera.CadenaEstadoTiquet & Format(Cabecera.ImporteBruto, "000000.000") & _
            Format(Cabecera.Impuestos, "000000.000") & Format(Cabecera.descuento, "000000.000") & _
            Format(Cabecera.documentoabono, "000000") & Format(Cabecera.LineasCanceladas, "000") & _
            Format(Cabecera.tipoTiquet, "0") & Format(Cabecera.cantidadEntre, "000000.000") & _
            Format(Cabecera.Cambio, "000000.000") & Format(Cabecera.tipoPago, "000")
            sParte = sAjusta(sParte)
            Print #nFich, sParte
            sParteCab = sParte
            ''''''''''''''''''''''''''''''
            nnFich = FreeFile()
            If Dir(App.Path & "\" & Format(Now(), "ddmmyy") & "tique.dat") <> "" Then
                Open App.Path & "\" & Format(Now(), "ddmmyy") & "tique.dat" For Append As #nnFich
            Else
                Open App.Path & "\" & Format(Now(), "ddmmyy") & "tique.dat" For Output As #nnFich
            End If
            
            'LINEAS
            nNumLin = 0
            MyRecAux.Open "select * from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF

                    nNumLin = nNumLin + 1
                    If MyRecAux.Fields("isventaporpeso") = True Then
                        Cabecera.Lines(nNumLin).EsPesado = True
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("peso")
                    Else
                        Cabecera.Lines(nNumLin).EsPesado = False
                        Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("unidades")
                    End If
                    Cabecera.Lines(nNumLin).FactorPeso = 1
                    Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("precio")
                    
                    Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                    Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("idarticulo")
                    
                    If nEPL <> 0 Then
                        'Actualizar Existencias....
                        'MyRec.Open "select * from dbo.existencias where IdArticulo =" & CStr(MyRecAux.Fields("idarticulo")) & " and isventaporpeso=" & MyRecAux.Fields("isventaporpeso"), MyConnObj, adOpenKeyset, adLockOptimistic
                        If MyRecAux.Fields("isventaporpeso") Then
                            MyRec.Open "select * from dbo.existencias where IdArticulo =" & CStr(MyRecAux.Fields("idarticulo")) & " and isventaporpeso=1", myConnTqt, adOpenKeyset, adLockOptimistic
                        Else
                            MyRec.Open "select * from dbo.existencias where IdArticulo =" & CStr(MyRecAux.Fields("idarticulo")) & " and isventaporpeso=0", myConnTqt, adOpenKeyset, adLockOptimistic
                        End If
                        
                        If MyRec.EOF Then
                            MyRec.AddNew
                            MyRec.Fields("idArticulo") = MyRecAux.Fields("idarticulo")
                            MyRec.Fields("isventaporpeso") = MyRecAux.Fields("isventaporpeso")
                            MyRec.Fields("existencias") = 0
                        End If
                        MyRec.Fields("existencias") = MyRec.Fields("existencias") - Cabecera.Lines(nNumLin).cantidad
                        MyRec.Fields("timestamp") = Now() 'Format(Now(), "yyyy-mm-dd hh:mm:ss") '"'" & Format(Now(), "yyyy-dd-mm") & " 00:00:00.000" & "'"
                        MyRec.Update
                        MyRec.Close
                        Set MyRec = Nothing
                        '''''''''''''''''''''''''''
                    End If
                    'ver qué se hace IVA As Integer
                    'ver qué se hace con TipoLinea As Byte
                    Cabecera.Lines(nNumLin).DescuentoPorcentaje = MyRecAux.Fields("descuento")  'porcentaje
                    Cabecera.Lines(nNumLin).ImporteLinea = MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                    Cabecera.Lines(nNumLin).ImporteBruto = MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                    'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                    If MyRecAux.Fields("idlineaabono") > 0 Then
                        Cabecera.Lines(nNumLin).LineaCancelada = True
                    Else
                        Cabecera.Lines(nNumLin).LineaCancelada = False
                    End If
                    If MyRecAux.Fields("precio") <> MyRecAux.Fields("preciodefecto") Then
                        Cabecera.Lines(nNumLin).SehaModificado = True
                    Else
                        Cabecera.Lines(nNumLin).SehaModificado = False
                    End If
                    '¿? ModificadoTemporal As Boolean
                    
                    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("idfichatrazabilidad")
                    'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                    Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                    
                    'Escribir línea
                    sParte = ""
                    If Cabecera.Lines(nNumLin).EsPesado = True Then
                        sParte = "W"
                    Else
                        sParte = "U"
                    End If
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).cantidad, "000000.000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).precio, "000000.000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).NumeroLinea, "000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).CodigoArticulo, "000000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).DescuentoPorcentaje, "000.000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).ImporteLinea, "000000.000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).ImporteBruto, "000000.000")
                    If Cabecera.Lines(nNumLin).LineaCancelada = True Then
                        sParte = sParte & "1"
                    Else
                        sParte = sParte & "0"
                    End If
                    If Cabecera.Lines(nNumLin).SehaModificado = False Then
                        sParte = sParte & "0"
                    Else
                        sParte = sParte & "1"
                    End If
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).FichaVacuno, "000000")
                    sParte = sParte & Format(Cabecera.Lines(nNumLin).DescuentoLinea, "000000.000")
                    sParte = sParte & Format(MyRecAux.Fields("lote"), "0000000000000000000000000")
                    sParte = sParte & Ajusta25(Mid(MyRecAux.Fields("descripcion"), 1, 25))
                    sParte = sAjusta(sParte)
                    Print #nFich, sParte
                    '''''''''''''''
                    
'Inicio   Fin     Contenido               Valores
'1        5       Número de Tiquet        00000 a 99999
'6        8       Número de Línea         001 a 999
'9        16      Fecha                   DD/MM/YY
'17       21      Hora                    HH:MM
'22       24      Sección Máquina         001 a 999
'25       28      Identificación Vendedor 0001 a 9999
'29       34      Código Artículo         000001 a 999999
'35       41      Precio Venta            Sin decimales 7 dig.
'42       42      Pesado / No pesado      0=Pesado, 1=No pesado
'43       49      Peso de la Venta        Sin decimales 7 dig.
'50       56      Unidades                Sin decimales 7 dig.
'57       63      Importe                 Sin decimales 7 dig.
'64       64      Tipo de Documento       (*)
'65       65      Linea Cancelada o no    0=No Cancelada
'                                         1=Cancelada    (***)
'66       90      Descriptivo artículo    25 Caracteres
'91       115     LOTE                    NºLote 25 Caracteres --> ES NUMÉRICO!!!!
'116      120     Código Cliente          00001 a 99999
'121      127     B.Importe               Importe Base.
'128      129     Forma de pago           (**)
'CR+LF                                    Fin Registro
'            ssParte = Format(Cabecera.Ntiquet, "00000") & Format(Cabecera.NVendedor, "0000") & _
'            Format(Cabecera.NMostrador, "000") & Format(Cabecera.NCliente, "000000") & _
'            Format(Ffecha, "ddmmyy")
                    'If Cabecera.Lines(nNumLin).LineaCancelada = False Then
                        ssParte = ""
                    '    ssParte = Format(Cabecera.Ntiquet, "00000") & Format(Cabecera.Lines(nNumLin).NumeroLinea, "000") & Format(Ffecha, "dd/mm/yy") & _
                    '            Mid(Cabecera.hora, 1, 5) & Format(Cabecera.NMostrador, "000") & Format(Cabecera.NVendedor, "0000") & _
                    '            Format(Cabecera.Lines(nNumLin).CodigoArticulo, "000000") & Format(Cabecera.Lines(nNumLin).precio * 100, "0000000")
                    '    If Cabecera.Lines(nNumLin).EsPesado = True Then
                    '        ssParte = ssParte & "0" & Format(Cabecera.Lines(nNumLin).cantidad * 1000, "0000000")
                    '    Else
                    '        ssParte = ssParte & "1" & Format(Cabecera.Lines(nNumLin).cantidad, "0000000")
                    '    End If
                    '    ssParte = ssParte & Format(Cabecera.Lines(nNumLin).ImporteLinea * 100, "0000000")
                    '    ssParte = ssParte & CStr(Cabecera.tipoTiquet) & "0" & Ajusta25(Mid(MyRecAux.Fields("descripcion"), 1, 25))
                    '    ssParte = ssParte & Format(MyRecAux.Fields("lote"), "0000000000000000000000000")
                    '    ssParte = ssParte & Format(Cabecera.NCliente, "00000") & Format(Cabecera.Lines(nNumLin).ImporteBruto * 100, "0000000")
                    '    ssParte = ssParte & Format(Cabecera.tipoPago, "00")
                        ssParte = sParteCab & sParte
                        Print #nnFich, ssParte
                    
                    'End If
                    
                    MyRecAux.MoveNext
                Loop
                
                Close #nFich
                Close #nnFich
                
                MyRecAux.Close
                Set MyRecAux = Nothing
                
                'Actualiza_Tabla_Cabecera MyBase, Cabecera
                'Do_Events
                'If Cabecera.Incluir_en_Base = True Then
                '    Actualiza_Tabla_Tickets MyBase, Cabecera
                '    Do_Events
                '    Actualiza_Tabla_Gtarti MyBase, Cabecera
                '    Do_Events
                '    Actualiza_Tabla_GtVend MyBase, Cabecera
                '    Do_Events
                '    actualiza_tabla_gtHora MyBase, Cabecera
                '    Do_Events
                '    Actualiza_Tabla_Gtsecs MyBase, Cabecera
                '    Do_Events
                '    actualiza_tabla_gtInfVen MyBase, Cabecera
                '    Do_Events
                '    Actualiza_Tabla_Log MyBase, Cabecera
                'End If
                
            End If



            '''''''''''''''
            'marcar...
            cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
            'cm = "update dbo.documento set iscapturado=1 where iddocumento=" & CStr(IdDoc)
            MyConnObj.Execute cm
            ''''''''''

            MyRecSet.Update

            MyRecSet.MoveNext
        
            'Call MonitorBarra
            'Call frmControl.Balanza_ComOK
            
            'If canCelar = 1 Or CancelHumano = True Then GoTo SalirLeeTique
            Do_Events
            
        Loop
    End If

    nResp = 0
    
SalirLeeTique:
''''''''
    If Err.Number <> 0 Then
        CadenadeLog "Error en LeeTickets:" & CStr(Err.Number) & Err.Description
        nResp = 1004
    End If
    On Error GoTo 0
    If nEPL <> 0 Then
        On Error Resume Next
        myConnTqt.Close
        Set myConnTqt = Nothing
        On Error GoTo 0
    End If

    On Error Resume Next
    MyRec.Close
    Set MyRec = Nothing
    On Error GoTo 0
    On Error Resume Next
    MyRecAux.Close
    Set MyRec = Nothing
    On Error GoTo 0
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    
    On Error GoTo 0

    LeeTickets = nResp
    
End Function

'/***********************************************/
Public Function Capa1A_SendItem_ScalePos(Item As Epel_Item) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim lEsAlta As Boolean
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim nPrice As Double
    
    lEsAlta = False
    
    PonDato "Capa1A_SendItem", 0
    
    If Item.Plu > 9999 Then
        Item.Plu = 0
        'PonDato "Se pasa PLU a 0 para código:" & CStr(Item.Code)
    End If
    
    On Error Resume Next
    PonDato "Familia...", 0
    MyRecAux.Open "select * from dbo.familia where idfamilia=" & CStr(Item.Family), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idfamilia") = Item.Family
        MyRecAux.Fields("descripcion") = "Familia " & CStr(Item.Family)
        If Dir(App.Path & "\images\families\" & CStr(Item.Family)) <> "" Then
            On Error GoTo noImgFam
            If FileLen(App.Path & "\images\families\" & CStr(Item.Family)) > 0 Then
            
                'Imagen
                strStream.type = adTypeBinary
                strStream.Open
                strStream.LoadFromFile App.Path & "\images\families\" & CStr(Item.Family)
                MyRecAux.Fields("imagen").Value = strStream.Read
                '''''''''''''''
                strStream.Flush
                strStream.Close
                Set strStream = Nothing
                ''''''''
            
            End If
noImgFam:
            If Err.Number <> 0 Then
                On Error GoTo 0
            End If
            
        End If
        MyRecAux.Fields("color") = "#FFCCCCCC"
        MyRecAux.Fields("codigo") = Item.Family
        MyRecAux.Update
        'si es alta...
        cm = "update dbo.familia set idglobal=NEWID() where idfamilia=" & CStr(Item.Family)
        MyConnObj.Execute cm
        ''''''''''''''
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    On Error Resume Next
    PonDato "Tarifa...", 0
    MyRecAux.Open "select * from dbo.tarifa where idtarifa=1", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idtarifa") = 1
        MyRecAux.Fields("descripcion") = "Tarifa Defecto"
        MyRecAux.Update
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    On Error Resume Next
    PonDato "Articulo...", 0
    MyRecAux.Open "select * from dbo.articulo where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        lEsAlta = True
    End If
    MyRecAux.Fields("idarticulo") = Item.Code
    MyRecAux.Fields("codigo") = Item.Code
    MyRecAux.Fields("descripcion") = Item.Text
    MyRecAux.Fields("impuestoventa") = 0 ' 155 + Item.vat
    MyRecAux.Fields("impuestotakeaway") = 0 '155 + Item.vat
    MyRecAux.Fields("impuestocompra") = 0
    'If Item.weight <> 0 Then
        MyRecAux.Fields("tipoventa") = Item.weight
    'Else
    '    MyRecAux.Fields("tipoventa") = 1
    'End If
    MyRecAux.Fields("taraventa") = Item.Tare / 1000
    MyRecAux.Fields("taraenvasado") = Item.Tare / 1000
    If Item.Caducity <> 0 Then
        MyRecAux.Fields("caducidad") = Item.Caducity
    End If
    MyRecAux.Fields("consumopreferente") = Item.Pref
    MyRecAux.Fields("pesoaproximado") = 0 'Item.nPeso
    MyRecAux.Fields("isdescatalogado") = 0
    MyRecAux.Fields("isidarticulodefinitivo") = 0
    MyRecAux.Fields("merma") = 0
    MyRecAux.Fields("tipopreseleccion") = 0
    MyRecAux.Fields("unidadespreseleccion") = 0
    MyRecAux.Fields("tipoetiqueta") = Item.Etq
    If lEsAlta Then
        MyRecAux.Fields("color") = "#FF9BBF5F"
    End If
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    If lEsAlta Then
        'si es alta...
        cm = "update dbo.articulo set idglobal=NEWID() where idarticulo=" & CStr(Item.Code)
        MyConnObj.Execute cm
        ''''''''''''''
    End If
'dbo.articuloimagen where idarticulo=" & CStr(Code) --> No pasan el dato. Si existe <Images\Items\code> se envia
'                           IdArticulo
'                           Imagen
    If Dir(App.Path & "\images\items\" & CStr(Item.Code)) <> "" Then
        On Error GoTo noImage
        If FileLen(App.Path & "\images\items\" & CStr(Item.Code)) > 0 Then
        
            MyRecAux.Open "select * from dbo.articuloimagen where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
            On Error Resume Next
            PonDato "Imagen Artículo...", 0
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo FIN_Capa1A_SendItem
            End If
            On Error GoTo FIN_Capa1A_SendItem
            
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            MyRecAux.Fields("idarticulo") = Item.Code
            
            'Imagen
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\images\items\" & CStr(Item.Code)
            MyRecAux.Fields("imagen").Value = strStream.Read
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
        
        End If
noImage:
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
    End If
'dbo.articulosfamilia where idarticulo=" & CStr(Code)
'                           IdFamilia
'                           IdArticulo
'                           Posicion
    On Error Resume Next
    PonDato "articulosfamilia...", 0
    MyRecAux.Open "select * from dbo.articulosfamilia where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("posicion") = 1000000
    End If
    MyRecAux.Fields("idarticulo") = Item.Code
    MyRecAux.Fields("idfamilia") = Item.Family
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
'dbo.articuloreferencia where idarticulo=" & CStr(Code)
'                               Referencia
'                               IdSeccion
'                               IdArticulo
    On Error Resume Next
    PonDato "Referencia...", 0
    If Item.Plu <> 0 Then
        MyRecAux.Open "select * from dbo.articuloreferencia where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Capa1A_SendItem
        End If
        On Error GoTo FIN_Capa1A_SendItem
        
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        End If
        MyRecAux.Fields("idarticulo") = Item.Code
        MyRecAux.Fields("idseccion") = Item.Sec
        MyRecAux.Fields("referencia") = Item.Plu
        MyRecAux.Update
        MyRecAux.Close
        Set MyRecAux = Nothing
    End If
'dbo.precio where idarticulo=" & CStr(Code)
'                   IdPrecio
'                   IdTarifa
'                   IdArticulo
'                   UnidadesMinimo
'                   UnidadesMaximo
'                   PrecioDefecto
'                   FechaInicioOferta
'                   FechaFinOferta
'                   PrecioOferta
    On Error Resume Next
    PonDato "PK precio", 0
    MyRecAux.Open "select COUNT(*) as numprecios from dbo.precio", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    nC = MyRecAux.Fields("numprecios")
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    On Error Resume Next
    PonDato "Precio...", 0
    MyRecAux.Open "select * from dbo.precio where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_SendItem
    End If
    On Error GoTo FIN_Capa1A_SendItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idprecio") = nC + 1
    End If
    MyRecAux.Fields("idarticulo") = Item.Code
    MyRecAux.Fields("idtarifa") = 1
    MyRecAux.Fields("unidadesminimo") = 0
    MyRecAux.Fields("unidadesmaximo") = 999999
    nPrice = Round(Item.Price / 100, 2)
    MyRecAux.Fields("preciodefecto") = nPrice
    MyRecAux.Fields("preciooferta") = 0
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    Resp = 0
    
FIN_Capa1A_SendItem:
    If Err.Number <> 0 Then
        CadenadeLog "Error en SendItem. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    
    PonDato "Articulo Cod: " & Item.Code & " Plu: " & Item.Plu & " Sec: " & Item.Sec & " Desc: " & Item.Text, Resp
    Capa1A_SendItem_ScalePos = Resp
End Function

Public Function Capa1A_GA_Erase_Totals_ScalePos(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para ScalePos  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
'dbo.acumuladoventas
'           IdBalanza
'           TipoDocumento
'           Fecha
'           ejemplo introducción Fecha...
'           -->cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:00.000" & "'" & " where iddocumento=" & CStr(IdDoc)
'           Importe
'           Operaciones
'           Unidades
'           Peso
'''''''''''''''''''''
'dbo.balanza
'           IdBalanza <-- Capturar para IdBalanza
'           NombreTerminal
'           Numero
'           DireccionIP
'           TipoConexion
'           DireccionIPServidor
'           IsConectadoCentral
'           DireccionIPCentral
'           NombreDatabaseCentral
'           ConexionSQL
    
    
'select A1.TipoDocumento tipdoc, sum(A2.unidades) uni, sum(A2.peso) pes, SUM(a2.importeconimpuestos) imp, COUNT(distinct(a2.IdTransaction)) ope from dbo.Documento A1, dbo.LineaDocumento A2 where (A1.IsFinalizado = 1 and a2.IdDocumento = a1.IdDocumento) group by A1.tipodocumento
   
    PonDato "Capa1A_GA_Erase_totals", 0
    On Error Resume Next
    PonDato "Identificación Balanza...", 0
    MyRecAux.Open "select * from dbo.balanza", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        CadenadeLog "No se ha podido abrir la conexión con dbo.Balanza..."
        Resp = 1000
        GoTo FIN_Capa1A_GA_Erase_Totals
    End If
    On Error GoTo FIN_Capa1A_GA_Erase_Totals
    If Not MyRecAux.EOF Then
        'ejemplo:FB9F5AA7-3AB9-46A4-8EDE-F1B57957C5B5
        'CadenadeLog "Captura idbalanza"
        IdEqp = MyRecAux.Fields("idbalanza")
    Else
        IdEqp = "00000000-0000-0000-0000-000000000000"
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    'On Error Resume Next
    On Error GoTo FIN_Capa1A_GA_Erase_Totals
    PonDato "Selección por tipo documento...", 0
    MyRecAux1.Open "select A1.TipoDocumento tipdoc, sum(A2.unidades) uni, sum(A2.peso) pes, SUM(a2.importeconimpuestos) imp, COUNT(distinct(a2.IdTransaction)) ope from dbo.Documento A1, dbo.LineaDocumento A2 where (A1.IsFinalizado = 1 and A2.IdDocumento = A1.IdDocumento) group by A1.tipodocumento", MyConnObj, adOpenKeyset, adLockOptimistic
    CadenadeLog "Tipos de documento seleccionados..."
    On Error GoTo FIN_Capa1A_GA_Erase_Totals
    If Not MyRecAux1.EOF Then
        MyRecAux1.MoveFirst
        Do While Not MyRecAux1.EOF
            'MyRecSet.Open "select * from dbo.acumuladoventas", MyConnObj, adOpenKeyset, adLockOptimistic
            'MyRecSet.AddNew
            CadenadeLog "añadido registro en Acumuladoventas."
            'MyRecSet.Fields("idbalanza") = IdEqp
            'MyRecSet.Fields("tipodocumento") = MyRecAux1.Fields("tipdoc")
            'MyRecSet.Fields("fecha") = CDate(Now()) 'Format(Now(), "yyyy-dd-mm") & " 00:00:00.000"
            'MyRecSet.Fields("importe") = MyRecAux1.Fields("imp")
            'MyRecSet.Fields("operaciones") = MyRecAux1.Fields("ope")
            'MyRecSet.Fields("unidades") = MyRecAux1.Fields("uni")
            'MyRecSet.Fields("peso") = MyRecAux1.Fields("pes")
            'MyRecSet.Update
            ''IdBalanza, TipoDocumento, Fecha, Importe, Operaciones, Unidades, Peso
            'cm = "insert into dbo.acumuladoventas values ('" & IdEqp & "'," & CStr(MyRecAux1.Fields("tipdoc")) & "," & _
            '                                                "CURRENT_TIMESTAMP" & "," & CStr(MyRecAux1.Fields("imp")) & "," & _
            '                                                CStr(MyRecAux1.Fields("ope")) & "," & CStr(MyRecAux1.Fields("uni")) & "," & _
            '                                                CStr(MyRecAux1.Fields("pes")) & ")"
    cm = "insert into dbo.acumuladoventas values ('" & _
         IdEqp & "'," & CStr(MyRecAux1.Fields("tipdoc")) & ",convert(date,'" & Format(Now(), "yyyy-mm-dd") & " 00:00:00.000" & "')," & Quitar_Coma(CStr(MyRecAux1.Fields("imp")), Len(CStr(MyRecAux1.Fields("imp")))) & "," & _
         CStr(MyRecAux1.Fields("ope")) & "," & CStr(MyRecAux1.Fields("uni")) & "," & Quitar_Coma(CStr(MyRecAux1.Fields("pes")), Len(CStr(MyRecAux1.Fields("pes")))) & ")"
            
            MyConnObj.Execute cm
          
            MyRecAux1.MoveNext
        Loop
    End If
    'MyRecSet.Close
    'Set MyRecSet = Nothing
    MyRecAux1.Close
    Set MyRecAux1 = Nothing
    
    PonDato "Iniciar Numeración si Procede...", 0
    If No_Reinit = 0 Then 'Iniciar numeración tiques...
        cm = "truncate table dbo.numeraciondocumento"
        MyConnObj.Execute cm
    End If
    PonDato "Borrado Documentos y Lineas Documento...", 0
    'cm = "truncate table dbo.documentoactivo" '
    'MyConnObj.Execute cm
    'cm = "truncate table dbo.MedioPagoDocumento" '
    'MyConnObj.Execute cm
    'cm = "truncate table dbo.lineadocumento"
    'MyConnObj.Execute cm
    'cm = "truncate table dbo.impuestodocumento" '
    'MyConnObj.Execute cm
    cm = "delete from dbo.documento"
    MyConnObj.Execute cm
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Capa1A_GA_Erase_Totals_ScalePos = Resp
End Function

Public Function Capa1A_SendVendor_ScalePos(IDV As Long, sName As String, nSection As Long, nKey As Long, sImgVnd As String, Status As Long, EraseVendor As Long) As Long
    '///////////////////////////////////////
    '// Envío de vendedores (ScalePos)    //
    '///////////////////////////////////////
'dbo.usuario
'       IdUsuario
'       IdGlobal
'       Nombre
'       PasswordRegistro
'       IsDescatalogado
'       Codigo
'1   1BF1F877-AD9E-465C-8914-F2BC946D4607    belinda NULL    0   1
'2   B90E23EE-85E4-41A7-9027-10A7BABFC956    catherine   NULL    0   2
'3   FCCB6063-F11F-4394-BE5B-7735DE890DE8    collin  NULL    0   3
'4   7807E852-B288-4747-9662-F5C2B15A400B    david   NULL    0   4
'5   AB519A61-712E-4424-AD15-137C89B1AFD2    elsa    NULL    0   5
'6   FA8E878A-918F-4632-8CCD-9B49B6A675DE    eva NULL    0   6
'7   6A5EF8A6-8C7E-40B4-97EE-179E80967243    george  NULL    0   7
'8   AAAC25F3-5245-4C0A-892E-4508EF121D57    jessica NULL    0   8
'9   B1DAE856-C1C7-4EE8-A2C6-1648ED4E3FB0    reese   NULL    0   9
'10  83E537E8-F1A1-45BD-81C1-95E221A9E280    sarah   NULL    0   10
'dbo.UsuarioImagen
'       IdUsuario
'       Imagen
'dbo.UsuariosSeccion
'       IdUsuario
'       IdSeccion
'       Posicion
'1   1   0
'2   1   1
'3   1   2
'4   1   3
'5   1   4
'6   1   5
'7   1   6
'8   1   7
'9   1   8
'10  1   9

    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    Dim lEsAlta As Boolean
    Dim cm As String
    
    On Error GoTo FIN_Error_SendVendor
    
    If EraseVendor = 1 Then
        On Error Resume Next
        cm = "delete dbo.usuario where idusuario=" & CStr(IDV)
        MyConnObj.Execute cm
        If Err.Number <> 0 Then
            On Error GoTo 0
            On Error Resume Next
        End If
        cm = "delete dbo.usuarioimagen where idusuario=" & CStr(IDV)
        MyConnObj.Execute cm
        If Err.Number <> 0 Then
            On Error GoTo 0
            On Error Resume Next
        End If
        cm = "delete dbo.usuariosseccion where idusuario=" & CStr(IDV)
        MyConnObj.Execute cm
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Error_SendVendor
        End If
    Else
        On Error Resume Next
        MyRecAux.Open "select * from dbo.usuario where idusuario=" & CStr(IDV), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Error_SendVendor
        End If
        On Error GoTo FIN_Error_SendVendor
        If MyRecAux.EOF Then
            lEsAlta = True
            MyRecAux.AddNew
        Else
            lEsAlta = False
        End If
'       IdUsuario
'       IdGlobal
'       Nombre
'       PasswordRegistro
'       IsDescatalogado
'       Codigo
        MyRecAux.Fields("idusuario") = IDV
        MyRecAux.Fields("nombre") = sName
        If lEsAlta Then
            MyRecAux.Fields("isdescatalogado") = 0
        End If
        MyRecAux.Fields("codigo") = IDV
        MyRecAux.Update
        If lEsAlta Then
            cm = "update dbo.usuario set idglobal=NEWID() where idusuario=" & CStr(IDV)
            MyConnObj.Execute cm
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
        
        On Error Resume Next
        MyRecAux.Open "select * from dbo.usuarioimagen where idusuario=" & CStr(IDV), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Error_SendVendor
        End If
        On Error GoTo FIN_Error_SendVendor
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        End If
        MyRecAux.Fields("idusuario") = IDV
        If Dir(App.Path & "\images\vendors\" & Trim(sImgVnd)) <> "" And IsNumeric(Trim(sImgVnd)) Then
            'Imagen
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\images\vendors\" & Trim(sImgVnd)
            MyRecAux.Fields("imagen").Value = strStream.Read
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
        Else
            If Dir(App.Path & "\vndimg.jpg") <> "" Then
                'Imagen
                strStream.type = adTypeBinary
                strStream.Open
                strStream.LoadFromFile App.Path & "\vndimg.jpg"
                MyRecAux.Fields("imagen").Value = strStream.Read
                '''''''''''''''
                strStream.Flush
                strStream.Close
                Set strStream = Nothing
                ''''''''
            Else
                'If IsNumeric(sImgVnd) Then
                '    If Val(sImgVnd) < 33 Then
                '        If Dir(App.Path & "\nn" & Format(Val(sImgVnd), "00") & ".jpg") <> "" Then
                '            If FileLen(App.Path & "\nn" & Format(Val(sImgVnd), "00") & ".jpg") > 0 Then
                '                'Imagen
                '                strStream.type = adTypeBinary
                '                strStream.Open
                '                strStream.LoadFromFile App.Path & "\nn" & Format(Val(sImgVnd), "00") & ".jpg"
                '                MyRecAux.Fields("imagen").Value = strStream.Read
                '                '''''''''''''''
                '                strStream.Flush
                '                strStream.Close
                '                Set strStream = Nothing
                '                ''''''''
                '            End If
                '
                '        End If
                '    End If
                'Else
                    If Trim(sImgVnd) <> "" Then
                        If Dir(sImgVnd) <> "" Then
                            If FileLen(sImgVnd) > 0 Then
                                'Imagen
                                strStream.type = adTypeBinary
                                strStream.Open
                                strStream.LoadFromFile Trim(sImgVnd)
                                MyRecAux.Fields("imagen").Value = strStream.Read
                                '''''''''''''''
                                strStream.Flush
                                strStream.Close
                                Set strStream = Nothing
                                ''''''''
                            End If
                        End If
                    End If
                'End If
            End If
        End If
        MyRecAux.Update
        
        MyRecAux.Close
        Set MyRecAux = Nothing
        Set MyRecAux = Nothing
        
        If lEsAlta Then
            On Error Resume Next
            MyRecAux.Open "select * from dbo.usuariosseccion where idusuario=" & CStr(IDV), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo FIN_Error_SendVendor
            End If
            On Error GoTo FIN_Error_SendVendor
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            MyRecAux.Fields("idusuario") = IDV
            MyRecAux.Fields("idseccion") = nSection
            MyRecAux.Fields("posicion") = nKey
            MyRecAux.Update
            
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If
    End If

    Resp = 0
    
FIN_Error_SendVendor:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_sendvendor. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1000
    Else
        If Resp = 0 Then
            If EraseVendor = 1 Then
                CadenadeLog "Baja Vendedor:" & CStr(IDV) & " Tecla:" & CStr(nKey) & " --> OK"
            Else
                CadenadeLog "Programación Vendedor:" & CStr(IDV) & " Tecla:" & CStr(nKey) & " --> OK"
            End If
        End If
    End If
    
    Capa1A_SendVendor_ScalePos = Resp

End Function

Public Function capa1A_sendVat_ScalePos(vat As Epel_Vat) As Long
'dbo.TipoImpuesto
'       IdTipoImpuesto
'       Descripcion
'       Porcentaje
'       AplicarSobreAcumulado
'155 IVA 18% 18.0000 0
'156 IVA 8%  8.0000  0
'157 IVA 4%  4.0000  0
'158 IVA 0%  0.0000  0
'dbo.Impuesto
'       IdImpuesto
'       Descripcion
'       IdTipoImpuesto1
'       IdTipoImpuesto2
'       IdTipoImpuesto3
'155 IVA 18% 155 NULL    NULL
'156 IVA 8%  156 NULL    NULL
'157 IVA 4%  157 NULL    NULL
'158 IVA 0%  158 NULL    NULL
    
'Se actuará sobre la Tabla dbo.tipoimpuesto
'   de forma que tipo 0 parámetro será IdImpuesto 155
'                     1                           156
'                     2                           157
'                     3                           158
'Public Type Epel_Vat
'    '//////////////////
'    '// Tipos de IVA //
'    '//////////////////
'    type1 As Long
'    type2 As Long
'    type3 As Long
'    type4 As Long
'    type5 As Long
'End Type
'-->String Parámetro: 18000800040000000000
    
    Dim MyRecAux As New ADODB.Recordset
    Dim Resp As Long
    Dim nImp155 As Double
    Dim nImp156 As Double
    Dim nImp157 As Double
    Dim nImp158 As Double
    Dim cm As String
    
    nImp155 = Round(vat.type1 / 100, 2)
    nImp156 = Round(vat.type2 / 100, 2)
    nImp157 = Round(vat.type3 / 100, 2)
    nImp158 = Round(vat.type4 / 100, 2)
    
    PonDato "Capa1A_SendVat", 0
    
    On Error Resume Next
    MyRecAux.Open "select * from dbo.tipoimpuesto", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1A_sendVat
    End If
    On Error GoTo FIN_Capa1A_sendVat
'dbo.TipoImpuesto
'       IdTipoImpuesto
'       Descripcion
'       Porcentaje
'       AplicarSobreAcumulado
    If Not MyRecAux.EOF Then
        cm = "update dbo.tipoimpuesto set porcentaje=" & CStr(nImp155) & ",descripcion='" & "IVA " & CStr(nImp155) & "' where idtipoimpuesto=155"
        MyConnObj.Execute cm
        cm = "update dbo.tipoimpuesto set porcentaje=" & CStr(nImp156) & ",descripcion='" & "IVA " & CStr(nImp156) & "' where idtipoimpuesto=156"
        MyConnObj.Execute cm
        cm = "update dbo.tipoimpuesto set porcentaje=" & CStr(nImp157) & ",descripcion='" & "IVA " & CStr(nImp157) & "' where idtipoimpuesto=157"
        MyConnObj.Execute cm
        cm = "update dbo.tipoimpuesto set porcentaje=" & CStr(nImp158) & ",descripcion='" & "IVA " & CStr(nImp158) & "' where idtipoimpuesto=158"
        MyConnObj.Execute cm
        
        cm = "update dbo.impuesto set descripcion='" & "IVA " & CStr(nImp155) & "' where idimpuesto=155"
        MyConnObj.Execute cm
        cm = "update dbo.impuesto set descripcion='" & "IVA " & CStr(nImp156) & "' where idimpuesto=156"
        MyConnObj.Execute cm
        cm = "update dbo.impuesto set descripcion='" & "IVA " & CStr(nImp157) & "' where idimpuesto=157"
        MyConnObj.Execute cm
        cm = "update dbo.impuesto set descripcion='" & "IVA " & CStr(nImp158) & "' where idimpuesto=158"
        MyConnObj.Execute cm
        
        Resp = 0
    Else
        Resp = 1002
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing

FIN_Capa1A_sendVat:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_sendvat. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1000
    End If
    capa1A_sendVat_ScalePos = Resp
    
End Function

'*************************************************
'//////////////////////////////////////
'//Eliminación (Borrado de Artículo...)
'///////////////////////////////////////
Public Function Capa1B_EliArt_ScalePos(ByVal Code As Long, ByVal Sec As Long, ByVal Plu As Long, ByVal cb As String) As Long
Dim cm As String
Dim MyRecAux As New ADODB.Recordset
    
    On Error GoTo EliartError
    
    If Code = 0 Then
        If cb <> "" And Sec = 0 Then
            MyRecAux.Open "select * from dbo.codigobarra where codigobarra='" & Trim(cb) & "'", MyConnObj, adOpenKeyset, adLockOptimistic
            If MyRecAux.EOF Then
                MyRecAux.Close
                Set MyRecAux = Nothing
                Capa1B_EliArt_ScalePos = 1004
                Exit Function
            Else
                Code = MyRecAux.Fields("idarticulo")
                MyRecAux.Close
                Set MyRecAux = Nothing
            End If
        
        Else
            If Sec <> 0 And Plu <> 0 Then
                MyRecAux.Open "select * from dbo.articuloreferencia where idseccion=" & CStr(Sec) & " and referencia=" & CStr(Plu), MyConnObj, adOpenKeyset, adLockOptimistic
                If MyRecAux.EOF Then
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                    Capa1B_EliArt_ScalePos = 1004
                    Exit Function
                Else
                    Code = MyRecAux.Fields("idarticulo")
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                End If
            Else
                Capa1B_EliArt_ScalePos = 1004
                Exit Function
            End If
        
        End If
    End If
    
    cm = "delete dbo.articuloimagen where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.articulo where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.articulosfamilia where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.articuloreferencia where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.precio where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.ArticuloDescripcion where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm

    cm = "delete dbo.CodigoBarra where idarticulo=" & CStr(Code)
    MyConnObj.Execute cm
    
    CadenadeLog "Eliminado Artículo código:" & CStr(Code)
    On Error GoTo 0
    Capa1B_EliArt_ScalePos = 0
    Exit Function

EliartError:
    CadenadeLog "Error eliminando Artículo:" & CStr(Code) & " Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error GoTo 0
    Capa1B_EliArt_ScalePos = 1004
    
End Function

Public Function Capa1B_Erase_Card_ScalePos(Cnumber As Long) As Long
'/////////////////////////////////////////////////
'// Borra los datos de una ficha de vacuno      //
'/////////////////////////////////////////////////
Dim cm As String

    On Error GoTo EraseCardError
    cm = "delete dbo.traza_campoficha where idficha=" & CStr(Cnumber)
    MyConnObj.Execute cm

    On Error GoTo 0
    Capa1B_Erase_Card_ScalePos = 0
    Exit Function
    
EraseCardError:
    CadenadeLog "Error eliminando Ficha Vacuno:" & CStr(Cnumber) & " Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error GoTo 0
    Capa1B_Erase_Card_ScalePos = 1004
    
End Function
Public Function Capa1B_SendCountry_ScalePos(cCode As Long, sName As String, EraseCountry As Long) As Long
'///////////////////////////////////////////////////
'// SE ENVIAN DATOS a "CAPÓN"...
'///////////////////////////////////////////////////
Dim cm As String
Dim nC As Integer
Dim nC1 As Integer
Dim nC2 As Integer
Dim sValor As String
Dim sComentario As String

    On Error GoTo errorPais
    
    cm = "delete dbo.traza_valorlookup"
    MyConnObj.Execute cm
    MyRecSet.Open "Select * from dbo.traza_valorlookup", MyConnObj, adOpenKeyset, adLockOptimistic
    
    nC = 1
    For nC2 = 1 To 999
        Select Case nC2
            Case 248
                sValor = "AALAND ISLANDS"
                sComentario = "248-AX-ALA-AALAND ISLANDS"
            Case 4
                sValor = "AFGHANISTAN"
                sComentario = "004-AF-AFG-AFGHANISTAN"
            Case 8
                sValor = "ALBANIA"
                sComentario = "008-AL-ALB-ALBANIA"
            Case 12
                sValor = "ALGERIA"
                sComentario = "012-DZ-DZA-ALGERIA"
            Case 16
                sValor = "AMERICAN SAMOA"
                sComentario = "016-AS-ASM-AMERICAN SAMOA"
            Case 20
                sValor = "ANDORRA"
                sComentario = "020-AD-AND-ANDORRA"
            Case 24
                sValor = "ANGOLA"
                sComentario = "024-AO-AGO-ANGOLA"
            Case 660
                sValor = "ANGUILLA"
                sComentario = "660-AI-AIA-ANGUILLA"
            Case 10
                sValor = "ANTARCTICA"
                sComentario = "010-AQ-ATA-ANTARCTICA"
            Case 28
                sValor = "ANTIGUA AND BARBUDA"
                sComentario = "028-AG-ATG-ANTIGUA AND BARBUDA"
            Case 32
                sValor = "ARGENTINA"
                sComentario = "032-AR-ARG-ARGENTINA"
            Case 51
                sValor = "ARMENIA"
                sComentario = "051-AM-ARM-ARMENIA"
            Case 533
                sValor = "ARUBA"
                sComentario = "533-AW-ABW-ARUBA"
            Case 36
                sValor = "AUSTRALIA"
                sComentario = "036-AU-AUS-AUSTRALIA"
            Case 40
                sValor = "AUSTRIA"
                sComentario = "040-AT-AUT-AUSTRIA"
            Case 31
                sValor = "AZERBAIJAN"
                sComentario = "031-AZ-AZE-AZERBAIJAN"
            Case 44
                sValor = "BAHAMAS"
                sComentario = "044-BS-BHS-BAHAMAS"
            Case 48
                sValor = "BAHRAIN"
                sComentario = "048-BH-BHR-BAHRAIN"
            Case 50
                sValor = "BANGLADESH"
                sComentario = "050-BD-BGD-BANGLADESH"
            Case 52
                sValor = "BARBADOS"
                sComentario = "052-BB-BRB-BARBADOS"
            Case 112
                sValor = "BELARUS"
                sComentario = "112-BY-BLR-BELARUS"
            Case 56
                sValor = "BELGIUM"
                sComentario = "056-BE-BEL-BELGIUM"
            Case 84
                sValor = "BELIZE"
                sComentario = "084-BZ-BLZ-BELIZE"
            Case 204
                sValor = "BENIN"
                sComentario = "204-BJ-BEN-BENIN"
            Case 60
                sValor = "BERMUDA"
                sComentario = "060-BM-BMU-BERMUDA"
            Case 64
                sValor = "BHUTAN"
                sComentario = "064-BT-BTN-BHUTAN"
            Case 68
                sValor = "BOLIVIA"
                sComentario = "068-BO-BOL-BOLIVIA"
            Case 70
                sValor = "BOSNIA AND HERZEGOWINA"
                sComentario = "070-BA-BIH-BOSNIA AND HERZEGOWINA"
            Case 72
                sValor = "BOTSWANA"
                sComentario = "072-BW-BWA-BOTSWANA"
            Case 74
                sValor = "BOUVET ISLAND"
                sComentario = "074-BV-BVT-BOUVET ISLAND"
            Case 76
                sValor = "BRAZIL"
                sComentario = "076-BR-BRA-BRAZIL"
            Case 86
                sValor = "BRITISH INDIAN OCEAN TERRITORY"
                sComentario = "086-IO-IOT-BRITISH INDIAN OCEAN TERRITORY"
            Case 96
                sValor = "BRUNEI DARUSSALAM"
                sComentario = "096-BN-BRN-BRUNEI DARUSSALAM"
            Case 100
                sValor = "BULGARIA"
                sComentario = "100-BG-BGR-BULGARIA"
            Case 854
                sValor = "BURKINA FASO"
                sComentario = "854-BF-BFA-BURKINA FASO"
            Case 108
                sValor = "BURUNDI"
                sComentario = "108-BI-BDI-BURUNDI"
            Case 116
                sValor = "CAMBODIA"
                sComentario = "116-KH-KHM-CAMBODIA"
            Case 120
                sValor = "CAMEROON"
                sComentario = "120-CM-CMR-CAMEROON"
            Case 124
                sValor = "CANADA"
                sComentario = "124-CA-CAN-CANADA"
            Case 132
                sValor = "CAPE VERDE"
                sComentario = "132-CV-CPV-CAPE VERDE"
            Case 136
                sValor = "CAYMAN ISLANDS"
                sComentario = "136-KY-CYM-CAYMAN ISLANDS"
            Case 140
                sValor = "CENTRAL AFRICAN REPUBLIC"
                sComentario = "140-CF-CAF-CENTRAL AFRICAN REPUBLIC"
            Case 148
                sValor = "CHAD"
                sComentario = "148-TD-TCD-CHAD"
            Case 152
                sValor = "CHILE"
                sComentario = "152-CL-CHL-CHILE"
            Case 156
                sValor = "CHINA"
                sComentario = "156-CN-CHN-CHINA"
            Case 162
                sValor = "CHRISTMAS ISLAND"
                sComentario = "162-CX-CXR-CHRISTMAS ISLAND"
            Case 166
                sValor = "COCOS (KEELING) ISLANDS"
                sComentario = "166-CC-CCK-COCOS (KEELING) ISLANDS"
            Case 170
                sValor = "COLOMBIA"
                sComentario = "170-CO-COL-COLOMBIA"
            Case 174
                sValor = "COMOROS"
                sComentario = "174-KM-COM-COMOROS"
            Case 180
                sValor = "CONGO, Democratic Republic of (was Zaire)"
                sComentario = "180-CD-COD-CONGO, Democratic Republic of (was Zaire)"
            Case 178
                sValor = "CONGO, Republic of"
                sComentario = "178-CG-COG-CONGO, Republic of"
            Case 184
                sValor = "COOK ISLANDS"
                sComentario = "184-CK-COK-COOK ISLANDS"
            Case 188
                sValor = "COSTA RICA"
                sComentario = "188-CR-CRI-COSTA RICA"
            Case 384
                sValor = "COTE D'IVOIRE"
                sComentario = "384-CI-CIV-COTE D'IVOIRE"
            Case 191
                sValor = "CROATIA (local name: Hrvatska)"
                sComentario = "191-HR-HRV-CROATIA (local name: Hrvatska)"
            Case 192
                sValor = "CUBA"
                sComentario = "192-CU-CUB-CUBA"
            Case 196
                sValor = "CYPRUS"
                sComentario = "196-CY-CYP-CYPRUS"
            Case 203
                sValor = "CZECH REPUBLIC"
                sComentario = "203-CZ-CZE-CZECH REPUBLIC"
            Case 208
                sValor = "DENMARK"
                sComentario = "208-DK-DNK-DENMARK"
            Case 262
                sValor = "DJIBOUTI"
                sComentario = "262-DJ-DJI-DJIBOUTI"
            Case 212
                sValor = "DOMINICA"
                sComentario = "212-DM-DMA-DOMINICA"
            Case 214
                sValor = "DOMINICAN REPUBLIC"
                sComentario = "214-DO-DOM-DOMINICAN REPUBLIC"
            Case 218
                sValor = "ECUADOR"
                sComentario = "218-EC-ECU-ECUADOR"
            Case 818
                sValor = "EGYPT"
                sComentario = "818-EG-EGY-EGYPT"
            Case 222
                sValor = "EL SALVADOR"
                sComentario = "222-SV-SLV-EL SALVADOR"
            Case 226
                sValor = "EQUATORIAL GUINEA"
                sComentario = "226-GQ-GNQ-EQUATORIAL GUINEA"
            Case 232
                sValor = "ERITREA"
                sComentario = "232-ER-ERI-ERITREA"
            Case 233
                sValor = "ESTONIA"
                sComentario = "233-EE-EST-ESTONIA"
            Case 231
                sValor = "ETHIOPIA"
                sComentario = "231-ET-ETH-ETHIOPIA"
            Case 238
                sValor = "FALKLAND ISLANDS (MALVINAS)"
                sComentario = "238-FK-FLK-FALKLAND ISLANDS (MALVINAS)"
            Case 234
                sValor = "FAROE ISLANDS"
                sComentario = "234-FO-FRO-FAROE ISLANDS"
            Case 242
                sValor = "FIJI"
                sComentario = "242-FJ-FJI-FIJI"
            Case 246
                sValor = "FINLAND"
                sComentario = "246-FI-FIN-FINLAND"
            Case 250
                sValor = "FRANCE"
                sComentario = "250-FR-FRA-FRANCE"
            Case 254
                sValor = "FRENCH GUIANA"
                sComentario = "254-GF-GUF-FRENCH GUIANA"
            Case 258
                sValor = "FRENCH POLYNESIA"
                sComentario = "258-PF-PYF-FRENCH POLYNESIA"
            Case 260
                sValor = "FRENCH SOUTHERN TERRITORIES"
                sComentario = "260-TF-ATF-FRENCH SOUTHERN TERRITORIES"
            Case 266
                sValor = "GABON"
                sComentario = "266-GA-GAB-GABON"
            Case 270
                sValor = "GAMBIA"
                sComentario = "270-GM-GMB-GAMBIA"
            Case 268
                sValor = "GEORGIA"
                sComentario = "268-GE-GEO-GEORGIA"
            Case 276
                sValor = "GERMANY"
                sComentario = "276-DE-DEU-GERMANY"
            Case 288
                sValor = "GHANA"
                sComentario = "288-GH-GHA-GHANA"
            Case 292
                sValor = "GIBRALTAR"
                sComentario = "292-GI-GIB-GIBRALTAR"
            Case 300
                sValor = "GREECE"
                sComentario = "300-GR-GRC-GREECE"
            Case 304
                sValor = "GREENLAND"
                sComentario = "304-GL-GRL-GREENLAND"
            Case 308
                sValor = "GRENADA"
                sComentario = "308-GD-GRD-GRENADA"
            Case 312
                sValor = "GUADELOUPE"
                sComentario = "312-GP-GLP-GUADELOUPE"
            Case 316
                sValor = "GUAM"
                sComentario = "316-GU-GUM-GUAM"
            Case 320
                sValor = "GUATEMALA"
                sComentario = "320-GT-GTM-GUATEMALA"
            Case 324
                sValor = "GUINEA"
                sComentario = "324-GN-GIN-GUINEA"
            Case 624
                sValor = "GUINEA-BISSAU"
                sComentario = "624-GW-GNB-GUINEA-BISSAU"
            Case 328
                sValor = "GUYANA"
                sComentario = "328-GY-GUY-GUYANA"
            Case 332
                sValor = "HAITI"
                sComentario = "332-HT-HTI-HAITI"
            Case 334
                sValor = "HEARD AND MC DONALD ISLANDS"
                sComentario = "334-HM-HMD-HEARD AND MC DONALD ISLANDS"
            Case 340
                sValor = "HONDURAS"
                sComentario = "340-HN-HND-HONDURAS"
            Case 344
                sValor = "HONG KONG"
                sComentario = "344-HK-HKG-HONG KONG"
            Case 348
                sValor = "HUNGARY"
                sComentario = "348-HU-HUN-HUNGARY"
            Case 352
                sValor = "ICELAND"
                sComentario = "352-IS-ISL-ICELAND"
            Case 356
                sValor = "INDIA"
                sComentario = "356-IN-IND-INDIA"
            Case 360
                sValor = "INDONESIA"
                sComentario = "360-ID-IDN-INDONESIA"
            Case 364
                sValor = "IRAN (ISLAMIC REPUBLIC OF)"
                sComentario = "364-IR-IRN-IRAN (ISLAMIC REPUBLIC OF)"
            Case 368
                sValor = "IRAQ"
                sComentario = "368-IQ-IRQ-IRAQ"
            Case 372
                sValor = "IRELAND"
                sComentario = "372-IE-IRL-IRELAND"
            Case 376
                sValor = "ISRAEL"
                sComentario = "376-IL-ISR-ISRAEL"
            Case 380
                sValor = "ITALY"
                sComentario = "380-IT-ITA-ITALY"
            Case 388
                sValor = "JAMAICA"
                sComentario = "388-JM-JAM-JAMAICA"
            Case 392
                sValor = "JAPAN"
                sComentario = "392-JP-JPN-JAPAN"
            Case 400
                sValor = "JORDAN"
                sComentario = "400-JO-JOR-JORDAN"
            Case 398
                sValor = "KAZAKHSTAN"
                sComentario = "398-KZ-KAZ-KAZAKHSTAN"
            Case 404
                sValor = "KENYA"
                sComentario = "404-KE-KEN-KENYA"
            Case 296
                sValor = "KIRIBATI"
                sComentario = "296-KI-KIR-KIRIBATI"
            Case 408
                sValor = "KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF"
                sComentario = "408-KP-PRK-KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF"
            Case 410
                sValor = "KOREA, REPUBLIC OF"
                sComentario = "410-KR-KOR-KOREA, REPUBLIC OF"
            Case 414
                sValor = "KUWAIT"
                sComentario = "414-KW-KWT-KUWAIT"
            Case 417
                sValor = "KYRGYZSTAN"
                sComentario = "417-KG-KGZ-KYRGYZSTAN"
            Case 418
                sValor = "LAO PEOPLE'S DEMOCRATIC REPUBLIC"
                sComentario = "418-LA-LAO-LAO PEOPLE'S DEMOCRATIC REPUBLIC"
            Case 428
                sValor = "LATVIA"
                sComentario = "428-LV-LVA-LATVIA"
            Case 422
                sValor = "LEBANON"
                sComentario = "422-LB-LBN-LEBANON"
            Case 426
                sValor = "LESOTHO"
                sComentario = "426-LS-LSO-LESOTHO"
            Case 430
                sValor = "LIBERIA"
                sComentario = "430-LR-LBR-LIBERIA"
            Case 434
                sValor = "LIBYAN ARAB JAMAHIRIYA"
                sComentario = "434-LY-LBY-LIBYAN ARAB JAMAHIRIYA"
            Case 438
                sValor = "LIECHTENSTEIN"
                sComentario = "438-LI-LIE-LIECHTENSTEIN"
            Case 440
                sValor = "LITHUANIA"
                sComentario = "440-LT-LTU-LITHUANIA"
            Case 442
                sValor = "LUXEMBOURG"
                sComentario = "442-LU-LUX-LUXEMBOURG"
            Case 446
                sValor = "MACAU"
                sComentario = "446-MO-MAC-MACAU"
            Case 807
                sValor = "MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF"
                sComentario = "807-MK-MKD-MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF"
            Case 450
                sValor = "MADAGASCAR"
                sComentario = "450-MG-MDG-MADAGASCAR"
            Case 454
                sValor = "MALAWI"
                sComentario = "454-MW-MWI-MALAWI"
            Case 458
                sValor = "MALAYSIA"
                sComentario = "458-MY-MYS-MALAYSIA"
            Case 462
                sValor = "MALDIVES"
                sComentario = "462-MV-MDV-MALDIVES"
            Case 466
                sValor = "MALI"
                sComentario = "466-ML-MLI-MALI"
            Case 470
                sValor = "MALTA"
                sComentario = "470-MT-MLT-MALTA"
            Case 584
                sValor = "MARSHALL ISLANDS"
                sComentario = "584-MH-MHL-MARSHALL ISLANDS"
            Case 474
                sValor = "MARTINIQUE"
                sComentario = "474-MQ-MTQ-MARTINIQUE"
            Case 478
                sValor = "MAURITANIA"
                sComentario = "478-MR-MRT-MAURITANIA"
            Case 480
                sValor = "MAURITIUS"
                sComentario = "480-MU-MUS-MAURITIUS"
            Case 175
                sValor = "MAYOTTE"
                sComentario = "175-YT-MYT-MAYOTTE"
            Case 484
                sValor = "MEXICO"
                sComentario = "484-MX-MEX-MEXICO"
            Case 583
                sValor = "MICRONESIA, FEDERATED STATES OF"
                sComentario = "583-FM-FSM-MICRONESIA, FEDERATED STATES OF"
            Case 498
                sValor = "MOLDOVA, REPUBLIC OF"
                sComentario = "498-MD-MDA-MOLDOVA, REPUBLIC OF"
            Case 492
                sValor = "MONACO"
                sComentario = "492-MC-MCO-MONACO"
            Case 496
                sValor = "MONGOLIA"
                sComentario = "496-MN-MNG-MONGOLIA"
            Case 500
                sValor = "MONTSERRAT"
                sComentario = "500-MS-MSR-MONTSERRAT"
            Case 504
                sValor = "MOROCCO"
                sComentario = "504-MA-MAR-MOROCCO"
            Case 508
                sValor = "MOZAMBIQUE"
                sComentario = "508-MZ-MOZ-MOZAMBIQUE"
            Case 104
                sValor = "MYANMAR"
                sComentario = "104-MM-MMR-MYANMAR"
            Case 516
                sValor = "NAMIBIA"
                sComentario = "516-NA-NAM-NAMIBIA"
            Case 520
                sValor = "NAURU"
                sComentario = "520-NR-NRU-NAURU"
            Case 524
                sValor = "NEPAL"
                sComentario = "524-NP-NPL-NEPAL"
            Case 528
                sValor = "NETHERLANDS"
                sComentario = "528-NL-NLD-NETHERLANDS"
            Case 530
                sValor = "NETHERLANDS ANTILLES"
                sComentario = "530-AN-ANT-NETHERLANDS ANTILLES"
            Case 540
                sValor = "NEW CALEDONIA"
                sComentario = "540-NC-NCL-NEW CALEDONIA"
            Case 554
                sValor = "NEW ZEALAND"
                sComentario = "554-NZ-NZL-NEW ZEALAND"
            Case 558
                sValor = "NICARAGUA"
                sComentario = "558-NI-NIC-NICARAGUA"
            Case 562
                sValor = "NIGER"
                sComentario = "562-NE-NER-NIGER"
            Case 566
                sValor = "NIGERIA"
                sComentario = "566-NG-NGA-NIGERIA"
            Case 570
                sValor = "NIUE"
                sComentario = "570-NU-NIU-NIUE"
            Case 574
                sValor = "NORFOLK ISLAND"
                sComentario = "574-NF-NFK-NORFOLK ISLAND"
            Case 580
                sValor = "NORTHERN MARIANA ISLANDS"
                sComentario = "580-MP-MNP-NORTHERN MARIANA ISLANDS"
            Case 578
                sValor = "NORWAY"
                sComentario = "578-NO-NOR-NORWAY"
            Case 512
                sValor = "OMAN"
                sComentario = "512-OM-OMN-OMAN"
            Case 586
                sValor = "PAKISTAN"
                sComentario = "586-PK-PAK-PAKISTAN"
            Case 585
                sValor = "PALAU"
                sComentario = "585-PW-PLW-PALAU"
            Case 275
                sValor = "PALESTINIAN TERRITORY, Occupied"
                sComentario = "275-PS-PSE-PALESTINIAN TERRITORY, Occupied"
            Case 591
                sValor = "PANAMA"
                sComentario = "591-PA-PAN-PANAMA"
            Case 598
                sValor = "PAPUA NEW GUINEA"
                sComentario = "598-PG-PNG-PAPUA NEW GUINEA"
            Case 600
                sValor = "PARAGUAY"
                sComentario = "600-PY-PRY-PARAGUAY"
            Case 604
                sValor = "PERU"
                sComentario = "604-PE-PER-PERU"
            Case 608
                sValor = "PHILIPPINES"
                sComentario = "608-PH-PHL-PHILIPPINES"
            Case 612
                sValor = "PITCAIRN"
                sComentario = "612-PN-PCN-PITCAIRN"
            Case 616
                sValor = "POLAND"
                sComentario = "616-PL-POL-POLAND"
            Case 620
                sValor = "PORTUGAL"
                sComentario = "620-PT-PRT-PORTUGAL"
            Case 630
                sValor = "PUERTO RICO"
                sComentario = "630-PR-PRI-PUERTO RICO"
            Case 634
                sValor = "QATAR"
                sComentario = "634-QA-QAT-QATAR"
            Case 638
                sValor = "REUNION"
                sComentario = "638-RE-REU-REUNION"
            Case 642
                sValor = "ROMANIA"
                sComentario = "642-RO-ROU-ROMANIA"
            Case 643
                sValor = "RUSSIAN FEDERATION"
                sComentario = "643-RU-RUS-RUSSIAN FEDERATION"
            Case 646
                sValor = "RWANDA"
                sComentario = "646-RW-RWA-RWANDA"
            Case 654
                sValor = "SAINT HELENA"
                sComentario = "654-SH-SHN-SAINT HELENA"
            Case 659
                sValor = "SAINT KITTS AND NEVIS"
                sComentario = "659-KN-KNA-SAINT KITTS AND NEVIS"
            Case 662
                sValor = "SAINT LUCIA"
                sComentario = "662-LC-LCA-SAINT LUCIA"
            Case 666
                sValor = "SAINT PIERRE AND MIQUELON"
                sComentario = "666-PM-SPM-SAINT PIERRE AND MIQUELON"
            Case 670
                sValor = "SAINT VINCENT AND THE GRENADINES"
                sComentario = "670-VC-VCT-SAINT VINCENT AND THE GRENADINES"
            Case 882
                sValor = "SAMOA"
                sComentario = "882-WS-WSM-SAMOA"
            Case 674
                sValor = "SAN MARINO"
                sComentario = "674-SM-SMR-SAN MARINO"
            Case 678
                sValor = "SAO TOME AND PRINCIPE"
                sComentario = "678-ST-STP-SAO TOME AND PRINCIPE"
            Case 682
                sValor = "SAUDI ARABIA"
                sComentario = "682-SA-SAU-SAUDI ARABIA"
            Case 686
                sValor = "SENEGAL"
                sComentario = "686-SN-SEN-SENEGAL"
            Case 891
                sValor = "SERBIA AND MONTENEGRO"
                sComentario = "891-CS-SCG-SERBIA AND MONTENEGRO"
            Case 690
                sValor = "SEYCHELLES"
                sComentario = "690-SC-SYC-SEYCHELLES"
            Case 694
                sValor = "SIERRA LEONE"
                sComentario = "694-SL-SLE-SIERRA LEONE"
            Case 702
                sValor = "SINGAPORE"
                sComentario = "702-SG-SGP-SINGAPORE"
            Case 703
                sValor = "SLOVAKIA"
                sComentario = "703-SK-SVK-SLOVAKIA"
            Case 705
                sValor = "SLOVENIA"
                sComentario = "705-SI-SVN-SLOVENIA"
            Case 90
                sValor = "SOLOMON ISLANDS"
                sComentario = "090-SB-SLB-SOLOMON ISLANDS"
            Case 706
                sValor = "SOMALIA"
                sComentario = "706-SO-SOM-SOMALIA"
            Case 710
                sValor = "SOUTH AFRICA"
                sComentario = "710-ZA-ZAF-SOUTH AFRICA"
            Case 239
                sValor = "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS"
                sComentario = "239-GS-SGS-SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS"
            Case 724
                sValor = "SPAIN"
                sComentario = "724-ES-ESP-SPAIN"
            Case 144
                sValor = "SRI LANKA"
                sComentario = "144-LK-LKA-SRI LANKA"
            Case 736
                sValor = "SUDAN"
                sComentario = "736-SD-SDN-SUDAN"
            Case 740
                sValor = "SURINAME"
                sComentario = "740-SR-SUR-SURINAME"
            Case 744
                sValor = "SVALBARD AND JAN MAYEN ISLANDS"
                sComentario = "744-SJ-SJM-SVALBARD AND JAN MAYEN ISLANDS"
            Case 748
                sValor = "SWAZILAND"
                sComentario = "748-SZ-SWZ-SWAZILAND"
            Case 752
                sValor = "SWEDEN"
                sComentario = "752-SE-SWE-SWEDEN"
            Case 756
                sValor = "SWITZERLAND"
                sComentario = "756-CH-CHE-SWITZERLAND"
            Case 760
                sValor = "SYRIAN ARAB REPUBLIC"
                sComentario = "760-SY-SYR-SYRIAN ARAB REPUBLIC"
            Case 158
                sValor = "TAIWAN"
                sComentario = "158-TW-TWN-TAIWAN"
            Case 762
                sValor = "TAJIKISTAN"
                sComentario = "762-TJ-TJK-TAJIKISTAN"
            Case 834
                sValor = "TANZANIA, UNITED REPUBLIC OF"
                sComentario = "834-TZ-TZA-TANZANIA, UNITED REPUBLIC OF"
            Case 764
                sValor = "THAILAND"
                sComentario = "764-TH-THA-THAILAND"
            Case 626
                sValor = "TIMOR-LESTE"
                sComentario = "626-TL-TLS-TIMOR-LESTE"
            Case 768
                sValor = "TOGO"
                sComentario = "768-TG-TGO-TOGO"
            Case 772
                sValor = "TOKELAU"
                sComentario = "772-TK-TKL-TOKELAU"
            Case 776
                sValor = "TONGA"
                sComentario = "776-TO-TON-TONGA"
            Case 780
                sValor = "TRINIDAD AND TOBAGO"
                sComentario = "780-TT-TTO-TRINIDAD AND TOBAGO"
            Case 788
                sValor = "TUNISIA"
                sComentario = "788-TN-TUN-TUNISIA"
            Case 792
                sValor = "TURKEY"
                sComentario = "792-TR-TUR-TURKEY"
            Case 795
                sValor = "TURKMENISTAN"
                sComentario = "795-TM-TKM-TURKMENISTAN"
            Case 796
                sValor = "TURKS AND CAICOS ISLANDS"
                sComentario = "796-TC-TCA-TURKS AND CAICOS ISLANDS"
            Case 798
                sValor = "TUVALU"
                sComentario = "798-TV-TUV-TUVALU"
            Case 800
                sValor = "UGANDA"
                sComentario = "800-UG-UGA-UGANDA"
            Case 804
                sValor = "UKRAINE"
                sComentario = "804-UA-UKR-UKRAINE"
            Case 784
                sValor = "UNITED ARAB EMIRATES"
                sComentario = "784-AE-ARE-UNITED ARAB EMIRATES"
            Case 826
                sValor = "UNITED KINGDOM"
                sComentario = "826-GB-GBR-UNITED KINGDOM"
            Case 840
                sValor = "UNITED STATES"
                sComentario = "840-US-USA-UNITED STATES"
            Case 581
                sValor = "UNITED STATES MINOR OUTLYING ISLANDS"
                sComentario = "581-UM-UMI-UNITED STATES MINOR OUTLYING ISLANDS"
            Case 858
                sValor = "URUGUAY"
                sComentario = "858-UY-URY-URUGUAY"
            Case 860
                sValor = "UZBEKISTAN"
                sComentario = "860-UZ-UZB-UZBEKISTAN"
            Case 548
                sValor = "VANUATU"
                sComentario = "548-VU-VUT-VANUATU"
            Case 336
                sValor = "VATICAN CITY STATE (HOLY SEE)"
                sComentario = "336-VA-VAT-VATICAN CITY STATE (HOLY SEE)"
            Case 862
                sValor = "VENEZUELA"
                sComentario = "862-VE-VEN-VENEZUELA"
            Case 704
                sValor = "VIET NAM"
                sComentario = "704-VN-VNM-VIET NAM"
            Case 92
                sValor = "VIRGIN ISLANDS (BRITISH)"
                sComentario = "092-VG-VGB-VIRGIN ISLANDS (BRITISH)"
            Case 850
                sValor = "VIRGIN ISLANDS (U.S.)"
                sComentario = "850-VI-VIR-VIRGIN ISLANDS (U.S.)"
            Case 876
                sValor = "WALLIS AND FUTUNA ISLANDS"
                sComentario = "876-WF-WLF-WALLIS AND FUTUNA ISLANDS"
            Case 732
                sValor = "WESTERN SAHARA"
                sComentario = "732-EH-ESH-WESTERN SAHARA"
            Case 887
                sValor = "YEMEN"
                sComentario = "887-YE-YEM-YEMEN"
            Case 894
                sValor = "ZAMBIA"
                sComentario = "894-ZM-ZMB-ZAMBIA"
            Case 716
                sValor = "ZIMBABWE"
                sComentario = "716-ZW-ZWE-ZIMBABWE"
            Case Else
                sComentario = ""
        End Select
        If sComentario <> "" Then
            
            MyRecSet.AddNew
            MyRecSet.Fields("idvalorlookup") = nC
            MyRecSet.Fields("idlookup") = 1
            MyRecSet.Fields("valor") = sValor
            MyRecSet.Fields("comentario") = sComentario
            MyRecSet.Update
            nC = nC + 1
            
        End If
    
    Next nC2
    
    For nC1 = 0 To 3
        
        MyRecSet.AddNew
        MyRecSet.Fields("idvalorlookup") = nC
        MyRecSet.Fields("idlookup") = 2
        Select Case nC1
        Case 0
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "SIN RAZA"
        MyRecSet.Fields("valor") = "SIN RAZA"
        
        Case 1
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "LECHERA"
        MyRecSet.Fields("valor") = "LECHERA"
        
        Case 2
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "MIXTA"
        MyRecSet.Fields("valor") = "MIXTA"
        
        Case 3
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "CARNE"
        MyRecSet.Fields("valor") = "CARNE"
        
        End Select
        
        MyRecSet.Update
        nC = nC + 1
        
    
    Next nC1
    For nC1 = 0 To 12
        
        MyRecSet.AddNew
        MyRecSet.Fields("idvalorlookup") = nC
        MyRecSet.Fields("idlookup") = 4

        Select Case nC1
        Case 0
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "SIN CATEGORIA"
        MyRecSet.Fields("valor") = "SIN CATEGORIA"
        
        Case 1
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "VACA JOVEN"
        MyRecSet.Fields("valor") = "VACA JOVEN"
        
        Case 2
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "TORO"
        MyRecSet.Fields("valor") = "TORO"
        
        Case 3
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "BUEY"
        MyRecSet.Fields("valor") = "BUEY"
        
        Case 4
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "VACA ADULTA"
        MyRecSet.Fields("valor") = "VACA ADULTA"
        
        Case 5
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "TERNERA"
        MyRecSet.Fields("valor") = "TERNERA"
        
        Case 6
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "BECERRO"
        MyRecSet.Fields("valor") = "BECERRO"
        
        Case 7
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "TERNERA V"
        MyRecSet.Fields("valor") = "TERNERA V"
        
        Case 8
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "TERNERA Z"
        MyRecSet.Fields("valor") = "TERNERA Z"
        
        Case 9
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "AÑOJO"
        MyRecSet.Fields("valor") = "AÑOJO"
        
        Case 10
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "NOVILLO"
        MyRecSet.Fields("valor") = "NOVILLO"
        
        Case 11
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "CEBÓN"
        MyRecSet.Fields("valor") = "CEBÓN"
        
        Case 12
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "ERAL"
        MyRecSet.Fields("valor") = "ERAL"
        
        End Select
        
        MyRecSet.Update
        nC = nC + 1
        
    
    Next nC1
    For nC1 = 1 To 2
        
        MyRecSet.AddNew
        MyRecSet.Fields("idvalorlookup") = nC
        MyRecSet.Fields("idlookup") = 3

        Select Case nC1
        Case 1
        MyRecSet.Fields("valor") = "MACHO"
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "M" & "-" & "MACHO"
        
        Case 2
        MyRecSet.Fields("valor") = "HEMBRA"
        MyRecSet.Fields("comentario") = Format(nC1, "000") & "-" & "H" & "-" & "HEMBRA"
        
        End Select
        
        MyRecSet.Update
        nC = nC + 1
        
    
    Next nC1

    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0
    CadenadeLog "Finalizada programación Paises, Categorias y Razas..."
    Capa1B_SendCountry_ScalePos = 0
    Exit Function
    
errorPais:
    CadenadeLog "Error en programación Paises, Categorias y Razas. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
    On Error GoTo 0
    Capa1B_SendCountry_ScalePos = 1004
    
End Function
Public Function Capa1B_SendCard_ScalePos(card As Epel_Card) As Long
'/////////////////////////////////////////////////
'// Envío de fichas de vacuno                   //
'/////////////////////////////////////////////////
Dim nC1 As Integer
Dim nResp As Long
Dim cm As String

'Public Type Epel_Card
'    Number As Long           ' Nº de ficha
'    Identifier As String     ' Identificador (texto), max. 20 caracteres
'    BirthCode As Long        ' Código país de nacimiento
'    BreedingCode As Long     ' Código país de crianza
'    SlaughterCode As Long    ' Código país de sacrificio
'    ButcheringCode As Long   ' Código país de despiece
'    ProductionCode As Long   ' Código país de producción
'    SlaughterReg As String   ' R.S.I. sacrificio (max. 20 caracteres)
'    ButcheringReg As String  ' R.S.I. despiece (max. 20 caracteres)
'    Category As Long         ' Categoría ( de 0 a 9)
'    Race As Long             ' Raza ( de 0 a 9)
'    SlaughterDate As String  ' Fecha Sacrificio (formato ddmmyy)
'    Age As Long              ' Edad (0 a 99) en años
'    Sex As String            ' Sexo, texto de un caracter
'    FreeText As String       ' Texto libre (max. 25 caracteres)
'    BirthDate As String      ' Fecha nacimiento (formato ddmmyy)
'    ButcheringDate As String ' Fecha despiece (formato ddmmyy)
'    SetNumber As String      ' Número de Lote (max. 20 caracteres)
'    Weight As Long           ' Peso (0 a 999999) en HectoGramos (Kg/10)
'End Type

    ' *******************
    ' Envío de datos
    ' *******************

    MyRecSet.Open "Select * from dbo.traza_campoficha where idficha=" & CStr(card.Number) & " order by idcampotipoficha", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        nResp = 1004
        GoTo salir
    Else
        
        On Error GoTo salir
        
        If Not (MyRecSet.EOF) Then 'si existe se borra y se crea de nuevo, si no solo se crea
            cm = "delete dbo.traza_campoficha where idficha=" & CStr(card.Number)
            MyConnObj.Execute cm
        End If
        

        For nC1 = 1 To 18
            
            MyRecSet.AddNew
            MyRecSet.Fields("idficha") = card.Number
            MyRecSet.Fields("idcampotipoficha") = nC1
            Select Case nC1
            Case 1
            MyRecSet.Fields("valor") = Trim(card.Identifier)
            Case 2
            MyRecSet.Fields("valor") = Trim(card.SetNumber)
            Case 3
            MyRecSet.Fields("valor") = CStr(card.ProductionCode) 'Debería ser el nombre del Pais... SE DEBERÁ buscar en la tabla de valores
                                                                 ' o modificar la estructura Epel_Card para que lleguen los nombres
                                                                 ' y no los códigos... De momento, de ejemplo, se deja así... Lo mismo
                                                                 ' para todos los campos en los que se introduzcan códigos. (*)
            Case 4
            MyRecSet.Fields("valor") = CStr(card.BirthCode) '(*)
            Case 5
            MyRecSet.Fields("valor") = Format(Mid(card.BirthDate, 1, 2) & "/" & Mid(card.BirthDate, 3, 2) & "/" & Mid(card.BirthDate, 5), "dd/mm/yyyy")
            Case 6
            MyRecSet.Fields("valor") = CStr(card.BreedingCode) '(*)
            Case 7
            MyRecSet.Fields("valor") = CStr(card.SlaughterCode) '(*)
            Case 8
            MyRecSet.Fields("valor") = Trim(card.SlaughterReg)
            Case 9
            MyRecSet.Fields("valor") = Format(Mid(card.SlaughterDate, 1, 2) & "/" & Mid(card.SlaughterDate, 3, 2) & "/" & Mid(card.SlaughterDate, 5), "dd/mm/yyyy")
            Case 10
            MyRecSet.Fields("valor") = Trim(card.ButcheringCode)
            Case 11
            MyRecSet.Fields("valor") = Trim(card.ButcheringReg)
            Case 12
            MyRecSet.Fields("valor") = Format(Mid(card.ButcheringDate, 1, 2) & "/" & Mid(card.ButcheringDate, 3, 2) & "/" & Mid(card.ButcheringDate, 5), "dd/mm/yyyy")
            Case 13
            MyRecSet.Fields("valor") = CStr(card.Category) '(*)
            Case 14
            MyRecSet.Fields("valor") = CStr(card.Race)  '(*)
            Case 15
            MyRecSet.Fields("valor") = CStr(card.Age)
            Case 16
            MyRecSet.Fields("valor") = Trim(card.Sex)
            Case 17
            MyRecSet.Fields("valor") = CStr((card.weight) * 10) 'Kilo
            Case 18 'día sacrificio??? --> Posiblemente el día de la semana (Lunes, Martes...) de momento dejamos el número
                    '                      de día de la fecha.
            MyRecSet.Fields("valor") = CStr(Day(Format(Mid(card.SlaughterDate, 1, 2) & "/" & Mid(card.SlaughterDate, 3, 2) & "/" & Mid(card.SlaughterDate, 5), "dd/mm/yyyy"))) '???
    
            End Select
        
            MyRecSet.Update
            
        Next nC1
        
    End If
                    
    nResp = 0
    
salir:
    If Err.Number <> 0 Then
        CadenadeLog "Error en programación Fichas. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        nResp = 1004
    End If
    On Error GoTo 0
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0
    Capa1B_SendCard_ScalePos = nResp

End Function
Public Function Query_Item_PluSec(ByVal nCode As Long, ByVal nSec As Long) As Long
    '**********************************************
    '* Obtiene PLU del artículo                   *
    '* a partir del Código y la Sección (interna) *
    '**********************************************
    ' Si no se localiza o error consulta devuelve 999999999
    Dim nResul As Long
    Dim MyRecTmp As New ADODB.Recordset
    MyRecTmp.Open "Select * from dbo.articuloreferencia where idarticulo=" & CStr(nCode) & " and idseccion=" & CStr(nSec), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        On Error GoTo 0
        nResul = 999999999
        GoTo FINItem_PluSec
    Else
        If Not MyRecTmp.EOF Then
            nResul = MyRecTmp.Fields("referencia")
        Else
            nResul = 999999999
        End If
    End If

FINItem_PluSec:
    On Error GoTo 0
    On Error Resume Next
    MyRecTmp.Close
    Set MyRecTmp = Nothing
    On Error GoTo 0
    
    Query_Item_PluSec = nResul
End Function

'Formas de Pago....
'1   Efectivo
'2   Tarjeta Crédito
'3   Tarjeta Débito
'4   EBT Cash
'5   Gift Card
'6   Cheque
'7   Pendiente
'14  Ticket Restaurant
'15  Chèque -Vacances
'16  Cheque Gourmet
'17  Restaurant Pass
'19  Cupón
'20  Others
'21  EBT FoodStamp
'22  Visa
'23  MasterCard
'24  American Express
'25  Traveller 's Check
'28  Multibanco
'30  Visa Electron
'31  Diners Club
'32  Big Pass
'33  Sodexho
'34  Pin
'35  Chipknip
'36  Cestaticket
'37  EC -Maestro
'38  Quick
'39  BPI -Express
'40  Gift Check
'41  Ripley
'42  Saga
'43  Otras tarjetas
'44  Cheque Comercial
'45  Cheque de Gobiernos
'46  CLAVE
'Public Type Epel_LinTick
'    '///////////////////////////////
'    '// Líneas de Ticket          //
'    '///////////////////////////////
'    amount As Long
'    Number As Long
'    Item_Code As Long
'    Item_Price As Long
'    Item_Pes As Long
'    positive As Long
'    Weighted As Long
'    Cancelled As Long
'    Discount As Long
'    BovineCard As Long
'End Type
'Public Type Epel_CabTick
'    '///////////////////////////////
'    '// Cabeceras de Ticket       //
'    '///////////////////////////////
'    type As Long '*** 0=> venta
'    NTicket As Long
'    Section As Long
'    Vendor As Long
'    Client As Long
'    Machine As Long '*** numMaq
'    nLines As Long '***numMax lin
'    positive As Long
'    amount As Long
'    Hour As String
'    date As String
'    Lines() As Epel_LinTick
'    Cancelled As Long
'    P1 As Long ' importe bruto
'    P2 As Long ' descuento
'    Task As Long ' 0--> Normal, 1 --> Tiquet de encargo
'    tipoPago As Long
'    cantidadEntre As Long
'    lote As String
'    lEncargo As Boolean
'    nLinCanceled As Integer
'End Type
Public Function Capa1B_QueryTicket_ScalePos(ByVal De_Que_Tipo As Long, Cabecera As Epel_CabTick) As Long
    '**********************
    ' de que tipo
    ' 0 --> Venta
    ' 1 --> Super,TPV
    ' 2 --> Envasado
    ' 3 --> Autoservicio
    ' 4 --> Venta Local
    ' 5 --> Super Local, TPV Local
    ' 6 --> Envasado Local
    ' 7 --> Autoservicio Local
    '**********************
    ' Convertir a:
    ' versión Julio (0=(Super y TPV), 1=(Venta y Mixto), 2=Envasado, 3=Autoservicio)
    '///////////////////////////////////////////////////
    '// Esta función pide los tiquets RED y LOCAL     //
    '// a balanzas Euroscale por protocolo  Binario   //
    '///////////////////////////////////////////////////
    
    'En primera versión se considera:
    '   ICG           Hydra+
    '  TotalNeto      Importe
    '  TotalBruto     Importe_bruto
    '  ImporteCambio  Cambio
    '                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecExt As New ADODB.Recordset
    
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim nCont As Long
    Dim Resp As Long
    Dim EsLocal As Boolean
    Dim DeQueTipo As Integer
    Dim Modo As Integer
    Modo = 0
    If De_Que_Tipo >= 1000 Then
        De_Que_Tipo = 1000 - De_Que_Tipo
        Modo = 1
    End If
    Resp = 0
    If De_Que_Tipo < 0 Or De_Que_Tipo > 7 Then
        Resp = 1002
        GoTo FIN_Capa1B_QueryTicket
    End If
    If De_Que_Tipo > 3 Then
        EsLocal = True
    Else
        EsLocal = False
    End If
    Select Case De_Que_Tipo
    Case 0
        DeQueTipo = 1
    Case 1
        DeQueTipo = 0
    Case 2 'envasado
        DeQueTipo = 2
    Case 3 'autoservicio
        DeQueTipo = 3
    Case 4
        DeQueTipo = 1
    Case 5
        DeQueTipo = 0
    Case 6 'envasado Local
        DeQueTipo = 2
    Case 7 'autoservicio Local
        DeQueTipo = 3
    End Select
    On Error Resume Next
    If Modo = 1 Then
        MyRecSet.Open "select * from dbo.documento where ( (isfinalizado=1) and tipodocumento=" & CStr(DeQueTipo) & " and iddocumento>" & CStr(myLastDoc) & ") order by numero", MyConnObj, adOpenKeyset, adLockOptimistic
    Else
        MyRecSet.Open "select * from dbo.documento where ( (isnull(iscapturado,0)=0 or iscapturado=0) and (isfinalizado=1) and tipodocumento=" & CStr(DeQueTipo) & ") order by numero", MyConnObj, adOpenKeyset, adLockOptimistic
    End If
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_QueryTicket
    End If
    On Error GoTo FIN_Capa1B_QueryTicket
    If Not MyRecSet.EOF Then
            Ffecha = MyRecSet.Fields("fecha")
            Fhora = Format(MyRecSet.Fields("hora"), "0000")
            'Fhora = Left(Fhora, 2) & ":" & Right(Fhora, 2) & ":00"
            Fhora = Left(Fhora, 2) & Right(Fhora, 2) & "00"
            IdDoc = MyRecSet.Fields("iddocumento")
            myLastDoc = MyRecSet.Fields("iddocumento")
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            ReDim Extra_lines(nNumLin)
            'recoger datos
            Cabecera.NTicket = MyRecSet.Fields("numero")
            Cabecera.Vendor = MyRecSet.Fields("idusuario")
            Cabecera.Section = MyRecSet.Fields("idseccion")
            Cabecera.Client = MyRecSet.Fields("idcontacto")
            'Cabecera.date = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.date = Format(Ffecha, "ddmmyy")
            Cabecera.Hour = Fhora
            Cabecera.Machine = MyRecSet.Fields("numeroterminal")
            Cabecera.amount = 100 * MyRecSet.Fields("totalneto")
            If Cabecera.amount >= 0 Then
                Cabecera.positive = 1
            Else
                Cabecera.positive = 0
            End If
            If IsNull(MyRecSet.Fields("iddocumentoabono")) Or MyRecSet.Fields("iddocumentoabono") = 0 Then
                Cabecera.Cancelled = 0
            Else
                Cabecera.Cancelled = 1
            End If
            'Cabecera.Task = 0 se utiliza para el % descuento
            Cabecera.Task = 100 * MyRecSet.Fields("porcentajedescuento")
            
            Cabecera.nLines = nNumLin
            Cabecera.P1 = 100 * MyRecSet.Fields("totalbruto")
            Cabecera.P2 = 100 * MyRecSet.Fields("totaldescuento")

            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where (Idlineaabono = -1 and iddocumento=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.nLinCanceled = nNumLin

            Select Case MyRecSet.Fields("tipodocumento")
            Case 0, 1
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 4
                Else
                    Cabecera.type = 0
                End If
            Case 2
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 6
                Else
                    Cabecera.type = 2
                End If
            Case 3
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 7
                Else
                    Cabecera.type = 3
                End If
            End Select

            Cabecera.tipoPago = MyRecSet.Fields("idmoneda")
            Cabecera.cantidadEntre = 100 * (MyRecSet.Fields("totalneto") + MyRecSet.Fields("importecambio"))
            Cabecera.lote = Space(20)
            Cabecera.lEncargo = False

            'LINEAS''

            nNumLin = 0
            MyRecAux.Open "select * from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF
                    nNumLin = nNumLin + 1
                    '''''''''''''''''''''
                    If MyRecAux.Fields("lote") <> 0 Then
                        Cabecera.lote = Format(MyRecAux.Fields("lote"), "00000000000000000000")
                    End If
                    '''''''''''''''''''''
                    Cabecera.Lines(nNumLin).amount = 100 * (MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos"))
                    Cabecera.Lines(nNumLin).Number = nNumLin
                    Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("idarticulo")
                    Cabecera.Lines(nNumLin).Item_Price = 100 * MyRecAux.Fields("precio")
                  
                    If MyRecAux.Fields("isventaporpeso") = True Then
                        Cabecera.Lines(nNumLin).Item_Pes = 1000 * MyRecAux.Fields("peso")
                        Cabecera.Lines(nNumLin).Weighted = 1
                    Else
                        Cabecera.Lines(nNumLin).Weighted = 0
                        Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("unidades")
                    End If
                    If Cabecera.Lines(nNumLin).amount >= 0 Then
                        Cabecera.Lines(nNumLin).positive = 1
                    Else
                        Cabecera.Lines(nNumLin).positive = 0
                    End If
                    
                    If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                        Cabecera.Lines(nNumLin).Cancelled = 1
                    Else
                        Cabecera.Lines(nNumLin).Cancelled = 0
                    End If
                    
                    Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("idfichatrazabilidad")
                    Cabecera.Lines(nNumLin).discount = 100 * MyRecAux.Fields("dtoslineaconimpuestos")
                                       
                    Extra_lines(nNumLin).porcentaje = 100 * MyRecAux.Fields("descuento")
                    
                    Extra_lines(nNumLin).idTique = Cabecera.NTicket
                    Extra_lines(nNumLin).idLinea = nNumLin
                    Extra_lines(nNumLin).idArticulo = Cabecera.Lines(nNumLin).item_code
                    MyRecExt.Open "select * from dbo.articuloreferencia where Idarticulo =" & CStr(Extra_lines(nNumLin).idArticulo), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Not MyRecExt.EOF Then
                        Extra_lines(nNumLin).idSeccion = MyRecExt.Fields("idseccion")
                        Extra_lines(nNumLin).idReferencia = MyRecExt.Fields("referencia")
                    Else
                        Extra_lines(nNumLin).idSeccion = 0
                        Extra_lines(nNumLin).idReferencia = 0
                    End If
                    MyRecExt.Close
                    Set MyRecExt = Nothing
                    MyRecExt.Open "select * from dbo.codigobarra where Idarticulo =" & CStr(Extra_lines(nNumLin).idArticulo), MyConnObj, adOpenKeyset, adLockOptimistic
                    If Not MyRecExt.EOF Then
                        Extra_lines(nNumLin).idBarras = Trim(MyRecExt.Fields("codigobarra"))
                    Else
                        Extra_lines(nNumLin).idBarras = Space(13)
                    End If
                    MyRecExt.Close
                    Set MyRecExt = Nothing
                    
                    
                    MyRecAux.MoveNext
                Loop
                'MyRecAux.Close
                'Set MyRecAux = Nothing
            End If '
            MyRecAux.Close
            Set MyRecAux = Nothing


            '''''''''''''''
            'marcar...
            cm = "update dbo.documento set iscapturado=1 where iddocumento=" & CStr(IdDoc)
            MyConnObj.Execute cm
            ''''''''''''

            MyRecSet.Update '
            
            Resp = 0
   
    Else
        Resp = 4
    End If
    MyRecSet.Close
    Set MyRecSet = Nothing

FIN_Capa1B_QueryTicket:
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueyTicket. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    Capa1B_QueryTicket_ScalePos = Resp
End Function

'Con esta función se puede recavar un tiquet en concreto, haya sido pedido antes o no. Esta función no marca el tiquet
'como pedido (por tanto, se puede volver a pedir con la función Query_Ticket)
'LType:
'Los mismos valores que en la función Query_ticket
'Lmode:
' Indica qué tipo de tiquet se desea solicitar :
'0 ' Numeración absoluta --> Para esta DLL se considera siempre Absoluta...
'1 ' Procedente de una sección en concreto
'2 ' Procedente de una balanza en concreto
'El modo que se indique ha de estar en consonancia con el modo en que esté organizada la numeración de tiquets
'del sistema (por balanza, sección o sistema)
'lMatch
' Dependiendo del valor de Lmode:
'"   Si se trata de "numeración absoluta", dejar a 0
'"   Si se trata de "numeración por sección", indicar el número de sección (1 a 99)
'"   Si se trata de "numeración por Balanza", indicar el número de balanza (1 a 99)
'lNumber
' Es el número de ticket que se solicita
'Cabecera
'Devuelve los datos de dicho tiquet
'Supongamos una balanza que ha generado tres tiquets, con números 00001, 00004, 00005
'Si solicito el tiquet 1 de dicha balanza (balanza 4 por ejemplo) en modo venta :
'Query_One_Ticket(0,2,4,1)
'Me devolverá el ticket 1
'Si solicito el ticket 2, me devolverá el ticket 4, que es el primero inmediatamente superior que existe
'Si solicito el 5, devolverá el 5
'Si solicito el 6 , devolverá el código 4 de error (no hay más tickets)
'La función devuelve el número de tiquet que se le pida, o bien el siguiente tiquet al número pedido.Si no hay ningún tiquet con número mayor, devuelve 4.
Public Function Capa1B_QueryOneTicket_ScalePos(ByVal De_Que_Tipo As Long, ByVal Lmode As Long, ByVal lMatch As Long, ByVal lNumber As Long, Cabecera As Epel_CabTick) As Long
    '**********************
    ' de que tipo
    ' 0 --> Venta
    ' 1 --> Super,TPV
    ' 2 --> Envasado
    ' 3 --> Autoservicio
    ' 4 --> Venta Local
    ' 5 --> Super Local, TPV Local
    ' 6 --> Envasado Local
    ' 7 --> Autoservicio Local
    '**********************
    ' Convertir a:
    ' versión Julio (0=(Super y TPV), 1=(Venta y Mixto), 2=Envasado, 3=Autoservicio)
    '///////////////////////////////////////////////////
    '// Esta función pide los tiquets RED y LOCAL     //
    '// a balanzas Euroscale por protocolo  Binario   //
    '///////////////////////////////////////////////////
    
    'En primera versión se considera:
    '   ICG           Hydra+
    '  TotalNeto      Importe
    '  TotalBruto     Importe_bruto
    '  ImporteCambio  Cambio
    '                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
    
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim nCont As Long
    Dim Resp As Long
    Dim EsLocal As Boolean
    Dim DeQueTipo As Integer
    
    Resp = 0
    If De_Que_Tipo < 0 Or De_Que_Tipo > 7 Then
        Resp = 1002
        GoTo FIN_Capa1B_QueryOneTicket
    End If
    If De_Que_Tipo > 3 Then
        EsLocal = True
    Else
        EsLocal = False
    End If
    Select Case De_Que_Tipo
    Case 0
        DeQueTipo = 1
    Case 1
        DeQueTipo = 0
    Case 2 'envasado
        DeQueTipo = 2
    Case 3 'autoservicio
        DeQueTipo = 3
    Case 4
        DeQueTipo = 1
    Case 5
        DeQueTipo = 0
    Case 6 'envasado Local
        DeQueTipo = 2
    Case 7 'autoservicio Local
        DeQueTipo = 3
    End Select
    On Error Resume Next
    MyRecSet.Open "select * from dbo.documento where ((isfinalizado=1) and tipodocumento=" & CStr(DeQueTipo) & " and numero>=" & CStr(lNumber) & ") order by numero", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_QueryOneTicket
    End If
    On Error GoTo FIN_Capa1B_QueryOneTicket
    If Not MyRecSet.EOF Then
            Ffecha = MyRecSet.Fields("fecha")
            Fhora = Format(MyRecSet.Fields("hora"), "0000")
            Fhora = Left(Fhora, 2) & ":" & Right(Fhora, 2) & ":00"
            IdDoc = MyRecSet.Fields("iddocumento")
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            'recoger datos
            Cabecera.NTicket = MyRecSet.Fields("numero")
            Cabecera.Vendor = MyRecSet.Fields("idusuario")
            Cabecera.Section = MyRecSet.Fields("idseccion")
            Cabecera.Client = MyRecSet.Fields("idcontacto")
            Cabecera.date = Format(Ffecha, "dd/mm/yyyy")
            Cabecera.Hour = Fhora
            Cabecera.Machine = MyRecSet.Fields("numeroterminal")
            Cabecera.amount = 100 * MyRecSet.Fields("totalneto")
            If Cabecera.amount >= 0 Then
                Cabecera.positive = 1
            Else
                Cabecera.positive = 0
            End If
            If IsNull(MyRecSet.Fields("iddocumentoabono")) Or MyRecSet.Fields("iddocumentoabono") = 0 Then
                Cabecera.Cancelled = 0
            Else
                Cabecera.Cancelled = 1
            End If
            Cabecera.Task = 0
            Cabecera.nLines = nNumLin
            Cabecera.P1 = 100 * MyRecSet.Fields("totalbruto")
            Cabecera.P2 = 100 * MyRecSet.Fields("totaldescuento")

            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where (Idlineaabono = -1 and iddocumento=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.nLinCanceled = nNumLin

            Select Case MyRecSet.Fields("tipodocumento")
            Case 0, 1
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 4
                Else
                    Cabecera.type = 0
                End If
            Case 2
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 6
                Else
                    Cabecera.type = 2
                End If
            Case 3
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 7
                Else
                    Cabecera.type = 3
                End If
            End Select

            Cabecera.tipoPago = MyRecSet.Fields("idmoneda")
            Cabecera.cantidadEntre = 100 * (MyRecSet.Fields("totalneto") + MyRecSet.Fields("importecambio"))
            Cabecera.lote = Space(20)
            Cabecera.lEncargo = False

            'LINEAS''

            nNumLin = 0
            MyRecAux.Open "select * from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF
                    nNumLin = nNumLin + 1
                    Cabecera.Lines(nNumLin).amount = MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos")
                    Cabecera.Lines(nNumLin).Number = nNumLin
                    Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("idarticulo")
                    Cabecera.Lines(nNumLin).Item_Price = 100 * MyRecAux.Fields("precio")
                  
                    If MyRecAux.Fields("isventaporpeso") = True Then
                        Cabecera.Lines(nNumLin).Item_Pes = 1000 * MyRecAux.Fields("peso")
                        Cabecera.Lines(nNumLin).Weighted = 1
                    Else
                        Cabecera.Lines(nNumLin).Weighted = 0
                        Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("unidades")
                    End If
                    If Cabecera.Lines(nNumLin).amount >= 0 Then
                        Cabecera.Lines(nNumLin).positive = 1
                    Else
                        Cabecera.Lines(nNumLin).positive = 0
                    End If
                    
                    If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                        Cabecera.Lines(nNumLin).Cancelled = 1
                    Else
                        Cabecera.Lines(nNumLin).Cancelled = 0
                    End If
                    

                    Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("idfichatrazabilidad")

                    MyRecAux.MoveNext
                Loop
            
            End If '
            
            MyRecAux.Close
            Set MyRecAux = Nothing

            '''''''''''''''Esta función no marca el Tique como pedido...
            ''marcar...
            ''cm = "update dbo.documento set iscapturado=1 where iddocumento=" & CStr(IdDoc)
            ''MyConnObj.Execute cm
            '''''''''''''
            ''MyRecSet.Update '
            '''''''''''''
            
            Resp = 0
   
    Else
        Resp = 4
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing

FIN_Capa1B_QueryOneTicket:
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryOneTicket. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    Capa1B_QueryOneTicket_ScalePos = Resp
    
End Function

Public Function Capa1B_QueryOneTicketMark_ScalePos(ByVal De_Que_Tipo As Long, ByVal Lmode As Long, ByVal lMatch As Long, ByVal lNumber As Long, Cabecera As Epel_CabTick) As Long
    '**********************
    ' de que tipo
    ' 0 --> Venta
    ' 1 --> Super,TPV
    ' 2 --> Envasado
    ' 3 --> Autoservicio
    ' 4 --> Venta Local
    ' 5 --> Super Local, TPV Local
    ' 6 --> Envasado Local
    ' 7 --> Autoservicio Local
    '**********************
    ' Convertir a:
    ' versión Julio (0=(Super y TPV), 1=(Venta y Mixto), 2=Envasado, 3=Autoservicio)
    '///////////////////////////////////////////////////
    '// Esta función pide los tiquets RED y LOCAL     //
    '// a balanzas Euroscale por protocolo  Binario   //
    '///////////////////////////////////////////////////
    
    'En primera versión se considera:
    '   ICG           Hydra+
    '  TotalNeto      Importe
    '  TotalBruto     Importe_bruto
    '  ImporteCambio  Cambio
    '                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
    
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim nnumlinCan As Variant
    Dim nCont As Long
    Dim Resp As Long
    Dim EsLocal As Boolean
    Dim DeQueTipo As Integer
    
    CadenadeLog "QueryOneTicketMark."
    Resp = 0
    If De_Que_Tipo < 0 Or De_Que_Tipo > 7 Then
        Resp = 1002
        GoTo FIN_Capa1B_QueryOneTicketMark
    End If
    If De_Que_Tipo > 3 Then
        EsLocal = True
    Else
        EsLocal = False
    End If
    Select Case De_Que_Tipo
    Case 0
        DeQueTipo = 1
    Case 1
        DeQueTipo = 0
    Case 2 'envasado
        DeQueTipo = 2
    Case 3 'autoservicio
        DeQueTipo = 3
    Case 4
        DeQueTipo = 1
    Case 5
        DeQueTipo = 0
    Case 6 'envasado Local
        DeQueTipo = 2
    Case 7 'autoservicio Local
        DeQueTipo = 3
    End Select
    On Error Resume Next
    PonDato "Documento", 0
    MyRecSet.Open "select * from dbo.documento where ((isfinalizado=1) and tipodocumento=" & CStr(DeQueTipo) & " and numero>=" & CStr(lNumber) & ") order by numero", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_QueryOneTicketMark
    End If
    On Error GoTo FIN_Capa1B_QueryOneTicketMark
    If Not MyRecSet.EOF Then
'       IdDocumento
'       IdGlobal
'       TipoDocumento
'       IdSeccion
'       NumeroTerminal
'       Numero
'       ImpuestosIncluidos
'       IdUsuario
'       IdContacto
'       IdMoneda
'       IdArqueo
'       Fecha
'       Hora
'       PorcentajeDescuento
'       TotalBruto
'       TotalDescuento
'       TotalImpuestos
'       TotalNeto
'       ImporteCambio
'       IdMonedaCambio
'       CotizacionCambio
'       IsFinalizado
'       IdDocumentoAbono
'       IdRegistroEntrada
'       IdPlantilla
'       IsLocal
'       IsSincronizado
'       NumCopias
'       IsCapturado
'       IdBalanza
            Ffecha = MyRecSet.Fields("fecha")
            Ffecha = Mid(Ffecha, 1, 2) & Mid(Ffecha, 4, 2) & Mid(Ffecha, 9, 2)
            Fhora = Format(MyRecSet.Fields("hora"), "0000")
            'Fhora = Left(Fhora, 2) & Right(Fhora, 2)
            IdDoc = MyRecSet.Fields("iddocumento")
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            nNumLin = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            ReDim Cabecera.Lines(nNumLin)
            'recoger datos
            Cabecera.NTicket = MyRecSet.Fields("numero")
            'CadenadeLog "Cabecera.NTicket=" & CStr(Cabecera.NTicket)
            Cabecera.Vendor = MyRecSet.Fields("idusuario")
            'CadenadeLog "Cabecera.Vendor=" & CStr(Cabecera.Vendor)
            Cabecera.Section = MyRecSet.Fields("idseccion")
            'CadenadeLog "Cabecera.Section=" & CStr(Cabecera.Section)
            Cabecera.Client = MyRecSet.Fields("idcontacto")
            'CadenadeLog "Cabecera.Client=" & CStr(Cabecera.Client)
            Cabecera.date = Ffecha 'Format(Ffecha, "dd/mm/yyyy")
            'CadenadeLog "Cabecera.date=" & Cabecera.date
            Cabecera.Hour = Fhora
            'CadenadeLog "Cabecera.Hour=" & Cabecera.Hour
            Cabecera.Machine = MyRecSet.Fields("numeroterminal")
            'CadenadeLog "Cabecera.Machine=" & CStr(Cabecera.Machine)
            Cabecera.amount = 100 * (MyRecSet.Fields("totalneto"))
            CadenadeLog "Cabecera.amount=" & CStr(Cabecera.amount)
            If Cabecera.amount >= 0 Then
                Cabecera.positive = 1
            Else
                Cabecera.positive = 0
                Cabecera.amount = Cabecera.amount '1.0.9 Abs(Cabecera.amount)
            End If
            CadenadeLog "Cabecera.positive=" & CStr(Cabecera.positive)
            If IsNull(MyRecSet.Fields("iddocumentoabono")) Or MyRecSet.Fields("iddocumentoabono") = 0 Then
                Cabecera.Cancelled = 0
            Else
                Cabecera.Cancelled = 1
            End If
            CadenadeLog "Cabecera.Cancelled=" & CStr(Cabecera.Cancelled)
            Cabecera.Task = 0
            'CadenadeLog "Cabecera.Task=" & CStr(Cabecera.Task)
            Cabecera.nLines = nNumLin
            'CadenadeLog "Cabecera.nLines=" & CStr(Cabecera.nLines)
            Cabecera.P1 = 100 * (MyRecSet.Fields("totalbruto"))
            'CadenadeLog "Cabecera.P1=" & CStr(Cabecera.P1)
            Cabecera.P2 = 100 * (MyRecSet.Fields("totaldescuento"))
            'CadenadeLog "Cabecera.P2=" & CStr(Cabecera.P2)
            MyRecAux.Open "select COUNT(*) as lineas from dbo.LineaDocumento where (idlineaabono = -1 and iddocumento=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
            nnumlinCan = MyRecAux.Fields("lineas")
            MyRecAux.Close
            Set MyRecAux = Nothing
            Cabecera.nLinCanceled = nnumlinCan
            'CadenadeLog "Cabecera.nLinCanceled=" & CStr(Cabecera.nLinCanceled)
            Select Case MyRecSet.Fields("tipodocumento")
            Case 0, 1
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 4
                Else
                    Cabecera.type = 0
                End If
            Case 2
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 6
                Else
                    Cabecera.type = 2
                End If
            Case 3
                If MyRecSet.Fields("islocal") = 1 Then
                    Cabecera.type = 7
                Else
                    Cabecera.type = 3
                End If
            End Select
            'CadenadeLog "Cabecera.type=" & CStr(Cabecera.type)
            Cabecera.tipoPago = MyRecSet.Fields("idmoneda")
            'CadenadeLog "Cabecera.tipoPago=" & CStr(Cabecera.tipoPago)
            Cabecera.cantidadEntre = 100 * (MyRecSet.Fields("totalneto")) + (100 * (MyRecSet.Fields("importecambio")))
            'CadenadeLog "Cabecera.cantidadEntre=" & CStr(Cabecera.cantidadEntre)
            Cabecera.lote = Space(20)
            'CadenadeLog "Cabecera.lote=" & Cabecera.lote
            Cabecera.lEncargo = False
            'CadenadeLog "Cabecera.lEncargo=" & "False"
            
            'LINEAS''

            nNumLin = 0
            On Error Resume Next
            'PonDato "Lineas...", 0
            MyRecAux.Open "select * from dbo.LineaDocumento where IdDocumento =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo FIN_Capa1B_QueryOneTicketMark
            End If
                
            If Not MyRecAux.EOF Then
                MyRecAux.MoveFirst
                Do While Not MyRecAux.EOF
                    'nNumLin = nNumLin + 1
                    Cabecera.Lines(nNumLin).amount = 100 * (MyRecAux.Fields("importeconimpuestos") - MyRecAux.Fields("dtoscabeceraconimpuestos"))   '- (100 * (MyRecAux.Fields("dtoslineaconimpuestos")))
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").amount=" & CStr(Cabecera.Lines(nNumLin).amount)
                    Cabecera.Lines(nNumLin).Number = nNumLin
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Number=" & CStr(Cabecera.Lines(nNumLin).Number)
                    Cabecera.Lines(nNumLin).item_code = MyRecAux.Fields("idarticulo")
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").item_code=" & CStr(Cabecera.Lines(nNumLin).Item_Code)
                    Cabecera.Lines(nNumLin).Item_Price = 100 * MyRecAux.Fields("precio")
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Item_Price=" & CStr(Cabecera.Lines(nNumLin).Item_Price)
                    If MyRecAux.Fields("isventaporpeso") = True Then
                        Cabecera.Lines(nNumLin).Item_Pes = 1000 * MyRecAux.Fields("peso")
                        Cabecera.Lines(nNumLin).Weighted = 1
                    Else
                        Cabecera.Lines(nNumLin).Weighted = 0
                        Cabecera.Lines(nNumLin).Item_Pes = MyRecAux.Fields("unidades")
                    End If
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Item_Pes=" & CStr(Cabecera.Lines(nNumLin).Item_Pes)
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").weighted=" & CStr(Cabecera.Lines(nNumLin).Weighted)
                    If Cabecera.Lines(nNumLin).amount >= 0 Then
                        Cabecera.Lines(nNumLin).positive = 1
                    Else
                        Cabecera.Lines(nNumLin).positive = 0
                        Cabecera.Lines(nNumLin).amount = Cabecera.Lines(nNumLin).amount '1.0.9 Abs(Cabecera.Lines(nNumLin).amount)
                        Cabecera.Lines(nNumLin).Item_Pes = (-1) * Cabecera.Lines(nNumLin).Item_Pes '(-1) * MyRecAux.Fields("unidades")
                    End If
                    CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").positive=" & CStr(Cabecera.Lines(nNumLin).positive)
                    CadenadeLog "Cabecera.lines(" & CStr(nNumLin) & ").amount=" & CStr(Cabecera.Lines(nNumLin).amount)
                    CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Item_Pes=" & CStr(Cabecera.Lines(nNumLin).Item_Pes)
                    If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                        Cabecera.Lines(nNumLin).Cancelled = 1
                    Else
                        Cabecera.Lines(nNumLin).Cancelled = 0
                    End If
                    CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Cancelled=" & CStr(Cabecera.Lines(nNumLin).Cancelled)

                    Cabecera.Lines(nNumLin).BovineCard = MyRecAux.Fields("idfichatrazabilidad")
                    'CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").BovineCard=" & CStr(Cabecera.Lines(nNumLin).BovineCard)
                    
                    Cabecera.Lines(nNumLin).discount = 100 * MyRecAux.Fields("dtoslineaconimpuestos")
                    'Cabecera.Lines(nNumLin).Discount = MyRecAux.Fields("descuento")
                    CadenadeLog "Cabecera.Lines(" & CStr(nNumLin) & ").Discount=" & CStr(Cabecera.Lines(nNumLin).discount)
                    
                    nNumLin = nNumLin + 1
                    MyRecAux.MoveNext
                Loop
            
            End If '
            
            MyRecAux.Close
            Set MyRecAux = Nothing

            '''''''''''''''Esta función marca el Tique como pedido si no estaba pedido
            ''''''''''''''' Si estaba pedido responde 6 (Resp=6)
            If (IsNull(MyRecSet.Fields("iscapturado")) Or MyRecSet.Fields("iscapturado") = 0) Then
                ''marcar...
                cm = "update dbo.documento set iscapturado=1 where iddocumento=" & CStr(IdDoc)
                MyConnObj.Execute cm
                Resp = 0
                CadenadeLog "No estaba pedido se marca como Pedido."
            Else
                Resp = 6
                CadenadeLog "Estaba Pedido se responde:6 y no se marca."
            End If
            
    Else
        Resp = 4
    End If
    
    MyRecSet.Close
    Set MyRecSet = Nothing

FIN_Capa1B_QueryOneTicketMark:
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryOneTicketMark. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    Capa1B_QueryOneTicketMark_ScalePos = Resp

End Function

' Peticion de lineas de vacuno para label vasco
Public Function QueryLineBovine_ScalePos(ByVal mFam As Integer, ByVal mBal As Integer, mLine As Epel_LinBovine) As Long
Dim Resp As Long
Dim MyRecAux As New ADODB.Recordset
Dim IdArt As Long
Dim IdFamAux As Long
Dim IdEqp As Variant

'PonDato "Identificación Balanza...", 0
'MyRecAux.Open "select * from dbo.balanza", MyConnObj, adOpenKeyset, adLockOptimistic
'If Err.Number <> 0 Then
'    CadenadeLog "No se ha podido abrir la conexión con dbo.Balanza..."
'    Resp = 1000
'    GoTo FIN_Capa1A_GA_Erase_Totals
'End If
'On Error GoTo FIN_Capa1A_GA_Erase_Totals
'If Not MyRecAux.EOF Then
'    'ejemplo:FB9F5AA7-3AB9-46A4-8EDE-F1B57957C5B5
'    'CadenadeLog "Captura idbalanza"
'    IdEqp = MyRecAux.Fields("idbalanza")
'Else
'    IdEqp = "00000000-0000-0000-0000-000000000000"
'End If
'MyRecAux.Close
'Set MyRecAux = Nothing
    
    Resp = 0
    On Error Resume Next
    ' Tengo que buscar una linea de la familia dada... es decir consultar por la familia del articulo
    MyRecAux.Open "select * from dbo.Familia where Codigo = " & mFam, MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo Fin_QueryLineBovine_ScalePos
    End If
    'CadenadeLog "Familia Pedida: " & mFam
    If Not MyRecAux.EOF Then
        IdFamAux = MyRecAux.Fields("IdFamilia")
        'CadenadeLog "IdFamilia Pedida: " & IdFamAux
        MyRecAux.Close
        Set MyRecAux = Nothing
    Else
        Resp = 1005
        MyRecAux.Close
        Set MyRecAux = Nothing
        GoTo Fin_QueryLineBovine_ScalePos
    End If
      
MyRecAux.Open "select * from dbo.balanza where numero=" & CStr(mBal), MyConnObj, adOpenKeyset, adLockOptimistic
If Err.Number <> 0 Then
    'CadenadeLog "No se ha podido abrir la conexión con dbo.Balanza..."
    Resp = 1000
    GoTo Fin_QueryLineBovine_ScalePos
End If
If Not MyRecAux.EOF Then
    'ejemplo:FB9F5AA7-3AB9-46A4-8EDE-F1B57957C5B5
    ''CadenadeLog "Captura idbalanza"
    IdEqp = MyRecAux.Fields("idbalanza")
Else
    'IdEqp = "00000000-0000-0000-0000-000000000000"
    Resp = 1000
    GoTo Fin_QueryLineBovine_ScalePos
End If
MyRecAux.Close
Set MyRecAux = Nothing
    
    'CadenadeLog IdEqp
    
    ' Una vez tenga los datos del articulo y de la linea, colocar en los campos correspondientes
    'MyRecAux.Open "select * from dbo.lineadocumento where IdArticulo in (select IdArticulo from dbo.articulosfamilia where IdFamilia = " & CStr(IdFamAux) & ") and(((SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(29)) and (SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(28))) or (SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) = CHAR(29) and dbo.LineaDocumento.IdLineaAbono>0)) and (IdDocumento in (select IdDocumento from dbo.documento where NumeroTerminal = " & CStr(mBal) & "))", MyConnObj, adOpenKeyset, adLockOptimistic
    MyRecAux.Open "select * from dbo.lineadocumento where IdArticulo in (select IdArticulo from dbo.articulosfamilia where IdFamilia = " & CStr(IdFamAux) & ") and(((SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(29)) and (SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(28))) or (SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) = CHAR(29) and dbo.LineaDocumento.IdLineaAbono>0)) and (dbo.LineaDocumento.idbalanza= '" & IdEqp & "')", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        'CadenadeLog "Error en Select principal..."
        Resp = 1000
        GoTo Fin_QueryLineBovine_ScalePos
    End If
    
    
    On Error GoTo Fin_QueryLineBovine_ScalePos
    If Not MyRecAux.EOF Then
        
        'CadenadeLog "Existen Lineas..."
        Sleep (40)
        
        mLine.IdNumLinea = MyRecAux.Fields("IdLineaDocumento")
        mLine.IdNumDocumento = MyRecAux.Fields("IdDocumento")
        mLine.line_abono = MyRecAux.Fields("IdLineaAbono") ' Linea abonada
        
        'CadenadeLog "1"
        
        IdArt = MyRecAux.Fields("IdArticulo")
        
        'CadenadeLog "2"
        
        mLine.line_amount = MyRecAux.Fields("importeconimpuestos") 'c2f 29-08-2012 Necesitan importe con impuestos MyRecAux.Fields("importe")
        mLine.line_price = MyRecAux.Fields("precio")
        mLine.line_tare = MyRecAux.Fields("tara")
        
        'CadenadeLog "3"
        
        If MyRecAux.Fields("isventaporpeso") = 0 Then
            mLine.line_weight = MyRecAux.Fields("unidades")
            mLine.item_weighted = "0"
        Else
            mLine.item_weighted = "1"
            mLine.line_weight = MyRecAux.Fields("peso")
        End If
        
        'CadenadeLog "4"
        
        mLine.item_descripcion = MyRecAux.Fields("descripcion")
        
        'CadenadeLog "5"
        
        If IsNull(MyRecAux("lote")) = True Then
            mLine.item_bovine = ""
        Else
            mLine.item_bovine = MyRecAux("lote")
        End If
        
        'CadenadeLog "6"
        
        MyRecAux.Close
        Set MyRecAux = Nothing
        
        ' Consultamos el articulo
        If IdArt <> 0 Then
            On Error Resume Next
            MyRecAux.Open "select codigo,caducidad from dbo.articulo where idarticulo = " & IdArt, MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo Fin_QueryLineBovine_ScalePos
            End If
            
            'CadenadeLog "Articulo: " & IdArt
            
            'On Error GoTo Fin_QueryLineBovine_ScalePos
            If Not MyRecAux.EOF Then
                'CadenadeLog "Caducidad:" & CStr(MyRecAux.Fields("caducidad"))
                mLine.item_cad = MyRecAux.Fields("caducidad")
                'CadenadeLog "Código:" & CStr(MyRecAux.Fields("codigo"))
                mLine.item_code = MyRecAux.Fields("codigo")
            Else
                mLine.item_cad = 0
                mLine.item_code = 0
            End If
            mLine.item_fam = mFam
            'CadenadeLog "Familia:" & CStr(mFam)
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If
                
        ' Leeemos modo de trabajo
        If mLine.IdNumDocumento <> 0 Then
            On Error Resume Next
            'CadenadeLog "select TipoDocumento from dbo.documento where iddocumento = " & mLine.IdNumDocumento
            MyRecAux.Open "select TipoDocumento from dbo.documento where iddocumento = " & mLine.IdNumDocumento, MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo Fin_QueryLineBovine_ScalePos
            End If
            
            'CadenadeLog "documento: " & mLine.IdNumDocumento
            On Error GoTo Fin_QueryLineBovine_ScalePos
            If Not MyRecAux.EOF Then
                mLine.scale_work = MyRecAux.Fields("tipodocumento")
            Else
                mLine.scale_work = 0
            End If
                    
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If
                
        ' De momento no se puede consultar
        mLine.scale_num = 0
        mLine.scale_step = 0
        ' --------------------------------
        
        ' Actualizamos el valor de descripcion
        ' Por ultimo marcar con un chr(12) el campo descriptivo de la línea
        On Error Resume Next
        MyRecAux.Open "select * from dbo.lineadocumento where IdLineaDocumento = " & mLine.IdNumLinea, MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo Fin_QueryLineBovine_ScalePos
        End If
        
        On Error GoTo Fin_QueryLineBovine_ScalePos
        If Not MyRecAux.EOF Then
            If mLine.line_abono > 0 Then
                MyRecAux.Fields("descripcion") = MyRecAux.Fields("descripcion") & Chr(28)
            Else
                MyRecAux.Fields("descripcion") = MyRecAux.Fields("descripcion") & Chr(29)
            End If
    'CadenadeLog "marcado de lineadocumento: " & mLine.IdNumLinea
            MyRecAux.Update
        End If
    Else
        'CadenadeLog "No existen Lineas..."
        Resp = 4
    End If
    
    MyRecAux.Close
    Set MyRecAux = Nothing

Fin_QueryLineBovine_ScalePos:
    If Err.Number <> 0 Then
        CadenadeLog "Error en QueryLineBovine_ScalePos. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    QueryLineBovine_ScalePos = Resp
    
End Function

' Programacion de lote de una ficha de vacuno
Public Function ProgLineBovine_ScalePos(vLineBov As Epel_LinBovine) As Long
Dim Resp As Long
Dim cm As String
    Resp = 0
    
    On Error GoTo Fin_PrgLineBovine_ScalePos
    cm = "update dbo.LineaDocumento set Lote='" & vLineBov.item_bovine & "' where IdLineaDocumento=" & CStr(vLineBov.IdNumLinea) & " and IdDocumento=" & CStr(vLineBov.IdNumDocumento)
    MyConnObj.Execute cm

Fin_PrgLineBovine_ScalePos:
    If Err.Number <> 0 Then
        CadenadeLog "Error en ProgLineBovine_ScalePos_ScalePos. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    ProgLineBovine_ScalePos = Resp
End Function


' Peticion de lineas de vacuno para label vasco
Public Function MarkLinesBovine_ScalePos(ByVal mBal As Integer) As Long
Dim Resp As Long
Dim cm As String
    On Error Resume Next
    Resp = 0
    cm = ""
    
    cm = "update dbo.lineadocumento set descripcion=descripcion + CHAR(29) where SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(29) and SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(28) and iddocumento in (select iddocumento from dbo.documento where numeroterminal=" & CStr(mBal) & ")"
    MyConnObj.Execute cm
    
    cm = "update dbo.lineadocumento set descripcion=descripcion + CHAR(28) where (dbo.lineadocumento.idlineaabono>0) and SUBSTRING(dbo.lineadocumento.Descripcion,LEN(dbo.lineadocumento.Descripcion),1) <> CHAR(28)  and iddocumento in (select iddocumento from dbo.documento where numeroterminal=" & CStr(mBal) & ")"
    MyConnObj.Execute cm
    

Fin_MarkLinesBovine_ScalePos:
    If Err.Number <> 0 Then
        CadenadeLog "Error en MarkLinesBovine_ScalePos. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    MarkLinesBovine_ScalePos = Resp
End Function

'Tablas a tocar para artículos...
'dbo.articuloimagen where idarticulo=" & CStr(Code) --> No pasan el dato. Si existe <Images\Items\code> se envia
'                           IdArticulo
'                           Imagen
'dbo.articulo where idarticulo=" & CStr(Code)
'                   IdArticulo
'                   Codigo
'                   IdGlobal
'                   Descripcion
'                   IdImpuestoVenta
'                   IdImpuestoTakeAway
'                   IdImpuestoCompra
'                   TipoVenta
'                   TaraVenta
'                   TaraEnvasado
'                   Caducidad
'                   ConsumoPreferente
'                   PesoAproximado
'                   Color
'                   IsDescatalogado
'                   IsIdArticuloDefinitivo
'                   Merma
'                   Ingredientes
'                   TipoPreseleccion
'                   UnidadesPreseleccion
'                   TipoEtiqueta
'                   IdPlantillaEtiqueta
'dbo.articulosfamilia where idarticulo=" & CStr(Code)
'                           IdFamilia
'                           IdArticulo
'                           Posicion
'dbo.articuloreferencia where idarticulo=" & CStr(Code)
'                               Referencia
'                               IdSeccion
'                               IdArticulo
'dbo.precio where idarticulo=" & CStr(Code)
'                   IdPrecio
'                   IdTarifa
'                   IdArticulo
'                   UnidadesMinimo
'                   UnidadesMaximo
'                   PrecioDefecto
'                   FechaInicioOferta
'                   FechaFinOferta
'                   PrecioOferta
'dbo.ArticuloDescripcion where idarticulo=" & CStr(Code)
'                               IdArticulo
'                               NumLinea
'                               Descripcion
'dbo.CodigoBarra where idarticulo=" & CStr(Code)
'                       CodigoBarra
'                       IdArticulo
'''¡¡¡¡DAR DE ALTA LA TARIFA 1 "Tarifa Defecto"
'dbo.Tarifa where IdTarifa=1
'                 IdTarifa
'                 Descripcion
'''¡¡¡¡DAR DE ALTA LA FAMILIA EN SU CORRESPONDIENTE TABLA...
'dbo.Familia where IdFamilia=X
'                   IdFamilia
'                   Descripcion
'                   IdGlobal
'                   Imagen  --> No pasan el dato. Si existe <Images\Families\IdFamilia> se envia
'                   Color
'                   IdTipoFichaTrazabilidad
'                   IdPlantillaEtiqueta
'                   Codigo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Type Epel_FullItem
'    Code As Long
'    Plu As Long
'    Sec As Long
'    SubSec As Long
'    Family As Long
'    Price As Long
'    weight As Long
'    Caducity As Long
'    Font As Long
'    Text As String
'    Tare As Long
'    Pref As Long
'    Etq As Long
'    vat As Long
'    EAN13 As String
'    Desc(20) As String
'    FDesc(10) As Long
'    Loss As Long
'    Presel As Long
'    Desc2040(20) As String
'    lMix As Boolean
'    nPeso As Long
'End Type
Public Function Capa1B_SendFullItem_ScalePos(Item As Epel_FullItem) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim lEsAlta As Boolean
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim nPrice As Double
    
    lEsAlta = False
    
    'PonDato "Capa1B_SendFullItem", 0
    If Item.Plu > 9999 Then
        Item.Plu = 0
        'PonDato "Se pasa PLU a 0 para código:" & CStr(Item.Code)
    End If
    On Error Resume Next
    'PonDato "Familia...", 0
    MyRecAux.Open "select * from dbo.familia where idfamilia=" & CStr(Item.Family), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        lEsAlta = True
    End If
    MyRecAux.Fields("idfamilia") = Item.Family

    If Dir(App.Path & "\images\families\" & CStr(Item.Family)) <> "" Then
        On Error GoTo noImgFam1
        If FileLen(App.Path & "\images\families\" & CStr(Item.Family)) > 0 Then
            'Imagen
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\images\families\" & CStr(Item.Family)
            MyRecAux.Fields("imagen").Value = strStream.Read
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
        End If
noImgFam1:
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
    Else
        If Dir(App.Path & "\fmlimg.jpg") <> "" Then
        On Error GoTo noImgFam2
        If FileLen(App.Path & "\fmlimg.jpg") > 0 Then
            
            'Imagen defecto
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\fmlimg.jpg"
            MyRecAux.Fields("imagen").Value = strStream.Read
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
            MyRecAux.Fields("color") = "#FFC2FF49" '"#FFCCCCCC"
            MyRecAux.Fields("descripcion") = "Familia " & CStr(Item.Family)
        End If
noImgFam2:
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
        End If

    End If
    If lEsAlta Then
        MyRecAux.Fields("codigo") = Item.Family
        MyRecAux.Fields("descripcion") = "Familia " & CStr(Item.Family)
    End If
    MyRecAux.Update
    'si es alta...
    If lEsAlta Then
        cm = "update dbo.familia set idglobal=NEWID() where idfamilia=" & CStr(Item.Family)
        MyConnObj.Execute cm
        ''''''''''''''
    End If
    lEsAlta = False
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    On Error Resume Next
    'PonDato "FamiliasSeccion...", 0
    MyRecAux.Open "select * from dbo.familiasseccion where idfamilia=" & CStr(Item.Family), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("posicion") = 1000000
    End If
    MyRecAux.Fields("idfamilia") = Item.Family
    MyRecAux.Fields("idseccion") = Item.Sec

    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    On Error Resume Next
    'PonDato "Tarifa...", 0
    MyRecAux.Open "select * from dbo.tarifa where idtarifa=1", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idtarifa") = 1
        MyRecAux.Fields("descripcion") = "Tarifa Defecto"
        MyRecAux.Update
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    On Error Resume Next
    'PonDato "Articulo...", 0
    MyRecAux.Open "select * from dbo.articulo where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        lEsAlta = True
    End If
    CadenadeLog "Añadido registro... Enrada Datos..."
'                   IdArticulo
'                   Codigo
'                   IdGlobal
'                   Descripcion
'                   IdImpuestoVenta
'                   IdImpuestoTakeAway
'                   IdImpuestoCompra
'                   TipoVenta
'                   TaraVenta
'                   TaraEnvasado
'                   Caducidad
'                   ConsumoPreferente
'                   PesoAproximado
'                   Color
'                   IsDescatalogado
'                   IsIdArticuloDefinitivo
'                   Merma
'                   Ingredientes
'                   TipoPreseleccion
'                   UnidadesPreseleccion
'                   TipoEtiqueta
'                   IdPlantillaEtiqueta
    
    MyRecAux.Fields("idarticulo") = Item.Code
    'CadenadeLog "Código:" & CStr(Item.Code)
    MyRecAux.Fields("codigo") = CStr(Item.Code)
    MyRecAux.Fields("descripcion") = Item.Text
    MyRecAux.Fields("idimpuestoventa") = 154 + Item.vat
    MyRecAux.Fields("idimpuestotakeaway") = 154 + Item.vat
    MyRecAux.Fields("idimpuestocompra") = 0
    If Item.weight <> 0 Then
        MyRecAux.Fields("tipoventa") = 0
    Else
        MyRecAux.Fields("tipoventa") = 1
    End If
    'CadenadeLog "Pesado/no Pesado:" & CStr(Item.weight)
    MyRecAux.Fields("taraventa") = Item.Tare / 1000
    MyRecAux.Fields("taraenvasado") = Item.Tare / 1000
    If Item.Caducity <> 0 Then
        MyRecAux.Fields("caducidad") = Item.Caducity
    End If
    MyRecAux.Fields("consumopreferente") = Item.Pref
    MyRecAux.Fields("pesoaproximado") = 0 'Item.nPeso
    If lEsAlta And IsNumeric(miartXtra.Imagen) = False Then
        MyRecAux.Fields("color") = "#FF9BBF5F"
        'CadenadeLog "#FF9BBF5F"
    End If
    If IsNumeric(miartXtra.Imagen) Then
        MyRecAux.Fields("color") = APaleta(Val(miartXtra.Imagen))
        'CadenadeLog APaleta(Val(miartXtra.Imagen))
    End If
    MyRecAux.Fields("isdescatalogado") = 0
    MyRecAux.Fields("isidarticulodefinitivo") = 0
    MyRecAux.Fields("merma") = Item.Loss
    If Item.weight <> 0 Then
        MyRecAux.Fields("tipopreseleccion") = 1
    Else
        MyRecAux.Fields("tipopreseleccion") = 0
    End If
    'MyRecAux.Fields("tipopreseleccion") = 0
    MyRecAux.Fields("unidadespreseleccion") = Item.Presel
    MyRecAux.Fields("tipoetiqueta") = Item.Etq
    'CadenadeLog "Etiqueta:" & CStr(Item.Etq)
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
    If lEsAlta Then
        'si es alta...
        cm = "update dbo.articulo set idglobal=NEWID() where idarticulo=" & CStr(Item.Code)
        MyConnObj.Execute cm
        ''''''''''''''
        'CadenadeLog "Código NewId"
    End If
'dbo.articuloimagen where idarticulo=" & CStr(Code) --> No pasan el dato. Si existe <Images\Items\code> se envia
'                           IdArticulo
'                           Imagen
    If Dir(App.Path & "\images\items\" & CStr(Item.Code)) <> "" Then
        On Error GoTo noImgArt1
        If FileLen(App.Path & "\images\items\" & CStr(Item.Code)) > 0 Then
            MyRecAux.Open "select * from dbo.articuloimagen where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
            On Error Resume Next
            PonDato "Imagen Artículo...", 0
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo FIN_Capa1B_SendFullItem
            End If
            On Error GoTo FIN_Capa1B_SendFullItem
            
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            MyRecAux.Fields("idarticulo") = Item.Code
            
            'Imagen
            strStream.type = adTypeBinary
            strStream.Open
            strStream.LoadFromFile App.Path & "\images\items\" & CStr(Item.Code)
            MyRecAux.Fields("imagen").Value = strStream.Read
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
            '''''''''''''''
            strStream.Flush
            strStream.Close
            Set strStream = Nothing
            ''''''''
        End If
noImgArt1:
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
    Else
        If Trim(miartXtra.Imagen) <> "" And IsNumeric(miartXtra.Imagen) = False Then
            If Dir(Trim(miartXtra.Imagen)) <> "" Then
                If FileLen(Trim(miartXtra.Imagen)) > 0 Then
                    MyRecAux.Open "select * from dbo.articuloimagen where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
                    On Error Resume Next
                    'PonDato "Imagen Artículo...", 0
                    If Err.Number <> 0 Then
                        Resp = 1000
                        GoTo FIN_Capa1B_SendFullItem
                    End If
                    On Error GoTo FIN_Capa1B_SendFullItem
                    
                    If MyRecAux.EOF Then
                        MyRecAux.AddNew
                    End If
                    MyRecAux.Fields("idarticulo") = Item.Code
                    
                    'Imagen
                    strStream.type = adTypeBinary
                    strStream.Open
                    strStream.LoadFromFile Trim(miartXtra.Imagen)
                    MyRecAux.Fields("imagen").Value = strStream.Read
                    MyRecAux.Update
                    MyRecAux.Close
                    Set MyRecAux = Nothing
                    '''''''''''''''
                    strStream.Flush
                    strStream.Close
                    Set strStream = Nothing
                    ''''''''
                End If
            End If
        Else
            If Dir(App.Path & "\itmimg.jpg") <> "" Then
                On Error GoTo noImgArt2
                If FileLen(App.Path & "\itmimg.jpg") > 0 Then
                MyRecAux.Open "select * from dbo.articuloimagen where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
                On Error Resume Next
                'PonDato "Imagen Artículo...", 0
                If Err.Number <> 0 Then
                    Resp = 1000
                    GoTo FIN_Capa1B_SendFullItem
                End If
                On Error GoTo FIN_Capa1B_SendFullItem
                
                If MyRecAux.EOF Then
                    MyRecAux.AddNew
                End If
                MyRecAux.Fields("idarticulo") = Item.Code
                
                'Imagen
                strStream.type = adTypeBinary
                strStream.Open
                strStream.LoadFromFile App.Path & "\itmimg.jpg"
                MyRecAux.Fields("imagen").Value = strStream.Read
                MyRecAux.Update
                MyRecAux.Close
                Set MyRecAux = Nothing
                '''''''''''''''
                strStream.Flush
                strStream.Close
                Set strStream = Nothing
                ''''''''
                End If
noImgArt2:
                If Err.Number <> 0 Then
                    On Error GoTo 0
                End If
            End If
        End If
    End If
    
'dbo.articulosfamilia where idarticulo=" & CStr(Code)
'                           IdFamilia
'                           IdArticulo
'                           Posicion
    On Error Resume Next
    'PonDato "articulosfamilia...", 0
    MyRecAux.Open "select * from dbo.articulosfamilia where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idfamilia") = Item.Family
        MyRecAux.Fields("posicion") = Item.Plu
    End If
    MyRecAux.Fields("idarticulo") = Item.Code

    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
'dbo.articuloreferencia where idarticulo=" & CStr(Code)
'                               Referencia
'                               IdSeccion
'                               IdArticulo
    On Error Resume Next
    'PonDato "Referencia...", 0
    If Item.Plu <> 0 Then
        MyRecAux.Open "select * from dbo.articuloreferencia where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Capa1B_SendFullItem
        End If
        '1.0.9 On Error GoTo FIN_Capa1B_SendFullItem
        On Error GoTo sigueREF
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        End If
        MyRecAux.Fields("idarticulo") = Item.Code
        MyRecAux.Fields("idseccion") = Item.Sec
        MyRecAux.Fields("referencia") = Item.Plu
        MyRecAux.Update
sigueREF:
        On Error GoTo 0
        MyRecAux.Close
        Set MyRecAux = Nothing
    End If
'dbo.precio where idarticulo=" & CStr(Code)
'                   IdPrecio
'                   IdTarifa
'                   IdArticulo
'                   UnidadesMinimo
'                   UnidadesMaximo
'                   PrecioDefecto
'                   FechaInicioOferta
'                   FechaFinOferta
'                   PrecioOferta
    On Error Resume Next
    'PonDato "PK precio", 0
    MyRecAux.Open "select COUNT(*) as numprecios from dbo.precio", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    nC = MyRecAux.Fields("numprecios")
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    On Error Resume Next
    'PonDato "Precio...", 0
    MyRecAux.Open "select * from dbo.precio where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        Resp = 1000
        GoTo FIN_Capa1B_SendFullItem
    End If
    On Error GoTo FIN_Capa1B_SendFullItem
    If MyRecAux.EOF Then
        MyRecAux.AddNew
        MyRecAux.Fields("idprecio") = nC + 1
    End If
    MyRecAux.Fields("idarticulo") = Item.Code
    MyRecAux.Fields("idtarifa") = 1
    MyRecAux.Fields("unidadesminimo") = 0
    MyRecAux.Fields("unidadesmaximo") = 999999
    nPrice = Round(Item.Price / 100, 2)
    MyRecAux.Fields("preciodefecto") = nPrice
    MyRecAux.Fields("preciooferta") = 0
    MyRecAux.Update
    MyRecAux.Close
    Set MyRecAux = Nothing
'dbo.ArticuloDescripcion where idarticulo=" & CStr(Code)
'                               IdArticulo
'                               NumLinea
'                               Descripcion
    nC = 0
    For nCont = 0 To 19
        If Len(Trim(Item.Desc(nCont))) > 0 Then
            nC = nC + 1
            On Error Resume Next
            'PonDato "Descriptivo:" & CStr(nC), 0
            MyRecAux.Open "select * from dbo.articulodescripcion where idarticulo=" & CStr(Item.Code) & " and numlinea=" & CStr(nC), MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                Resp = 1000
                GoTo FIN_Capa1B_SendFullItem
            End If
            On Error GoTo FIN_Capa1B_SendFullItem
            
            If MyRecAux.EOF Then
                MyRecAux.AddNew
            End If
            MyRecAux.Fields("idarticulo") = Item.Code
            MyRecAux.Fields("numlinea") = nC
            MyRecAux.Fields("descripcion") = Trim(Item.Desc(nCont))
            MyRecAux.Update
            MyRecAux.Close
            Set MyRecAux = Nothing
        End If
    Next nCont
'dbo.CodigoBarra where idarticulo=" & CStr(Code)
'                       CodigoBarra
'                       IdArticulo
    'If Len(Trim(Item.EAN13)) >= 12 Then
        Item.EAN13 = Trim(Item.EAN13) 'Mid(Item.EAN13, 1, 12) & Chr(48 + DigitoControl(Mid(Item.EAN13, 1, 12)))
        
        '1.0.9
        'On Error Resume Next
        ''PonDato "Código Barras...", 0
        'MyRecAux.Open "select * from dbo.codigobarra where idarticulo=" & CStr(Item.Code), MyConnObj, adOpenKeyset, adLockOptimistic
        'If Err.Number <> 0 Then
        '    Resp = 1000
        '    GoTo FIN_Capa1B_SendFullItem
        'End If
        'On Error GoTo FIN_Capa1B_SendFullItem
        'If MyRecAux.EOF Then
        '    MyRecAux.AddNew
        'End If
        'MyRecAux.Fields("idarticulo") = Item.Code
        'MyRecAux.Fields("codigobarra") = Item.EAN13
        'MyRecAux.Update
        'MyRecAux.Close
        'Set MyRecAux = Nothing
        ''''''''''''''''''''''''
        On Error Resume Next
        MyRecAux.Open "select * from dbo.codigobarra where codigobarra='" & Item.EAN13 & "'", MyConnObj, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            Resp = 1000
            GoTo FIN_Capa1B_SendFullItem
        End If
        If MyRecAux.EOF Then
            MyRecAux.AddNew
        Else
            If MyRecAux.Fields("idarticulo") <> Item.Code Then
                cm = "delete dbo.codigobarra where codigobarra='" & Item.EAN13 & "'"
                MyConnObj.Execute cm
                MyRecAux.AddNew
            End If
        End If
        MyRecAux.Fields("idarticulo") = Item.Code
        MyRecAux.Fields("codigobarra") = Item.EAN13
        MyRecAux.Update
        MyRecAux.Close
        Set MyRecAux = Nothing
        
    'End If
    
'TRAMOS''''''''''''''''
    'CadenadeLog "prc1:" & CStr(miartXtra.PRC1) & " prc2:" & CStr(miartXtra.PRC2)
    If (miartXtra.TRM1 <> 0 Or miartXtra.TRM2 <> 0) And (miartXtra.PRC1 > miartXtra.PRC2) Then
        cm = "delete dbo.precio where idarticulo=" & CStr(Item.Code) & " and unidadesminimo<>0"
        'CadenadeLog cm
        MyConnObj.Execute cm
        cm = "update dbo.precio set unidadesmaximo=" & Quitar_Coma_S(CStr(Round((miartXtra.TRM1) / 1000, 3))) & " where idarticulo=" & CStr(Item.Code) & " and unidadesminimo=0"
        MyConnObj.Execute cm
        If miartXtra.TRM2 <> 0 Then
            cm = "insert into dbo.precio (idprecio,idarticulo,idtarifa,unidadesminimo,preciooferta,preciodefecto," & _
                  "unidadesmaximo) values ( (select max(ISNULL(idprecio,0)) + 1 from dbo.precio)," & CStr(Item.Code) & ",1," & Quitar_Coma_S(CStr(Round((miartXtra.TRM1) / 1000, 3))) & ",0," & Quitar_Coma_S(CStr(Round((miartXtra.PRC1) / 100, 2))) & _
                  "," & Quitar_Coma_S(CStr(Round((miartXtra.TRM2) / 1000, 3))) & ")"
            'CadenadeLog "A:" & cm
            MyConnObj.Execute cm
            cm = "insert into dbo.precio (idprecio,idarticulo,idtarifa,unidadesminimo,preciooferta,preciodefecto," & _
                  "unidadesmaximo) values ((select max(ISNULL(idprecio,0)) + 1 from dbo.precio)," & CStr(Item.Code) & ",1," & Quitar_Coma_S(CStr(Round((miartXtra.TRM2) / 1000, 3))) & ",0," & Quitar_Coma_S(CStr(Round((miartXtra.PRC2) / 100, 2))) & _
                  ",999999)"
            'CadenadeLog "B:" & cm
            MyConnObj.Execute cm
        Else
            cm = "insert into dbo.precio (idprecio,idarticulo,idtarifa,unidadesminimo,preciooferta,preciodefecto," & _
                  "unidadesmaximo) values ((select max(ISNULL(idprecio,0)) + 1 from dbo.precio)," & CStr(Item.Code) & ",1," & Quitar_Coma_S(CStr(Round((miartXtra.TRM1) / 1000, 3))) & ",0," & Quitar_Coma_S(CStr(Round((miartXtra.PRC1) / 100, 2))) & _
                  ",999999)"
            'CadenadeLog "C:" & cm
            MyConnObj.Execute cm
        End If
''''''''''''''''''''''
    End If
    
    Resp = 0
    
FIN_Capa1B_SendFullItem:
    If Err.Number <> 0 Then
        CadenadeLog "Error en SendFullItem. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If
    
    PonDato "Articulo Cod: " & Item.Code & " Plu: " & Item.Plu & " Sec: " & Item.Sec & " Desc: " & Item.Text, Resp
    Capa1B_SendFullItem_ScalePos = Resp
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////
'**********************************************************************************************
' Peticiones de datos de totales a ficheros
'**********************************************************************************************
'///////////////////////////////////////////////////////////////////////////////////////////////
Public Function F_Query_All_Tickets_ScalePos(lType As Long, Mode As Long, MyPath As String, Optional ByVal lPago As Boolean) As Long
    Dim Arch As Integer
    Dim Resp As Long
    Dim Resul As Long
    Dim MiTicket As Epel_CabTick
    Dim Buf1 As String
    Dim Bucle As Integer
    Dim MyFlag As Boolean
    Dim Berror As Long
    Dim NumTickGB As Long
    NumTickGB = 1
    If Dir(App.Path & "\cancel.tqt") <> "" Then
        On Error Resume Next
        Kill App.Path & "\cancel.tqt"
        On Error GoTo 0
    End If
    Do_Cancel = False
    Resp = 0
    Resul = 0
    '1 º check acceso al fichero
    If Mode = 1 Then
        MyFlag = True
    Else
        MyFlag = False
    End If
    Arch = FreeFile()
    On Error Resume Next
    Open MyPath For Append As #Arch
    If Err.Number = 0 Then
        On Error GoTo 0
        Do Until Resul <> 0
            For Berror = 1 To 3
                If Mode = 1 Then
                    Resul = Capa1B_QueryTicket_ScalePos(1000 + lType, MiTicket)
                Else
                    Resul = Capa1B_QueryTicket_ScalePos(lType, MiTicket)
                End If
                If Resul = 0 Or Resul >= 1000 Then
                    If Resul >= 1000 Then Resp = 1000
                    Exit For
                End If
            Next Berror
            If Resul = 0 Then
                '*****************
                ' introduce datos de la cabecera
                '*****************
                NumTickGB = MiTicket.NTicket + 1
                Buf1 = "C"
                Buf1 = Buf1 & MiTicket.type & Format(MiTicket.NTicket, "00000")
                Buf1 = Buf1 & Format(MiTicket.Section, "00") & Format(MiTicket.Machine, "00")
                Buf1 = Buf1 & Format(MiTicket.Vendor, "0000") & Format(MiTicket.Client, "00000")
                Buf1 = Buf1 & MiTicket.date & MiTicket.Hour
                Buf1 = Buf1 & Format(MiTicket.nLines, "0000") & Format(MiTicket.Cancelled, "0")
                Buf1 = Buf1 & Format(MiTicket.positive, "0") & Format(MiTicket.amount, "000000000")
                Buf1 = Buf1 & Format(MiTicket.P1, "000000000") & Format(MiTicket.P2, "000000000")
                Buf1 = Buf1 & Format(MiTicket.Task, "0")
                Buf1 = sAjusta(Buf1)
                On Error Resume Next
                Print #Arch, Buf1
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Resp = 1101
                    Exit Do
                Else
                    On Error GoTo 0
                    '*******************
                    ' introduce las líneas
                    '*******************
                    For Bucle = 1 To MiTicket.nLines
                        Buf1 = "L"
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Number, "0000")
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).positive
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).Cancelled
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).item_code, "000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Price, "0000000")
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).Weighted
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Pes, "000000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).amount, "000000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).BovineCard, "00000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).discount, "00000")
                        
                        On Error Resume Next
                        Buf1 = sAjusta(Buf1)
                        Print #Arch, Buf1
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            Resp = 1101
                            Exit For
                        Else
                            On Error GoTo 0
                            '******************
                            EnviaEvento.Raise_TicketReceived MiTicket.NTicket, MiTicket.Section, MiTicket.Machine, Do_Cancel
                            '*******************
                            If Dir(App.Path & "\cancel.tqt") <> "" Then
                                Do_Cancel = 4
                            End If
                            If Do_Cancel <> False Then
                                CadenadeLog "Recibida orden de cancelación de recepción de tiquets"
                                Resp = 1102
                                Exit For
                            End If
                        End If
                    Next Bucle
                    If Resp = 1101 Or Resp = 1102 Then Exit Do
                End If
            End If
        Loop
    Else
        On Error GoTo 0
        Resp = 1100
        Resul = 1100
    End If
    Close #Arch
    F_Query_All_Tickets_ScalePos = Resul
    If Resul >= 1000 Then
        EnviaEvento.Raise_ComError Resul
    Else
        EnviaEvento.Raise_ComOK
    End If
End Function


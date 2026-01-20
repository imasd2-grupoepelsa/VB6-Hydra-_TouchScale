Attribute VB_Name = "EpelOcx"
Option Explicit

Public MyConnObj As New ADODB.Connection
Public MyRecSet As New ADODB.Recordset
Public Estado As Long
Public lMarcaTime As Boolean
Public sAddIP As String
Public sPortBal As String
Public EstadoBal As Integer
Public cancelar As Integer
Public Type Tipo_Grupo_IVA
    porcentaje As Double
    Base As Double
    Importe As Double
End Type
Public Type Tipo_TqCodBar
    Tipo As Integer
    NBalanza As Integer
    NMostrador As Integer
    NTicket As Long
    NVendedor As Long
    Importe As Double
End Type
Public Type Tipo_Linea
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
    porcentaje As Double
End Type


Public Type tipo_cabecera
     Ntiquet As Long
     NVendedor As Long
     NMostrador As Long
     NCliente As Variant
     Fecha As Variant
     hora As Variant
     NBalanza As Long
     ImporteTotal As Double
     nLineas As Long
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
     Trainning As Boolean
     Lines() As Tipo_Linea
     base0 As Double
     por0 As Double
     imp0 As Double
     base1 As Double
     por1 As Double
     imp1 As Double
     base2 As Double
     por2 As Double
     imp2 As Double
     base3 As Double
     por3 As Double
     imp3 As Double
     base4 As Double
     por4 As Double
     imp4 As Double
     factura As String
     EAN13 As String
End Type
Public Type Tipo_NombreTiquet
    Balanza As Integer
    seccion As Integer
    tiquet As Long
    modo As Integer
    cliente As Long
    vendedor As Integer
End Type

Public Type Epel_FullItem
    Code As Long
    Plu As Long
    Sec As Long
    SubSec As Long
    Family As Long
    Price As Long
    weight As Long
    caducity As Long
    Font As Long
    tare As Long
    pref As Long
    etq As Long
    vat As Long
    Loss As Long
    Presel As Long
    nPeso As Long
    Traza As Long
    Etiqueta2 As Long
    Retirada As Long
    FamilyOrder As Long
    lMix As Boolean
    lTraza As Boolean
    sendTraza As Boolean
    EAN13 As String
    sNameShort As String
    sOper As String
    sText1 As String
    sText2 As String
    sLongText As String
End Type
Public Function Get_Status() As Long
    Get_Status = MyConnObj.State
End Function


Public Function Configure(ByVal MachineID As Long, ByVal TypeConf As Long, ByVal sConfigure As String) As Long
    Dim Resp As Long
    If Dir(App.Path & "\marcatime") <> "" Then
        lMarcaTime = True
    Else
        lMarcaTime = False
    End If
    'CadenadeLog "Recibida orden de configuración.Máquina : " & MachineID & ".Tipo conexión : " & TypeConf & ".Cadena Conexión : " & sConfigure
    Resp = Capa0_Configure_Touch(MachineID, TypeConf, sConfigure)
    Configure = Resp
End Function

Public Sub reset()
    If MyConnObj.State Then
        CadenadeLog "Estado de la conexión:" & CStr(MyConnObj.State) & " Se cierra conexión Existente previamente."
        On Error Resume Next
        MyConnObj.Close
        Set MyConnObj = Nothing
        If Err.Number <> 0 Then
            CadenadeLog "Error cerrando conexión previamente abierta:" & CStr(Err.Number) & " descr.:" & Err.Description
        End If
        On Error GoTo 0
    End If
End Sub

Public Function Capa0_Configure_Touch(ByVal MachineID As Long, ByVal TypeConf As Long, sConfigure As String) As Long
Dim Bytes(5) As String
Dim Contador As Integer
Dim Bucle As Integer
Dim Resp As Boolean
Dim Puerto As Long
Dim sIp As String
Dim nResp As Integer
    'If MyConnObj.State Then
    '    CadenadeLog "Estado de la conexión:" & CStr(MyConnObj.State) & " Se cierra conexión Existente previamente."
    '    On Error Resume Next
    '    MyConnObj.Close
    '    Set MyConnObj = Nothing
    '    If Err.Number <> 0 Then
    '        CadenadeLog "Error cerrando conexión previamente abierta:" & CStr(Err.Number) & " descr.:" & Err.Description
    '    End If
    '    On Error GoTo 0
    'End If
    
    sIp = ""
    For Contador = 0 To 4
        Bytes(Contador) = ""
    Next Contador
    Contador = 0
    Resp = False
    For Bucle = 1 To Len(sConfigure)
        Select Case Asc(Mid(sConfigure, Bucle, 1))
            Case 48 To 57
                Bytes(Contador) = Bytes(Contador) & Mid(sConfigure, Bucle, 1)
            Case Asc(".")
                Contador = Contador + 1
            Case Asc(":")
                Contador = 4
            Case Else
                Exit For
        End Select
        If Contador > 4 Then
            Exit For
        End If
    Next Bucle
    If Bytes(0) <> "" And Bytes(1) <> "" And Bytes(2) <> "" And Bytes(3) <> "" Then
        Bytes(0) = Val(Bytes(0))
        Bytes(1) = Val(Bytes(1))
        Bytes(2) = Val(Bytes(2))
        Bytes(3) = Val(Bytes(3))
        If Val(Bytes(0)) < 256 And Val(Bytes(1)) < 256 And Val(Bytes(2)) < 256 And Val(Bytes(3)) < 256 Then
            If Bytes(4) <> "" Then
                Puerto = Bytes(4)
            Else
                Puerto = 3306
            End If
            sPortBal = CStr(Puerto)
            If Puerto >= 1 And Puerto <= 65535 Then
                sIp = CStr(Bytes(0)) & "." & CStr(Bytes(1)) & "." & CStr(Bytes(2)) & "." & CStr(Bytes(3))
                Resp = True
            End If
        Else
            Resp = False
        End If
    Else
        Resp = False
    End If
    If sIp <> "" And Resp = True Then
        On Error Resume Next
        
        nResp = ElPing(sIp)
        If nResp = 1 Then
        
            Resp = False
            'CadenadeLog "Cadena de conexión(Uid y Pass no reales): DRIVER={MySQL ODBC 5.1 Driver};Uid=X;password=Y;DATABASE=pcscale;Data Source=" & sIp
            If Dir(App.Path & "\odbc53.txt") <> "" Then
                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sIp & "; Port=" & CStr(Puerto) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
            Else
                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sIp & "; Port=" & CStr(Puerto) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
            End If
            
            If Err.Number <> 0 Then
                
                CadenadeLog "Error: " & Err.Number & " Descr.: " & Err.Description
                
                On Error GoTo 0
                Resp = False
                On Error Resume Next
                MyConnObj.Close
                Set MyConnObj = Nothing
                On Error GoTo 0
                Estado = 0
            Else
                Resp = True
                Estado = 1
            End If
            On Error GoTo 0
        
        Else
            CadenadeLog "Sin ping con:" & sIp
            Resp = False
        End If
    End If
    If Resp Then
        Capa0_Configure_Touch = 0
        'CadenadeLog "Conexión realizada con:" & sIp
    Else
        Capa0_Configure_Touch = 3
    End If
End Function

Public Sub Switch_M(ByVal MiP As String)

Dim MyFich As Integer
Dim s As String
Dim BufDetiene As Variant
Dim valorip As String
Dim sTi As String
Dim nResp As Integer

    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    EstadoBal = 1000
    On Error GoTo 0
    
    MiP = Val(Left(MiP, 3)) & "." & Val(Mid(MiP, 5, 3)) & "." & Val(Mid(MiP, 9, 3)) & "." & Val(Right(MiP, 3))
    

    '''''''''''''''''''''''''''''''
    'Conectar con sAddIp y sPortBal
    '''''''''''''''''''''''''''''''
    On Error Resume Next

    CadenadeLog "Cadena de Conexión" & "-->" & MiP & "," & sPortBal
    Do_Events

    nResp = ElPing(MiP)
    If nResp = 1 Then
    
        If Dir(App.Path & "\odbc53.txt") <> "" Then
            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        Else
            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        End If
        
        CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
        CadenadeLog "Error:" & CStr(Err.Number)
        
        If Err.Number <> 0 Then

            On Error GoTo 0
            EstadoBal = 1000
            GoTo finMSWITCH
        Else
            EstadoBal = 0
        End If
    
    Else
        'frmmonitor.MostrarDato "No Ping:" & MiP
        EstadoBal = 1000
    End If
                

finMSWITCH:
    If EstadoBal <> 0 Then
        cancelar = 1
    Else
        cancelar = 0
    End If

End Sub

Public Sub Switch_mode(ByVal Scaletype As Integer)
    Dim MyC9 As String
    Dim MyResp As Long
    Dim UsuarioOK As Long
    Dim MyNum As String
    Dim BufDetiene As Boolean
    Dim MyNumMaq As Integer
    Dim MyIP As String
    Dim valorip As String
    Dim nST As Long
    
    If Dir(App.Path & "\cancel.tqt") <> "" Then Kill App.Path & "\cancel.tqt"
    cancelar = 0

    On Error Resume Next
    MyResp = 0
    'nST = Get_Status() '.balanza
    'If nST <> 1 Or (Err.Number <> 0) Then
    '    If (Err.Number <> 0) Then
    '        On Error GoTo 0
    '    End If
    If MyConnObj.State = 1 Or MyConnObj.Errors.Count > 0 Then
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
        MyResp = Configure(11, 0, sAddIP) '.balanza
    'End If
    On Error GoTo 0
    If MyResp = 0 Then
       EstadoBal = 11
       cancelar = 0
       sArrSTATUS(nSerEnCur) = "Con conexión."
    Else
       reset
       EstadoBal = 1000
       cancelar = 1
       sArrSTATUS(nSerEnCur) = "Sin conexión."
    End If
End Sub


Public Sub LeeTickets(Optional Tabla As Integer)
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    Dim MyRecImp As New ADODB.Recordset
    'Tratar directamente con tabla balanza
    'Dim Mybase As DAO.Database
    'Dim myRec As DAO.Recordset
    'Dim MisEqp As New ADODB.Recordset
    
    Dim nLR As Integer
    
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim nCont As Long
    Dim dLImporte As Double
    'Dim nLocal As Integer
    Dim sCap As String
    Dim nFichIn As Integer
    Dim sIn As String
    Dim nC As Integer
    Dim MyS As String
    
    'nLocal = 1
    
    On Error GoTo FIN
    Switch_mode 11
    If EstadoBal <> 11 Then GoTo SalirLeeTique
    'Tratar directamente con tabla balanza
    'Set Mybase = DAO.OpenDatabase(Base_General)
    'If Dir(App.Path & "\tqslaves.txt") <> "" Then
    '    MisEqp.Open "select * from pcscale.scales where id<>0", MyConnObj, adOpenKeyset, adLockOptimistic
    'Else
    '    MisEqp.Open "select * from pcscale.scales where type='S'", MyConnObj, adOpenKeyset, adLockOptimistic
    'End If
    'If MisEqp.EOF Then
    '    GoTo SalirLeeTique
    'Else
    '    MisEqp.MoveFirst
    '    Do While Not MisEqp.EOF
    '
    '        If MisEqp.Fields("type") = "S" Then
    '            Switch_mode 11
    '        Else
    '            Switch_M MisEqp.Fields("ip") & ":3306"
    '        End If
            
            'CadenadeLog "paso 3"
            
            If EstadoBal = 11 Then
            
                For nLR = 0 To 1 ' Red/Local
            
    '            If MisEqp.Fields("type") <> "S" And nLR = 0 Then GoTo SigueNlr
                If nLR = 0 Then

                    If IsNull(Tabla) Or Tabla = 0 Then
                        If Dir(App.Path & "\noencargo.txt") <> "" Then
                            If lMarcaTime Then
                                MyRecSetTmp.Open "select * from pcscale.htickets where (CapturedDateTime='0000-00-00 00:00:00' or  CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                            Else
                                MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or isnull(captured)) and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                            End If
                        Else
                            If lMarcaTime Then
                                MyRecSetTmp.Open "select * from pcscale.htickets where (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                            Else
                                MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or isnull(captured)) and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                            End If
                        End If
                    Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.packetlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.packetlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.sservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.sservicelabels where customer<>'9999999999' order by id"
                    End If
                    End If
                Else
            'MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
            
                    If IsNull(Tabla) Or Tabla = 0 Then
                    'MyRecSetTmp.Open "select * from pcscale.localhtickets where captured='' and status='C' and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    If Dir(App.Path & "\noencargo.txt") <> "" Then
                        MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and TckType<>'O' and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    Else
                        MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    End If
                    'CadenadeLog "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='L' order by id"
                    Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id"
                    End If
                    End If
                End If
            'CadenadeLog "paso 5"
                If Not MyRecSetTmp.EOF Then
                Do While Not MyRecSetTmp.EOF
                    'CadenadeLog "paso 6"
                    Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
                    Fhora = Replace(Mid(Ffecha, 12, 8), ":", "")
                    
                    Ffecha = Mid(Ffecha, 1, 10)
                    'CadenadeLog "paso 7"
                    IdDoc = MyRecSetTmp.Fields("id")
                    If IsNull(MyRecSetTmp.Fields("CapturedDateTime")) Then
                        sCap = "0000-00-00 00:00:00"
                    Else
                        sCap = MyRecSetTmp.Fields("CapturedDateTime")
                    End If
                    sCap = Mid(sCap, 1, 17) & "11"
                    'CadenadeLog "paso 8"
                    ''''medio de pago
                    '1=efectivo
                    '2=t.crédito
                    '3=cheque
                    Cabecera.tipoPago = 1
                    Cabecera.cantidadEntre = 0
                    
                    If IsNull(Tabla) Or Tabla = 0 Then
'Select Case !TYPTIC --> Cabecera.tipoTiquet
'    Case 1
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
'    Case 2
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
'    Case 3
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
'    Case 4
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
'    Case 5
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
'    Case 6
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
'    Case 7
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
'    Case 8
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
'    Case 12
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Descart."
'End Select
                    
                        'If nLR = 0 Then
                        '''ver
                            If MyRecAux.State <> 0 Then
                                MyRecAux.Close
                                Set MyRecAux = Nothing
                            End If
                            MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc)
                            If Not MyRecAux.EOF Then
                                'CadenadeLog "paso 9"
                                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                                'CadenadeLog "paso 10"
                            End If
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'CadenadeLog "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc)
                        '    If Not MyRecAux.EOF Then
                        '        'CadenadeLog "paso 11"
                        '        Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                        '        Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                        '        'CadenadeLog "paso 12"
                        '    End If
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                    End If
        '''''''
                    'Impuestos
                    Cabecera.base0 = 0
                    Cabecera.por0 = 0
                    Cabecera.imp0 = 0
                    Cabecera.base1 = 0
                    Cabecera.por1 = 0
                    Cabecera.imp1 = 0
                    Cabecera.base2 = 0
                    Cabecera.por2 = 0
                    Cabecera.imp2 = 0
                    Cabecera.base3 = 0
                    Cabecera.por3 = 0
                    Cabecera.imp3 = 0
                    Cabecera.base4 = 0
                    Cabecera.por4 = 0
                    Cabecera.imp4 = 0
        ''''''''''
                    
                    '''''''''''''''''
                    If IsNull(Tabla) Or Tabla = 0 Then
                        
                        'If nLR = 0 Then
                            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "paso 13"
                            nNumLin = MyRecAux.Fields("lineas")
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            ReDim Cabecera.Lines(nNumLin)
                            'MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                            'On Error Resume Next
                            ''CadenadeLog "paso 14"
                            'dLImporte = MyRecAux.Fields("importe")
                            ''CadenadeLog "paso 15"
                            'If Err.Number <> 0 Then
                            '    On Error GoTo 0
                            '    On Error GoTo FIN
                            '    dLImporte = 0
                            'End If
                            'MyRecAux.Close
                            'Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    'CadenadeLog "paso 16"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '    ReDim Cabecera.Lines(nNumLin)
                        '    'MyRecAux.Open "select sum(amount) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'On Error Resume Next
                        '    'dLImporte = MyRecAux.Fields("importe")
                        '    ''CadenadeLog "paso 17"
                        '    'If Err.Number <> 0 Then
                        '    '    On Error GoTo 0
                        '    '    On Error GoTo FIN
                        '    '    dLImporte = 0
                        '    'End If
                        '    'MyRecAux.Close
                        '    'Set MyRecAux = Nothing
                        '
                        'End If
                    Else
                        nNumLin = 1
                        'CadenadeLog "paso 18"
                        ReDim Cabecera.Lines(nNumLin)
                        'dLImporte = MyRecSetTmp.Fields("amount")
                        'CadenadeLog "paso 19"
                    End If
                    
                    'recoger datos
                    Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
                    Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
                    Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
                    'CadenadeLog "paso 20"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        If IsNull(MyRecSetTmp.Fields("invoice")) Or Len(MyRecSetTmp.Fields("invoice")) = 0 Then
                            Cabecera.factura = Space(15)
                        Else
                            Cabecera.factura = Trim(MyRecSetTmp.Fields("invoice"))
                            'CadenadeLog "paso 21"
                        End If
                        
                        Cabecera.EAN13 = Trim(MyRecSetTmp.Fields("ean13"))
                        
                        'CadenadeLog "paso 22"
                        If MyRecSetTmp.Fields("customer") <> "" Then
                            If IsNull(MyRecSetTmp.Fields("customer")) Then
                                Cabecera.NCliente = 0
                            Else
                                On Error Resume Next
                                Cabecera.NCliente = MyRecSetTmp.Fields("customer")
                                If Err.Number <> 0 Then
                                    On Error GoTo 0
                                    Cabecera.NCliente = 0
                                End If
                                On Error GoTo FIN
                            End If
                        Else
                            Cabecera.NCliente = 0
                        End If
                        'CadenadeLog "paso 23"
                    Else
                        Cabecera.factura = Space(15)
                        Cabecera.NCliente = 0
                    End If
                    'CadenadeLog "paso 24"
                    
                    Cabecera.Fecha = Format(Ffecha, "ddmmyy")
                    Cabecera.hora = Fhora
                    Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
                    If Tabla <> 0 Then
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("amount")
                    Else
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("netamount")
                    End If
                    
                    
                    'CadenadeLog "paso 25"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.nLineas = nNumLin
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = nNumLin
                        Cabecera.LineasMensaje = nNumLin
                    Else
                        Cabecera.nLineas = 1
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = 1
                        Cabecera.LineasMensaje = 1
                    End If
                    'cabecera.codigo ???
                    
                    Cabecera.EstadoTiquet = 16
                    'CadenadeLog "paso 26"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If MyRecSetTmp.Fields("NetStat") = "L" Then
                            Cabecera.CadenaEstadoTiquet = "Local"
                            Cabecera.tipoTiquet = 4
                        Else
                            Cabecera.CadenaEstadoTiquet = "Red"
                            Cabecera.tipoTiquet = 1
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'CadenadeLog "paso 27"
                        'If nLR = 0 Then
                        '    'CadenadeLog "paso 28"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 29"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    'CadenadeLog "paso 30"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 31"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                        'CadenadeLog "paso 32"
                        Cabecera.Trainning = False
                        If MyRecSetTmp.Fields("tcktype") = "U" Then
                            Cabecera.tipoTiquet = 12
                            Cabecera.Trainning = True
                        End If
                        If MyRecSetTmp.Fields("status") = "R" Then
                            Cabecera.tipoTiquet = 13
                            Cabecera.EstadoTiquet = 64

                        End If
                        'CadenadeLog "paso 33"
                        If MyRecSetTmp.Fields("status") = "X" Then
                            Cabecera.tipoTiquet = 14
                            Cabecera.EstadoTiquet = 64
                            
                        End If
                        'CadenadeLog "paso 34"
                        'If Tabla <> 0 Then
                        '    Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                        'Else
                            Cabecera.ImporteBruto = MyRecSetTmp.Fields("grossamount")
                        'End If
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Round(Cabecera.ImporteBruto - Cabecera.ImporteTotal, 2) 'más abajo se hace como porcentaje...
                        Cabecera.documentoabono = 0
                    
                    Else
                        'CadenadeLog "paso 35"
                        If nLR = 0 Then
                            Cabecera.CadenaEstadoTiquet = "Red"
                        Else
                            Cabecera.CadenaEstadoTiquet = "Local"
                        End If
                        
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If Tabla = 1 Then
                            Cabecera.tipoTiquet = 2
                        Else
                            Cabecera.tipoTiquet = 3
                        End If
                        
                        'Cabecera.CadenaEstadoTiquet = "Red"
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        'Cabecera.ImporteBruto = dLImporte
                        Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal
                        Cabecera.documentoabono = 0
                        'Cabecera.Cambio = 0
                        'Cabecera.descuento = 0
                        'Cabecera.LineasCanceladas = 0
                    End If
                    'CadenadeLog "paso 36"
                    If nLR = 0 Then
                        If Tabla = 0 Or IsNull(Tabla) Then
                            ''frmmonitor.MostrarDato CargaCadena(868) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            'frmmonitor.MostrarDato CargaCadena(1412) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Etiqueta Envasado:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            'frmmonitor.MostrarDato CargaCadena(1413) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Etiqueta Autoservicio:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    Else
                        If Tabla = 0 Or IsNull(Tabla) Then
                            'frmmonitor.MostrarDato CargaCadena(1414) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Ticket Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            'frmmonitor.MostrarDato CargaCadena(1415) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Etiqueta Envasado Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            'frmmonitor.MostrarDato CargaCadena(1416) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "Etiqueta Autoservicio Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    
                    End If
                    'CadenadeLog "paso 37"
                    If Tabla = 0 Or IsNull(Tabla) Then
                        'If nLR = 0 Then
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '
                        'End If
                        'CadenadeLog "paso 38"
                        'CadenadeLog "7"
                        
                        Cabecera.LineasCanceladas = MyRecSetTmp.Fields("canceledlines")
                        
                        'también se vé afectado si es cancelado o no pero de mamento cerrado...
                        'Cabecera.EstadoTiquet = 16
                        '********************
                        ' este dato se introduce
                        ' tras pasar por actualiza_tabla_cabecera
                        'Incluir_en_Base As Boolean
                        'Cabecera.tipoPago = myrecsettmp.Fields("idmoneda")
                        'Cabecera.cantidadEntre = myrecsettmp.Fields("totalneto") + myrecsettmp.Fields("importecambio")
                        
                        Cabecera.Cambio = Round(Cabecera.ImporteTotal - Cabecera.cantidadEntre, 2) ' myrecsettmp.Fields("importecambio")
                        Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")
                        'CadenadeLog "paso 39"
                        'lote As String
                        'lEncargo As Boolean
            
            
                        'LINEAS
            
                        If MyRecAux.State <> 0 Then
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                        End If
                        nNumLin = 0
                        'If nLR = 0 Then
                            MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc) & " and linetype<>0", MyConnObj, adOpenKeyset, adLockOptimistic
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localltickets where idhticket =" & CStr(IdDoc) & " and linetype<>0", MyConnObj, adOpenKeyset, adLockOptimistic
                        'End If
                        'CadenadeLog "8"
                        'CadenadeLog "paso 40"
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
                                'CadenadeLog "paso 41"
                                nNumLin = nNumLin + 1
                                If MyRecAux.Fields("saleform") = 1 Then
                                    Cabecera.Lines(nNumLin).EsPesado = True
                                    Cabecera.Lines(nNumLin).cantidad = Round(MyRecAux.Fields("weight") * 1000, 0)
                                Else
                                    Cabecera.Lines(nNumLin).EsPesado = False
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                End If
                                'CadenadeLog "paso 42"
                                'CadenadeLog "8.4"
                                
                                Cabecera.Lines(nNumLin).FactorPeso = 1
                                Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                                'CadenadeLog "paso 43"
                                'CadenadeLog "8.5"
                                
                                Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                                Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                                
                                'CadenadeLog "paso 44"
                                'ver qué se hace IVA As Integer
                                Cabecera.Lines(nNumLin).IVA = MyRecAux.Fields("vatcode")
                                Cabecera.Lines(nNumLin).porcentaje = MyRecAux.Fields("vatpercent")
                                'CadenadeLog "paso 45"
                                'CadenadeLog "8.6"
                                
                                Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("ticketdiscount") + MyRecAux.Fields("linediscount") 'porcentaje debido a la cabecera + porcentaje en linea
                                'CadenadeLog "paso 46"
                                Cabecera.Lines(nNumLin).ImporteBruto = MyRecAux.Fields("amount") 'Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2) 'MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                                'CadenadeLog "paso 47"
                                Cabecera.Lines(nNumLin).ImporteLinea = Round(MyRecAux.Fields("amount") - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                'Cabecera.Lines(nNumLin).ImporteLinea = Round(Cabecera.Lines(nNumLin).ImporteLinea - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                
                                'CadenadeLog "paso 48"
                                'CadenadeLog "8.7"
                                
                                'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                                If MyRecAux.Fields("linetype") = 0 Then
                                    Cabecera.Lines(nNumLin).LineaCancelada = True
                                Else
                                    Cabecera.Lines(nNumLin).LineaCancelada = False
                                End If
                                'CadenadeLog "paso 49"
                                'ver qué se hace con TipoLinea As Byte
                                'STV (positivo o negativo)
                                'Linea(Bucle).ModificadoTemporal = False
                                If MyRecAux.Fields("amount") > 0 Then
                                    'Cabecera.Lines(nNumLin).positive = 1
                                     Cabecera.Lines(nNumLin).FactorPeso = 1
                                     Cabecera.Lines(nNumLin).TipoLinea = 0
                                
                                Else
                                    'Cabecera.Lines(nNumLin).positive = 0
                                     Cabecera.Lines(nNumLin).FactorPeso = -1
                                     Cabecera.Lines(nNumLin).TipoLinea = 2
                                    
                                End If
                                'CadenadeLog "paso 50"
            '                    If Cabecera.Lines(nNumLin).positive = 1 Then
            '                         Cabecera.Lines(nNumLin).FactorPeso = 1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 0
            '                    Else
            '                         Cabecera.Lines(nNumLin).FactorPeso = -1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 2
            '                    End If
                                'STP (pesado o no pesado)
                                If Cabecera.Lines(nNumLin).EsPesado Then
                                     'Cabecera.Lines(nNumLin).Weighted = 1
                                Else
                                     'Cabecera.Lines(nNumLin).Weighted = 0
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                                End If
                                
                                'CadenadeLog "paso 51"
                                ' STC (cancelado o no cancelado)
                                If Cabecera.Lines(nNumLin).LineaCancelada Then
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                                End If
                                'CadenadeLog "paso 52"
                                'If MyRecAux.Fields("precio") <> MyRecAux.Fields("preciodefecto") Then
                                '    Cabecera.Lines(nNumLin).SehaModificado = True
                                'Else
                                    Cabecera.Lines(nNumLin).SehaModificado = False
                                'End If
                                '¿? ModificadoTemporal As Boolean
                                'CadenadeLog "paso 53"
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
                                
                                MyRecAux.MoveNext
                            Loop
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            'CadenadeLog "paso 54"
                            
                            nFichIn = FreeFile()
                            sIn = Trim(Cabecera.EAN13)
                            If sIn = "" Or Len(sIn) = 0 Or Len(sIn) <> 13 Then
                                '2YYTTTTIIIIIC.in
                                If Cabecera.ImporteTotal > 0 Then
                                    If Cabecera.NMostrador = 1 Then
                                        sIn = "2" & "00" & Format(Cabecera.Ntiquet, "0000") & Format(Cabecera.ImporteTotal, "00000")
                                    Else
                                        sIn = "2" & Format(Cabecera.NMostrador, "00") & Format(Cabecera.Ntiquet, "0000") & Format(Cabecera.ImporteTotal, "00000")
                                    End If
                                Else
                                    If Cabecera.NMostrador = 1 Then
                                        sIn = "2" & "00" & Format(Cabecera.Ntiquet, "0000") & "-" & Format(Abs(Cabecera.ImporteTotal), "0000")
                                    Else
                                        sIn = "2" & Format(Cabecera.NMostrador, "00") & Format(Cabecera.Ntiquet, "0000") & "-" & Format(Abs(Cabecera.ImporteTotal), "0000")
                                    End If
                                End If
                                sIn = sIn & Format(DigitoControl(sIn), "0") & ".in"
                            Else
                                sIn = sIn & ".in"
                            End If
                            Open "C:\tiquetbal\" & sIn For Output As #nFichIn
                            MyS = sIn
                            If Cabecera.NMostrador = 1 Then
                                sIn = "H " & Format(Cabecera.Ntiquet, "00000000000000000000") & Format(Cabecera.nLineas - Cabecera.LineasCanceladas, "0000") & _
                                       "00" & Cabecera.Fecha & Cabecera.hora & Format(Cabecera.NVendedor, "00000")
                            Else
                                sIn = "H " & Format(Cabecera.Ntiquet, "00000000000000000000") & Format(Cabecera.nLineas - Cabecera.LineasCanceladas, "0000") & _
                                       Format(Cabecera.NMostrador, "00") & Cabecera.Fecha & Cabecera.hora & Format(Cabecera.NVendedor, "00000")
                            End If
                            Print #nFichIn, sIn
                            For nC = 1 To nNumLin
                                sIn = "I  " & Format(Cabecera.Lines(nC).CodigoArticulo, "0000000000000000") & " " & Format(Cabecera.Lines(nC).cantidad, "00000") & " "
                                If Cabecera.Lines(nC).EsPesado = True Then
                                    sIn = sIn & "0"
                                Else
                                    sIn = sIn & "1"
                                End If
                                'If Cabecera.Lines(nC).ImporteLinea > 0 Then
                                If Cabecera.Lines(nC).precio > 0 Then
                                    'sIn = sIn & Format(Cabecera.Lines(nC).ImporteLinea * 100, "0000000") & "0000000000000"
                                    sIn = sIn & Format(Cabecera.Lines(nC).precio * 100, "0000000") & "0000000000000"
                                Else
                                    'sIn = sIn & "-" & Format(Abs(Cabecera.Lines(nC).ImporteLinea) * 100, "000000") & "0000000000000"
                                    sIn = sIn & "-" & Format(Abs(Cabecera.Lines(nC).precio) * 100, "000000") & "0000000000000"
                                End If
                                Print #nFichIn, sIn
                            Next nC
                            
                            
                            
                            Close #nFichIn
                            CadenadeLog "Depositado fichero Tique:" & MyS
                            
                            'Actualiza_Tabla_Cabecera Mybase, Cabecera
                            ''CadenadeLog "paso 55"
                            Do_Events
                            'If Cabecera.Incluir_en_Base = True Then
                            '    Actualiza_Tabla_Tickets Mybase, Cabecera
                            '    Do_Events
                            '    Actualiza_Tabla_Gtarti Mybase, Cabecera
                            '    Do_Events
                            '    Actualiza_Tabla_GtVend Mybase, Cabecera
                            '    Do_Events
                            '    actualiza_tabla_gtHora Mybase, Cabecera
                            '    Do_Events
                            '    Actualiza_Tabla_Gtsecs Mybase, Cabecera
                            '    Do_Events
                            '    actualiza_tabla_gtInfVen Mybase, Cabecera
                            '    Do_Events
                            '    Actualiza_Tabla_Log Mybase, Cabecera
                            'End If
                            'CadenadeLog "paso 56"
                        End If
            
            
            
                        '''''''''''''''
                        'marcar...
                        If lMarcaTime Then
                            'cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
                            cm = "update pcscale.htickets set CapturedDateTime='" & sCap & "' where id=" & CStr(IdDoc)
                        Else
                            'cm = "update pcscale.localhtickets set captured='*' where id=" & CStr(IdDoc)
                            cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
                        End If
                        MyConnObj.Execute cm
                        ''''''''''
                    
                    Else 'REVISAR...
                    
                    
                            'CadenadeLog "paso 58"
                            nNumLin = 1
                            If MyRecSetTmp.Fields("saleform") = 1 Then
                                Cabecera.Lines(nNumLin).EsPesado = True
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            Else
                                Cabecera.Lines(nNumLin).EsPesado = False
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            End If
                            'CadenadeLog "paso 59"
                            'CadenadeLog "8.4"
                            
                            Cabecera.Lines(nNumLin).FactorPeso = 1
                            Cabecera.Lines(nNumLin).precio = MyRecSetTmp.Fields("price")
                            
                            'CadenadeLog "8.5"
                            'CadenadeLog "paso 60"
                            Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                            Cabecera.Lines(nNumLin).CodigoArticulo = MyRecSetTmp.Fields("item")
                            'ver qué se hace IVA As Integer
                            Cabecera.Lines(nNumLin).IVA = 0
        
                            'CadenadeLog "paso 61"
                            'CadenadeLog "8.6"
                            
                            Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("linediscount") 'porcentaje
                            Cabecera.Lines(nNumLin).ImporteLinea = MyRecSetTmp.Fields("amount") 'MyRecSetTmp.Fields("importeconimpuestos") - MyRecSetTmp.Fields("dtoscabeceraconimpuestos")
                            Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecSetTmp.Fields("weight") * MyRecSetTmp.Fields("price"), 2) 'MyRecSetTmp.Fields("importe") - MyRecSetTmp.Fields("dtoscabecera")
                            'CadenadeLog "paso 62"
                            'CadenadeLog "8.7"
                            
                            'If MyRecSetTmp.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                            If MyRecSetTmp.Fields("linetype") = 0 Then
                                Cabecera.Lines(nNumLin).LineaCancelada = True
                            Else
                                Cabecera.Lines(nNumLin).LineaCancelada = False
                            End If
                            'CadenadeLog "paso 63"
                            'ver qué se hace con TipoLinea As Byte
                            'STV (positivo o negativo)
                            'Linea(Bucle).ModificadoTemporal = False
                            If MyRecSetTmp.Fields("amount") > 0 Then
                                'Cabecera.Lines(nNumLin).positive = 1
                                 Cabecera.Lines(nNumLin).FactorPeso = 1
                                 Cabecera.Lines(nNumLin).TipoLinea = 0
                            
                            Else
                                'Cabecera.Lines(nNumLin).positive = 0
                                 Cabecera.Lines(nNumLin).FactorPeso = -1
                                 Cabecera.Lines(nNumLin).TipoLinea = 2
                                
                            End If
                            'CadenadeLog "paso 64"
        '                    If Cabecera.Lines(nNumLin).positive = 1 Then
        '                         Cabecera.Lines(nNumLin).FactorPeso = 1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 0
        '                    Else
        '                         Cabecera.Lines(nNumLin).FactorPeso = -1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 2
        '                    End If
                            'STP (pesado o no pesado)
                            If Cabecera.Lines(nNumLin).EsPesado Then
                                 'Cabecera.Lines(nNumLin).Weighted = 1
                            Else
                                 'Cabecera.Lines(nNumLin).Weighted = 0
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                            End If
                            'CadenadeLog "paso 65"
                            
                            ' STC (cancelado o no cancelado)
                            If Cabecera.Lines(nNumLin).LineaCancelada Then
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                            End If
                            'CadenadeLog "paso 66"
                            'If MyRecSetTmp.Fields("precio") <> MyRecSetTmp.Fields("preciodefecto") Then
                            '    Cabecera.Lines(nNumLin).SehaModificado = True
                            'Else
                                Cabecera.Lines(nNumLin).SehaModificado = False
                            'End If
                            '¿? ModificadoTemporal As Boolean
                            'CadenadeLog "paso 67"
                            'CadenadeLog "8.8"
                            'c2f 1.1.1 aclarar ficha Vacuno...
                            'If IsNull(MyRecSetTmp.Fields("eartag")) Or MyRecSetTmp.Fields("eartag") = "" Then
                                Cabecera.Lines(nNumLin).FichaVacuno = 0
                            'Else
                            '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecSetTmp.Fields("eartag")
                            'End If
                            'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                            'Cabecera.descuento = Cabecera.descuento + MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'CadenadeLog "8.9"
                    
                        'CadenadeLog "paso 68"
                        'Actualiza_Tabla_Cabecera Mybase, Cabecera
                        'Do_Events
                        'If Cabecera.Incluir_en_Base = True Then
                        '    Actualiza_Tabla_Tickets Mybase, Cabecera
                        '    Do_Events
                        '    Actualiza_Tabla_Gtarti Mybase, Cabecera
                        '    Do_Events
                        '    Actualiza_Tabla_GtVend Mybase, Cabecera
                        '    Do_Events
                        '    actualiza_tabla_gtHora Mybase, Cabecera
                        '    Do_Events
                        '    Actualiza_Tabla_Gtsecs Mybase, Cabecera
                        '    Do_Events
                        '    actualiza_tabla_gtInfVen Mybase, Cabecera
                        '    Do_Events
                        '    Actualiza_Tabla_Log Mybase, Cabecera
                        'End If
                        

                        'marcar...
                        'If nLR = 0 Then
                        '    If Tabla = 1 Then
                        '        cm = "update pcscale.packetlabels set customer='9999999999' where id=" & CStr(IdDoc)
                        '    Else
                        '        cm = "update pcscale.sservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                        '    End If
                        'Else
                        '    If Tabla = 1 Then
                        '        cm = "update pcscale.localpacketlabels set customer='9999999999' where id=" & CStr(IdDoc)
                        '    Else
                        '        cm = "update pcscale.localsservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                        '    End If
                        'End If
                        'MyConnObj.Execute cm
                        ''''''''''
                    
                    End If
                    
                    MyRecSetTmp.Update
        
                    MyRecSetTmp.MoveNext
                
                    
                    If cancelar = 1 Then GoTo SalirLeeTique
                    Do_Events
                    
                Loop
                End If
            
                If MyRecSetTmp.State <> 0 Then
                    MyRecSetTmp.Close
                    Set MyRecSetTmp = Nothing
                End If
            
SigueNlr:

                Next nLR
            
            End If
            'CadenadeLog "paso 73"
    '        MisEqp.MoveNext
            
    '    Loop
        
    '    MisEqp.Close
    '    Set MisEqp = Nothing
        
    'End If
    
SalirLeeTique:
''''''''
    On Error GoTo 0
    On Error GoTo FIN
    If MyRecSetTmp.State <> 0 Then
        MyRecSetTmp.Close
        Set MyRecSetTmp = Nothing
    End If
FIN:
    If Err.Number <> 0 Then
        CadenadeLog "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
        On Error GoTo 0
        On Error GoTo FIN1
        If MyConnObj.State <> 0 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            Do_Events
            Sleep (200)
            EstadoBal = 1000
        End If
    End If
FIN1:
    On Error GoTo 0
    Do_Events
End Sub

Public Function Erase_Totals(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para TouchScale  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
   
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0

    CadenadeLog "Capa1A_GA_Erase_totals"
'*******
    
    
    CadenadeLog "Iniciar Numeración si Procede..."
    If No_Reinit = 0 Then 'Iniciar numeración tiques...
        On Error Resume Next
        cm = "delete from pcscale.localnumtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "delete from pcscale.numtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "delete from pcscale.erasetickets"
        MyConnObj.Execute cm
        On Error GoTo 0
    End If
    CadenadeLog "Borrado Documentos y Lineas Documento..."
    On Error Resume Next
    cm = "delete from pcscale.HTickets where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.sservicelabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.packetlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    'Añadididas 1.0.4
    cm = "delete from pcscale.localerasetickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localhtickets where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketpresets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpaytickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localsservicelabels"
    MyConnObj.Execute cm
    ''''''''''''''''''''''''''''''''
    On Error GoTo 0
    On Error Resume Next
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Erase_Totals = Resp
End Function

Public Function Item_OrderInFamily_T(ByVal nFam As Long, ByVal nMethod As Long) As Long
'nFam == número de familia
'nMethod --> 0=Código Artículo, 1=Plu Artículo, 2=Name, 3=shortname
'SELECT code,family,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=3 order by shortname
Dim Resp As Long
Dim cm As String
Dim MyRecDes As New ADODB.Recordset
Dim n As Long
    CadenadeLog "Ordenación teclas en familia:" & CStr(nFam)
    Resp = 9999
    On Error GoTo Sal_T
    Select Case nMethod
    Case 0
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 1
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by plu", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 2
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by name", MyConnObj, adOpenKeyset, adLockOptimistic
    Case 3
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by shortname", MyConnObj, adOpenKeyset, adLockOptimistic
    Case Else
    MyRecDes.Open "SELECT code,family,name,shortname,familyorder,@rownum := @rownum + 1 AS rank FROM items, (SELECT @rownum := 0) r where family=" & CStr(nFam) & " order by code", MyConnObj, adOpenKeyset, adLockOptimistic
    End Select
    Resp = 10000
    n = 0
    If Not MyRecDes.EOF Then
        MyRecDes.MoveFirst
        Do While Not MyRecDes.EOF
            Resp = 10001
            'cm = "update items set familyorder=" & CStr(MyRecDes.Fields("rank")) & " where code=" & CStr(MyRecDes.Fields("code")) & " and family=" & CStr(nFam)
            'CadenadeLog cm
            MyConnObj.Execute cm
            n = n + 1
            MyRecDes.MoveNext
        Loop
    End If
    Resp = 10002
    MyRecDes.Close
    Set MyRecDes = Nothing
Sal_T:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'Resp = 10000
    Else
        Resp = n
    End If
    Item_OrderInFamily_T = Resp
End Function

Public Function Capa1B_SendFullItem_TouchScale(Optional ByVal lSoloG As Boolean) As Long
    Dim MyRecAux As New ADODB.Recordset
    Dim cm As String
    Dim nF As Integer
    Dim lEsAlta As Boolean
    Dim nC As Long
    Dim nCont As Integer
    Dim Resp As Long
    Dim nPrice As Double
    Dim sIngredientes As String
    Dim lCodFam As Boolean
    Dim s(5) As String
    Dim sEtapa As String
    Dim nE As Integer
    Dim Item As Epel_FullItem
    Dim nContAlta201 As Integer
    Dim nContAlta202 As Integer
    Dim nContAlta203 As Integer
    Dim nContAlta204 As Integer
    Dim nContAlta205 As Integer
    Dim nContAlta206 As Integer
    Dim nContAlta207 As Integer
    Dim nContAlta208 As Integer
    Dim nContAlta209 As Integer
    Dim nContAlta7002 As Integer
    
    Switch_mode 11
    If EstadoBal <> 11 Then
        'CadenadeLog "no hay conexión con Servidora..."
        GoTo FIN_NO
    End If
    MyRecAux.Open "select * from pcscale.scales where (type='S' or type='A') and ip='" & Mid(sAddIP, 1, Len(sAddIP) - 5) & "'", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyRecAux.Close
        Set MyRecAux = Nothing
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    
    'counters
    MyRecAux.Open "select * from pcscale.counters where code=1", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyConnObj.Execute "replace into counters (code,name,description) values (1,'Panadería','Panadería')"
        If Err.Number <> 0 Then
            GoTo FIN_Capa1B_SendFullItem
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecAux.Open "select * from pcscale.counters where code=2", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyConnObj.Execute "replace into counters (code,name,description) values (2,'Frutería','Frutería')"
        If Err.Number <> 0 Then
            GoTo FIN_Capa1B_SendFullItem
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecAux.Open "select * from pcscale.counters where code=4", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyConnObj.Execute "replace into counters (code,name,description) values (4,'Charcutería','Charcutería')"
        If Err.Number <> 0 Then
            GoTo FIN_Capa1B_SendFullItem
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecAux.Open "select * from pcscale.counters where code=6", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyConnObj.Execute "replace into counters (code,name,description) values (6,'Carnicería','Carnicería')"
        If Err.Number <> 0 Then
            GoTo FIN_Capa1B_SendFullItem
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
    MyRecAux.Open "select * from pcscale.counters where code=8", MyConnObj, adOpenKeyset, adLockOptimistic
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    If MyRecAux.EOF Then
        MyConnObj.Execute "replace into counters (code,name,description) values (8,'Pescadería','Pescadería')"
        If Err.Number <> 0 Then
            GoTo FIN_Capa1B_SendFullItem
        End If
    End If
    MyRecAux.Close
    Set MyRecAux = Nothing
        
    'families
    MyConnObj.Execute "replace into families (code,name,label) values (201,'CARNES FRESCAS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (202,'PLATOS COCINADOS Y PRECOCINADOS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (203,'CHARCUTERIA','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (204,'CONGELADOS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (205,'DERIVADOS LACTEOS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (206,'FRUTAS Y HORTALIZAS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (207,'PANADERIA Y PASTELERIA','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (208,'PESCADOS MARISCOS Y CEFALOPODOS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (209,'QUESOS','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into families (code,name,label) values (7002,'CONSUMIBLES','-')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    'code128, ean13
    '2YYTTTTIIIII
    MyConnObj.Execute "replace into barcodeean13 (numbertype,number,networksales,networktotalslb,localsales,localtotalslb,localtotal,networklineslb,networkss,networkpacket1,networkpacket2,locallineslb,localss,localpacket1,localpacket2) values ('C',1,'200TTTTIIIII','200TTTTIIIII','200TTTTIIIII','200TTTTIIIII','200TTTTIIIII','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii','200CCCCiiiii')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into barcodeean13 (numbertype,number,networksales,networktotalslb,localsales,localtotalslb,localtotal,networklineslb,networkss,networkpacket1,networkpacket2,locallineslb,localss,localpacket1,localpacket2) values ('C',2,'2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into barcodeean13 (numbertype,number,networksales,networktotalslb,localsales,localtotalslb,localtotal,networklineslb,networkss,networkpacket1,networkpacket2,locallineslb,localss,localpacket1,localpacket2) values ('C',4,'2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into barcodeean13 (numbertype,number,networksales,networktotalslb,localsales,localtotalslb,localtotal,networklineslb,networkss,networkpacket1,networkpacket2,locallineslb,localss,localpacket1,localpacket2) values ('C',6,'2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into barcodeean13 (numbertype,number,networksales,networktotalslb,localsales,localtotalslb,localtotal,networklineslb,networkss,networkpacket1,networkpacket2,locallineslb,localss,localpacket1,localpacket2) values ('C',8,'2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYTTTTIIIII','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii','2YYCCCCiiiii')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into code128 (counter,code1,code2,code3) values (1,'250TTTTTTPPPPPiiiiiWWWWW00','232TTTTTTPPPPPiiiiiWWWWW00','257TTTTTTPPPPPiiiiiWWWWW00')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into code128 (counter,code1,code2,code3) values (2,'250TTTTTTPPPPPiiiiiWWWWWYY','232TTTTTTPPPPPiiiiiWWWWWYY','257TTTTTTPPPPPiiiiiWWWWWYY')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into code128 (counter,code1,code2,code3) values (4,'250TTTTTTPPPPPiiiiiWWWWWYY','232TTTTTTPPPPPiiiiiWWWWWYY','257TTTTTTPPPPPiiiiiWWWWWYY')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into code128 (counter,code1,code2,code3) values (6,'250TTTTTTPPPPPiiiiiWWWWWYY','232TTTTTTPPPPPiiiiiWWWWWYY','257TTTTTTPPPPPiiiiiWWWWWYY')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    MyConnObj.Execute "replace into code128 (counter,code1,code2,code3) values (8,'250TTTTTTPPPPPiiiiiWWWWWYY','232TTTTTTPPPPPiiiiiWWWWWYY','257TTTTTTPPPPPiiiiiWWWWWYY')"
    If Err.Number <> 0 Then
        GoTo FIN_Capa1B_SendFullItem
    End If
    
    If lSoloG Then GoTo FIN_NO
    nContAlta201 = 0
    nContAlta202 = 0
    nContAlta203 = 0
    nContAlta204 = 0
    nContAlta205 = 0
    nContAlta206 = 0
    nContAlta207 = 0
    nContAlta208 = 0
    nContAlta209 = 0
    nContAlta7002 = 0
    
    On Error GoTo FIN_Capa1B_SendFullItem
    For nE = 0 To nEnvios - 1
        Item.sOper = myRec(nE).MyS(0)
        Item.Code = myRec(nE).MyS(1)
        If Item.sOper <> "B" And Item.Code < 9999 Then
            Item.FamilyOrder = myRec(nE).MyS(1)
            Item.Family = myRec(nE).MyS(2)
            Item.Price = myRec(nE).MyS(3)
            Item.sNameShort = myRec(nE).MyS(4)
            Item.sText1 = myRec(nE).MyS(5)
            Item.sText2 = myRec(nE).MyS(6)
            Item.weight = myRec(nE).MyS(7)
            If IsNull(myRec(nE).MyS(8)) Or Trim(myRec(nE).MyS(8)) = "" Then
                Item.etq = 0
            Else
                Item.etq = Val(myRec(nE).MyS(8))
            End If
            Item.caducity = myRec(nE).MyS(9)
            Item.sLongText = myRec(nE).MyS(10)
            Item.Plu = 0
            Item.Sec = 0
            Item.vat = 0
            Item.pref = 0
            Item.tare = 0
            Item.nPeso = 0
            Item.lMix = False
            Item.EAN13 = ""
            'If lNoLog = False Then
                sEtapa = "Presentación datos..."
                CadenadeLog "----------------------------------------------------------------------->"
                
                CadenadeLog "Código:" & CStr(Item.Code)
                CadenadeLog "Plu:" & CStr(Item.Plu)
                CadenadeLog "Sec.:" & CStr(Item.Sec)
                CadenadeLog "Sub. Sec.:" & CStr(Item.SubSec)
                CadenadeLog "Familia:" & CStr(Item.Family)
                CadenadeLog "Precio:" & CStr(Item.Price)
                CadenadeLog "W:" & CStr(Item.weight)
                CadenadeLog "Caducidad:" & CStr(Item.caducity)
                CadenadeLog "Descriptivo:" & Item.sNameShort
                CadenadeLog "Tara:" & CStr(Item.tare)
                CadenadeLog "Preferente:" & CStr(Item.pref)
                CadenadeLog "Etiqueta:" & CStr(Item.etq)
                CadenadeLog "Iva:" & CStr(Item.vat)
                CadenadeLog "EAN:" & Item.EAN13
                CadenadeLog "Mixto:" & CStr(Item.lMix)
                CadenadeLog "Peso:" & CStr(Item.nPeso)
                CadenadeLog "FamilyOrder:" & CStr(Item.FamilyOrder)
                CadenadeLog "Operación:" & Item.sOper
                CadenadeLog "<-----------------------------------------------------------------------"
            'End If
            
            If IsNull(Item.nPeso) Then
                Item.nPeso = 0
            End If
            If (Item.nPeso < 0) Or (Item.nPeso > 15000) Then
                Item.nPeso = 0
            End If
            
            lEsAlta = False
            lCodFam = False
            
            On Error Resume Next

            'sEtapa = "select * from pcscale.items where code=" & Item.Code
            
            MyRecAux.Open "select * from pcscale.items where code=" & Item.Code, MyConnObj, adOpenKeyset, adLockOptimistic
            If Err.Number <> 0 Then
                GoTo FIN_Capa1B_SendFullItem
            End If
            On Error GoTo FIN_Capa1B_SendFullItem
            lEsAlta = False
            If MyRecAux.EOF Then
                sEtapa = "No existe se añade registro..."
                MyRecAux.AddNew
                lEsAlta = True
                If Item.Family = 201 Then nContAlta201 = nContAlta201 + 1
                If Item.Family = 202 Then nContAlta202 = nContAlta202 + 1
                If Item.Family = 203 Then nContAlta203 = nContAlta203 + 1
                If Item.Family = 204 Then nContAlta204 = nContAlta204 + 1
                If Item.Family = 205 Then nContAlta205 = nContAlta205 + 1
                If Item.Family = 206 Then nContAlta206 = nContAlta206 + 1
                If Item.Family = 207 Then nContAlta207 = nContAlta207 + 1
                If Item.Family = 208 Then nContAlta208 = nContAlta208 + 1
                If Item.Family = 209 Then nContAlta209 = nContAlta209 + 1
                If Item.Family = 7002 Then nContAlta7002 = nContAlta7002 + 1
                
            End If
            sEtapa = "code, Text..."
            MyRecAux.Fields("code") = Item.Code
            MyRecAux.Fields("shortname") = Item.sNameShort
            MyRecAux.Fields("name") = Item.sNameShort

            'cm = ""
            'For nCont = 0 To 19
            '    cm = cm & Trim(Item.Desc2040(nCont))
            'Next nCont
            'If Trim(cm) <> "" Then
            '    MyRecAux.Fields("text") = Mid(cm, 1, 1024)
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text") = ""
            'End If
            'cm = ""
            'For nCont = 0 To 19
            '    cm = cm & Trim(Item.Desc(nCont))
            'Next nCont
            's(1) = Mid(cm, 1, 100)
            's(2) = Mid(cm, 101, 100)
            's(3) = Mid(cm, 201, 100)
            's(4) = Mid(cm, 301, 100)
            's(5) = Mid(cm, 401)
            'If Trim(s(1)) <> "" Then
            '    MyRecAux.Fields("text1") = Trim(s(1))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text1") = ""
            'End If
            'If Trim(s(2)) <> "" Then
            '    MyRecAux.Fields("text2") = Trim(s(2))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text2") = ""
            'End If
            'If Trim(s(3)) <> "" Then
            '    MyRecAux.Fields("text3") = Trim(s(3))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text3") = ""
            'End If
            'If Trim(s(4)) <> "" Then
            '    MyRecAux.Fields("text4") = Trim(s(4))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text4") = ""
            'End If
            'If Trim(s(5)) <> "" Then
            '    MyRecAux.Fields("text5") = Trim(s(5))
            'Else
            '    If lEsAlta Then MyRecAux.Fields("text5") = ""
            'End If
            
            MyRecAux.Fields("text1") = Item.sText1
            MyRecAux.Fields("text2") = Item.sText2
            MyRecAux.Fields("text") = Item.sLongText
            MyRecAux.Fields("family") = Item.Family
            MyRecAux.Fields("price") = Round(Item.Price / 100, 2) 'Item.Price
            If Item.weight = 1 Then
                MyRecAux.Fields("saleform") = 1
            Else
                MyRecAux.Fields("saleform") = 0
            End If
            MyRecAux.Fields("plu") = Item.Plu
            MyRecAux.Fields("counter") = Item.Sec
            MyRecAux.Fields("vat") = Item.vat
            If lEsAlta Then
                MyRecAux.Fields("discount") = 0
                MyRecAux.Fields("qlevel") = ""
            End If
            ''''''''''''
            If Trim(Item.EAN13) <> "" Then
                MyRecAux.Fields("ean13") = Trim(Item.EAN13)
            Else
                MyRecAux.Fields("ean13") = ""
            End If
            If Trim(Item.EAN13) <> "" Then
                MyRecAux.Fields("ean13print") = Mid(Trim(Item.EAN13), 1, 12)
            Else
                MyRecAux.Fields("ean13print") = ""
            End If
            
            MyRecAux.Fields("expiring") = Item.caducity
            MyRecAux.Fields("bestbefore") = Item.pref
            MyRecAux.Fields("tare") = Item.tare / 1000
            'If lEsAlta Then
            If Item.etq <> 0 Then
                If Item.etq = 16 Then
                    MyRecAux.Fields("label") = "-"
                Else
                    MyRecAux.Fields("label") = CStr(Item.etq) '& ".xml"
                End If
            Else
                MyRecAux.Fields("label") = ""
            End If
            'If lEsAlta Then MyRecAux.Fields("icon") = ""
            'If Trim(miartXtra.Imagen) <> "" Then
            '    If InStr(1, miartXtra.Imagen, ".") <> 0 Then
            '        MyRecAux.Fields("icon") = Trim(miartXtra.Imagen)
            '    End If
            ''Else
            ''    MyRecAux.Fields("icon") = ""
            'End If
            'If Trim(Item.DescTraza(40)) <> "" Then
            '        MyRecAux.Fields("icon") = Trim(Item.DescTraza(40))
            'End If
            MyRecAux.Fields("icon") = Format(Item.Code, "0000") & ".png"
            If Item.lMix Then
                MyRecAux.Fields("mixed") = 1
            Else
                MyRecAux.Fields("mixed") = 0
            End If
            MyRecAux.Fields("mixedweight") = Item.nPeso / 1000
            
            'If (lVer51) Then
            '    If Item.FamilyOrder > 0 Then
                    
            If lEsAlta Then
                    MyRecAux.Fields("familyorder") = 0 'Item.Code
                    MyRecAux.Fields("onkey") = 1
            End If
            '    End If
            'End If
            
            MyRecAux.Update
            
            MyRecAux.Close
            Set MyRecAux = Nothing
            
        Else
            If Item.Code > 9999 Then
                CadenadeLog "Registro rechazado:" & CStr(Item.Code)
            Else
                If Item.sOper = "B" Then
                    MyRecAux.Open "delete from pcscale.items where code=" & Item.Code, MyConnObj, adOpenKeyset, adLockOptimistic
                End If
            End If
        End If
    Next nE
    
    If nContAlta201 > 0 Then Resp = Item_OrderInFamily_T(201, 0)
    If nContAlta202 > 0 Then Resp = Item_OrderInFamily_T(202, 0)
    If nContAlta203 > 0 Then Resp = Item_OrderInFamily_T(203, 0)
    If nContAlta204 > 0 Then Resp = Item_OrderInFamily_T(204, 0)
    If nContAlta205 > 0 Then Resp = Item_OrderInFamily_T(205, 0)
    If nContAlta206 > 0 Then Resp = Item_OrderInFamily_T(206, 0)
    If nContAlta207 > 0 Then Resp = Item_OrderInFamily_T(207, 0)
    If nContAlta208 > 0 Then Resp = Item_OrderInFamily_T(208, 0)
    If nContAlta209 > 0 Then Resp = Item_OrderInFamily_T(209, 0)
    If nContAlta7002 > 0 Then Resp = Item_OrderInFamily_T(7002, 0)
    
    CadenadeLog "Fin de envío Servidora:" & sAddIP
    
FIN_NO:
    Resp = 0

FIN_Capa1B_SendFullItem:
    If Err.Number <> 0 Then
        CadenadeLog "Error en SendFullItem_touchScale. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    'CadenadeLog "Articulo Cod: " & Item.Code & " Plu: " & Item.Plu & " Sec: " & Item.Sec & " Desc: " & Item.sNameShort
    Capa1B_SendFullItem_TouchScale = Resp
End Function



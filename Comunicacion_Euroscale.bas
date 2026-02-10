Attribute VB_Name = "Comunicacion_Euroscale"
Option Explicit
'*********************************************************************
' Comunicación de datos de la base a Sistemas Euroscale.
' Este módulo contempla el trabajo para subredes (grupos de balanzas
' con NetIDs diferentes). En general, para comunicar, se trabaja siempre del mismo
' modo :
' 1º) Se consultan los sistemas existentes con Mue_Sistemas, la cual rellena
'     la estructura "SisEur" con los datos de los sistemas existentes (si no
'     hay configuración multieuroscale, siempre devuelve un solo sistema)
' 2º) Se realiza un ciclo por los diferentes sistemas, enviando a cada uno
'     el que corresponde
'*********************************************************************
Private Type TipoBarras
    Tipo As String
    numero As Integer
    estado As String
End Type
Public Type Typ_Tot_Sec    ' Totales por sección
    nSec As Integer
    seccion() As Integer
    Importe() As Double
    Tiquets_red() As Long
    tiquets_Local() As Long
End Type
'*************************************
' Comunicación para sistemas
' con múltiples servidoras (NetID)
'*************************************
Public Type St_Euroscale    ' Esta estructura define un sistema
    IP As String            ' IP del sistema
    NSecciones As Integer   ' Nº de Secciones que pertenecen a este sistema
    Secciones() As Integer  ' Números de sección
    Smaq() As Integer       ' Números de sección máquina
End Type
Public Type St_MultiEuroscale   ' esta estructura engloba todos los sitemas
    NSistemas As Integer        ' Nº de sistemas existentes
    Sistemas() As St_Euroscale  ' Datos de cada sistema
End Type
Public SisEur As St_MultiEuroscale  ' Variable global con los datos de los sistemas


Public Function MUE_SecToSys(ByVal nSec As Long) As Long
'********************************
' obtiene el sistema al que
' pertenece un mostrador.
'********************************
    Dim Resp As Long
    Dim B1 As Long, B2 As Long
    Resp = -1
    MUE_Sistemas
    For B1 = 0 To SisEur.NSistemas - 1
        For B2 = 0 To SisEur.Sistemas(B1).NSecciones - 1
            If nSec = SisEur.Sistemas(B1).Secciones(B2) Then
                Resp = B2
                Exit For
            End If
        Next B2
        If Resp <> -1 Then Exit For
    Next B1
    MUE_SecToSys = Resp
End Function

Public Function Check_Dns_File() As String
    Dim Resp As String
    Dim MyPath As String
    Dim Arch As Integer
    Resp = ""
    If Not MultiEuroscale Then
        If Not HayMulti Then
            MyPath = App.Path
        Else
            If TiendaActual < 100 Then
                MyPath = App.Path & "\T" & Format(TiendaActual, "00")
            Else
                MyPath = App.Path & "\T" & Format(TiendaActual, "000")
            End If
        End If
        If Dir(MyPath, vbDirectory) <> "" Then
            If Dir(MyPath & "\dns.txt") <> "" Then
                Arch = FreeFile()
                Open MyPath & "\" & "dns.txt" For Input As #Arch
                If Not EOF(Arch) Then Line Input #Arch, Resp
                Close #Arch
            End If
        End If
    End If
    Resp = Trim(Resp)
    Check_Dns_File = Resp
End Function
'******************************************************
'   Obtiene de la base los sistemas Euroscale
'   existentes así como sus direcciones IP
'   almacenándolas en la estructura SisEur
'******************************************************
Public Sub MUE_Sistemas()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim bucle As Integer
    '''''''''''''

    Set Base = OpenDatabase(Base_General)
    SisEur.NSistemas = 0
    If MultiEuroscale Then
        Set Registro = Base.OpenRecordset( _
                       "select distinct(sec_ip) from seccion where (enviardatos=" & Chr(34) & "GA" & Chr(34) & _
                     " or enviardatos=" & Chr(34) & "TODOS" & Chr(34) & ") and borrado=false")
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    If Not IsNull(.Fields(0)) Then
                        If Trim(.Fields(0)) <> "" Then
                            SisEur.NSistemas = SisEur.NSistemas + 1
                            ReDim Preserve SisEur.Sistemas(SisEur.NSistemas)
                            SisEur.Sistemas(SisEur.NSistemas - 1).IP = .Fields(0)
                            SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones = 0
                            Set Reg2 = Base.OpenRecordset( _
                                       "select codi_ident,secc_maqui from seccion where (enviardatos=" & Chr(34) & "GA" & Chr(34) & _
                                     " or enviardatos=" & Chr(34) & "TODOS" & Chr(34) & ") and borrado=false and sec_ip=" & Chr(34) & Registro.Fields(0) & Chr(34))
                            Reg2.MoveFirst
                            Do Until Reg2.EOF
                                SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones = SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones + 1
                                ReDim Preserve SisEur.Sistemas(SisEur.NSistemas - 1).Secciones(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones)
                                SisEur.Sistemas(SisEur.NSistemas - 1).Secciones(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones - 1) = Reg2.Fields("codi_ident")
                                ReDim Preserve SisEur.Sistemas(SisEur.NSistemas - 1).Smaq(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones)
                                SisEur.Sistemas(SisEur.NSistemas - 1).Smaq(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones - 1) = Reg2.Fields("secc_maqui")
                                Reg2.Movenext
                            Loop
                        End If
                    End If
                    .Movenext
                Loop
            End If
        End With
        Registro.Close
        Set Registro = Nothing
    Else
        If leeIp Or (Not tcpip) Then
            SisEur.NSistemas = 1
            ReDim SisEur.Sistemas(1)
            SisEur.Sistemas(0).IP = sAddIp
            Set Base = OpenDatabase(Base_General)
            Set Reg2 = Base.OpenRecordset("select * from seccion where borrado=false and enviardatos=" & _
                                          Chr(34) & "GA" & Chr(34) & " or enviardatos=" & Chr(34) & "TODOS" & Chr(34))
            Do Until Reg2.EOF
                SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones = SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones + 1
                ReDim Preserve SisEur.Sistemas(SisEur.NSistemas - 1).Secciones(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones)
                SisEur.Sistemas(SisEur.NSistemas - 1).Secciones(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones - 1) = Reg2.Fields("codi_ident")
                ReDim Preserve SisEur.Sistemas(SisEur.NSistemas - 1).Smaq(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones)
                SisEur.Sistemas(SisEur.NSistemas - 1).Smaq(SisEur.Sistemas(SisEur.NSistemas - 1).NSecciones - 1) = Reg2.Fields("secc_maqui")
                Reg2.Movenext
            Loop
            Reg2.Close
            Set Reg2 = Nothing
        End If
    End If
    Do_Events
    Set Base = Nothing
    Do_Events
End Sub

Public Function MUE_CadenaSQL(ByVal NSistema As Integer) As String
'********************************************
' genera un fragmento de cada SQL para
' filtraje de los datos que pertenecen
' a una sección concreta. Poniendo como
' condición el mostrador (codi_ident)
'********************************************
    Dim bucle As Integer
    Dim StrSQL As String

    StrSQL = "("
    If SisEur.Sistemas(NSistema).NSecciones = 0 Then
        StrSQL = StrSQL & "codi_ident=0"
    End If
    For bucle = 0 To SisEur.Sistemas(NSistema).NSecciones - 1
        StrSQL = StrSQL & " codi_ident=" & Format(SisEur.Sistemas(NSistema).Secciones(bucle))
        If (bucle <> SisEur.Sistemas(NSistema).NSecciones - 1) Then
            StrSQL = StrSQL & " or"
        End If
        'If Bucle = 0 And SisEur.Sistemas(NSistema).NSecciones = 2 Then
        '    strSQL = strSQL & " or"
        'End If
    Next bucle
    StrSQL = StrSQL & ")"
    MUE_CadenaSQL = StrSQL
End Function
Public Function MUE_CadenaSQLMaq(ByVal NSistema As Integer) As String
'********************************************
' genera un fragmento de cada SQL para
' filtraje de los datos que pertenecen
' a una sección concreta. Poniendo como
' condición la sección máquina (secc_maqui)
'********************************************
    Dim bucle As Integer
    Dim StrSQL As String
    StrSQL = "("
    If SisEur.Sistemas(NSistema).NSecciones = 0 Then
        StrSQL = StrSQL & "secc_maqui=0"
    End If
    For bucle = 0 To SisEur.Sistemas(NSistema).NSecciones - 1
        StrSQL = StrSQL & " secc_maqui=" & Format(SisEur.Sistemas(NSistema).Smaq(bucle))
        If (bucle <> SisEur.Sistemas(NSistema).NSecciones - 1) Then
            StrSQL = StrSQL & " or"
        End If
        'If Bucle = 0 And SisEur.Sistemas(NSistema).NSecciones = 2 Then
        '    strSQL = strSQL & " or"
        'End If
    Next bucle
    StrSQL = StrSQL & ")"
    MUE_CadenaSQLMaq = StrSQL
End Function
'TouchScale --> MUE_LanzaDatos
'Public Function MUE_LanzaDatos(TipoDato As Integer, Modificaciones As Boolean, Optional MiCodigo As Long) As Boolean
'    Dim Bucle As Integer
'    Dim Marcar As Boolean
'    'Dim AhoBase as dao.database
'    'Dim ahoReg as dao.recordset
'    '*******************************************************
'    ' esta función realiza el ciclo de envíos por cada
'    ' sección, para datos asociados a una sección concreta.
'    '
'    ' TipoDato :
'    '   0 --> Artículos
'    '   1 --> Códigos de barras
'    '   2 --> Cabeceras, Leyendas, Textos Libres
'    '   3 --> Publicidad y ofertas
'    '   4 --> Teclas
'    '   5 --> Vendedores
'    '   6 --> Tramos
'    '*******************************************************
'    MUE_Sistemas
'    'CadenadeLog "mue_lanzadatos..."
'    If SisEur.NSistemas > 0 Then
'        For Bucle = 0 To SisEur.NSistemas - 1
'
'
'            If Bucle = SisEur.NSistemas - 1 Then Marcar = True
'            If CancelHumano Then Exit For
'            canCelar = 0
'            Select Case TipoDato
'                Case 0
'                    MUE_EnviaArticuloNW Bucle, Modificaciones, Marcar
'                    'If AhorraMas Then 'c2f 1.7.25
'                    '    Set AhoBase = OpenDatabase(App.Path & "\dbasetouch.mdb")
'                    '    Set ahoReg = AhoBase.OpenRecordset("select * from articulo where borrado")
'                    '    If Not (ahoReg.EOF) Then
'                    '        ahoReg.MoveFirst
'                    '        Do While Not (ahoReg.EOF)
'                    '            ahoReg.Edit
'                    '            ahoReg.Delete
'                    '            ahoReg.Update
'                    '            ahoReg.Movenext
'                    '        Loop
'                    '
'                    '    End If
'                    '    ahoReg.Close
'                    '    AhoBase.Close
'                    'End If
'                Case 1
'                    MUE_EnviaCodigosBarras Bucle, Modificaciones, Marcar
'                Case 2
'                    If MiCodigo = 0 Then
'                        'MUE_EnviaCabecerasLeyendas Bucle, Modificaciones, False
'                    Else
'                        'MUE_EnviaCabecerasLeyendas Bucle, Modificaciones, True
'                    End If
'                Case 3
'                    'MUE_EnviaPublicidad Bucle, Modificaciones
'                Case 4
'                    'MUE_EnviaTeclas Bucle, Modificaciones
'                Case 5
'                    MUE_EnviaVendedor Bucle, Modificaciones
'                Case 6
'                    'MUE_EnviaTramos Bucle, Marcar, Modificaciones
'            End Select
'
'        Next Bucle
'    End If
'End Function

'Public Function MUE_LanzaDatosComunes(TipoDato As Integer, Modificaciones As Boolean, Optional Data_Aditional As Long) As Boolean
'    Dim Bucle As Integer
'    Dim Marcar As Boolean
'    Dim ErrorPrevio As Boolean
'    Dim FinDia As Boolean
'    '*******************************************************
'    ' esta función realiza el ciclo de envíos por cada
'    ' sección, para datos no asociados a una sección en concreto.
'    '
'    ' TipoDato :
'    '   0 --> IVAS
'    '   1 --> Clientes
'    '   2 --> Países
'    '   3 --> Fichas de Vacuno
'    '   4 --> Recepción totales Fichas de Vacuno
'    '   5 --> Recepción de Pedido
'    '   6 --> Recepción de inventario
'    '*******************************************************
'    ErrorPrevio = False
'    Marcar = False
'    MUE_Sistemas
'    If SisEur.NSistemas > 0 Then
'        For Bucle = 0 To SisEur.NSistemas - 1
'            If Bucle = SisEur.NSistemas - 1 Then
'                If Not ErrorPrevio Then Marcar = True
'            End If
'            If CancelHumano Then Exit For
'            canCelar = 0
'            Select Case TipoDato
'                Case 0
'                    'MUE_EnviaIVAS Bucle, Modificaciones, Marcar
'                Case 1
'                    MUE_EnviaCLientes Bucle, Modificaciones, Marcar
'                Case 2
'                    'MUE_EnviarPaises Bucle, Modificaciones, Marcar
'                Case 3
'                    'If VacunoD Then
'                        MUE_EnviarFichasVacuno Bucle, Modificaciones, Marcar
'                    'Else
'                    '    MUE_EnviarFichasVacunoC Bucle, Modificaciones, Marcar
'                    'End If
'                Case 4
'                    'MUE_RecibirTotalesFichasVacuno Bucle
'                Case 5 ' pedido
'                    FinDia = False
'                    If Data_Aditional = 1 Then FinDia = True
'                    'MUE_SolicitaPedido Bucle, FinDia
'                Case 6 ' inventario
'                    FinDia = False
'                    If Data_Aditional = 1 Then FinDia = True
'                    'MUE_SolicitaInventario Bucle, FinDia
'            End Select
'            If canCelar = 1 Then ErrorPrevio = True
'        Next Bucle
'    End If
'End Function

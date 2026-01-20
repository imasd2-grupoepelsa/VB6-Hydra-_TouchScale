Attribute VB_Name = "Principal"
Option Explicit

Public Sub Main()
    Dim MyFich As Integer
    Dim MyResul As Long
    Dim MyProc As Long
    Dim sCommand As String
    
    If App.PrevInstance Then End
    
    lNoCoger = False
    lNoConnect = False
    frmControl.tmrpluasc.Enabled = False
    lFinPluAsc = False
    nTocaLocal = 0
    lSerial = False
    sCfgSerial = ""
    nMaqType = 0
    
    If Dir(App.Path & "\ndplu.txt") <> "" Then
        MyFich = FreeFile()
        Open App.Path & "\ndplu.txt" For Input As #MyFich
        Line Input #MyFich, sCommand
        nDelayPluAsc = Val(sCommand)
        Close #MyFich
    Else
        nDelayPluAsc = 2
        MyFich = FreeFile()
        Open App.Path & "\ndplu.txt" For Output As #MyFich
        Print #MyFich, CStr(nDelayPluAsc)
        Close #MyFich
    End If
    
    If Dir(App.Path & "\logtiqah.txt") <> "" Then
        If FileLen(App.Path & "\logtiqah.txt") > 5242880 Then
            On Error Resume Next
            Kill App.Path & "\logtiqah.txt"
            On Error GoTo 0
        End If
    End If
    
    var.version = tools.versionar
    var.modo = 0
    var.UsaEuro = True
    LeeINI
    frmControl.cSysTray1.TrayTip = "TIQAH v" & var.version
    frmControl.cSysTray2.TrayTip = "TIQAH v" & var.version
    frmControl.cSysTray3.TrayTip = "TIQAH v" & var.version
    
    frmControl.cSysTray1.InTray = True
    frmControl.TimeTQT.Interval = MyInterval
    frmControl.TimeTQT.Enabled = True
End Sub
Public Sub Do_Events()
    Sleep 2
    DoEvents
End Sub
Public Sub LeeINI()
    Dim Arch As Integer
    Dim Buf1 As String, Buf2 As String
    Dim nCTQ As Integer
    
    If Dir(App.Path & "\tiqahtouch.ini") = "" Then
        On Error GoTo ErrorCritico1
        Arch = FreeFile()
        Open App.Path & "\tiqahtouch.ini" For Output As #Arch
        If Dir(App.Path & "\infoser.txt") <> "" Then
            Write #Arch, "Path de SC10, ruta de red al concentrador              ", "z:\balanzas"
            Write #Arch, "Intervalo de recogida en minutos (mayor que cero   )   ", "63" '2
            Write #Arch, "Máximo número de tiquets por Recogida (0 no hay límite)", "0"
            Write #Arch, "Intervalo de retardo en milisegundos, entre copias     ", "5000"
            Write #Arch, "Nombre del Fichero de Tiquets Individual               ", "BBTTTT" '5
            Write #Arch, "Ruta del directorio almacen de ficheros de exportacion ", "c:\Balanzas" '6
            Write #Arch, "Dirección IP/Port Balanza                              ", "127.000.000.001:6000" '7
            Write #Arch, "Nombre del Fichero de Tiquets General                  ", "tqgen" '8
            Write #Arch, "Hora de Borrado de Totales                             ", "22:30:00" '9
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
            Write #Arch, "Recepción Tique Venta RED                              ", "1"
            Write #Arch, "Recepción Tique Super RED                              ", "0"
            Write #Arch, "Recepción Tique Envasado RED                           ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio RED                      ", "0"
            Write #Arch, "Recepción Tique Venta Local                            ", "1"
            Write #Arch, "Recepción Tique Super Local                            ", "0"
            Write #Arch, "Recepción Tique Envasado Local                         ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio Local                    ", "0"
            Write #Arch, "Número peticiones Red antes de Locales                 ", "200"
        Else
            Write #Arch, "Tipo Sistema                                           ", "TS"
            Write #Arch, "Intervalo de recogida (60 + segundos)                  ", "64" '2
            Write #Arch, "Máximo número de tiquets por Recogida (0 no hay límite)", "0"
            Write #Arch, "Intervalo de retardo en milisegundos, entre copias     ", "5000"
            Write #Arch, "Nombre del Fichero de Tiquets Individual               ", "XBBTTTT" '5
            Write #Arch, "Ruta del directorio almacen de ficheros de exportacion ", "exporta" '6
            Write #Arch, "Dirección IP/Port Balanza                              ", "127.000.000.001:3306" '7
            Write #Arch, "Nombre del Fichero de Tiquets General                  ", "tqgen" '8
            Write #Arch, "Hora de Borrado de Totales                             ", "00:00:00" '9
            Write #Arch, "Recepción Tique Venta Red sin Cobrar                   ", "1"
            Write #Arch, "Recepción Tique Venta Red Cobrado                      ", "1"
            Write #Arch, "Recepción Tique Envasado RED                           ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio RED                      ", "0"
            Write #Arch, "Recepción Tique Venta Local sin cobrar                 ", "0"
            Write #Arch, "Recepción Tique Venta Local Cobrado                    ", "0"
            Write #Arch, "Recepción Tique Envasado Local                         ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio Local                    ", "0"
            Write #Arch, "Número peticiones Red antes de Locales                 ", "200"
        End If
        Close #Arch
        On Error GoTo 0
    End If
    
    On Error GoTo ErrorCritico2
    Arch = FreeFile()
    Open App.Path & "\tiqahtouch.ini" For Input As #Arch
    Input #Arch, Buf1, Buf2 '1
    var.PathSC10 = Buf2
    If UCase(var.PathSC10) = "SC10" Then nMaqType = 1
    If UCase(var.PathSC10) = "SP" Then nMaqType = 10
    If UCase(var.PathSC10) = "TS" Then nMaqType = 11
    
    Input #Arch, Buf1, Buf2 '2
    If IsNumeric(Buf2) Then
        TimeAuto = Val(Buf2)
        TimeAutoSeg = 0
        sInftxt1 = CStr(TimeAuto)
        If TimeAuto <= 5 Then
            sInftxt1 = sInftxt1 & " Minutos."
        Else
            If TimeAuto > 60 Then
                TimeAuto = TimeAuto - 60
                TimeAutoSeg = TimeAuto
                sInftxt1 = CStr(TimeAutoSeg) & " Segundos."
            Else
                sInftxt1 = sInftxt1 & " Segundos."
            End If

        End If
        If TimeAuto < 1 Then
            TimeAuto = 10
            lNoCoger = True
            sInftxt1 = "10 segundos (Si Reco.txt)"
            'CadenadeLog "Tiempo de recogida automática incorrecto"
            'GoTo ErrorCritico2
        End If
    Else
        CadenadeLog "Tiempo de recogida automática incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '3
    If IsNumeric(Buf2) Then
        NumMax = Val(Buf2)
        If NumMax < 0 Then
            CadenadeLog "Número máximo de tiquets incorrecto"
            GoTo ErrorCritico2
        End If
    Else
        CadenadeLog "Número máximo de tiquets incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '4
    If IsNumeric(Buf2) Then
        MiliTiquet = Val(Buf2)
        If MiliTiquet < 0 Then
            CadenadeLog "Número máximo de tiquets incorrecto"
            GoTo ErrorCritico2
        End If
    Else
        CadenadeLog "Número máximo de tiquets incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '5
    var.TqNombre = Buf2
    sInftxt2 = Buf2
    
    Input #Arch, Buf1, Buf2 '6
    If Buf2 = "" Then
        Buf2 = App.Path
    End If
    var.PathExpor = Buf2
    sInftxt3 = Buf2
    '
    Input #Arch, Buf1, Buf2 '7
    sAddIP = Buf2
    sInftxt4 = Buf2
    If UCase(Mid(sAddIP, 1, 3)) = "COM" Then
        lSerial = True
        sCOMMX = Mid(sAddIP, 1, 4)
        sCfgSerial = Mid(sAddIP, 4, 1)
        sCfgSerial = Mid(sAddIP, 6) & "," & sCfgSerial
    End If
    '
    Input #Arch, Buf1, Buf2 '8
    sTiqGen = Buf2
    sInftxt5 = Buf2
    
    sTiqGen = sTiqGen & Format(Day(Now), "00")
    If Month(Now) < 10 Then
        sTiqGen = sTiqGen & CStr(Month(Now))
    Else
        sTiqGen = sTiqGen & Chr(Asc("A") + Month(Now) - 10)
    End If
    Input #Arch, Buf1, Buf2 '9
    sHoraFin = Buf2
    sInftxt6 = Buf2
    dHoraFin = CDate(sHoraFin)
    
    On Error GoTo INIERR
    For nCTQ = 0 To 7
        Input #Arch, Buf1, Buf2
        TipoTQT(nCTQ) = Buf2
    Next nCTQ
    
    Input #Arch, Buf1, Buf2
    
    nContLocal = Val(Buf2)
    
INIERR:
    If Err.Number <> 0 Then
        For nCTQ = 0 To 7
        Select Case nCTQ
            Case 0
            TipoTQT(nCTQ) = "1"
            Case 1
            TipoTQT(nCTQ) = "0"
            Case 2
            TipoTQT(nCTQ) = "0"
            Case 3
            TipoTQT(nCTQ) = "0"
            Case 4
            TipoTQT(nCTQ) = "1"
            Case 5
            TipoTQT(nCTQ) = "0"
            Case 6
            TipoTQT(nCTQ) = "0"
            Case 7
            TipoTQT(nCTQ) = "0"
        End Select
        
        Next nCTQ
        nContLocal = 200
        
    End If
    
    Close #Arch
    On Error GoTo 0
    Exit Sub
ErrorCritico1:
    MsgBox "No se puede escribir el fichero de configuración TIQAHTOUCH.INI", vbCritical, "Recogida Contínua de Tiquets"
    End
ErrorCritico2:
    MsgBox "fichero de configuración TIQAHTOUCH.INI incorrecto", vbCritical, "Recogida Contínua de Tiquets"
    End
End Sub

Public Function FrmSC10_BufTiquet() As Boolean
    '***************************************
    '* pasa los tiquets AH a carpeta local *
    '***************************************
    Dim BuclePath As Long
    Dim ListaArch() As String
    Dim Arch1 As Integer
    Dim Arch2 As Integer
    Dim Buf As String
    Dim CopiaLenta As Boolean
    Dim CuentaEvents As Integer
    Dim CuentaTiquets As Long
    Dim Retardo As Integer
    cancelar = 0
    CopiaLenta = False
    'frmmonitor.MostrarDato "Obteniendo tickets de la red..."
    Do_Events
    If Dir(App.Path & "\copiavb.txt") <> "" Then CopiaLenta = True
    If Dir(App.Path & "\ticksc10", vbDirectory) = "" Then MkDir App.Path & "\ticksc10"
    On Error GoTo errorred
    frmControl.File1.Path = PathSC10
    frmControl.File1.Pattern = "T*.*"
    frmControl.File1.Refresh
    On Error GoTo 0
    If frmControl.File1.ListCount > 0 Then
        frmControl.File1.ListIndex = 0
        ReDim ListaArch(frmControl.File1.ListCount)
        For BuclePath = 0 To frmControl.File1.ListCount - 1
            ListaArch(BuclePath) = frmControl.File1.List(BuclePath)
        Next BuclePath
    End If
    For BuclePath = 0 To frmControl.File1.ListCount - 1
        If NumMax > 0 Then
            If CuentaTiquets > NumMax Then
                Exit For
            Else
                CuentaTiquets = CuentaTiquets + 1
            End If
        End If
        If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
            Do_Pausa
            FrmSC10_BufTiquet = False
            Exit Function
        End If
        '**********************
        ' comprueba formato tiquet
        If Len(ListaArch(BuclePath)) = 12 And _
        StrConv(Left(ListaArch(BuclePath), 1), vbUpperCase) = "T" And _
        IsNumeric(Mid(ListaArch(BuclePath), 2, 5)) Then
            frmMonitor.Caption = ListaArch(BuclePath)
            On Error GoTo errorred
            If Not CopiaLenta Then FileCopy PathSC10 & "\" & ListaArch(BuclePath), _
            App.Path & "\ticksc10\" & ListaArch(BuclePath)
            If CopiaLenta Then
                Arch1 = FreeFile()
                On Error GoTo errorred
                Open PathSC10 & "\" & ListaArch(BuclePath) For Input As #Arch1
                On Error GoTo 0
                Arch2 = FreeFile()
                Open App.Path & "\ticksc10\" & ListaArch(BuclePath) For Output As #Arch2
                Do Until EOF(Arch1)
                    On Error GoTo errorred
                    Line Input #Arch1, Buf
                    On Error GoTo 0
                    Print #Arch2, Buf
                Loop
                Close #Arch1
                Close #Arch2
                Arch1 = 0
                Arch2 = 0
            End If
            Kill PathSC10 & "\" & ListaArch(BuclePath)
            On Error GoTo 0
            
            'c2f 1.3.1 ... Se anula de momento
            'CuentaEvents = 0
            'frmControl.Ret = False
            'frmControl.TmRet.Interval = MiliTiquet
            'frmControl.TmRet.Enabled = True
            'Do Until frmControl.Ret = True
            '    Sleep (1)
            '    CuentaEvents = CuentaEvents + 1
            '    If CuentaEvents > 100 Then
            '        If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then Exit Do
            '        CuentaEvents = 0
            '        DoEvents
            '    End If
            'Loop
            'frmControl.TmRet.Enabled = False
            ''''''''''''''''''''''''''''''''''
            
            'CuentaEvents = CuentaEvents + 1
            'If CuentaEvents >= 10 Then
            '    Do_Events
            '    CuentaEvents = 0
            'End If
            
            
        End If
    Next BuclePath
    frmMonitor.Caption = "Recogida Continua de Tiquets"
    FrmSC10_BufTiquet = True
    Exit Function
errorred:
    frmMonitor.Caption = "Recogida Continua de Tiquets"
    CadenadeLog "Sin acceso a ruta de red"
    If Arch1 <> 0 Then Close Arch1
    If Arch2 <> 0 Then Close Arch2
    'frmmonitor.MostrarDato CargaCadena(854) & " " & PathSC10
    'frmmonitor.MostrarDato CargaCadena(544)
    'frmEpelsa.LblSC10.Caption = CargaCadena(544)
    FrmSC10_BufTiquet = False
End Function
Public Function lMod() As Boolean
    Dim Mybase As DAO.Database
    Dim nR As DAO.Recordset
    Dim l As Boolean
    
    l = False
    
    Set Mybase = OpenDatabase(App.Path & "\dbasetouch.mdb")
    Set nR = Mybase.OpenRecordset("select plu,tran_plu from articulo where tran_plu<>'*'")
    If Not nR.EOF Then
        l = True
    End If
    
    nR.Close
    Set nR = Nothing
    Mybase.Close
    Set Mybase = Nothing
    
    lMod = l
End Function
Public Function AbrirBase() As DAO.Database
    Dim Mybase As DAO.Database
    Dim Continuar As Boolean
    Dim Bucle As Integer
    Continuar = False
    Do Until Continuar
        On Error Resume Next
        Set Mybase = OpenDatabase(App.Path & "\dbasetouch.mdb")
        If Err.Number = 0 Then Continuar = True
        Do_Events
        On Error GoTo 0
        '1.4.1...
        'For Bucle = 1 To 10
        '    Sleep (100)
        '    DoEvents
        'Next Bucle
        '....
    Loop
    Set AbrirBase = Mybase
End Function
Public Sub FrmSC10_Coge_tiquets_AH()
Dim MyPath As String
Dim MiContador As Integer
Dim Archivo As Integer
Dim MiRecibe As String
Dim BuclePath As Long
Dim PedirTiquet As String
Dim Cabecera As tipo_cabecera
Dim Linea() As Tipo_Linea
Dim Repetir As Boolean
Dim NroLineas As Integer
Dim ToTalLineas As Integer
Dim ErrorLineas As Boolean
Dim Buffer As String
Dim Base As DAO.Database
Dim RegArt As DAO.Recordset
Dim ForzarTodos As Boolean
Dim Contador As Long
Dim ListaArch() As String
Dim ok As Boolean
    frmControl.cSysTray1.InTray = False
    frmControl.cSysTray2.InTray = True
    If Dir(App.Path & "\sc10tiq.txt") <> "" Then
        ForzarTodos = True
        MyPath = App.Path & "\ticksc10"
        ok = FrmSC10_BufTiquet
        ok = True
        If Not ok Then
            frmControl.cSysTray2.InTray = False
            frmControl.cSysTray1.InTray = True
            Exit Sub
        End If
    Else
        MyPath = PathSC10
    End If
    Cabecera.Ntiquet = 0
    Repetir = True
    Set Base = AbrirBase()
    On Error GoTo errorred
    frmControl.File1.Path = MyPath
    frmControl.File1.Pattern = "T*.*"
    frmControl.File1.Refresh
    On Error GoTo 0
    If frmControl.File1.ListCount > 0 Then
        frmControl.File1.ListIndex = 0
        ReDim ListaArch(frmControl.File1.ListCount)
        For BuclePath = 0 To frmControl.File1.ListCount - 1
            ListaArch(BuclePath) = frmControl.File1.List(BuclePath)
        Next BuclePath
    End If
    Contador = 0
    For BuclePath = 0 To frmControl.File1.ListCount - 1
        '**********************
        ' comprueba formato tiquet
        Do_Events
        If Not ForzarTodos Then
            If (NumMax > 0) Then
                If Contador > NumMax Then
                    frmControl.cSysTray2.InTray = False
                    frmControl.cSysTray1.InTray = True
                    Exit For
                End If
            End If
        End If
        If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
            Base.Close
            frmControl.cSysTray2.InTray = False
            frmControl.cSysTray1.InTray = True
            Do_Pausa
            Exit Sub
        End If
        If Len(ListaArch(BuclePath)) = 12 And _
        StrConv(Left(ListaArch(BuclePath), 1), vbUpperCase) = "T" And _
        IsNumeric(Mid(ListaArch(BuclePath), 2, 5)) Then
            If SC10_PROCESARTOTALVENTA Then
                If Not ForzarTodos Then Contador = Contador + 1
                Archivo = FreeFile()
                On Error Resume Next
                Open MyPath & "\" & ListaArch(BuclePath) For Input As Archivo
                Line Input #Archivo, MiRecibe
                On Error GoTo 0
                If MiRecibe = "" Then
                    Exit For
                End If
                '***
                ' si STS=0, lo procesa (tiquet cerrado)
                '***
                If Mid(MiRecibe, 53, 1) = "0" Then
                    Repetir = True
                    '********************
                    ' Toma los datos de Cabecera
                    '********************
                    Cabecera.NLineas = 0
                    Cabecera.Ntiquet = Mid(MiRecibe, 4, 5)
                    'frmmonitor.MostrarDato CargaCadena(865) & " " & Cabecera.Ntiquet
                    'frmEpelsa.LblSC10.Caption = CargaCadena(866) & " " & frmControl.File1.FileName
                    Cabecera.NMostrador = Mid(MiRecibe, 19, 2)
                    Cabecera.NVendedor = Mid(MiRecibe, 22, 4)
                    Cabecera.NBalanza = Mid(MiRecibe, 27, 2)
                    'Cabecera.NLineas = Mid(MiRecibe, 30, 2)
                    'ReDim Linea(Cabecera.NLineas + 1)
                    Cabecera.LineasMensaje = Cabecera.NLineas
                    Cabecera.LineaInicial = 1
                    'Cabecera.LineaFinal = Cabecera.NLineas
                    Cabecera.LineasCanceladas = Mid(MiRecibe, 33, 2)
                    Cabecera.ImporteTotal = Mid(MiRecibe, 36, 7)
                    Cabecera.ImporteBruto = Cabecera.ImporteTotal
                    If IsNumeric(Mid(MiRecibe, 44, 2)) And IsNumeric(Mid(MiRecibe, 46, 2)) Then
                        Cabecera.hora = TimeSerial(Mid(MiRecibe, 44, 2), Mid(MiRecibe, 46, 2), 0)
                    Else
                        Cabecera.hora = TimeSerial(0, 0, 0)
                    End If
                    Cabecera.NCliente = Mid(MiRecibe, 57, 6)
                    If Mid(MiRecibe, 70, 2) = "  " Then
                       Cabecera.Fecha = DateSerial(Mid(MiRecibe, 68, 2), Mid(MiRecibe, 66, 2), Mid(MiRecibe, 64, 2))
                    Else
                       Cabecera.Fecha = DateSerial(Mid(MiRecibe, 70, 2), Mid(MiRecibe, 67, 2), Mid(MiRecibe, 64, 2))
                    End If
                    Cabecera.tipoTiquet = Mid(MiRecibe, 2, 1)
                    If Cabecera.tipoTiquet <> "0" Then
                        Cabecera.EstadoTiquet = 17
                        Cabecera.CadenaEstadoTiquet = "Local"
                        Cabecera.tipoTiquet = 5
                    Else
                        Cabecera.EstadoTiquet = 16
                        Cabecera.CadenaEstadoTiquet = "Red"
                        Cabecera.tipoTiquet = 1
                    End If
                    Cabecera.descuento = 0
                    '///cas.v117
                    Cabecera.tipoPago = 0
                    Cabecera.cantidadEntre = 0
                    Cabecera.Cambio = 0
                    'If Cabecera.tipoPago = 255 Then Cabecera.Cambio = 0
                    '///cas.v117
                    '*******************
                    ' Si el tipo de tiquet es el que queremos,pide las lineas
                    '*******************
                    If Mid(MiRecibe, 53, 1) = 0 Then
                        MiContador = 1
                        Do Until EOF(Archivo)
                            Line Input #Archivo, MiRecibe
                            Cabecera.NLineas = MiContador
                            Cabecera.LineaFinal = Cabecera.NLineas
                            Cabecera.LineasMensaje = Cabecera.NLineas
                            ReDim Preserve Linea(Cabecera.NLineas + 1)
                            'frmmonitor.MostrarDato CargaCadena(867) & " " & MiContador
                            '******************************
                            ' recoge los datos de la(s) linea(s)
                            '******************************
                            'STV (positivo o negativo)
                            If Mid(MiRecibe, 51, 1) = "0" Then
                                Linea(MiContador).FactorPeso = 1
                                Linea(MiContador).TipoLinea = 0
                            Else
                                Linea(MiContador).FactorPeso = -1
                                Linea(MiContador).TipoLinea = 2
                            End If
                            'STP (pesado o no pesado)
                            If Mid(MiRecibe, 53, 1) = "0" Then
                                Linea(MiContador).EsPesado = True
                            Else
                                Linea(MiContador).EsPesado = False
                                Linea(MiContador).TipoLinea = Linea(MiContador).TipoLinea + 1
                            End If
                            ' STC (cancelado o no cancelado)
                            If Mid(MiRecibe, 55, 1) = "0" Then
                                Linea(MiContador).LineaCancelada = False
                            Else
                                Linea(MiContador).LineaCancelada = True
                                Linea(MiContador).TipoLinea = Linea(MiContador).TipoLinea + 16
                            End If
                            Linea(MiContador).NumeroLinea = MiContador
                            Linea(MiContador).DescuentoLinea = 0
                            Linea(MiContador).CodigoArticulo = Mid(MiRecibe, 26, 6)
                            Linea(MiContador).precio = Mid(MiRecibe, 33, 7)
                            Linea(MiContador).ImporteLinea = Mid(MiRecibe, 57, 7)
                            'If linea(MiContador).LineaCancelada = True Then
                            '    linea(MiContador).ImporteLinea = linea(MiContador).ImporteLinea * (-1)
                            'End If
                            Linea(MiContador).cantidad = Mid(MiRecibe, 41, 9)
                            If Linea(MiContador).EsPesado = True Then
                                Linea(MiContador).cantidad = Linea(MiContador).cantidad / 1000
                            End If
                            '***************************************
                            ' verifica si se ha modificado el precio
                            '***************************************
                            Set RegArt = Base.OpenRecordset _
                            ("select * from articulo where codigo=" & _
                            Linea(MiContador).CodigoArticulo & " and borrado=false")
                            If RegArt.EOF Then
                                Linea(MiContador).SehaModificado = False
                                Linea(MiContador).IVA = 0
                            Else
                                Linea(MiContador).IVA = RegArt!tipo_iva
                                
                                If (Linea(MiContador).precio / 100) <> RegArt!Euros Then
                                    Linea(MiContador).SehaModificado = True
                                Else
                                    Linea(MiContador).SehaModificado = False
                                End If
                            End If
                            RegArt.Close
                            MiContador = MiContador + 1
                        Loop
                        '***********************************************
                        ' si se han recibido todas las líneas lo procesa
                        '***********************************************
                        Close #Archivo
                        If Cabecera.NLineas = MiContador - 1 Then
                            Repetir = False
                            On Error Resume Next
                            Kill MyPath & "\" & ListaArch(BuclePath)
                            On Error GoTo 0
                            Cabecera.Incluir_en_Base = True
                            db.Actualiza_Tabla_Cabecera Base, Cabecera
                            If Cabecera.Incluir_en_Base = True And SC10_PROCESARTOTALVENTA Then
                                db.Actualiza_Tabla_Tickets Base, Cabecera
                                db.Actualiza_Tabla_Gtarti Base, Cabecera
                                db.actualiza_tabla_gtInfVen Base, Cabecera
                                db.Actualiza_Tabla_Gtsecs Base, Cabecera
                                db.Actualiza_Tabla_GtVend Base, Cabecera
                                db.Actualiza_Tabla_Log Base, Cabecera
                            End If
                        End If
                    End If
                Else
                    Close #Archivo
                    On Error Resume Next
                    Kill MyPath & "\" & ListaArch(BuclePath)
                    On Error GoTo 0
                End If
            Else
                On Error Resume Next
                Kill MyPath & "\" & ListaArch(BuclePath)
                On Error GoTo 0
                Close #Archivo
            End If
        End If
    Next BuclePath
errorred:
    frmControl.cSysTray2.InTray = False
    frmControl.cSysTray1.InTray = True
    CadenadeLog "Sin acceso a ruta de Red"
    Base.Close
End Sub

Public Sub Do_Pausa()
    Dim MyBuf As Boolean
    Dim lCont As Long
    Dim MyICO As Integer
    
    If frmControl.cSysTray1.InTray = True Then
        MyICO = 1
    Else
        MyICO = 2
    End If
    frmControl.cSysTray1.InTray = False
    frmControl.cSysTray2.InTray = False
    frmControl.cSysTray3.InTray = True
    frmMonitor.Caption = "PAUSA"
    CadenadeLog "Pausa solicitada por Hydra Basic"
    If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
        Do Until MyBuf
            On Error GoTo 0
            On Error Resume Next
            Kill App.Path & "\dirtmp\tqpausa"
            If Err.Number = 0 Then
                MyBuf = True
            End If
        Loop
    End If
    On Error GoTo 0
    Do Until Not MyBuf
        Sleep (10)
        lCont = lCont + 1
        Do_Events
        If Dir(App.Path & "\dirtmp\tqreanuda") <> "" Then MyBuf = False
        If lCont > 1000 Then MyBuf = False
        On Error Resume Next
        If MyBuf = False Then Kill App.Path & "\dirtmp\tqreanuda"
        On Error GoTo 0
    Loop
    '
    frmControl.cSysTray3.InTray = False
    If MyICO = 1 Then
        frmControl.cSysTray1.InTray = True
    Else
        frmControl.cSysTray2.InTray = True
    End If
    frmMonitor.Caption = "Recogida contínua de Tiquets"
    CadenadeLog "Fin de Pausa"
End Sub



'\\\\\\
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Euroscale=true OR SC10=false
'/////////////////////////////
Public Sub Taquion_EnviaTiquetsPendientes(Euroscale_o_SC10 As Boolean)
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim rstAux As DAO.Recordset
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
Dim Bucle As Integer
Dim RegIVA As DAO.Recordset
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

Dim sPes_Uni As String

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim sSQL As String

CadenadeLog "Escritura Tiques pendientes..."

    Fichero1 = FreeFile()
    'If Euroscale_o_SC10 Then
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If
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
    sSQL = "SELECT * " & _
           "FROM " & MiTabla1 & " " & _
           "WHERE reserv3=" & Chr(34) & "NO" & Chr(34) & " or reserv3=" & Chr(34) & "--" & Chr(34) & " " & _
           "ORDER BY val(nume)"
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
                        NombreF.modo = 1
                    Case 2
                        NombreF.modo = 2
                    Case 3
                        NombreF.modo = 3
                        
                        If Tipo = "A" Or Tipo = "B" Then
                            Tipo = "F"
                        Else
                            Tipo = "E"
                        End If
                        
                    Case 4
                        NombreF.modo = 4
                    Case 5
                        NombreF.modo = 1
                    Case 6
                        NombreF.modo = 2
                    Case 7
                        NombreF.modo = 3
                    Case 8
                        NombreF.modo = 4
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
                If Dir(App.Path & "\nscli.txt") <> "" Then
                    If IsNull(!numlote) Then
                        sCadAd = Space(20)
                    Else
                        sCadAd = Trim(!numlote) & Space(20 - Len(Trim(!numlote)))
                    End If
                    sCadAd = sCadAd & Format(Val(!codcli), "000000")
                Else
                    sCadAd = ""
                End If
                sCadfP = ""
                If Dir(App.Path & "\nsfpa.txt") <> "" Then
                    If IsNull(!tipoPago) Then
                        sCadfP = "0"
                    Else
                        sCadfP = CStr(!tipoPago)
                    End If
                Else
                    sCadfP = ""
                End If
                
                ReDim BufferLineas(1)
                BufferLineas(0) = MiEnvio & sCadAd & sCadfP
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
                    If Dir(App.Path & "\transpluasc.exe") <> "" Then
                        NombreTiquet = Mid(NombreTiquet, 1, 1) & Mid(Cod_Bar, 3, 2) & Mid(NombreTiquet, 4)
                    End If
                    '''''''''''''''''''''
                End If
                '//Inco ver si se activa cabecera If Dir(App.Path & "\sncli.txt") = "" Then
                    MiCuentaLineas = 0
                '//End If
                'NombreTiquet = Codigo_bal & Format(!nume, "0000") & "." & _
                'Left(fecha, 2) & Hex$(Val(Mid(fecha, 3, 2)))
                NombreTiquet = NombreTiquet & "." & Left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
                rst.Edit
                .Fields("reserv3") = "--"
                .Update
                Do_Events
            Else
                HayCabecera = False
            End If
        End With
        Set rstAux = db.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
        Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & rst!nume & Chr(34) & _
                 " and numbal=" & rst!Balanza & " and date=" & Chr(34) & rst!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & rst!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & rst!TYPTIC & Chr(34) & " order by val(numlin), lincan") 'c2f 1.6.12 - 1
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
                        If Dir(App.Path & "\transpluasc.exe") <> "" Then
                        MiEnvio = Format(rst!nume, "0000") & Num_lin & sVmS & "0" & sVm & Format(Val(!Vendor), "0000") & _
                                Cod_Bar & Cod_Cod & Cod_Plu & Format(Abs(rst!Importe), "00000.00") & Format(Abs(!price), "00000.00") & _
                                sPes_Uni & Format(Abs(!amount), "00000.00") & Signo & Tipo_Linea & LineaAnulada & Mid(CStr(!IVA), 1, 1) & _
                                Tipo & Format(!D_FECHA, "ddmmyyyy") & Format(!D_HORA, "hhmm")
                        
                        Else
                            'MiEnvio = Format(rst!nume, "0000") & Num_lin & sVmS & sVm & Format(Val(!Vendor), "0000") & _
                            '        Cod_Bar & Cod_Cod & Cod_Plu & Format(Abs(rst!Importe), "00000.00") & Format(Abs(!price), "00000.00") & _
                            '        sPes_Uni & Format(Abs(!amount), "00000.00") & Signo & Tipo_Linea & LineaAnulada & Mid(CStr(!IVA), 1, 1) & _
                            '        Tipo & Format(!D_FECHA, "ddmmyyyy") & Format(!D_HORA, "hhmm")
                            MiEnvio = Format(rst!nume, "0000") & Num_lin & sVmS & sVm & Format(Val(!Vendor), "0000") & _
                                    Cod_Bar & Cod_Cod & Cod_Plu & precio & Format(Abs(!price), "00000.00") & _
                                    sPes_Uni & Format(Abs(!amount), "00000.00") & Signo & Tipo_Linea & LineaAnulada & Mid(CStr(!IVA), 1, 1) & _
                                    Tipo & Format(!D_FECHA, "ddmmyyyy") & Format(!D_HORA, "hhmm")
                            
                            If Dir(App.Path & "\fichalote.txt") <> "" Then
                                scadAD1 = Format(Val(!reserv1), "0000")
                            End If
                        End If
                        MiEnvio = sin_Coma(MiEnvio)
                        ReDim Preserve BufferLineas(MiCuentaLineas + 1)
                        ReDim Preserve BufferLineasAD1(MiCuentaLineas + 1)
                        BufferLineas(MiCuentaLineas) = MiEnvio
                        BufferLineasAD1(MiCuentaLineas) = scadAD1
                        MiCuentaLineas = MiCuentaLineas + 1
                        rstAux.Edit
                        .Fields("reserv3") = "SI"
                        .Update
                        .MoveNext
                        Do_Events
                    End With
                 Loop
                 
                 If HayCabecera Then
                     Fichero1 = FreeFile()
                     On Error GoTo 0
                     Open App.Path & "\tiquet.buffer" For Output As #Fichero1
                     Fichero2 = FreeFile()
                     Open var.PathExpor & "\" & sTiqGen For Append As #Fichero2
                     For BucleLineas = 0 To MiCuentaLineas - 1
                        Print #Fichero1, BufferLineas(BucleLineas) & sCadAd & sCadfP & BufferLineasAD1(BucleLineas)
                        Print #Fichero2, BufferLineas(BucleLineas) & sCadAd & sCadfP & BufferLineasAD1(BucleLineas)
                     Next BucleLineas
                     Close #Fichero1
                     Close #Fichero2
                     'frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905) & " " & Cod_Bar
                     'FileCopy App.Path & "\tiquet.buffer", App.Path & "\" & NombreTiquet
                     FileCopy App.Path & "\tiquet.buffer", var.PathExpor & "\" & NombreTiquet
                 
                 End If
                 
            Else
                rst.Edit
                rst.Fields("reserv3") = "SI"
                rst.Update
            End If
        End With
        rst.MoveNext
        Loop
    End If
    db.Close
    Exit Sub
ErrorPath:
    MsgBox "Error al acceder a " & App.Path
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Euroscale=true OR SC10=false
'/////////////////////////////
Public Sub Taquion_EnviaTiquetsPendientes_ah(Euroscale_o_SC10 As Boolean)
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim rstAux As DAO.Recordset
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
Dim Bucle As Integer
Dim RegIVA As DAO.Recordset
Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim NombreTiquet As String
Dim BufferLineas() As String
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
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim sSQL As String
    Fichero1 = FreeFile()
    'If Euroscale_o_SC10 Then
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If
    Set db = AbrirBase()
    '*******************
    ' carga tipos de IVA
    '*******************
    Set RegIVA = db.OpenRecordset("SELECT * FROM iva")
    For Bucle = 0 To 4
        RegIVA.FindFirst "tipo=" & (Bucle + 1)
        If RegIVA.EOF Or RegIVA.NoMatch Then
            MiIva(Bucle).porcentaje = 0
        Else
            MiIva(Bucle).porcentaje = RegIVA!Valor
        End If
    Next Bucle
    '*******************
    ' /carga tipos de IVA
    '*******************
    sSQL = "SELECT * " & _
           "FROM " & MiTabla1 & " " & _
           "WHERE reserv3=" & Chr(34) & "NO" & Chr(34) & " or reserv3=" & Chr(34) & "--" & Chr(34) & " " & _
           "ORDER BY val(nume)"
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
                        NombreF.modo = 1
                    Case 2
                        NombreF.modo = 2
                    Case 3
                        NombreF.modo = 3
                    Case 4
                        NombreF.modo = 4
                    Case 5
                        NombreF.modo = 1
                    Case 6
                        NombreF.modo = 2
                    Case 7
                        NombreF.modo = 3
                    Case 8
                        NombreF.modo = 4
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
                ReDim BufferLineas(1)
                BufferLineas(0) = MiEnvio
                MiCuentaLineas = 1
                NombreTiquet = ObtenNombre(NombreF)
                'NombreTiquet = Codigo_bal & Format(!nume, "0000") & "." & _
                'Left(fecha, 2) & Hex$(Val(Mid(fecha, 3, 2)))
                NombreTiquet = NombreTiquet & "." & Left(Fecha, 2) & Hex$(Val(Mid(Fecha, 3, 2)))
                rst.Edit
                .Fields("reserv3") = "--"
                .Update
            Else
                HayCabecera = False
            End If
        End With
        Set rstAux = db.OpenRecordset("select * from " & MiTabla2 & " where reserv3=" & _
        Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & rst!nume & Chr(34) & _
                 " and numbal=" & rst!Balanza & " and date=" & Chr(34) & rst!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & rst!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & rst!TYPTIC & Chr(34) & " order by val(numlin), lincan") 'c2f 1.6.12 - 1
        With rstAux
            If Not .EOF Then
                .MoveFirst
                 For BucleIVA = 0 To 4
                    MiIva(BucleIVA).Base = 0
                    MiIva(BucleIVA).Importe = 0
                 Next BucleIVA
                 Do Until .EOF
                    With rstAux
                        Num_lin = Format(!numlin, "000")
                        Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        Cod_Cod = Format(!Code, "000000")
                        If !price >= 0 Then
                            Precio_Uni = !price
                            Precio_Uni = Quitar_Coma(Precio_Uni, 8)
                        Else
                            Precio_Uni = (-1) * !price
                            Precio_Uni = Quitar_Coma(Precio_Uni, 7)
                            Precio_Uni = "-" & Precio_Uni
                        End If
                        If !IVA >= 0 And !IVA < 5 Then
                            MiIva(!IVA).Importe = MiIva(!IVA).Importe + !amount
                            MiIva(!IVA).Base = MiIva(!IVA).Base + PRsinIVA(!amount, MiIva(!IVA).porcentaje)
                        End If
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            Else
                                Pes_Uni = (-1) * !Weight
                                Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            End If
                            Tipo_Linea = "P"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                Pes_Uni = Quitar_Coma(Pes_Uni, 7)
                            Else
                                Pes_Uni = (-1) * !units
                                Pes_Uni = "-" & Quitar_Coma(Pes_Uni, 6)
                            End If
                            
                            Tipo_Linea = "U"
                        End If
                        If !amount < 0 Then
                            Total = (-1) * !amount
                            Total = "-" & Quitar_Coma(Total, 7)
                            Signo = "-"
                        Else
                            Total = !amount
                            Total = Quitar_Coma(Total, 8)
                            Signo = "+"
                        End If
                        If !lincan = 0 Then
                            LineaAnulada = "N"
                        Else
                            LineaAnulada = "S"
                        End If
                        MiEnvio = "L" & Codigo_mos & _
                        Cod_Bar & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Total & Signo & _
                        Tipo_Linea & LineaAnulada & Cod_Cod & Format(Val(!IVA), "0")
                        ' Envia la cadena
                        ReDim Preserve BufferLineas(MiCuentaLineas + 1)
                        BufferLineas(MiCuentaLineas) = MiEnvio
                        MiCuentaLineas = MiCuentaLineas + 1
                        rstAux.Edit
                        .Fields("reserv3") = "SI"
                        .Update
                        .MoveNext
                    End With
                 Loop
                 Set rstAux = db.OpenRecordset("select * from " & MiTabla2 & " where reserv3<>" & _
                 Chr(34) & "SI" & Chr(34) & " and nume=" & Chr(34) & rst!nume & Chr(34) & _
                 " and numbal=" & rst!Balanza & " and date=" & Chr(34) & rst!Fecha & Chr(34) _
                 & " and hour=" & Chr(34) & rst!hora & Chr(34) _
                 & " and typtic=" & Chr(34) & rst!TYPTIC & Chr(34) & " order by val(numlin)")
                 If rstAux.EOF Then
                     rst.Edit
                     rst.Fields("reserv3") = "SI"
                     rst.Update
                     Fichero1 = FreeFile()
                     On Error GoTo 0
                     Open App.Path & "\tiquet.buffer" For Output As #Fichero1
                     If HayCabecera Then
                        BufferIVA = ""
                        BufferIVA2 = ""
                        PrecioSinIVA = 0
                        PrecioConIVA = 0
                        For BucleIVA = 0 To 4
                            If MiIva(BucleIVA).Base >= 0 Then
                                B1 = MiIva(BucleIVA).Base
                                BufferIVA = BufferIVA & Quitar_Coma(B1, 8)
                            Else
                                B1 = (-1) * MiIva(BucleIVA).Base
                                BufferIVA = BufferIVA & "-" & Quitar_Coma(B1, 7)
                            End If
                            PrecioSinIVA = PrecioSinIVA + MiIva(BucleIVA).Base
                            '
                            If MiIva(BucleIVA).Importe >= 0 Then
                                B2 = MiIva(BucleIVA).Importe
                                BufferIVA2 = BufferIVA2 & Quitar_Coma(B2, 8)
                            Else
                                B2 = (-1) * MiIva(BucleIVA).Importe
                                BufferIVA2 = BufferIVA2 & "-" & Quitar_Coma(B2, 7)
                            End If
                            PrecioConIVA = PrecioConIVA + MiIva(BucleIVA).Importe
                        Next BucleIVA
                        If PrecioSinIVA > 0 Then
                            B1 = PrecioSinIVA
                            BufferLineas(0) = BufferLineas(0) & Quitar_Coma(B1, 8) & BufferIVA
                        Else
                            B1 = PrecioSinIVA * (-1)
                            BufferLineas(0) = BufferLineas(0) & "-" & Quitar_Coma(B1, 7) & BufferIVA
                        End If
                        BufferLineas(0) = BufferLineas(0) & BufferIVA2
                     End If
                     For BucleLineas = 0 To MiCuentaLineas - 1
                        Print #Fichero1, BufferLineas(BucleLineas)
                     Next BucleLineas
                     Close #Fichero1
                     'frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905) & " " & Cod_Bar
                     'FileCopy App.Path & "\tiquet.buffer", App.Path & "\" & NombreTiquet
                     FileCopy App.Path & "\tiquet.buffer", var.PathExpor & "\" & NombreTiquet
                     On Error GoTo 0
                 End If
            Else
                rst.Edit
                rst.Fields("reserv3") = "SI"
                rst.Update
            End If
        End With
        rst.MoveNext
        Loop
    End If
    db.Close
    Exit Sub
ErrorPath:
    MsgBox "Error al acceder a " & App.Path
End Sub


Public Sub TrataBal_BDP()
Dim nFichOrg As Integer
Dim nFichDes As Integer
Dim sOrg As String
Dim sDes As String
Dim sPre As String
Dim sPreS As String
Dim sPes As String
Dim nL As Integer

    nFichOrg = FreeFile()
    Open App.Path & "\bal.dat" For Input As #nFichOrg
    nFichDes = FreeFile()
    Open App.Path & "\bal.tmp" For Output As #nFichDes
    
'            sDes = Format(Val(Mid(sOrg, 1, 13)), "000000") 'codigo
'            sDes = sDes & Format(Val(Mid(sOrg, 16, 2)), "00") 'seccion
'            sDes = sDes & Mid(sOrg, 19, 25)
'            sPre = Right(sOrg, 8)
'            sPre = Trim(sPre)
'            sPreS = ""
'            For nL = 1 To Len(sPre)
'                If Mid(sPre, nL, 1) <> "." Then
'                    sPreS = sPreS & Mid(sPre, nL, 1)
'                End If
'
'            Next nL
'            sDes = sDes & Format(Val(sPreS), "0000000")
    
    Do While Not EOF(nFichOrg)
        Line Input #nFichOrg, sOrg
        sOrg = Trim(sOrg)
        If Len(sOrg) >= 59 Then
            sPes = Mid(sOrg, 19, 1)
            If sPes = "1" Then
                sPes = "U"
            Else
                sPes = "W"
            End If
            sDes = Format(Val(Mid(sOrg, 1, 13)), "000000") 'codigo
            sDes = sDes & Format(Val(Mid(sOrg, 16, 2)), "00") 'seccion
            sDes = sDes & sPes
            If sPes = "W" Then
                If (Mid(sOrg, 19, 1) = "0") Then
                    sDes = sDes & Mid(sOrg, 20, 24)
                Else
                    sDes = sDes & Mid(sOrg, 19, 24)
                End If
            Else
                sDes = sDes & Mid(sOrg, 20, 24)
            End If
            sPre = Right(sOrg, 8)
            sPre = Trim(sPre)
            sPreS = ""
            For nL = 1 To Len(sPre)
                If Mid(sPre, nL, 1) <> "." Then
                    sPreS = sPreS & Mid(sPre, nL, 1)
                End If
                
            Next nL
            sDes = sDes & Format(Val(sPreS), "0000000")
        End If
        Print #nFichDes, sDes
        
    
    Loop

    Close #nFichOrg
    Close #nFichDes
    
    Kill App.Path & "\bal.dat"
    Sleep (500)
    FileCopy App.Path & "\bal.tmp", App.Path & "\bal.dat"
    
End Sub


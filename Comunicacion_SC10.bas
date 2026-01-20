Attribute VB_Name = "Comunicacion_SC10"
'/////////////////////////////////////////////////////
'// Funciones de Comunicación ASCII exclusivas SC10 //
'/////////////////////////////////////////////////////
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private ComprobandoEstado As Boolean
Public FrmSC10_MostrarMensajes As Boolean
Public CuentaLineas As Long
Private Type tipoAuxCodBar
    Numero As Long
    numeroActual As Long
    seccion() As Integer
    enviadoRed() As Boolean
    enviadoLocal() As Boolean
    borrado As Boolean
End Type
Private Type TipoBarras
    Tipo As String
    Numero As Integer
    estado As String
End Type
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
    canCelar = 0
    CopiaLenta = False
    frmMonitor.MostrarDato "Obteniendo tickets de la red..."
    Do_Events
    If Dir(App.Path & "\copiavb.txt") <> "" Then CopiaLenta = True
    If Dir(App.Path & "\ticksc10", vbDirectory) = "" Then MkDir App.Path & "\ticksc10"
    On Error GoTo errorred
    frmControl.File1.Path = pathSC10
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
        '**********************
        ' comprueba formato tiquet
        
        If Len(ListaArch(BuclePath)) = 12 And _
        StrConv(Left(ListaArch(BuclePath), 1), vbUpperCase) = "T" And _
        IsNumeric(Mid(ListaArch(BuclePath), 2, 5)) Then
            If Not CopiaLenta Then FileCopy pathSC10 & "\" & ListaArch(BuclePath), _
            App.Path & "\ticksc10\" & ListaArch(BuclePath)
            
            If CopiaLenta Then
                Arch1 = FreeFile()
                On Error GoTo errorred
                Open pathSC10 & "\" & ListaArch(BuclePath) For Input As #Arch1
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
            
            Kill pathSC10 & "\" & ListaArch(BuclePath)
            
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 10 Then
                Do_Events
                CuentaEvents = 0
                If canCelar = 1 Then Exit For
            End If
        End If
    Next BuclePath
    If canCelar = 0 Then
        FrmSC10_BufTiquet = True
    Else
        FrmSC10_BufTiquet = False
    End If
    Exit Function
errorred:
    If Arch1 <> 0 Then Close Arch1
    If Arch2 <> 0 Then Close Arch2
    frmMonitor.MostrarDato CargaCadena(854) & " " & pathSC10
    frmMonitor.MostrarDato CargaCadena(544)
    frmEpelsa.LblSC10.Caption = CargaCadena(544)
    FrmSC10_BufTiquet = False
End Function
Public Sub FrmSC10_Coge_tiquets_AH(De_Que_Tipo As Integer, Optional NumMax As Integer)
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
    Dim Base As Database
    Dim RegArt As Recordset
    Dim Contador As Long
    Dim ListaArch() As String
    '*******************
    If Dir(App.Path & "\sc10tiq.txt") <> "" Then
        MyPath = App.Path & "\ticksc10"
        If Not FrmSC10_BufTiquet Then Exit Sub
    Else
        MyPath = pathSC10
    End If
    '**********************
    Cabecera.Ntiquet = 0
    Repetir = True
    frmMonitor.MostrarInfo CargaCadena(829)
    frmEpelsa.LblSC10.Caption = CargaCadena(829)
    Set Base = OpenDatabase(Base_General)
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
        If NumMax > 0 Then
            If Contador > NumMax Then Exit For
        End If
        If Len(ListaArch(BuclePath)) = 12 And _
        StrConv(Left(ListaArch(BuclePath), 1), vbUpperCase) = "T" And _
        IsNumeric(Mid(ListaArch(BuclePath), 2, 5)) Then
            If SC10_ProcesarTotalVenta Then
                Contador = Contador + 1
                Archivo = FreeFile()
                On Error Resume Next
                Open MyPath & "\" & ListaArch(BuclePath) For Input As Archivo
                Line Input #Archivo, MiRecibe
                On Error GoTo 0
                If MiRecibe = "" Then
                    frmMonitor.MostrarDato CargaCadena(544)
                    frmEpelsa.LblSC10.Caption = CargaCadena(544)
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
                    frmMonitor.MostrarDato CargaCadena(865) & " " & Cabecera.Ntiquet
                    frmEpelsa.LblSC10.Caption = CargaCadena(866) & " " & frmControl.File1.FileName
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
                            frmMonitor.MostrarDato CargaCadena(867) & " " & MiContador
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
                            '*************************
                            ' verifica si se ha modificado
                            ' el precio
                            '*************************
                            Set RegArt = Base.OpenRecordset _
                            ("select * from articulo where codigo=" & _
                            Linea(MiContador).CodigoArticulo & " and borrado=false")
                            If RegArt.EOF Then
                                Linea(MiContador).SehaModificado = False
                                Linea(MiContador).IVA = 0
                            Else
                                Linea(MiContador).IVA = RegArt!tipo_iva
                                If UsaEuro Then
                                    If (Linea(MiContador).precio / 100) <> RegArt!Euros Then
                                        Linea(MiContador).SehaModificado = True
                                    Else
                                        Linea(MiContador).SehaModificado = False
                                    End If
                                Else
                                    If (Linea(MiContador).precio / (10 ^ decimales)) <> RegArt!precio Then
                                        Linea(MiContador).SehaModificado = True
                                    Else
                                        Linea(MiContador).SehaModificado = False
                                    End If
                                End If
                            End If
                            RegArt.Close
                            '*************************
                            MiContador = MiContador + 1
                        Loop
                        '**************
                        ' si se han recibido
                        ' todas las líneas
                        ' lo procesa
                        '**************
                        Close #Archivo
                    
                    
                        If Cabecera.NLineas = MiContador - 1 Then
                            frmMonitor.MostrarDato CargaCadena(868) & " " & Cabecera.Ntiquet & " " & CargaCadena(869) & " --> OK"
                            Repetir = False
                            On Error Resume Next
                            Kill MyPath & "\" & ListaArch(BuclePath)
                            On Error GoTo 0
                            '*******************
                            ' Actualiza las tablas de totales
                            '*******************
                            Cabecera.Incluir_en_Base = True
                            Actualiza_Tabla_Cabecera Base, Cabecera, "SC10", Linea
                            If Cabecera.Incluir_en_Base = True And SC10_ProcesarTotalVenta Then
                                Actualiza_Tabla_Tickets Base, Cabecera, Linea, "SC10"
                                If Cabecera.Incluir_en_Base = True Then
                                    Actualiza_Tabla_Gtarti Base, Cabecera, Linea
                                    Actualiza_Tabla_Gtsecs Base, Cabecera, Linea, True
                                    Actualiza_Tabla_GtVend Base, Cabecera, Linea, True
                                    Actualiza_Tabla_Log Base, Cabecera, Linea
                                End If
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
                frmMonitor.MostrarDato CargaCadena(815) & " Ticket : " & ListaArch(BuclePath)
                Kill MyPath & "\" & ListaArch(BuclePath)
                On Error GoTo 0
                Close #Archivo
            End If
        End If
    Next BuclePath
    frmMonitor.MostrarDato CargaCadena(540)
    frmEpelsa.LblSC10.Caption = CargaCadena(540)
    CerrarBase Base
    Exit Sub
errorred:
    CerrarBase Base
    frmMonitor.MostrarDato CargaCadena(854) & " " & MyPath
    frmMonitor.MostrarDato CargaCadena(544)
    frmEpelsa.LblSC10.Caption = CargaCadena(544)
End Sub
Public Sub FrmSC10_envia_BaseMaestra(SoloModificaciones As Boolean)
    Dim LA_Fecha As Date
    Dim LA_Hora As Date
    Dim MyID As Integer
    MyID = 1
    If TipoC9 Then MyID = 2
    On Error Resume Next
    If Dir(MiruTa & "\orden.dat") <> "" Then Kill MiruTa & "\orden.dat"
    On Error GoTo 0
    Dim Base As Database
    If Not SC10Rs232 And Not TipoC9 Then
        CuentaLineas = 1
    End If
    If MyID = 1 Then
        If canCelar = 0 Then
            'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 9)
            Do_Events
            CG_Envia_Seccion MyID, SoloModificaciones
            If canCelar = 1 Then Exit Sub
        End If
    End If
    If MyID = 1 And (Not SC10Rs232) And (Not TipoC9) Then
        If canCelar = 0 Then
            'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 677)
            Do_Events
            CG_EnviaSubSec 1, SoloModificaciones
            If canCelar = 1 Then Exit Sub
        End If
    End If
    If canCelar = 0 And (Not SC10Rs232) And (Not TipoC9) Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 678)
        Do_Events
        CG_EnviaFamilia 1, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 11)
        Do_Events
        CG_envia_equipos MyID, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 12)
        Do_Events
        CG_EnviaArticulo 1, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 14)
        Do_Events
        CG_EnviaPublicidad MyID, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 21)
        Do_Events
        CG_EnviaCabecerasLeyendas MyID, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 23)
        Do_Events
        CG_EnviaCodigosBarras MyID, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 And (Not TipoC9) Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 13)
        Do_Events
        CG_EnviaVendedor 1, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
    If canCelar = 0 Then
        'frmMonitor.MostrarDato CargaCadena( 830) & " " & CargaCadena( 22)
        Do_Events
        CG_EnviaTeclas 1, SoloModificaciones
        If canCelar = 1 Then Exit Sub
    End If
End Sub

Public Sub FrmSC10_SC10_ComprobarEstado()
    Dim OrdenDat As Integer
    Dim Longitud As Long
    Dim Bucle As Integer
    Dim Funcionando As Boolean
    Dim Buffer As String
    Dim EspMax As Long
    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = False
    Funcionando = False
    lCogeTiquet = False
    frmMonitor.Label1.Caption = CargaCadena(30)  '"Comprobando Estado..."
    ComprobandoEstado = True
    frmEpelsa.Enabled = False
    On Error GoTo errorred
    
    If Dir(pathSC10 & "\resul.dat") <> "" Then
        Longitud = FileLen(pathSC10 & "\resul.dat")
        EspMax = 0
        frmMonitor.MostrarDato CargaCadena(30)  '"Comprobando Estado..."
        Do While EspMax < 40
            Do_Events
            Sleep (100)
            EspMax = EspMax + 1
        Loop
        If Longitud <> FileLen(pathSC10 & "\resul.dat") Then
            frmMonitor.MostrarDato CargaCadena(860)  '"El concentrador está ocupado procesando Información"
            lCogeTiquet = True
            HabilitarSalir
            On Error GoTo 0
            Exit Sub
        End If
    End If
    '******
    ' Intenta Enviar una orden
    '******
    If Dir(pathSC10 & "\resul.dat") <> "" Then
        Kill pathSC10 & "\resul.dat"
    End If
    
    OrdenDat = FreeFile()
    Open pathSC10 & "\orden.dat" For Output As OrdenDat
    If Dir(App.Path & "\sercom.txt") = "" Then
        Print #OrdenDat, "000124"
        Print #OrdenDat, "000224"
        Print #OrdenDat, "0003260|0|"
        Print #OrdenDat, "000429"
    Else
        Print #OrdenDat, "24"
        Print #OrdenDat, "24"
        Print #OrdenDat, "260|0|"
        Print #OrdenDat, "29"
    
    End If
    Close #OrdenDat
    frmMonitor.MostrarDato CargaCadena(881)
    
    If Dir(App.Path & "\sercom.txt") = "" Then
    For Bucle = 1 To 7
        frmMonitor.MostrarDato CargaCadena(882) & " " & Bucle
        EspMax = 0
        Do While EspMax < 60
           Do_Events
           Sleep (100)
           EspMax = EspMax + 1
        Loop
        If Dir(pathSC10 & "\resul.dat") <> "" Then
            Longitud = FileLen(pathSC10 & "\resul.dat")
            Funcionando = True
            Exit For
        End If
    Next Bucle
    Else
        Funcionando = True
        frmMonitor.MostrarDato CargaCadena(886)  '"El concentrador admite Ordenes de Red ..."
        Exit Sub
    End If
    If Funcionando = False Then
        frmMonitor.MostrarDato CargaCadena(883)  '"El concentrador no Procesa Ordenes"
        frmMonitor.MostrarDato CargaCadena(884)  '"Puede que el programa del concentrador no "
        frmMonitor.MostrarDato CargaCadena(885)  '"esté funcionando o no admita órdenes de red"
        HabilitarSalir
        On Error GoTo 0
        Exit Sub
    End If
    If Funcionando = True Then
        frmMonitor.MostrarDato CargaCadena(886)  '"El concentrador admite Ordenes de Red ..."
        EspMax = 0
        Do While EspMax < 40
             Do_Events
             Sleep (100)
             EspMax = EspMax + 1
        Loop
        OrdenDat = FreeFile()
        Open pathSC10 & "\resul.dat" For Input As OrdenDat
        Input #OrdenDat, Buffer
        If Mid(Buffer, 6, 1) = "1" Then
            frmMonitor.MostrarDato CargaCadena(879)  '"Red de Balanzas Conectada"
        Else
            frmMonitor.MostrarDato CargaCadena(880)  '"Red de Balanzas Desconectada"
        End If
        Input #OrdenDat, Buffer
        Input #OrdenDat, Buffer
        Input #OrdenDat, Buffer
        frmMonitor.MostrarDato CargaCadena(434) & " : " & Mid(Buffer, 15, 2) & "/" & Mid(Buffer, 17, 2) & "/" & Mid(Buffer, 19, 2)
        frmMonitor.MostrarDato CargaCadena(435) & " : " & Mid(Buffer, 6, 2) & ":" & Mid(Buffer, 8, 2)
        Close #OrdenDat
        HabilitarSalir
        On Error GoTo 0
    End If
    lCogeTiquet = True
    Exit Sub
errorred:
    On Error GoTo 0
    lCogeTiquet = True
    frmMonitor.MostrarDato CargaCadena(266) & " -->NOK"
    frmMonitor.MostrarDato CargaCadena(887) & " ( " & pathSC10 & " )"
    frmMonitor.MostrarDato CargaCadena(888)
    HabilitarSalir
    
End Sub
Public Function FrmSC10_EnviaDirectorioRedSC10() As Boolean
    '
    ' Esta función comprueba que el fichero resul.dat
    ' No está variando, y si es así copia el orden.dat
    '
    Dim Resul As String
    Dim tam1 As Long
    Dim tam2 As Long
    Dim Cambio As Boolean
    Dim Bucle As Integer
    Dim EspMax As Integer
    Dim conxf As Integer
    Dim sErr As String
    frmControl.Balanza.Reset
    frmControl.EstadoBal = 1000
    frmMonitor.Command1.Enabled = False
    frmMonitor.cmdmonitorstop.Enabled = False
    Cambio = False
    frmMonitor.MostrarDato CargaCadena(983)
    frmMonitor.MostrarInfo CargaCadena(889)
    frmMonitor.MostrarDato "...................................................."
    Do_Events
    frmEpelsa.LblSC10.Caption = CargaCadena(889)  '"Enviando datos a Unidad de Red"
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v144
    conxf = FreeFile()
    On Error Resume Next
    Open pathSC10 & "\conx.sys" For Binary Access Read Lock Read As #conxf
    If Err.Number <> 0 Then sErr = Err.Description
    On Error GoTo 0
    On Error GoTo NoAbierto
    If sErr <> "" Then Err.Raise 1000
    Close #conxf
    Kill pathSC10 & "\conx.sys"
    '//////////////////////////////////////////////////////////cas.v144
    Resul = Dir(pathSC10 & "\resul.dat")
    If Resul <> "" Then
        tam1 = FileLen(pathSC10 & "\resul.dat")
        tam2 = tam1
        If FileLen(MiruTa & "\orden.dat") <> 0 Then
            Do While EspMax < 30
               Do_Events
               Sleep (100)
               EspMax = EspMax + 1
            Loop
        End If
        tam2 = FileLen(pathSC10 & "\resul.dat")
        If tam1 <> tam2 Then
            Cambio = True
        Else
            Kill (pathSC10 & "\resul.dat")
            Kill (pathSC10 & "\orden.dat")
        End If
    End If
    
    If Dir(App.Path & "\sercom.txt") <> "" Then
        Sleep (3000)
        Cambio = False
    End If
    
    If Cambio Then
        On Error GoTo 0
        FrmSC10_EnviaDirectorioRedSC10 = False
        If FrmSC10_MostrarMensajes Then
            If Not UsaGamaAlta Then MsgBox CargaCadena(859)
        End If
        frmMonitor.MostrarDato CargaCadena(860)
        frmEpelsa.LblSC10.Caption = Format(Now, "hh:mm:ss") & CargaCadena(860)
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
        Kill MiruTa & "\orden.dat"
        HabilitarSalir
        Exit Function
    End If
    On Error GoTo NoAbierto
    Resul = Dir(pathSC10 & "\orden.dat")
    If Resul <> "" Then Kill pathSC10 & "\orden.dat"
    If FileLen(MiruTa & "\orden.dat") > 0 Then
        On Error GoTo NoAbierto
        HazCopia MiruTa & "\orden.dat", pathSC10 & "\orden.dat"
        If Dir(pathSC10 & "\orden.dat") <> "" Then
            frmMonitor.MostrarDato CargaCadena(890) & " --> OK"
            frmMonitor.MostrarDato ""
            frmMonitor.MostrarDato ""
            frmMonitor.Label1.Caption = "SC10 --> OK"
            For Bucle = 1 To 3
                frmMonitor.Label1.BackColor = vbGreen
                Do_Events
                Sleep (500)
                frmMonitor.Label1.BackColor = vbWhite
                Do_Events
                Sleep (500)
            Next Bucle
            frmEpelsa.LblSC10.Caption = Format(Now, "HH:MM:SS") & CargaCadena(890) & " --> OK"
            frmMonitor.Label1.Caption = CargaCadena(991) '"FICHERO CORRECTAMENTE TRANSMITIDO"
        Else
            frmMonitor.Label1.Caption = CargaCadena(992)
        End If
        HabilitarSalir
    End If
    On Error GoTo 0
    Kill MiruTa & "\orden.dat"
    frmMonitor.cmdmonitorstop.Enabled = False
    HabilitarSalir
    FrmSC10_EnviaDirectorioRedSC10 = True
    Exit Function
NoAbierto:
    FrmSC10_EnviaDirectorioRedSC10 = False
    frmMonitor.MostrarDato CargaCadena(544)
    If sErr <> "" Then
        frmMonitor.MostrarDato sErr  '"No se ha encontrado la ruta de acceso"
        sErr = ""
    End If
    frmEpelsa.LblSC10.Caption = Format(Now, "hh:mm:ss") & CargaCadena(544)
    If FrmSC10_MostrarMensajes Then
        If Not UsaGamaAlta Then MsgBox CargaCadena(854) & " " & pathSC10
    End If
    frmMonitor.Label1.Caption = CargaCadena(992) '"ERROR AL TRANSMITIR EL FICHERO"
    frmMonitor.cmdmonitorstop.Enabled = False
    HabilitarSalir
    Err.Clear
    On Error GoTo 0
    Exit Function
End Function
Private Sub HazCopia(Origen As String, Destino As String)
    Dim MiProceso As Long
    Dim MiResultado As Long
    Dim MiAppPath As String
    Dim sBuf As String * 260
    Dim sBuf2 As String * 260
    Dim i As Long
    Dim NombreCorto As String
    i = GetShortPathName(pathSC10, sBuf, Len(sBuf))
    If i = 0 Then Exit Sub
    NombreCorto = Left$(sBuf, i)
    '
    i = GetShortPathName(App.Path, sBuf2, Len(sBuf2))
    If i = 0 Then Exit Sub
    MiAppPath = Left$(sBuf2, i)
    '
    ChDir App.Path
    MiResultado = STILL_ACTIVE
    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Environ("COMSPEC") & " /C " & MiAppPath & "\xcopy.exe orden.dat " & NombreCorto & " /Y /D", vbHide))
    Do_Events
    If MiProceso <> 0 Then
        Do While MiResultado = STILL_ACTIVE
            Do_Events
            GetExitCodeProcess MiProceso, MiResultado
        Loop
    End If
    
    'Dim Arch1 As Integer
    'Dim Arch2 As Integer
    'Dim bUF As String
    'Arch1 = FreeFile()
    'Open Origen For Input As #Arch1
    'Arch2 = FreeFile()
    'Open Destino For Output Lock Read Write As #Arch2
    'Do Until EOF(Arch1)
    '    Line Input #Arch1, bUF
    '    Print #Arch2, bUF
    'Loop
    'Close #Arch1
    'Close #Arch2
End Sub
Public Function SC10_CheckOrden() As Long
    
    
    Dim Arch As Integer
    Dim MyCont As Long
    Dim Buf As String
    MyCont = 1
    frmControl.Balanza.Reset
    If Dir(App.Path & "\orden.dat") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\orden.dat" For Input As #Arch
        Do Until EOF(Arch)
            Line Input #Arch, Buf
            MyCont = MyCont + 1
            If MyCont > 9999 Then MyCont = 1
        Loop
        Close #Arch
    End If
    SC10_CheckOrden = MyCont
End Function

    
Public Sub Marcar_SC10Path(Que_base As Integer)
    '//////////////////////////////////////////////////////////////////////////////
    '// Función para marcar los datos que se han conseguido enviar a SC10        //
    '// a través de Path de Red                                                  //
    '//////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////
    ' Que_Base :
    ' 0 --> Base Maestra
    ' 1 --> Secciones
    ' 2 --> Subsecciones
    ' 3 --> Familias
    ' 4 --> Artículos
    ' 5 --> Publicidad
    ' 6 --> Cab./Ley.
    ' 7 --> Cod. Bar.
    ' 8 --> Vendedores
    ' 9 --> Teclas
    ' 10 --> Equipos
    '/////////////////////////////////////////////
    Dim Base As Database
    Dim BaseDbf As Database
    Dim Registro As Recordset
    Set Base = OpenDatabase(Base_General)
    If Que_base = 0 Or Que_base = 1 Then
        Set Registro = Base.OpenRecordset("select * from seccion")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_secSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_sec = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 2 Then
        Set Registro = Base.OpenRecordset("select * from subsec")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_sub = "*"
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 3 Then
        Set Registro = Base.OpenRecordset("select * from familia")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_fam = "*"
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 4 Then
        Set Registro = Base.OpenRecordset("select * from articulo")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_pluSC10 = "*"
                !tran_textosc10 = "*"
                !tran_tx1SC10 = "*"
                !tran_cbsc10 = "*"
                If !borrado = True And Not UsaGamaAlta Then
                    !tran_pluSC10 = "*"
                    !tran_textosc10 = "*"
                    !tran_tx1SC10 = "*"
                    !tran_cbsc10 = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 5 Then
        Set Registro = Base.OpenRecordset("select * from publi")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_pubSC10 = "*"
                !tran_ofeSC10 = "*"
                !tran_actSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_pub = "*"
                    !tran_ofe = "*"
                    !tran_act = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
        
    If Que_base = 0 Or Que_base = 6 Then
        Set Registro = Base.OpenRecordset("select * from cabley")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_cabSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_cab = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 7 Then
        Set Registro = Base.OpenRecordset("select * from codbar")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_barSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_bar = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 8 Then
        Set Registro = Base.OpenRecordset("select * from vendedor")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_vendSC10 = "*"
                !tran_asignSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_vend = "*"
                    !tran_asign = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
    End If
    
    If Que_base = 0 Or Que_base = 9 Then
        Set Registro = Base.OpenRecordset("select * from teclas")
        With Registro
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                Edit_Record Registro
                !tran_tecSC10 = "*"
                If Not UsaGamaAlta And !borrado = True Then
                    !tran_tec = "*"
                End If
                .Update
                .Movenext
            Loop
        End With
        Registro.Close
        
    End If
    
    If Que_base = 0 Or Que_base = 10 Then
        Set Registro = Base.OpenRecordset("select * from equipos")
        ' Marca como enviados
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do While Not .EOF
                    Edit_Record Registro
                    !tran_eqp = "*"
                    .Update
                    .Movenext
                Loop
                .MoveFirst
            End If
           ' Borra los marcados para borrar
           If Not .EOF Then .MoveFirst
            Do While Not .EOF
                If !borrado = True Then
                    Delete_Record Registro
                    .Movenext
                Else
                    .Movenext
                End If
            Loop
        End With
    End If
    CerrarBase Base
End Sub

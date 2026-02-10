VERSION 5.00
Begin VB.Form FrmHlink 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   6960
   ClientTop       =   2160
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TmrTimeOut 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   1680
   End
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "FrmHlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyProc As Long
Private MyResul As Long
Private Timeout As Boolean
Private MyPath As String
Private MyOrder As String
Private Const FileIn = "\plugincall"
Private Const FileOut = "\mastercall"
Private Const FileEnd = "\masterend"
Private Sub Form_Load()
    If Dir(App.Path & "\dirtmptouch", vbDirectory) = "" Then
        MkDir App.Path & "\dirtmptouch"
    End If
    MyPath = App.Path & "\dirtmptouch"
    If Dir(Miruta & "\hydplugin.exe") <> "" Then
        DataPlugin = TestMessage()
    Else
        DataPlugin.mode = ""
        DataPlugin.type = ""
    End If
End Sub
Private Sub ClearPipes()
    If Dir(MyPath & FileIn) <> "" Then
        Kill MyPath & FileIn
    End If
    If Dir(MyPath & FileOut) <> "" Then
        Kill MyPath & FileOut
    End If
    If Dir(MyPath & FileEnd) <> "" Then
        Kill MyPath & FileEnd
    End If
    If Dir(MyPath & "\stopnowcall") <> "" Then
        Kill MyPath & "\stopnowcall"
    End If
End Sub
Private Function TestMessage() As TPlugin
    Dim MyDesc As Integer
    Dim Buffer As String
    Dim Piece As String
    Dim MyData As TPlugin
    Dim Cont As Integer
    Dim Zona As Integer
    ClearPipes
    Timeout = False
    FrmHlink.TmrTimeOut.Interval = 300
    FrmHlink.TmrTimeOut.Enabled = True
    FrmHlink.TmrTimeOut.Enabled = True
    WPlugin "TEST"
    ChDir (Miruta)
    Shell (Miruta & "\hydplugin.exe")
    Do Until (Dir(MyPath & FileIn) <> "") Or Timeout
        Do_Events
    Loop
    FrmHlink.TmrTimeOut.Enabled = False
    If Dir(MyPath & FileIn) <> "" Then
        MyDesc = FreeFile()
        Open MyPath & FileIn For Input As MyDesc
        Line Input #MyDesc, Buffer
        Close MyDesc
        Kill MyPath & FileIn
    End If
    If Buffer <> "" Then
        If Right(Buffer, 1) <> "#" Then Buffer = Buffer & "#"
        Zona = 1
        Cont = 1
        Do Until Cont > Len(Buffer)
            Piece = ""
            Do Until Mid(Buffer, Cont, 1) = "#"
                Piece = Piece & Mid(Buffer, Cont, 1)
                Cont = Cont + 1
            Loop
            Select Case Zona
            Case 1
                MyData.mode = Piece
            Case 2
                MyData.type = Piece
            Case 3
                MyData.NombreMenu = Piece
            Case Else
                If Zona Mod 2 = 0 Then

                    MyData.NombreEntrada(MyData.Nentradas) = Piece
                    MyData.Nentradas = MyData.Nentradas + 1
                Else
                    MyData.accion(MyData.Nentradas - 1) = Piece
                End If
            End Select

            Cont = Cont + 1
            Zona = Zona + 1
        Loop
        If MyData.mode = "X" Then
            PlugOculta = True
        Else
            PlugOculta = False
        End If
        TestMessage = MyData
    Else
        MyData.mode = ""
        MyData.type = ""
        TestMessage = MyData
    End If
End Function
Public Sub STOPALL()
    Dim MyDesc As Integer
    MyDesc = FreeFile()
    Open (MyPath & FileEnd) For Output As #MyDesc
    Close #MyDesc
End Sub

Public Sub Start(Status As Integer)
    ClearPipes
    If Dir(Miruta & "\hydplugin.exe") <> "" Then
        WPlugin "START#" & Format(PlugStatus, "000")
        MyResul = STILL_ACTIVE
        ChDir (Miruta)
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Miruta & "\hydplugin.exe", vbNormalFocus))
        FrmHlink.Tmr1.Interval = 100
        FrmHlink.Tmr1.Enabled = True
        PlugActivo = True
    End If
    Timeout = False
    FrmHlink.TmrTimeOut.Interval = 7000
    FrmHlink.TmrTimeOut.Enabled = True
    Do Until Dir(MyPath & FileOut) = "" Or Timeout
        Do_Events
    Loop
End Sub
Private Sub Tmr1_Timer()
    Dim MyDesc As Integer
    Dim Buffer As String
    Dim TryAgain As Integer
    Buffer = ""
    '*******************************
    ' checks message
    '*******************************

    If Dir(MyPath & FileIn) <> "" Then
        TryAgain = 0
        Do While TryAgain < 4
            On Error Resume Next
            MyDesc = FreeFile()
            Open MyPath & FileIn For Input As MyDesc
            If Err.Number <> 0 Then
                TryAgain = TryAgain + 1
                Sleep (100)
                Do_Events
            Else
                TryAgain = 6
            End If
        Loop
        If TryAgain >= 6 Then
            Line Input #MyDesc, Buffer
            Close MyDesc
            Kill MyPath & FileIn
        End If
    End If
    If Dir(MyPath & "\stopnowcall") <> "" Then
        canCelar = 1
        On Error Resume Next
        Kill MyPath & "\stopnowcall"
        On Error GoTo 0
    End If
    If Buffer <> "" Then CheckOrder (Buffer)
    GetExitCodeProcess MyProc, MyResul
    If MyResul <> STILL_ACTIVE Then
        lCogeTiquet = True
        PlugActivo = False
        Tmr1.Enabled = False
        PlugStatus = 0
    End If
End Sub
Public Function RPlugin() As String


    RPlugin = MyOrder
    MyOrder = ""

End Function
Public Function WPlugin(dato As String)
    Dim MyDesc As Integer
    Dim Buffer As String
    Dim TryAgain As Integer
    TryAgain = 0
    Do While TryAgain < 5
        On Error Resume Next
        MyDesc = FreeFile()
        Open MyPath & FileOut & ".tmp" For Output As MyDesc
        If Err.Number <> 0 Then
            Sleep (100)
            Do_Events
            TryAgain = TryAgain + 1
        Else
            TryAgain = 6
        End If
        On Error GoTo 0
    Loop
    If TryAgain >= 6 Then
        Print #MyDesc, dato
        Close MyDesc
        TryAgain = 0
        Do While TryAgain < 5
            On Error Resume Next
            Name MyPath & FileOut & ".tmp" As MyPath & FileOut
            If Err.Number <> 0 Then
                Sleep (100)
                Do_Events
                TryAgain = TryAgain + 1
            Else
                TryAgain = 6
            End If
            On Error GoTo 0
        Loop
    End If
    On Error GoTo 0
End Function
Private Sub TmrTimeOut_Timer()
    Timeout = True
End Sub
Private Sub CheckOrder(orden As String)
    Dim Valor As String
    Dim Error As String
    Error = "001"
    Select Case orden
    Case ""
    Case Else
        If Len(orden) > 4 Then
            If InStr(1, orden, "#") = 4 Then
                If IsNumeric(left(orden, 3)) Then
                    PlugStatus = left(orden, 3)
                    Valor = Mid(orden, 5)
                    Error = RealizaAccion(Valor)
                End If
            End If
        End If
    End Select
    WPlugin Error
End Sub
Public Function RealizaAccion(Valor As String) As String
    Dim accion As Integer
    Dim Adicional As String
    Dim Salida As String
    Dim V1 As Integer
    'c2f/caspiunza **
    Dim ErrorEnvio As Integer
    Dim ErrorRecibe As Integer
    Dim nContErr As Integer

    ErrorEnvio = 0
    ErrorRecibe = 0

    If DebugActivo Then CadenadeLog "Recibida Orden del plugin --> " & Valor
    Salida = "000"
    If IsNumeric(left(Valor, 2)) Then
        accion = Val(left(Valor, 2))
        If Len(Valor) > 2 Then
            Adicional = Mid(Valor, 3)
        Else
            Adicional = ""
        End If
    Else
        accion = 1000
    End If
    Select Case accion
    Case 1
        ErrorEnvio = 1
        frmEpelsa.mnuBaseMaestra_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 2
        ErrorEnvio = 1

        If Dir(App.Path & "\gigante.txt") <> "" Then
            nContErr = 0
            Do While nContErr < 5

                frmEpelsa.mnuEnvio_modificaciones_Click
                If canCelar = 0 Then
                    nContErr = 5
                Else
                    nContErr = nContErr + 1
                End If
            Loop
            salir_programa
        Else
            frmEpelsa.mnuEnvio_modificaciones_Click
            If HayMulti Then
                If OrdenMulti <> "02" Then    '1.8.7
                    Do Until Multi_Salir
                        Do_Events
                    Loop
                Else
                    salir_programa
                End If
            End If
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 3
        ErrorEnvio = 1
        'frmEpelsa.mnuEnviarSecciones_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 4
        ErrorEnvio = 1
        'frmEpelsa.mnuenviarSubsecciones_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 5
        ErrorEnvio = 1
        'frmEpelsa.mnuenviarFamilias_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 6
        ErrorEnvio = 1
        'frmEpelsa.mnu_enviarEquipos_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 7
        ErrorEnvio = 1
        frmEpelsa.mnuEnviarPaises_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 8
        ErrorEnvio = 1
        frmEpelsa.mnuEnviarFichasVacuno_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 9
        ErrorEnvio = 1
        frmEpelsa.mnu_envio_articulos_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 10
        ErrorEnvio = 1
        'frmEpelsa.mnu_envio_publicidad_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 11
        ErrorEnvio = 1
        'frmEpelsa.mnu_envio_cabley_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 12
        ErrorEnvio = 1
        frmEpelsa.mnu_envio_codigobar_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 13
        ErrorEnvio = 1
        frmEpelsa.mnu_envio_vendedor_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 14
        ErrorEnvio = 1
        'frmEpelsa.mnu_envio_teclas_rapidas_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 15
        ErrorRecibe = 1
        MostrarMonitor
        lCogeTiquet = False
        FrmExportar.AN_RecogerTotales True
        lCogeTiquet = True
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorRecibe = 2

    Case 16
        ErrorRecibe = 1
        frmEpelsa.mnuTotalFichaVAcuno_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorRecibe = 2

    Case 17
        If Adicional <> "" Or Dir(App.Path & "\multiimp.ord") <> "" Then
            Select Case RealizarImportacion(Adicional)
            Case 0
                Salida = "000"
            Case 1
                Salida = "002"
            Case 2
                Salida = "003"
            Case 10
                Salida = "004"
            End Select
        Else
            Salida = "001"
        End If
        If Dir(App.Path & "\multiimp.ord") <> "" Then
            Kill App.Path & "\multiimp.ord"
            Multi_Salir = True
        End If

        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
    Case 18
        frmControl.ini_import

        frmControl.ControlImport1.FamFijas = FamiliasFijas
        frmControl.ControlImport1.SubFijas = SubSeccionesFijas
        frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
        frmControl.ControlImport1.PathBase = Base_General
        frmControl.ControlImport1.Euros = True
        frmControl.ControlImport1.PathFile = Miruta
        frmControl.ControlImport1.PathNTQ = Miruta
        frmControl.ControlImport1.ActivarLog = DebugActivo
        frmControl.ControlImport1.MostrarFormulario = True
        frmControl.ControlImport1.Fichero = "fichas.dat"
        frmControl.ControlImport1.FamVacuno = FamiliaEspecial
        frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
        frmControl.ControlImport1.MuestraMensajes = False
        frmControl.ControlImport1.FactorEuro = 1
        frmControl.ControlImport1.Importar_Background "fichas.dat"
        Do Until frmControl.ControlImport1.Exito <> 4
            Do_Events
        Loop
        Unload frmControl
        If Dir(App.Path & "\asign.tst") <> "" Then
            Call AsignTec
        End If
        Corregir_Nulos

    Case 19
        ErrorRecibe = 1
        MostrarMonitor
        lCogeTiquet = False
        FrmExportar.AN_RecogerTotales False
        lCogeTiquet = True
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorRecibe = 2

    Case 20
        Espera_Mostrar FrmSeccionyMaquina
    Case 21
        Espera_Mostrar FrmArticulo2
    Case 22
        Espera_Mostrar FrmVendedor2
    Case 23
        'Espera_Mostrar FrmPublicidad2
    Case 24
        'Espera_Mostrar FrmCabley2
    Case 25
        Espera_Mostrar FrmCodBar2

    Case 26
        'Espera_Mostrar FrmTeclas2
    Case 27
        'Espera_Mostrar FrmConserv386
    Case 28
        'Espera_Mostrar FrmConservSC10
    Case 29
        Espera_Mostrar FrmPaises
    Case 30
        frmFichasVacuno.Show
        Do While frmFichasVacuno.Visible = True
            Do_Events
        Loop
        Unload frmFichasVacuno
    Case 31
        Espera_Mostrar FrmCambiosPrecio
    Case 32
        Espera_Mostrar FrmIVA
    Case 33
        ErrorEnvio = 1
        frmEpelsa.mnu_Envio_IVA_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 34
        RealizarImportacion (Adicional)

        frmEpelsa.mnuEnvio_modificaciones_Click
        If Dir(App.Path & "\multiimp.ord") <> "" Then
            Kill App.Path & "\multiimp.ord"
            Multi_Salir = True
        End If
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
    Case 35
        'AN_FechaHora Format(Now, "dd/mm/yy"), Format(Now, "hh:mm")
        Unload frmMonitor
    Case 36
        ErrorRecibe = 1
        'FrmConsultaArt.Show
        'FrmConsultaArt.Command1(0).Caption = CargaCadena(288)
        'FrmConsultaArt.CmbFiles.TexT = Mid(Adicional, 2)
        'FrmConsultaArt.ConsultaLosArticulos False
        'If canCelar = 0 Then FrmConsultaArt.CmdExportar_Click
        '**************
        ' sabeco : copia a carpeta CP
        'If Sabeco Then
        '    ChDir MiruTa
        '    Shell "findia.exe /CP"
        'End If
        '**************
        If canCelar = 1 Then ErrorRecibe = 2

    Case 37
        ErrorEnvio = 1
        If IsNumeric(Adicional) Then
            MostrarMonitor
            AN_BorrarSeccion (Adicional)
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 50
        Salida = "000" & pathred
    Case 51
        Salida = "000" & PathOrden
    Case 52
        If pathred <> "" Then
            V1 = FrmResturarBackup.Realizar_Backup(False)
            Select Case V1
            Case 0
                Salida = "000"
            Case 1
                Salida = "006"
            Case 2
                Salida = "007"
            Case 3
                Salida = "008"
            End Select
        Else
            Salida = "006"
        End If
    Case 53
        'If SoyMaster Or SoyReserva Then MS_EnviaBaseRemota True
    Case 54
        Select Case ExportaTotalesDat(False, Adicional)
        Case True
            Salida = "000"
        Case Else
            Salida = "009"
        End Select
    Case 55
        If Len(Adicional) = 3 Then
            If IsNumeric(left(Adicional, 1)) Then
                Select Case Val(left(Adicional, 1))
                Case 0
                Case 1
                    ' por sección
                Case 2
                    ' por balanza
                End Select
            End If
        End If
    Case 56
        Espera_Mostrar FrmTextoLibre
    Case 57
        Espera_Mostrar FrmClientes
    Case 58
        'frmEpelsa.mnu_envio_textos_Click
        'If HayMulti Then
        '    Do Until Multi_Salir
        '        Do_Events
        '    Loop
        'End If
        ErrorEnvio = 1
    Case 59
        ErrorEnvio = 1
        frmEpelsa.mnuEnvioCLientes_Click
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorEnvio = 2

    Case 60
        ' mostrar listados
        Salida = "000"
        If Len(Adicional) >= 2 Then
            Select Case left(Adicional, 2)
            Case "01"
                Aux_Ejecuta "rpmostrador"
            Case "02"
                Aux_Ejecuta "rpequipos"
            Case "03"
                Aux_Ejecuta "rpvendor"
            Case "04"
                Aux_Ejecuta "rpart"
            Case "05"
                If Len(Adicional) = 5 Then
                    Aux_Ejecuta "rpartsec", Val(Mid(Adicional, 3))
                Else
                    Salida = "001"
                End If
            Case "06"
                Aux_Ejecuta "rpcprecios"
            Case Else
                Salida = "001"
            End Select
        Else
            Salida = "001"
        End If
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
    Case 61
        Espera_Mostrar FrmConfiguracion2
    Case 62
        frmEpelsa.configurator_Click
    Case 63
        Espera_Mostrar FrmExportConf
    Case 64
        Espera_Mostrar FrmLineaComandos
    Case 65
        frmEpelsa.mnuseuridad_Click
    Case 66
        'Espera_Mostrar FrmBaseVacuno
    Case 67
        Espera_Mostrar FrmConsultaTiquets
    Case 68
        MostrarMonitor
        ErrorRecibe = 1
        lCogeTiquet = False
        FrmExportar.AN_RecogerTotales True
        If EstadoBal < 1000 Then AN_BaseMaestra True
        lCogeTiquet = True
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorRecibe = 2

    Case 69
        ErrorRecibe = 1
        lCogeTiquet = False
        'FrmExportar.AN_PedidoInventario
        lCogeTiquet = True
        If HayMulti Then
            Do Until Multi_Salir
                Do_Events
            Loop
        End If
        If canCelar = 1 Then ErrorRecibe = 2

    Case 70
        'Espera_Mostrar FrmTipoTARA
    Case 71
        'Espera_Mostrar FrmTextoECO
    Case 72
        Espera_Mostrar FrmFamilias    ' 'FrmFamilias

        'cazani exportar csv
    Case 80
        'ExportarFichero(ElBoton As Integer, Optional MiArchivo As String, Optional MiFecha As Date, Optional Continua As Boolean)
        frmWAIT.Label1 = "Exportando:" & "cabcsv.csv"
        frmWAIT.Label2 = Base_General
        Load frmWAIT
        frmWAIT.Visible = True
        frmWAIT.Refresh
        Do_Events
        FrmExportar.ExportarFichero 0, "cabcsv.csv"

    Case 1000
        Salida = "001"
    End Select
    If (accion = 19) Or (accion > 0 And accion < 17) Then

        If canCelar <> 0 Then
            Salida = "002"
        Else
            Salida = "000"
        End If
        Unload frmMonitor
    End If

    'c2f/caspiunza**
    'llamada a resumentxt
    If ErrorEnvio <> 0 Then
        Call ResumenTXT(10000 + ErrorEnvio)
    End If
    If ErrorRecibe <> 0 Then
        Call ResumenTXT(ErrorRecibe)
    End If
    ''''''''''''''''''''''

    RealizaAccion = Salida
End Function

Public Sub Espera_Mostrar(Formulario As Form)
    Formulario.Show
    Do While Formulario.Visible = True
        Do_Events
    Loop
    Unload Formulario

End Sub

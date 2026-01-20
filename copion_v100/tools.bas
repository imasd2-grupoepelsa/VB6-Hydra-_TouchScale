Attribute VB_Name = "tools"
Option Explicit
'''''''''''''''
'''''''''''''''
Public Function LeerParametrosInicio()
    Dim Fichero As Integer
    Dim Buffer As String
    Dim Contador As Long
    Dim Reescribe As Boolean
    Dim ConfiguracionINI(150) As String
    '
    ' Meter valores por defecto si no existe mng.ini
    '
    Reescribe = False
    'If Dir(MiruTa & "\hydratouch.ini") = "" Then
    '    Los_parametros_Defecto
    '    If AhorraMas Then Los_Parametros_Defecto_Ahorramas
    '    If Sabeco Then Los_Parametros_Defecto_Sabeco
    '    If Alcampo Then Los_parametros_Defecto_Alcampo
    '    If AHOLD Then Los_parametros_Defecto_Ahold
    '    CrearPassSec
    'Else
    '
    ' Leer ntq.ini
    '
        On Error Resume Next
        Fichero = FreeFile
        Open HYDRA_INI For Input As Fichero
        If Err.Number <> 0 Then
            MsgBox "Error accediendo a hydratouch.ini", vbCritical
            End
        End If
        On Error GoTo 0
        Contador = 0
        Do Until EOF(Fichero)
            Input #Fichero, ConfiguracionINI(Contador), Buffer
            Contador = Contador + 1
        Loop
        Close Fichero
        '*************
        ' sec
        '*************
        'Buffer = DesEncripta(ConfiguracionINI(0))
        'If Not IsNumeric(Left(Buffer, 8)) Then
        '    CadenadeLog CargaCadena(851)
        '    MsgBox CargaCadena(851), vbCritical
        '    End
        'Else
        '    If (Val(Left(Buffer, 4)) + 3) = Val(Mid(Buffer, 5, 4)) Then
        '        HaySeguridad = False
        '    Else
        '        If (Val(Left(Buffer, 4)) + 1) = Val(Mid(Buffer, 5, 4)) Then
        '            HaySeguridad = True
        '        Else
        '            MsgBox CargaCadena(851), vbCritical
        '            End
        '        End If
        '    End If
        'End If
        ' Idioma
        If Not IsNumeric(ConfiguracionINI(1)) Then
            Select Case ConfiguracionINI(1)
                Case "Spanish"
                    id = 0
                Case "English"
                    id = 1000
                Case "French"
                    id = 2000
                Case "German"
                    id = 3000
            End Select
            Reescribe = True
        Else
            id = Val(ConfiguracionINI(1))
        End If
        Empresa = ConfiguracionINI(2)
        UsaGamaAlta = ConfiguracionINI(3)
        '********************************
        ' ML100 --> Si el peso está
        ' entre 15 y 30 kg, sólo
        ' proporciona dos decimales
        ' de peso, y hay que corregirlo
        '*****************************
        If Dir(MiruTa & "\ml100.ini") <> "" Then
            ML100 = True
        End If
        usaSC10 = ConfiguracionINI(4)
        UsaFichasVacuno = ConfiguracionINI(5)
        UsaGruposdeConservacion = ConfiguracionINI(6)
        gedit = "0"
        If Dir(var.rutaSYS & "\gedit", vbDirectory) <> "" Then
            If Dir(var.rutaSYS & "\gedit\gedit.exe") <> "" Then
                gedit = "1"
            Else
                gedit = "0"
            End If
        End If
        If ConfiguracionINI(8) = "#TRUE#" Then
            tcpip = True
        Else
            tcpip = False
        End If
        cngvelocidad = ConfiguracionINI(9)
        sPortBal = ConfiguracionINI(10)
        pathSC10 = ConfiguracionINI(11)
        'descAuto = ConfiguracionINI(12)
        'If descAuto Then
        '    lCogeTiquet = True
        '    cuenTaTqt = 1
        '    descautotime = ConfiguracionINI(13)
        '    If Right(descautotime, 3) = "sec" Then
        '        Form1.timetqt.Interval = 1000
        '    Else
        '        Form1.timetqt.Interval = 60000
        '    End If
        'Else
        '    lCogeTiquet = False
        '    descautotime = ""
        '    Form1.timetqt.Enabled = False
        '    Form1.timetqt.Interval = 0
        'End If
        GA_RecibirTotalVenta = ConfiguracionINI(14)
        GA_RecibirTotalSuper = ConfiguracionINI(15)
        GA_RecibirTotalEnvasado = ConfiguracionINI(16)
        GA_RecibirTotalAutoservicio = ConfiguracionINI(17)
        GA_RecibirTotalVentaL = ConfiguracionINI(18)
        GA_RecibirTotalSuperL = ConfiguracionINI(19)
        GA_RecibirTotalEnvasadoL = ConfiguracionINI(20)
        GA_RecibirTotalAutoservicioL = ConfiguracionINI(21)
        GA_BorrarTotal = ConfiguracionINI(22)
        SC10_BorrarTotal = ConfiguracionINI(23)
        SC10_ProcesarTotalVenta = ConfiguracionINI(24)
        SC10_ProcesarTotalSuper = ConfiguracionINI(25)
        SC10_ProcesarTotalEnvasado = ConfiguracionINI(26)
        SC10_ProcesarTotalAutoservicio = ConfiguracionINI(27)
        SC10_Borrar_Ticket_dat = ConfiguracionINI(28)
        SC10_Borrar_Eti_dat = ConfiguracionINI(29)
        RecibirTotalVacuno = ConfiguracionINI(30)
        BorrarTotalVacuno = ConfiguracionINI(31)
        decimales = ConfiguracionINI(32)
        SimboloMonetario = ConfiguracionINI(33)
        ValorEuro = ConfiguracionINI(34)
        If ConfiguracionINI(35) = "Nacional" Then
            UsaEuro = False
        Else
            UsaEuro = True
            decimales = 2
        End If
        
        If ConfiguracionINI(36) = "1" Then
            UsaBackup = True
        Else
            UsaBackup = False
        End If
        pathred = ConfiguracionINI(37)
        FamiliaEspecial = ConfiguracionINI(38)
        EtiquetaEspecial = ConfiguracionINI(39)
        PathDBF = ConfiguracionINI(40)
        decimales = ConfiguracionINI(41)
        SimboloMonetario = ConfiguracionINI(42)
        If ConfiguracionINI(43) = "1" Then
            Usa20Lineas = True
        Else
            Usa20Lineas = False
        End If
        '
        '
        ' ConfiguracionINI(44)
        '
        '
        If ConfiguracionINI(45) <> "" Then
            SC10Rs232 = ConfiguracionINI(45)
        Else
            SC10Rs232 = False
        End If
        If ConfiguracionINI(46) <> "" Then
            SC10ParametrosSerie = ConfiguracionINI(46)
        Else
            SC10ParametrosSerie = "19200,E,8,1,2"
        End If
        If ConfiguracionINI(47) <> "" Then
            EnvioDirecto = ConfiguracionINI(47)
        Else
            EnvioDirecto = False
        End If
        FamiliasFijas = False
        If ConfiguracionINI(132) = "1" Then FamiliasFijas = True
        SubseccionesFijas = False
        If ConfiguracionINI(134) = "1" Then SubseccionesFijas = True
        '*********************************
        ' Parámetros de hora de cierre
        ' automática
        '*********************************
        'CI_HORA = ConfiguracionINI(48)
        'CI_MINUTO = ConfiguracionINI(49)
        'RE_HORA = ConfiguracionINI(50)
        'RE_MINUTO = ConfiguracionINI(51)
        'If CI_HORA < 24 And CI_MINUTO < 60 Then
        '    Form1.TmrCierre.Interval = 10000
        '    Form1.TmrCierre.Enabled = True
        'Else
        '    Form1.TmrCierre.Enabled = False
        'End If
        'If RE_HORA < 24 And RE_MINUTO < 60 Then
        '    Form1.TmrGranTotal.Interval = 10000
        '    Form1.TmrGranTotal.Enabled = True
        'Else
        '    Form1.TmrGranTotal.Enabled = False
        'End If
        '*********************************
        'CI_HOST = ConfiguracionINI(52)
        'CI_TCP = ConfiguracionINI(53)
        'CI_UDP = ConfiguracionINI(54)
        
        'StoreFlow = ConfiguracionINI(55)
        'Taquion = ConfiguracionINI(56)
        '*******************************+
        'If Taquion Then
        '    If Dir(MiruTa & "\asorden", vbDirectory) = "" Then
        '        On Error Resume Next
        '        MkDir MiruTa & "\asorden"
        '        On Error GoTo 0
        '    End If
        '    If Dir(MiruTa & "\asorden\off.ord") <> "" Then
        '        On Error Resume Next
        '        Kill MiruTa & "\asorden\off.ord"
        '        On Error GoTo 0
        '    End If
        '    If Dir(MiruTa & "\asorden\gtotal.ord") <> "" Then
        '        On Error Resume Next
        '        Kill MiruTa & "\asorden\gtotal.ord"
        '        On Error GoTo 0
        '    End If
        '    Form1.TmrCierre.Interval = 10000
        '    Form1.TmrCierre.Enabled = True
        '    Form1.TmrGranTotal.Interval = 10000
        '    Form1.TmrGranTotal.Enabled = True
        'End If
        'TaquionFichero = ConfiguracionINI(57)
        'TaquionTiquets = ConfiguracionINI(58)
        'Iconificar = ConfiguracionINI(59)
        'IconificarInicio = ConfiguracionINI(60)
        'LogArticulos = ConfiguracionINI(61)
        'DebugActivo = ConfiguracionINI(62)
        'ComandoLibre(0) = ConfiguracionINI(63)
        'EjecutarLibre(0) = ConfiguracionINI(64)
        'ComandoLibre(1) = ConfiguracionINI(65)
        'EjecutarLibre(1) = ConfiguracionINI(66)
        'ComandoLibre(2) = ConfiguracionINI(67)
        'EjecutarLibre(2) = ConfiguracionINI(68)
        'ComandoLibre(3) = ConfiguracionINI(69)
        'EjecutarLibre(3) = ConfiguracionINI(70)
        'ComandoLibre(4) = ConfiguracionINI(71)
        'EjecutarLibre(4) = ConfiguracionINI(72)
        'UsarPantallas = ConfiguracionINI(73)
        'PuertoPantallas = ConfiguracionINI(74)
        UsaDual = ConfiguracionINI(75)
        '***************************
        ' a partir de aquí mantener
        ' compatibilidad de versiones
        '***************************
        If ConfiguracionINI(76) <> "" Then
            DiasBorrado = ConfiguracionINI(76)
        Else
            DiasBorrado = 0
        End If
        'If ConfiguracionINI(77) <> "" Then
        '    BorrarBackup = ConfiguracionINI(77)
        'Else
        '    BorrarBackup = 0
        'End If
        'If ConfiguracionINI(78) <> "" Then
        '    Exp_Fecha = ConfiguracionINI(78)
        'Else
        '    Exp_Fecha = 0
        'End If
        'If ConfiguracionINI(79) <> "" Then
        '    Exp_hora = ConfiguracionINI(79)
        'Else
        '    Exp_hora = 0
        'End If
        'Exp_Mas = ConfiguracionINI(80)
        'Exp_Menos = ConfiguracionINI(81)
        'If ConfiguracionINI(82) <> "" Then
        '    If ConfiguracionINI(82) = "1" Then
        '        Exp_Absoluto = True
        '    Else
        '        Exp_Absoluto = False
        '    End If
        'Else
        '    Exp_Absoluto = False
        'End If
        'If Trim(ConfiguracionINI(83)) <> "" Then
        '    Exp_Path = ConfiguracionINI(83)
        'Else
        '    Exp_Path = MiruTa
        'End If
        'If ConfiguracionINI(84) = "" Then
        '    Exp_Diario = False
        'Else
        '    Exp_Diario = ConfiguracionINI(84)
        'End If
        'If ConfiguracionINI(85) = "" Then
        '    TqNombre = "BBTTTT"
        'Else
        '    TqNombre = ConfiguracionINI(85)
        'End If
        'TqDirecto = True
        'If ConfiguracionINI(86) <> "" Then
        '    If ConfiguracionINI(86) = "0" Then TqDirecto = False
        'End If
        'If ConfiguracionINI(87) <> "" Then
        '    If ConfiguracionINI(87) = "1" Then
        '        VacunoD = True
        '    Else
        '        VacunoD = False
        '    End If
        'Else
        '    VacunoD = False
        'End If
        'If ConfiguracionINI(88) <> "" Then
        '    If ConfiguracionINI(88) = "1" Then
        '        bunTicket = True
        '    Else
        '        bunTicket = False
        '    End If
        'Else
        '    bunTicket = False
        'End If
        'sunTicket = ConfiguracionINI(89)
        'If ConfiguracionINI(90) <> "" Then
        '    If ConfiguracionINI(90) = "1" Then
        '        GA_Reinit = True
        '    Else
        '        GA_Reinit = False
        '    End If
        'Else
        '    GA_Reinit = True
        'End If
        If ConfiguracionINI(91) <> "" Then
            TipoC9 = ConfiguracionINI(91)
            If TipoC9 Then SC10Rs232 = True
        Else
            TipoC9 = False
        End If
        If ConfiguracionINI(92) <> "" Then
            ModoC9 = ConfiguracionINI(92)
        End If
        'CheckNoTiquet = False
        'If ConfiguracionINI(93) <> "" Then
        '    If ConfiguracionINI(93) <> "0" Then CheckNoTiquet = True
        'End If
        'ClienteD = False
        'If ConfiguracionINI(94) <> "" Then
        '    If ConfiguracionINI(94) <> "0" Then ClienteD = True
        'End If
        'If UsarPantallas Then Load FrmRemoto
        '***************
        'If ConfiguracionINI(95) <> "" Then
        '    PuertoModem = Val(ConfiguracionINI(95))
        '    If PuertoModem < 0 Or PuertoModem > 4 Then PuertoModem = 1
        'Else
        '    PuertoModem = 1
        'End If
        'If ConfiguracionINI(96) <> "" Then
        '    VelocidadModem = Val(ConfiguracionINI(96))
        '    If VelocidadModem < 0 Or VelocidadModem > 4 Then VelocidadModem = 1
        'Else
        '    VelocidadModem = 1
        'End If
        'If ConfiguracionINI(97) <> "" Then
        '    TelefonoModem = ConfiguracionINI(97)
        '    If Len(TelefonoModem) > 50 Then TelefonoModem = ""
        'Else
        '    TelefonoModem = ""
        'End If
        'If ConfiguracionINI(98) <> "" Then
        '    CadModem1 = ConfiguracionINI(98)
        'Else
        '    CadModem1 = ""
        'End If
        'If ConfiguracionINI(99) <> "" Then
        '    CadModem2 = ConfiguracionINI(99)
        'Else
        '    CadModem2 = ""
        'End If
        'If ConfiguracionINI(100) <> "" Then
        '    CadModem3 = ConfiguracionINI(100)
        'Else
        '    CadModem3 = ""
        'End If
        'If ConfiguracionINI(101) = "1" Then
        '    ColgarModem = True
        'Else
        '    If ConfiguracionINI(101) = "" Then
        '        ColgarModem = True
        '    Else
        '        ColgarModem = False
        '    End If
        'End If
        'If ConfiguracionINI(102) = "1" Then
        '    UsaModem = True
        'Else
        '    UsaModem = False
        'End If
        'If ConfiguracionINI(103) = "1" Then
        '    Pasarela = True
        'Else
        '    Pasarela = False
        'End If
        'If ConfiguracionINI(104) = "" Then
        '    PasarelaHost = "127.000.000.001"
        'Else
        '    PasarelaHost = ConfiguracionINI(104)
        'End If
        'If ConfiguracionINI(105) = "" Then
        '    PasarelaPuerto = 32340
        'Else
        '    PasarelaPuerto = Val(ConfiguracionINI(105))
        'End If
        'AbsorverCambios = False
        'If ConfiguracionINI(106) = "1" Then AbsorverCambios = True
    'End If
    'CargaIdiomas
    'If Reescribe Then EscribirParametrosInicio
End Function
'''''''
'''''''


Attribute VB_Name = "ParametrosdeInicio"
Option Explicit

Public Sub EscribirParametrosInicio()
' Esta función depende de FrmConfiguracion2
    Dim Fichero As Integer
    Dim ParametrosSerie As String
    Dim SC10ParSerie As String
    Dim MiValor As Integer
    Dim MiValor2 As Integer
    Dim bucle As Integer
    Dim MyModeloV4 As Integer
    On Error GoTo 0
    ParametrosSerie = FrmConfiguracion2.CmbVelocidad.TexT & ","
    ParametrosSerie = ParametrosSerie & FrmConfiguracion2.CmbParidad.TexT & ","
    ParametrosSerie = ParametrosSerie & FrmConfiguracion2.CmbBits.TexT & ","
    ParametrosSerie = ParametrosSerie & "1" & ","
    ParametrosSerie = ParametrosSerie & Right(FrmConfiguracion2.CmbCOM.TexT, 1)
    SC10ParSerie = FrmConfiguracion2.CMBSC10Baudios.TexT & ","
    SC10ParSerie = SC10ParSerie & FrmConfiguracion2.CmbSC10Paridad.TexT & ","
    SC10ParSerie = SC10ParSerie & FrmConfiguracion2.CmbSC10bits.TexT & ","
    SC10ParSerie = SC10ParSerie & "1" & ","
    SC10ParSerie = SC10ParSerie & Right(FrmConfiguracion2.CmbSC10Port.TexT, 1)
    Check_FileAttr Hydra_INI
    Fichero = FreeFile
    Open Hydra_INI For Output As Fichero
    '**********
    ' sec
    '**********
    Randomize Time
    MiValor = Int((4500 * Rnd) + 1)
    If HaySeguridad Then
        MiValor2 = MiValor + 1
    Else
        MiValor2 = MiValor + 3
    End If
    Write #Fichero, Encripta(Format(MiValor, "0000") & Format(MiValor2, "0000")), "NEVER EDIT NOR REMOVE THIS LINE"
    Write #Fichero, id, "Idioma "
    'Write #Fichero, FrmConfiguracion2.TxtEmpresa.TexT, "Nombre de la Empresa"
    Write #Fichero, Replace(FrmConfiguracion2.TxtEmpresa.TexT, vbCrLf, "||"), "Nombre de la Empresa"
    Write #Fichero, FrmConfiguracion2.ChkGamaAlta.Value, "Utilizar TouchScale"
    Write #Fichero, FrmConfiguracion2.ChkSC10.Value, "Utilizar Concentrador SC-10"
    Write #Fichero, FrmConfiguracion2.ChkFichaVacuno.Value, "Utilizar Fichas de Vacuno"
    Write #Fichero, FrmConfiguracion2.ChkGruposConservacion.Value, "Utilizar Grupos de Conservacion"
    Write #Fichero, FrmConfiguracion2.ChkGedit.Value, "Utilizar Gedit"
    Write #Fichero, FrmConfiguracion2.OptTCPIP, "Ethernet(#TRUE#) o RS-232(#FALSE#)"
    Write #Fichero, ParametrosSerie, "Parametros del Puerto Serie RS-232 TouchScale"
    Write #Fichero, FrmConfiguracion2.TxtPuerto.TexT, "Puerto TCP/IP"
    Write #Fichero, FrmConfiguracion2.TxtPathSC10, "Ruta de Red del Concentrador SC-10"
    Write #Fichero, FrmConfiguracion2.ChkContinua.Value, "Recepción contínua de Tiquets"
    If FrmConfiguracion2.ChkContinua.Value = vbChecked Then
        Write #Fichero, FrmConfiguracion2.TxtIntervalo.TexT, "Intervalo de Recepción de Tiquets"
    Else
        Write #Fichero, "", "Intervalo de Recepción de Tiquets"
    End If
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(0).Value, "TouchScale:Recibir Totales Venta"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(1).Value, "TouchScale:Recibir Totales Super"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(2).Value, "TouchScale:Recibir Totales Envasado"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(3).Value, "TouchScale:Recibir Totales Autoservicio"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(4).Value, "TouchScale:Recibir Totales Venta Local"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(5).Value, "TouchScale:Recibir Totales Super Local"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(6).Value, "TouchScale:Recibir Totales Envasado Local"
    Write #Fichero, FrmConfiguracion2.ChkTotalesGamaAlta(7).Value, "TouchScale:Recibir Totales Autoservicio Local"
    Write #Fichero, FrmConfiguracion2.ChkBorrarTotalGA.Value, "TouchScale:Borrar Totales"
    Write #Fichero, FrmConfiguracion2.ChkBorrarTotalSC10.Value, "SC10:Borrar Totales"
    Write #Fichero, FrmConfiguracion2.ChkTotalesSC10(0).Value, "SC10:Recibir Totales Venta"
    Write #Fichero, FrmConfiguracion2.ChkTotalesSC10(1).Value, "SC10:Recibir Totales Super"
    Write #Fichero, FrmConfiguracion2.ChkTotalesSC10(2).Value, "SC10:Recibir Totales Envasado"
    Write #Fichero, FrmConfiguracion2.ChkTotalesSC10(3).Value, "SC10:Recibir Totales Autoservicio"
    Write #Fichero, FrmConfiguracion2.ChkBorrarTiquet.Value, "SC10:Borrar Ticket.dat"
    Write #Fichero, FrmConfiguracion2.ChkBorrarEti.Value, "SC10:Borrar Eti.dat"
    Write #Fichero, FrmConfiguracion2.ChkTotalesVacuno.Value, "Recibir Totales de Fichas de Vacuno"
    Write #Fichero, FrmConfiguracion2.ChkBorrarVacuno.Value, "Borrar Totales de Fichas de Vacuno"
    Write #Fichero, FrmConfiguracion2.TxtDecimales.TexT, "Decimales de la Moneda"
    Write #Fichero, FrmConfiguracion2.TxtSimbolo.TexT, "Símbolo de Moneda Nacional"
    Write #Fichero, 0, "Reservado"
    Write #Fichero, FrmConfiguracion2.CmbMoneda.TexT, "Moneda a Utilizar (Nacional/Euro)"
    Write #Fichero, FrmConfiguracion2.ChkBackup.Value, "Realizar Backup de la Base"
    Write #Fichero, FrmConfiguracion2.TxtBackup.TexT, "Ruta de Backup"
    Write #Fichero, FrmConfiguracion2.TxtFamiliaVacuno.TexT, "Familia Especial"
    Write #Fichero, FrmConfiguracion2.TxtEtiquetaVacuno.TexT, "Etiqueta Especial"
    Write #Fichero, FrmConfiguracion2.TxtDbteclas.TexT, "Ubicación de DBTECLAS.DBF"
    Write #Fichero, FrmConfiguracion2.TxtDecimales.TexT, "Nro. de Decimales moneda nacional"
    Write #Fichero, FrmConfiguracion2.TxtSimbolo.TexT, "Símbolo de la moneda Nacional"
    Write #Fichero, FrmConfiguracion2.Chk20Lineas.Value, "Usar 20 líneas de Descriptivo"
    Write #Fichero, "0", "Reservado"
    Write #Fichero, FrmConfiguracion2.OptSC10Comu(1).Value, "Comunicación Rs232 con SC10"
    Write #Fichero, SC10ParSerie, "Parametros del Puerto Serie RS-232 SC10"

    Write #Fichero, True, "Reservado"

    If IsNumeric(FrmConfiguracion2.TxtCI_HORA) Then
        Write #Fichero, FrmConfiguracion2.TxtCI_HORA, "Hora de Cierre"
        Write #Fichero, FrmConfiguracion2.txtCI_MINUTO, "Minuto de Cierre"
    Else
        Write #Fichero, 25, "Hora de Cierre"
        Write #Fichero, 61, "Minuto de Cierre"
    End If
    If IsNumeric(FrmConfiguracion2.TxtRE_HORA) Then
        Write #Fichero, FrmConfiguracion2.TxtRE_HORA, "Hora de Recogida"
        Write #Fichero, FrmConfiguracion2.TxtRE_MINUTO, "Minuto de Recogida"
    Else
        Write #Fichero, 25, "Hora de Recogida"
        Write #Fichero, 61, "Minuto de Recogida"
    End If
    Write #Fichero, FrmConfiguracion2.txtCI_HOST, "Host - Enlace Socket"
    Write #Fichero, FrmConfiguracion2.txtCI_TCP, "Puerto TCP - Enlace Socket"
    Write #Fichero, FrmConfiguracion2.txtCI_UDP, "Puerto UDP - Enlace Socket"
    Write #Fichero, FrmConfiguracion2.ChkStoreFlow.Value, "Enlace Socket"
    Write #Fichero, FrmConfiguracion2.ChkTaquion.Value, "Enlace ASCII"
    Write #Fichero, FrmConfiguracion2.TxtTaquionFichero.TexT, "Enlace ASCII - fichero de importación contínua"
    Write #Fichero, FrmConfiguracion2.ChkTaquionTiquets, "Enlace ASCII - Dejar ficheros de tiquets"
    Write #Fichero, FrmConfiguracion2.ChkIconificar.Value, "Iconificar en barra de estado"
    Write #Fichero, FrmConfiguracion2.ChkIconificarInicio.Value, "Iconificar al inicio"
    Write #Fichero, FrmConfiguracion2.ChkLogArticulos.Value, "Log de Artículos"
    Write #Fichero, FrmConfiguracion2.ChklogSistema.Value, "Log del Sistema"
    Write #Fichero, FrmConfiguracion2.TxtComando(0).TexT, "Comando Libre - Nombre"
    Write #Fichero, FrmConfiguracion2.TxtEjecutar(0).TexT, "Comando Libre - Linea de comandos"
    Write #Fichero, FrmConfiguracion2.TxtComando(1).TexT, "Comando Libre - Nombre"
    Write #Fichero, FrmConfiguracion2.TxtEjecutar(1).TexT, "Comando Libre - Linea de comandos"
    Write #Fichero, FrmConfiguracion2.TxtComando(2).TexT, "Comando Libre - Nombre"
    Write #Fichero, FrmConfiguracion2.TxtEjecutar(2).TexT, "Comando Libre - Linea de comandos"
    Write #Fichero, FrmConfiguracion2.TxtComando(3).TexT, "Comando Libre - Nombre"
    Write #Fichero, FrmConfiguracion2.TxtEjecutar(3).TexT, "Comando Libre - Linea de comandos"
    Write #Fichero, FrmConfiguracion2.TxtComando(4).TexT, "Comando Libre - Nombre"
    Write #Fichero, FrmConfiguracion2.TxtEjecutar(4).TexT, "Comando Libre - Linea de comandos"
    Write #Fichero, FrmConfiguracion2.ChkPantallas.Value, "Activar Pantallas Remotas"
    Write #Fichero, FrmConfiguracion2.TxtPuertoPantallas.TexT, "Puerto de Pantallas Remotas"
    Write #Fichero, UsaDual, "Utilizar Moneda distinta del Euro"
    Write #Fichero, FrmConfiguracion2.Txtdias.TexT, "Borrar totales más antiguos de ... días"
    Write #Fichero, FrmConfiguracion2.TxtBackupErase.TexT, "Borrar copias de seguridad más antíguas de ... días"
    Write #Fichero, FrmConfiguracion2.CmbFormato(0).ListIndex, "Exportación - Formato de Fecha"
    Write #Fichero, FrmConfiguracion2.CmbFormato(1).ListIndex, "Exportación - Formato de Hora"
    Write #Fichero, FrmConfiguracion2.TxtPositivo(0).TexT, "Exportación - Signo positivo"
    Write #Fichero, FrmConfiguracion2.TxtPositivo(1).TexT, "Exportación - Signo negativo"
    Write #Fichero, FrmConfiguracion2.ChkAbsoluto.Value, "Exportación - Utilizar valores absolutos"
    Write #Fichero, FrmConfiguracion2.TxtPathExport.TexT, "Exportación - Path"
    Write #Fichero, FrmConfiguracion2.Optdiario(0).Value, "Exportación - Sólo Diario"
    Write #Fichero, FrmConfiguracion2.TxtNombreTicket.TexT, "Nombre Fichero Tiquets"
    If FrmConfiguracion2.ChkTqDirecto.Value = vbChecked Then
        Write #Fichero, "1", "ASCII : envío directo"
    Else
        Write #Fichero, "0", "ASCII : envío directo"
    End If
    If FrmConfiguracion2.Option1(1).Value = True Then
        Write #Fichero, "1", "Vacuno versión D o posterior"
    Else
        Write #Fichero, "0", "Vacuno versión D o posterior"
    End If
    'If FrmConfiguracion2.Option2(1).Value = True Then
    '    Write #Fichero, "1", "Ticket ASCII en fichero único"
    'Else
    Write #Fichero, "0", "Ticket ASCII en fichero único"
    'End If
    Write #Fichero, FrmConfiguracion2.TxtAllTicket.TexT, "Nombre Fichero Tiquets Único"
    Write #Fichero, FrmConfiguracion2.ChkGaReinit.Value, "Reiniciar Numeración de Tiquets TouchScale"
    Write #Fichero, FrmConfiguracion2.OptSC10(1).Value, "Concentrador C9 o C10"
    Write #Fichero, FrmConfiguracion2.CmbTipoC9.ListIndex, "Número de Secciones C9 o C10"
    Write #Fichero, FrmConfiguracion2.ChkNoTiquet.Value, "Chequeo Rígido de Tiquets"
    Write #Fichero, FrmConfiguracion2.ChkClientes.Value, "Versión D : clientes"
    Write #Fichero, FrmConfiguracion2.CmbModemPuerto.ListIndex + 1, "Modem --> Puerto"
    Write #Fichero, FrmConfiguracion2.CmbModemVelocidad.ListIndex + 1, "Modem --> Velocidad"
    Write #Fichero, FrmConfiguracion2.TxtModemTel.TexT, "Modem --> Teléfono"
    Write #Fichero, FrmConfiguracion2.TxtModemCadena(0).TexT, "Modem --> Cadena"
    Write #Fichero, FrmConfiguracion2.TxtModemCadena(1).TexT, "Modem --> Cadena"
    Write #Fichero, FrmConfiguracion2.TxtModemCadena(2).TexT, "Modem --> Cadena"
    Write #Fichero, FrmConfiguracion2.ChkColgarModem.Value, "Modem --> Colgar Automáticamente"
    Write #Fichero, FrmConfiguracion2.ChkModem.Value, "Modem --> Sí(1) o No(2)"
    Write #Fichero, FrmConfiguracion2.ChkPasarela.Value, "Usa Pasarela Remota"
    Write #Fichero, FrmConfiguracion2.TxtPasarelaHost.TexT, "Host Pasarela"
    Write #Fichero, FrmConfiguracion2.TxtPasarelaPort.TexT, "Puerto Pasarela"
    Write #Fichero, FrmConfiguracion2.ChkAbsorverCambios.Value, "Absorver Cambios de Precio de Balanzas"
    Write #Fichero, FrmConfiguracion2.ChkPlu0.Value, "Rechazar artículos con Precio cero"
    Write #Fichero, FrmConfiguracion2.ChkSolicitarPedido, "Solicitar pedido en recogida contínua"
    Write #Fichero, FrmConfiguracion2.ChkBorrarPedido, "borrar pedido tras cada recogida"
    Write #Fichero, FrmConfiguracion2.ChkPedidoFinalDia, "Borrar pedido a fin de día"
    Write #Fichero, FrmConfiguracion2.ChkSolicitarInventario, "Solicitar inventario en recogida contínua"
    Write #Fichero, FrmConfiguracion2.ChkBorrarInventario, "borrar inventario tras cada recogida"
    Write #Fichero, FrmConfiguracion2.ChkInventarioFinal, "borrar inventario a fin de día"
    Write #Fichero, FrmConfiguracion2.ChkSolicitarPedidoFin, "Solicitar pedido a fin de día"
    Write #Fichero, FrmConfiguracion2.ChkSolicitarInventarioFin, "solicitar inventario a fin de día"
    Write #Fichero, FrmConfiguracion2.ChkEcoNet.Value, "Equipos EcoNet / V8"
    Write #Fichero, FrmConfiguracion2.OptECO(0).Value, "Tipo EcoNet (true) o V8 (false)"
    Write #Fichero, FrmConfiguracion2.ChkEcoLabel.Value, "Equipos Marte IC-ILC"
    Write #Fichero, FrmConfiguracion2.TxtEquipoV8.TexT, "Número equipo V8/Marte IC-ILC conectado"
    Write #Fichero, FrmConfiguracion2.TxtEquipoLabel.TexT, "Número equipo Marte IC-ILC conectado"
    Write #Fichero, FrmConfiguracion2.CmbPortV8.ListIndex, "puerto RS-232 Marte IC-ILC/V8"
    Write #Fichero, FrmConfiguracion2.CmbPortLabel.ListIndex, "puerto RS-232 Marte IC-ILC"
    Write #Fichero, FrmConfiguracion2.ChkEcoPrint.Value, "Equipos IV4 LCD / IV4"
    For bucle = 0 To 3    'c2f iv4 ic+ 3
        If FrmConfiguracion2.OptIV4(bucle) = True Then
            MyModeloV4 = bucle
        End If
    Next bucle
    'c2f iv4 IC+ ver como actuar en caso de mymodelov4
    Write #Fichero, MyModeloV4, "Tipo IV4"
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Write #Fichero, FrmConfiguracion2.CmbPortPrint.ListIndex, "puerto RS-232 Marte IC-ILC"
    Write #Fichero, FrmConfiguracion2.ChkEstiloInterfaz.Value, "Estilo de la interfaz"
    Write #Fichero, FrmConfiguracion2.ChkEcoSoloUna.Value, "Marte/IV4 - Comunicar sólo con una máquina"
    Write #Fichero, FrmConfiguracion2.ChkPedidoIV4.Value, "IV4 TQ - Solicitar pedido/inventario al recibir totales"
    Write #Fichero, FrmConfiguracion2.ChkEnvioContinuo.Value, "Envío contínuo + Recoger tiquets contínuo"
    Write #Fichero, FrmConfiguracion2.ChkMultiEuroscale.Value, "Trabajo con múltiples servidoras TouchScale"
    Write #Fichero, FrmConfiguracion2.Chk4Lineas.Value, "SC10 - 4 líneas de descriptivo"
    Write #Fichero, FrmConfiguracion2.ChkFamilias.Value, "Familias fijas"
    Write #Fichero, Val(FrmConfiguracion2.TxtRetries.TexT), "Reintentos de comunicación"
    Write #Fichero, FrmConfiguracion2.ChkSubSec.Value, "Subsecciones fijas"
    Write #Fichero, FrmConfiguracion2.OptSC10Comu(2).Value, "Comunicación TCP/IP con SC10"
    Write #Fichero, FrmConfiguracion2.TxtIPSC10.TexT, "IP SC10"
    Write #Fichero, FrmConfiguracion2.TxtTCPSC10.TexT, "Port SC10"
    Write #Fichero, FrmConfiguracion2.chkC9Cod.Value, FrmConfiguracion2.OptSC10(1).Caption & " " & FrmConfiguracion2.chkC9Cod.Caption
    Write #Fichero, FrmConfiguracion2.chkLin2040.Value, "version G: Lineas 20 - 40"
    '2.0.9
    Write #Fichero, FrmConfiguracion2.option3.Value, "Comunicación Marte 9600"
    Write #Fichero, FrmConfiguracion2.option4.Value, "Comunicación Marte 19200"
    Write #Fichero, FrmConfiguracion2.CheckMarteNew.Value, "Marte Nuevo Protocolo"
    Write #Fichero, FrmConfiguracion2.CheckMarteAlfa.Value, "Marte Alfanumérica"
    Write #Fichero, FrmConfiguracion2.Check49T.Value, "Marte Alfanumérica 49T"
    Write #Fichero, FrmConfiguracion2.Check66T.Value, "Marte Alfanumérica 66T"
    ''''''2.0.23
    Write #Fichero, FrmConfiguracion2.ChkL000.Value, "Versión 00L-??? o posterior"
    ''''''''''''
    Close #Fichero
fin:
End Sub
Public Sub LeerParametrosInicio()
    Dim nMyFich As Integer
    Dim sNumdec As String
    Dim sNum100 As String
    Dim Fichero As Integer
    Dim Buffer As String
    Dim Contador As Long
    Dim Reescribe As Boolean
    Dim W_5_2_2 As Boolean
    Dim ConfiguracionINI(151) As String
    Dim lReW As Boolean

    W_5_2_2 = False

    '
    ' Meter valores por defecto si no existe mng.ini
    '
    Reescribe = False
    If Dir(Hydra_INI) = "" Then
        Los_parametros_Defecto
        'If AhorraMas Then Los_Parametros_Defecto_Ahorramas
        'If Sabeco Then Los_Parametros_Defecto_Sabeco
        'If AHOLD Then Los_parametros_Defecto_Ahold
        'If Supeco Then Los_Parametros_Defecto_Supeco

        'If Dir(App.Path & "\hydraserver.exe") <> "" Then Los_Parametros_Defecto_StoreFlow

        CrearPassSec
    Else
        '
        ' Leer ntq.ini
        '
        Fichero = FreeFile
        Open Hydra_INI For Input As Fichero
        Contador = 0
        Do Until EOF(Fichero)
            Input #Fichero, ConfiguracionINI(Contador), Buffer
            Contador = Contador + 1
        Loop
        Close Fichero
        If Contador > 116 And Contador < 119 Then W_5_2_2 = True

        '*************
        ' sec
        '*************
        Buffer = DesEncripta(ConfiguracionINI(0))
        If Not IsNumeric(left(Buffer, 8)) Then
            CadenadeLog CargaCadena(851)
            MsgBox CargaCadena(851), vbCritical
            End
        Else
            If (Val(left(Buffer, 4)) + 3) = Val(Mid(Buffer, 5, 4)) Then
                HaySeguridad = False
            Else
                If (Val(left(Buffer, 4)) + 1) = Val(Mid(Buffer, 5, 4)) Then
                    HaySeguridad = True
                Else
                    MsgBox CargaCadena(851), vbCritical
                    End
                End If
            End If
        End If
        ' Idioma
        If Not HayMulti Then
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
        End If
        'Grecia
        If id = 6000 Then
            Fichero = FreeFile()
            Open App.Path & "\greek.txt" For Output As #Fichero
            Print #Fichero, Chr(130) & Chr(131) & Chr(135)
            Close #Fichero

        End If
        '''''''
        Empresa = ConfiguracionINI(2)
        UsaGamaAlta = ConfiguracionINI(3)

        '**************************
        ' ML100 --> Si el peso está
        ' entre 15 y 30 kg, sólo
        ' proporciona dos decimales
        ' de peso, y hay que corregirlo
        '******************************
        If Dir(Miruta & "\ml100.ini") <> "" Then
            ML100 = True
            nMyFich = FreeFile()
            Open App.Path & "\ml100.ini" For Input As #nMyFich
            On Error Resume Next
            Line Input #nMyFich, sNumdec
            Line Input #nMyFich, sNum100
            Close #nMyFich
            If Err.Number <> 0 Then
                sNumdec = "03"
                sNum100 = "00"
            End If
            On Error GoTo 0
            If Len(sNumdec) > 1 Then
                nNumDecML100 = Val(sNumdec)
            Else
                nNumDecML100 = 3
            End If
            If Len(sNum100) > 1 Then
                nNumBalML100 = Val(sNum100)
            Else
                nNumBalML100 = 0
            End If
        End If
        usaSC10 = ConfiguracionINI(4)
        UsaFichasVacuno = ConfiguracionINI(5)
        UsaGruposdeConservacion = ConfiguracionINI(6)
        gedit = "0"
        If ConfiguracionINI(8) = "#TRUE#" Then
            tcpip = True
        Else
            tcpip = False
        End If
        cngvelocidad = ConfiguracionINI(9)
        sPortBal = ConfiguracionINI(10)
        pathSC10 = ConfiguracionINI(11)
        descAuto = ConfiguracionINI(12)
        If descAuto Then
            lCogeTiquet = True
            cuenTaTqt = 1
            descautotime = ConfiguracionINI(13)
            If Right(descautotime, 3) = "sec" Then
                frmControl.timetqt.Interval = 1000
            Else
                frmControl.timetqt.Interval = 60000
            End If
        Else
            lCogeTiquet = False
            descautotime = ""
            frmControl.timetqt.Enabled = False
            frmControl.timetqt.Interval = 0
        End If
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
        '****************************
        ' obsoleto
        'ValorEuro = ToDouble(ConfiguracionINI(34))
        ValorEuro = 1
        '****************************
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
        ' 47 --> sIN USO ACTUALMENTE
        ' If ConfiguracionINI(47) <> "" Then

        '*********************************
        ' Parámetros de hora de cierre
        ' automática
        '*********************************
        CI_HORA = ConfiguracionINI(48)
        CI_MINUTO = ConfiguracionINI(49)
        RE_HORA = ConfiguracionINI(50)
        RE_MINUTO = ConfiguracionINI(51)
        If Not StoreFlow Then
            If CI_HORA < 24 And CI_MINUTO < 60 Then
                frmControl.TmrCierre.Interval = 10000
                frmControl.TmrCierre.Enabled = True
            Else
                frmControl.TmrCierre.Enabled = False
            End If
            If RE_HORA < 24 And RE_MINUTO < 60 Then
                frmControl.TmrGranTotal.Interval = 10000
                frmControl.TmrGranTotal.Enabled = True
            Else
                frmControl.TmrGranTotal.Enabled = False
            End If
        End If
        '*********************************
        CI_HOST = ConfiguracionINI(52)
        CI_TCP = ConfiguracionINI(53)
        CI_UDP = ConfiguracionINI(54)

        'StoreFlow = ConfiguracionINI(55)
        If Dir(App.Path & "\hydraserver.exe") <> "" Then
            StoreFlow = True
        Else
            StoreFlow = False
        End If
        Taquion = ConfiguracionINI(56)
        If StoreFlow Then Taquion = False
        '*******************************+
        If Taquion Then
            If Dir(Miruta & "\asorden", vbDirectory) = "" Then
                On Error Resume Next
                MkDir Miruta & "\asorden"
                On Error GoTo 0
            End If
            If Dir(Miruta & "\asorden\off.ord") <> "" Then
                On Error Resume Next
                Kill Miruta & "\asorden\off.ord"
                On Error GoTo 0
            End If
            If Dir(Miruta & "\asorden\gtotal.ord") <> "" Then
                On Error Resume Next
                Kill Miruta & "\asorden\gtotal.ord"
                On Error GoTo 0
            End If
            frmControl.TmrCierre.Interval = 10000
            frmControl.TmrCierre.Enabled = True
            frmControl.TmrGranTotal.Interval = 10000
            frmControl.TmrGranTotal.Enabled = True
        End If
        TaquionFichero = ConfiguracionINI(57)
        TaquionTiquets = ConfiguracionINI(58)
        Iconificar = ConfiguracionINI(59)
        IconificarInicio = ConfiguracionINI(60)
        LogArticulos = ConfiguracionINI(61)
        DebugActivo = ConfiguracionINI(62)
        ComandoLibre(0) = ConfiguracionINI(63)
        EjecutarLibre(0) = ConfiguracionINI(64)
        ComandoLibre(1) = ConfiguracionINI(65)
        EjecutarLibre(1) = ConfiguracionINI(66)
        ComandoLibre(2) = ConfiguracionINI(67)
        EjecutarLibre(2) = ConfiguracionINI(68)
        ComandoLibre(3) = ConfiguracionINI(69)
        EjecutarLibre(3) = ConfiguracionINI(70)
        ComandoLibre(4) = ConfiguracionINI(71)
        EjecutarLibre(4) = ConfiguracionINI(72)
        UsarPantallas = ConfiguracionINI(73)
        PuertoPantallas = ConfiguracionINI(74)
        UsaDual = ConfiguracionINI(75)
        '***************************
        ' a partir de aquí mantener
        ' compatibilidad de versiones
        '***************************
        DiasBorrado = 0
        BorrarBackup = 0
        Exp_Fecha = 0
        Exp_hora = 0
        Exp_Absoluto = False
        Exp_Path = Miruta
        Exp_Diario = False
        TqNombre = "BBTTTT"
        TqDirecto = True
        If ConfiguracionINI(76) <> "" Then DiasBorrado = ConfiguracionINI(76)
        If Dir(App.Path & "\upper.txt") <> "" And DiasBorrado = 0 Then
            lReW = True
            DiasBorrado = 5
        Else
            lReW = False
        End If

        If ConfiguracionINI(77) <> "" Then BorrarBackup = ConfiguracionINI(77)
        'If ConfiguracionINI(79) <> ""        Then
        Exp_Fecha = ConfiguracionINI(78)
        'If ConfiguracionINI(80) <> "" Then
        Exp_hora = ConfiguracionINI(79)
        Exp_Mas = ConfiguracionINI(80)
        Exp_Menos = ConfiguracionINI(81)
        If ConfiguracionINI(82) <> "" Then
            If ConfiguracionINI(82) = "1" Then
                Exp_Absoluto = True
            Else
                Exp_Absoluto = False
            End If
        End If
        If Trim(ConfiguracionINI(83)) <> "" Then Exp_Path = ConfiguracionINI(83)
        'If ConfiguracionINI(84) <> "" Then
        Exp_Diario = ConfiguracionINI(84)
        If ConfiguracionINI(85) <> "" Then TqNombre = ConfiguracionINI(85)
        If ConfiguracionINI(86) <> "" Then
            If ConfiguracionINI(86) = "0" Then TqDirecto = False
        End If
        If ConfiguracionINI(87) <> "" Then
            If ConfiguracionINI(87) = "1" Then
                VacunoD = True
            Else
                VacunoD = False
            End If
        Else
            VacunoD = False
        End If
        If ConfiguracionINI(88) <> "" Then
            If ConfiguracionINI(88) = "1" Then
                bunTicket = True
            Else
                bunTicket = False
            End If
        Else
            bunTicket = False
        End If
        sunTicket = ConfiguracionINI(89)
        If ConfiguracionINI(90) <> "" Then
            If ConfiguracionINI(90) = "1" Then
                GA_Reinit = True
            Else
                GA_Reinit = False
            End If
        Else
            GA_Reinit = True
        End If
        If ConfiguracionINI(91) <> "" Then
            TipoC9 = ConfiguracionINI(91)
            If TipoC9 Then SC10Rs232 = True
        Else
            TipoC9 = False
        End If
        If ConfiguracionINI(92) <> "" Then
            ModoC9 = ConfiguracionINI(92)
        End If
        CheckNoTiquet = False
        If ConfiguracionINI(93) <> "" Then
            If ConfiguracionINI(93) <> "0" Then CheckNoTiquet = True
        End If
        ClienteD = False
        If ConfiguracionINI(94) <> "" Then
            If ConfiguracionINI(94) <> "0" Then ClienteD = True
        End If
        '***************
        If ConfiguracionINI(95) <> "" Then
            PuertoModem = Val(ConfiguracionINI(95))
            If PuertoModem < 0 Or PuertoModem > 4 Then PuertoModem = 1
        Else
            PuertoModem = 1
        End If
        If ConfiguracionINI(96) <> "" Then
            VelocidadModem = Val(ConfiguracionINI(96))
            If VelocidadModem < 0 Or VelocidadModem > 4 Then VelocidadModem = 1
        Else
            VelocidadModem = 1
        End If
        If ConfiguracionINI(97) <> "" Then
            TelefonoModem = ConfiguracionINI(97)
            If Len(TelefonoModem) > 50 Then TelefonoModem = ""
        Else
            TelefonoModem = ""
        End If
        If ConfiguracionINI(98) <> "" Then
            CadModem1 = ConfiguracionINI(98)
        Else
            CadModem1 = ""
        End If
        If ConfiguracionINI(99) <> "" Then
            CadModem2 = ConfiguracionINI(99)
        Else
            CadModem2 = ""
        End If
        If ConfiguracionINI(100) <> "" Then
            CadModem3 = ConfiguracionINI(100)
        Else
            CadModem3 = ""
        End If
        If ConfiguracionINI(101) = "1" Then
            ColgarModem = True
        Else
            If ConfiguracionINI(101) = "" Then
                ColgarModem = True
            Else
                ColgarModem = False
            End If
        End If
        If ConfiguracionINI(102) = "1" Then
            UsaModem = True
        Else
            UsaModem = False
        End If
        If ConfiguracionINI(103) = "1" Then
            Pasarela = True
        Else
            Pasarela = False
        End If
        If ConfiguracionINI(104) = "" Then
            PasarelaHost = "127.000.000.001"
        Else
            PasarelaHost = ConfiguracionINI(104)
        End If
        If ConfiguracionINI(105) = "" Then
            PasarelaPuerto = 32340
        Else
            PasarelaPuerto = Val(ConfiguracionINI(105))
        End If
        AbsorverCambios = False
        If ConfiguracionINI(106) = "1" Then AbsorverCambios = True
        If ConfiguracionINI(107) = "1" Then
            RechazarPlu0 = True
        Else
            'If Not Sabeco Then
            RechazarPlu0 = False
            'Else
            '    RechazarPlu0 = True
            'End If
        End If
        '///////////////////////////////////////
        '// pedido / inventario               //
        '///////////////////////////////////////
        If ConfiguracionINI(108) = "" Or ConfiguracionINI(108) = "0" Then
            SolicitarPedido = False
        Else
            SolicitarPedido = True
        End If
        If ConfiguracionINI(109) = "" Or ConfiguracionINI(109) = "0" Then
            BorrarPedido = False
        Else
            BorrarPedido = True
        End If
        If ConfiguracionINI(110) = "" Or ConfiguracionINI(110) = "1" Then
            'If Not Sabeco Then
            BorrarPedidoFin = True
            'Else
            '    BorrarPedidoFin = False
            'End If
        Else
            BorrarPedidoFin = False
        End If
        If ConfiguracionINI(111) = "" Or ConfiguracionINI(111) = "0" Then
            SolicitarInventario = False
        Else
            SolicitarInventario = True
        End If
        If ConfiguracionINI(112) = "" Or ConfiguracionINI(112) = "0" Then
            BorrarInventario = False
        Else
            BorrarInventario = True
        End If
        If ConfiguracionINI(113) = "" Or ConfiguracionINI(113) = "1" Then
            'If Not Sabeco Then
            BorrarInventarioFin = True
            'Else
            '    BorrarInventarioFin = False
            'End If
        Else
            BorrarInventarioFin = False
        End If
        If ConfiguracionINI(114) = "" Or ConfiguracionINI(114) = "0" Then
            SolicitarPedidoFin = False
        Else
            SolicitarPedidoFin = True
        End If
        If ConfiguracionINI(115) = "" Or ConfiguracionINI(115) = "0" Then
            SolicitarInventarioFin = False
        Else
            SolicitarInventarioFin = True
        End If
        UsaECO1 = False
        TipoV8 = False
        UsaECO2 = False
        '2.0.9
        leco9600 = False
        leco19200 = False
        lMarteNew = False
        lMarteAlfa = False
        l49T = False
        l66T = False
        '''''''
        MaquinaV8 = 1
        MaquinaLabel = 1
        PuertoV8 = 1
        PuertoLabel = 1
        UsaEcoPrint = False
        TipoEcoPrint = 0
        PuertoEcoPrint = 1
        EstiloInterfaz = 0
        EcoSoloUna = False
        PedidoIV4 = False
        EnvioContinuo = False
        MultiEuroscale = False
        SC104Lineas = False
        FamiliasFijas = True
        SubSeccionesFijas = False
        CommRetries = 3
        SC10TCPIP = False
        SC10IP = "127.0.0.1"
        SC10tcp = 6000
        If ConfiguracionINI(116) <> "" Then UsaECO1 = ConfiguracionINI(116)
        If W_5_2_2 Then
            SC104Lineas = UsaECO1
            UsaECO1 = False
        End If
        TipoV8 = True
        'If ConfiguracionINI(117) <> "" Then TipoV8 = Not CBool(ConfiguracionINI(117))
        If ConfiguracionINI(118) <> "" Then UsaECO2 = ConfiguracionINI(118)
        If ConfiguracionINI(119) <> "" Then MaquinaV8 = Val(ConfiguracionINI(119))
        If ConfiguracionINI(120) <> "" Then MaquinaLabel = Val(ConfiguracionINI(120))
        If ConfiguracionINI(121) <> "" Then PuertoV8 = Val(ConfiguracionINI(121)) + 1
        If ConfiguracionINI(122) <> "" Then PuertoLabel = Val(ConfiguracionINI(122)) + 1
        If ConfiguracionINI(123) <> "" Then UsaEcoPrint = ConfiguracionINI(123)
        If ConfiguracionINI(124) <> "" Then TipoEcoPrint = ConfiguracionINI(124)
        If ConfiguracionINI(125) <> "" Then PuertoEcoPrint = ConfiguracionINI(125) + 1
        If ConfiguracionINI(126) <> "" Then EstiloInterfaz = ConfiguracionINI(126)
        If ConfiguracionINI(127) = "1" Then EcoSoloUna = True
        If ConfiguracionINI(128) = "1" Then PedidoIV4 = True
        If ConfiguracionINI(129) = "1" Then EnvioContinuo = True
        If ConfiguracionINI(130) = "1" Then MultiEuroscale = True
        If ConfiguracionINI(131) = "1" Then SC104Lineas = True
        If ConfiguracionINI(132) = "1" Then FamiliasFijas = True
        FamiliasFijas = True
        If ConfiguracionINI(133) <> "" Then CommRetries = Val(ConfiguracionINI(133))
        If ConfiguracionINI(134) = "1" Then SubSeccionesFijas = True
        If ConfiguracionINI(135) <> "" Then SC10TCPIP = ConfiguracionINI(135)
        If ConfiguracionINI(136) <> "" Then SC10IP = ConfiguracionINI(136)
        If ConfiguracionINI(137) <> "" Then SC10tcp = Val(ConfiguracionINI(137))
        If ConfiguracionINI(138) <> "" Then var.C9Cod = CInt(ConfiguracionINI(138))
        versionG = False
        If (ConfiguracionINI(139) = "1") Then versionG = True
        '2.0.9
        If Contador > 140 Then
            If (ConfiguracionINI(140) = "1") Then leco9600 = True
            If (ConfiguracionINI(141) = "1") Then leco19200 = True
            If (ConfiguracionINI(142) = "1") Then lMarteNew = True
            If (ConfiguracionINI(143) = "1") Then lMarteAlfa = True
            If (ConfiguracionINI(144) = "1") Then l49T = True
            If (ConfiguracionINI(145) = "1") Then l66T = True
        Else
            leco9600 = False
            leco19200 = False
            lMarteNew = False
            lMarteAlfa = False
            l49T = False
            l66T = False
            If Dir(App.Path & "\marte.new") <> "" Then
                lMarteNew = True
            End If
            If Dir(App.Path & "\eco9600.txt") <> "" Then
                leco9600 = True
                Kill App.Path & "\eco9600.txt"
            End If
            If Dir(App.Path & "\eco19200.txt") <> "" Then
                leco19200 = True
                Kill App.Path & "\eco19200.txt"
            End If

        End If
        '2.0.23
        If Contador > 145 Then
            If ConfiguracionINI(146) = "" Then
                lReW = True
            End If
            If (ConfiguracionINI(146) = "1") Then
                lVL000 = True
                FrmConfiguracion2.ChkL000.Value = vbChecked
            Else
                lVL000 = False
                FrmConfiguracion2.ChkL000.Value = vbUnchecked
            End If
        Else
            lVL000 = False
            FrmConfiguracion2.ChkL000.Value = vbUnchecked
            lReW = True
        End If
        ''''''''''''
        If SC10TCPIP Then SC10Rs232 = True
        If CommRetries < 3 Then CommRetries = 3
    End If
    'If Sabeco Then RechazarPlu0 = True
    If Not HayMulti Then CargaIdiomas
    If (Reescribe Or lReW) Then EscribirParametrosInicio
    Lee_Multi_INI
End Sub
Private Sub Lee_Multi_INI()
    Dim Arch As Integer
    Dim Buffer(9) As String
    Dim bucle As Integer
    Dim Buf As String
    If Dir(App.Path & "\hydramulti.ini") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\hydramulti.ini" For Input As #Arch
        For bucle = 0 To 8
            Input #Arch, Buffer(bucle), Buf
        Next bucle
        Close #Arch
        'c2f 2.0.6-
        'se toman de hydratouch.ini
        Exp_Fecha = Buffer(3)
        Exp_hora = Buffer(4)
        Exp_Mas = Buffer(5)
        Exp_Menos = Buffer(6)
        Exp_Absoluto = Buffer(7)
        Exp_Diario = Buffer(8)
        ''''''''''''
    End If
End Sub
'Private Sub Los_Parametros_Defecto_Ahorramas()
'            Empresa = "Ahorramas"
'            usaSC10 = True
'            UsaFichasVacuno = True
'            UsaGruposdeConservacion = True
'            pathSC10 = "z:\balanzas"
'            GA_RecibirTotalVenta = False
'            GA_RecibirTotalVentaL = False
'            SC10_Borrar_Ticket_dat = True
'            SC10_Borrar_Eti_dat = True
'            FamiliaEspecial = ""
'            EtiquetaEspecial = ""
'            '2.0.39 ... a partir de esta versión no se sincroniza dbteclas.dbf...
'            'PathDBF = "C:\Gestion"
'            PathDBF = ""
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            Usa20Lineas = True
'            SC10Rs232 = False
'            SC10ParametrosSerie = "19200,E,8,1,2"
'
'            'Iconificar = True
'            'IconificarInicio = True
'            'TaquionFichero = ""
'            'Taquion = True
'            'descAuto = True
'            'descautotime = "45  sec"
'            SC10_BorrarTotal = True
'            'SC10_ProcesarTotalVenta = True
'            SC10_ProcesarTotalVenta = False
'            DiasBorrado = 7
'End Sub
'Private Sub Los_parametros_Defecto_Ahold()
'            Empresa = "Dinosol"
'            UsaGamaAlta = False
'            usaSC10 = True
'            gedit = 0
'            tcpip = True
'            pathSC10 = "z:\sc10"
'            frmControl.timetqt.Enabled = False
'            frmControl.timetqt.Interval = 0
'            GA_RecibirTotalVenta = False
'            GA_RecibirTotalVentaL = False
'            SC10_BorrarTotal = True
'            SC10_ProcesarTotalVenta = True
'            SC10_Borrar_Ticket_dat = True
'            SC10_Borrar_Eti_dat = True
'            UsaBackup = False
'            pathred = ""
'            FamiliaEspecial = ""
'            EtiquetaEspecial = ""
'            PathDBF = ""
'            Usa20Lineas = True
'            SC10Rs232 = False
'            SC10ParametrosSerie = "19200,E,8,1,2"
'            DebugActivo = True
'            LogArticulos = True
'            DiasBorrado = 10
'End Sub
Private Sub Los_parametros_Defecto()
    id = 0
    HaySeguridad = False
    Empresa = ""
    UsaGamaAlta = True
    usaSC10 = False
    UsaFichasVacuno = False
    UsaGruposdeConservacion = False
    gedit = 0
    tcpip = True
    cngvelocidad = "19200,E,8,1,1"
    sPortBal = 3306
    pathSC10 = "z:\sc10"
    descAuto = False
    descautotime = ""
    frmControl.timetqt.Enabled = False
    frmControl.timetqt.Interval = 0
    GA_RecibirTotalVenta = True
    GA_RecibirTotalSuper = False
    GA_RecibirTotalEnvasado = False
    GA_RecibirTotalAutoservicio = False
    GA_RecibirTotalVentaL = True
    GA_RecibirTotalSuperL = False
    GA_RecibirTotalEnvasadoL = False
    GA_RecibirTotalAutoservicioL = False
    GA_BorrarTotal = True
    SC10_BorrarTotal = False
    SC10_ProcesarTotalVenta = True
    SC10_ProcesarTotalSuper = False
    SC10_ProcesarTotalEnvasado = False
    SC10_ProcesarTotalAutoservicio = False
    SC10_Borrar_Ticket_dat = False
    SC10_Borrar_Eti_dat = False
    RecibirTotalVacuno = False
    BorrarTotalVacuno = False
    decimales = 0
    SimboloMonetario = "Pta."
    '********************
    ' Obsoleto
    ValorEuro = 1
    '********************
    UsaEuro = True
    UsaBackup = False
    pathred = ""
    FamiliaEspecial = "0"
    EtiquetaEspecial = 0
    PathDBF = ""
    Usa20Lineas = True
    SC10Rs232 = False
    SC10ParametrosSerie = "19200,E,8,1,2"
    CI_HORA = 25
    CI_MINUTO = 61
    RE_HORA = 25
    RE_MINUTO = 61
    CI_HOST = "127.000.000.001"
    CI_TCP = 5580
    CI_UDP = 6001
    If Dir(App.Path & "\hydraserver.exe") <> "" Then
        StoreFlow = True
    Else
        StoreFlow = False
    End If
    Taquion = False
    TaquionFichero = "artcon.dat"
    TaquionTiquets = True

    Iconificar = False
    IconificarInicio = False
    DebugActivo = False
    LogArticulos = False
    ComandoLibre(0) = ""
    EjecutarLibre(0) = ""
    ComandoLibre(1) = ""
    EjecutarLibre(1) = ""
    ComandoLibre(2) = ""
    EjecutarLibre(2) = ""
    ComandoLibre(3) = ""
    EjecutarLibre(3) = ""
    ComandoLibre(4) = ""
    EjecutarLibre(4) = ""
    UsarPantallas = False
    PuertoPantallas = 32337
    DiasBorrado = 0
    Exp_Fecha = 0
    Exp_hora = 0
    Exp_Mas = "+"
    Exp_Menos = "-"
    Exp_Absoluto = False
    Exp_Path = Miruta
    Exp_Diario = True
    bunTicket = False
    sunTicket = ""
    GA_Reinit = True
    CheckNoTiquet = False
    PuertoModem = 1
    VelocidadModem = 1
    TelefonoModem = ""
    CadModem1 = ""
    CadModem2 = ""
    CadModem3 = ""
    ColgarModem = True
    Pasarela = False
    PasarelaHost = "127.000.000.001"
    PasarelaPuerto = 32340
    AbsorverCambios = False
    BorrarPedidoFin = True
    BorrarInventarioFin = True
    CommRetries = 3
    FamiliasFijas = True
    SubSeccionesFijas = False
End Sub

'Private Sub Los_Parametros_Defecto_Sabeco()
'    Exp_Fecha = 6
'    Exp_hora = 1
'    Exp_Mas = "0"
'    Exp_Menos = "1"
'    Exp_Absoluto = True
'    Exp_Diario = True
'   BorrarPedidoFin = False
'   BorrarInventarioFin = False
'End Sub
'Private Sub Los_Parametros_Defecto_StoreFlow()
'
'    RE_HORA = 22
'    RE_MINUTO = 15
'    CI_HORA = 22
'    CI_MINUTO = 15
'    descAuto = True
'    descautotime = "10  sec"
'    GA_BorrarTotal = True
'    SC10_BorrarTotal = True
'    RechazarPlu0 = True
'    FamiliaEspecial = ""
'    EtiquetaEspecial = ""
'    BorrarBackup = "3"
'    DiasBorrado = "3"
'    SC10Rs232 = True
'    TipoC9 = False
'    UsaBackup = True
'    pathred = App.Path & "\backup"
'    AbsorverCambios = True
'End Sub
'Private Sub Los_Parametros_Defecto_Supeco()
'    Exp_Path = App.Path
'    Exp_Fecha = 4
'    Exp_hora = 1
'    Exp_Diario = False
'    bunTicket = True
'    sunTicket = "supeco.tiq"
'    descAuto = True
'    descautotime = "20  sec"
'    Iconificar = True
'    IconificarInicio = True
'    TqDirecto = False
'    Taquion = True
'    TaquionFichero = "art.dat"
'    UsaBackup = True
'    pathred = App.Path & "\backup"
'    BorrarBackup = "3"
'    DiasBorrado = "3"
'    FamiliaEspecial = ""
'    EtiquetaEspecial = ""
'    ClienteD = True
'    VacunoD = True
'End Sub

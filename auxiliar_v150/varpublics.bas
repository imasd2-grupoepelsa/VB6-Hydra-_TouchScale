Attribute VB_Name = "varpublics"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public sPathNGT As String
Public sOrdenNGT As String
Public nlastGT As Long
Public lngt As Boolean

Public Filtro0(8) As Integer

'/////////////////////////////////////////////
'/////////////////////////////////////////////
'/////////////////////////////////////////////
' Modificaciones Pendientes
'LSB --> Mostradores
'2   --> Equipos
'3   --> Subsecciones
'4   --> Familias
'5   --> Paises
'6   --> Fichas
'7   --> Iva
'8   --> Artículos
'9   --> Vendedores
'10  --> Publicidad
'11  --> Cabeceras y Leyendas
'12  --> Teclas
'13  --> Códigos de Barras
'14
'15
'MSB
'///
' checknotiquet, si está activado,
' tras los 5 intentos de recepción
' con descuadre, no borra totales
Global CheckNoTiquet As Boolean
'******************************
Global TipoC9 As Boolean
Global ModoC9 As Integer
Public NombreTienda As String
Global VacunoD As Boolean
'*********************
' cancelhumano, para
' que no se cancele
' recepción de tiquets
Public CancelHumano As Boolean
'************************
Public HayComandos As Boolean
Global ModificacionesPendientes(3) As Byte
'**********
' Ordenes remotas
' desde pantallas
' clientes
Global DiasBorrado As Integer
Global UsuarioRemoto As Tipo_User
Global UsarPantallas As Boolean
Public PuertoPantallas As Long
Public OrdenRemota As String
Public RespuestaRemota As String
'***********
' Plugin
Public Type TPlugin
    Mode As String
    type As String
    Nentradas As Integer
    NombreMenu As String
    NombreEntrada(10) As String
    accion(10) As String
End Type
Public DataPlugin As TPlugin
Public PlugActivo As Boolean
Public PlugStatus As Integer
Public PlugOculta As Boolean
'**********
' Seguridad
Public Type Tipo_User
    Nombre As String
    Password As String
    Permisos(8) As Byte
    Asociado As Integer
End Type
Dim CantidadUsuarios As Integer
Global usuarios() As Tipo_User
Global UsuarioActual As Tipo_User
Global HaySeguridad As Boolean
'*****************************
Global DebugActivo As Boolean
Global LogArticulos As Boolean
'*****************************
' comandos libres
Public ComandoLibre(5) As String
Public EjecutarLibre(5) As String
'******************
' Master / Slave
Global SoyMaster As Boolean
Global SoySlave As Boolean
Global SoyReserva As Boolean
    Public MS_LPort As Integer
    Public MS_RPath As String
    Public MS_RHost As String
    Public MS_Salir As Boolean
    Public MS_Lock As Boolean
'''públicas
'********************************************
' ESPACIO RESERVADO
' INTRODUCIR AQUI LAS VARIABLES ESPECÍFICAS
' QUE DEFINEN A UN CLIENTE
' las variables serán booleanas, y actuarán
' como interruptor de diversas acciones
' En el arranque del programa, lee la existencia
' de ficheros de configuración específicos,
' determinando el valor (1/0) de estos
' interruptores
'********************************************
Global AhorraMas As Boolean
    Public PathDBF As String
    Public RecibirTotalVacuno As Boolean
    Public BorrarTotalVacuno As Boolean
    Public GA_RecibirTotalVenta As Boolean
    Public GA_RecibirTotalSuper As Boolean
    Public GA_RecibirTotalEnvasado As Boolean
    Public GA_RecibirTotalAutoservicio As Boolean
    Public GA_RecibirTotalVentaL As Boolean
    Public GA_RecibirTotalSuperL As Boolean
    Public GA_RecibirTotalEnvasadoL As Boolean
    Public GA_RecibirTotalAutoservicioL As Boolean
    Public GA_BorrarTotal As Boolean
    Public GA_Reinit As Boolean
    Public SimboloMonetario As String

Global Taquion As Boolean
    Public TaquionFichero As String
    Public TaquionTiquets As Boolean
    Public TqStatus As Integer
    Public TqNombre As String
    Public TqDirecto As Boolean
    Public bunTicket As Boolean
    Public sunTicket As String
Global Alcampo As Boolean

Global Sabeco As Boolean
'********************************************
' FIN ESPACIO RESERVADO
'********************************************

Public EL_BS As String
Public EnvioDirecto As Boolean
Public FamiliaEspecial As String
Public EtiquetaEspecial As String
Public SC10Rs232 As Boolean
Public SC10ParametrosSerie As String
Public NumMaxTiquets As Integer
Public Silencio As Boolean
Public Usa20Lineas As Boolean
Public UsaGruposdeConservacion As Boolean
Public UsaFichasVacuno As Boolean
Public MiruTa As String
Public id
Public Ascbin

Public puerto As Integer

Public cngvelocidad As String
Public pathred As String
Public pathlocal As String
Public UsaGamaAlta As Boolean
' *************************
' Ahorramas

' **************************
' Daniel Campos
' 19-7-2001
' Añade una variable para el path de red de SC10
Public pathSC10 As String
' Fin de corrección 19-7-2001



Public tcpip As Boolean

Public decimales
Public user As String
Public usersec As String
Public Asgr As String
Public falloBorrado As Boolean
Public usersecven As String
Public usersecvenflt As String
Public usec As String
Public admin As Boolean
Public LoginSucceeded As Boolean
Public esgedit As String
Public usaSC10 As Boolean
'C2F(1)... variable para almacenar dirección IP Balanza
'      ... comunicaciones binarias
Public sAddIp As String
Public sPortBal As String
Public lIp As Boolean
Public globalrecibe As String
Public lSend As Boolean
Public lRecS As Boolean
Public PathOrden As String
Public lEsEnvBin As Boolean
Public sBinSx As String
Public nSeccmaqui As Integer
Public terMinaip As Integer
Public canCelar As Integer

Public UsaBackup As Boolean
Public BorrarBackup As Integer
Public lCogeTiquet As Boolean
Public cgdtiquet As Boolean
Public descAuto As Boolean
Public descautotime As String
Public cuenTaTqt As Integer
Public cuenta     As Integer
Public Empresa As String
Public gedit As String

Public lReinit As Boolean

Public SC10_ProcesarTotalVenta As Boolean
Public SC10_ProcesarTotalSuper As Boolean
Public SC10_ProcesarTotalEnvasado As Boolean
Public SC10_ProcesarTotalAutoservicio As Boolean
Public SC10_BorrarTotal As Boolean
Public SC10_Borrar_Ticket_dat As Boolean
Public SC10_Borrar_Eti_dat As Boolean
Public ModoRecogidaContinua As Boolean

Public Iconificar As Boolean
Public IconificarInicio As Boolean
'****************
' Euro
Public UsaDual As Boolean
Public UsaEuro As Boolean
Public ValorEuro As Double
''''''''''''
'*******************************
Global SuspenderSC10 As Boolean
Global suspenderGA As Boolean
Global LockBase As Boolean
Global Base_General As String


Global PausaTiquets As Boolean
'**************************
' hora y minuto de cierre
'**************************
Global CI_HORA As Integer
Global CI_MINUTO As Integer
'**************************
' hora y minuto de recogida de totales
'**************************
Global RE_HORA As Integer
Global RE_MINUTO As Integer
Global RE_TOTAL As Boolean
'**************************
' Ajuste de Hora
'**************************
Global CI_DobleHora As Boolean
Global RE_DobleHora As Boolean
'******************
' enlace Corte Inglés - Store Flow
'
Global StoreFlow As Boolean
Global CI_HayEuroscale As Boolean

Global CI_HOST As String
Global CI_TCP As Long
Global CI_UDP As Long
Global CI_ProximaRecogida As Boolean

Global CI_Datagrama As String
Global CI_Paquete As String
Global CI_Modificaciones As Boolean



Global CI_Salir As Boolean
'*********************
' CI_BufferTiquet almacena
' la cadena con el tiquet a enviar
' a storeflow
Global CI_BufferTiquet As String
'*************************
' CI_Exito, determina si se pudo
' enviar un tiquet a Storeflow
'*************************
Global CI_Exito As Boolean
Global CI_Autonomo As Boolean
'*************************
' Estado del sistema
' 1 --> ORDEN / Esperando UDP "ORDEN"
' 2 --> ORDEN / Recibido datagrama, Abriendo TCP
' 3 --> ORDEN / TCP abierto, enviando trama "11"
' 4 --> ORDEN / Enviado "11", esperando la orden del StoreFlow
' 5 --> ORDEN / Recibida la orden del storeflow, Procesando
' *************************
'

Global CI_EstadoSistema As Integer
'*****************************
' Exportación
'*****************************
Global Exp_Fecha As Integer
Global Exp_hora As Integer
Global Exp_Mas As String
Global Exp_Menos As String
Global Exp_Absoluto As Boolean
Global Exp_Path As String
Global Exp_Diario As Boolean
'*****************************
' ML100 --> Si el peso está
' entre 15 y 30 kg, sólo
' proporciona dos decimales
' de peso, y hay que corregirlo
'*****************************
Global ML100 As Boolean
Global FamiliasFijas As Boolean
Global SubseccionesFijas As Boolean

Global version As String

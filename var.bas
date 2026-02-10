Attribute VB_Name = "var"
Option Explicit
'//////////////
'\\\\\\\\\\\\\\
Public Const WHOIAM = "HYDRATOUCH.EXE"

Public salir As Integer
Public myPER1 As Date
Public myPER2 As Date
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
'''''''''''''''''''''''''''''''''''''''''''''''
'Api FindWindowEx
Public Declare Function FindWindowEx Lib "user32" _
                                     Alias "FindWindowExA" _
                                     (ByVal hWndParent As Long, _
                                      ByVal hWndChildAfter As Long, _
                                      ByVal lpClassName As String, _
                                      ByVal lpWindowName As String) As Long

'Api GetParent
Public Declare Function GetParent Lib "user32" _
                                  (ByVal hWnd As Long) As Long

'Api SendMessage
Public Declare Function SendMessage Lib "user32" _
                                    Alias "SendMessageA" _
                                    (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

'/////////////////////////////////////////////version de programa

''''''''''''''''''''''''''''''''''''''''''''''
Public OrdenMulti As String
Public lVL000 As Boolean
'Global Const miver As String = "Release 1.3.4"
Public miver As String
Public EstadoBal As Integer
Public MyIP As String
Public FamMaqui(99) As Integer

Public MultiIP As String
Public Modem_Marcando As Boolean
Public Exclusivo_Fin As Boolean
Public Enlazado As Boolean
Public AutoLanzado As Boolean
Public MaquinaDesconectada As Integer
Public MaquinaConectada As Integer
Public PuertoConectado As Integer

'1.error en form configuracion.
'corrige problema de desasociacion de usuario
'a mostrador
'2.se incorporan funciones de nuevos totales
'de vendedores
'/////////////
'*************
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
'*************************************
' checknotiquet, si está activado,
' tras los 5 intentos de recepción
' con descuadre, no borra totales
' borrar pedido en cada recepción de tiquets
'////////////////////////////////////////
' // Pedido / Inventario
Global EnvioContinuo As Boolean
Global BorrarPedido As Boolean
Global BorrarInventario As Boolean
' solicitar pedido de forma continua
Global SolicitarPedido As Boolean
Global SolicitarInventario As Boolean
' solicitar pedido al final del día
Global SolicitarPedidoFin As Boolean
Global SolicitarInventarioFin As Boolean
' borrar pedido al final del día
Global BorrarPedidoFin As Boolean
Global BorrarInventarioFin As Boolean
' //
'////////////////////////////////////////
Global RechazarPlu0 As Boolean
Global AbsorverCambios As Boolean
Global AHOLD As Boolean
Global CheckNoTiquet As Boolean
'******************************
Global TipoC9 As Boolean
Global C9Cod As Integer
Global ModoC9 As Integer
Public NombreTienda As String
Global VacunoD As Boolean
Global ClienteD As Boolean
Global versionG As Boolean
'*************************
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
Public MultiEuroscale As Boolean    ' Varias servidoras (una por sección)
Public PuertoPantallas As Long
Public OrdenRemota As String
Public RespuestaRemota As String
'leclerc
Public Type RegMaj
    Tipo As String
    seccion As String
    codigo As String
    des1 As String
    precio As String
    paridad As String
    WGH As String
    ean As String
    cad As String
    cad2 As String
    desresto As String
    tipotara As String
    tara As String
    parlinea As String
    pesofijo As String
    tecla As String
End Type

Public Type RegMajNew
    Tipo As String
    seccion As String
    codigo As String
    des1 As String
    des2 As String
    des3 As String
    des4 As String
    precio As String
    'paridad As String
    WGH As String
    ean As String
    cad As String
    cad2 As String
    desresto As String
    tipotara As String
    tara As String
    cod_ingre As String
    'parlinea As String
    pesofijo As String
    'tecla As String
    promocion As String
    Etiqueta As String
End Type

'*******************************
' Plugin
Public Type TPlugin
    mode As String
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
'***************************
' Seguridad
Public Type Tipo_User
    Nombre As String
    password As String
    Permisos(8) As Byte
    Asociado As Integer
End Type
Dim CantidadUsuarios As Integer
Global usuarios() As Tipo_User
Global UsuarioActual As Tipo_User
Global HaySeguridad As Boolean
Global DebugActivo As Boolean
Global LogArticulos As Boolean
'*****************************
' enlace Supeco
Global Supeco As Boolean
Global Supeco_Path As String    ' path para depositar tickets
Global Supeco_Art As String    ' path para recoger artibal
'*****************
' comandos libres
Public ComandoLibre(5) As String
Public EjecutarLibre(5) As String
'******************
' Master / Slave
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

'Global Sabeco As Boolean
'**********************
' FIN ESPACIO RESERVADO
'**********************
Public FamiliaEspecial As String
Public EtiquetaEspecial As String
Public SC10IP As String
Public SC10tcp As Long
Public SC10TCPIP As Boolean
Public SC10Rs232 As Boolean
Public SC10ParametrosSerie As String
Public SC104Lineas As Boolean
Public NumMaxTiquets As Integer
Public Silencio As Boolean
Public Usa20Lineas As Boolean
Public UsaGruposdeConservacion As Boolean
Public UsaFichasVacuno As Boolean
Public Miruta As String
Public id
Public cngvelocidad As String
Public pathred As String
Public UsaGamaAlta As Boolean
Public UsaECO1 As Boolean
Public TipoV8 As Boolean
Public PuertoV8 As Integer
Public PuertoLabel As Integer
Public MaquinaV8 As Integer
Public UsaECO2 As Boolean
'2.0.9
Public leco9600 As Boolean
Public leco19200 As Boolean
Public lMarteNew As Boolean
Public lMarteAlfa As Boolean
Public l49T As Boolean
Public l66T As Boolean
''''''
Public MaquinaLabel As Integer
Public UsaEcoPrint As Boolean
Public TipoEcoPrint As Integer
Public PuertoEcoPrint As Integer
Public EcoSoloUna As Boolean
Public PedidoIV4 As Boolean
' ************************
' Ahorramas
' ************************
' Daniel Campos
' 19-7-2001
' Añade una variable para el path de red de SC10
Public pathSC10 As String
' Fin de corrección 19-7-2001
Public tcpip As Boolean
Public decimales
Public admin As Boolean
Public LoginSucceeded As Boolean
Public usaSC10 As Boolean
'C2F(1)... variable para almacenar dirección IP Balanza
'      ... comunicaciones binarias
Public sAddIp As String
Public sPortBal As String
Public lIp As Boolean
Public globalrecibe As String
Public PathOrden As String

Public sBinSx As String
Public nSeccMaqui As Integer
Public canCelar As Integer
Public UsaBackup As Boolean
Public BorrarBackup As Integer
Public lCogeTiquet As Boolean
Public cgdtiquet As Boolean
Public descAuto As Boolean
Public descautotime As String
Public cuenTaTqt As Integer
Public Empresa As String
Public gedit As String
Public SC10_ProcesarTotalVenta As Boolean
Public SC10_ProcesarTotalSuper As Boolean
Public SC10_ProcesarTotalEnvasado As Boolean
Public SC10_ProcesarTotalAutoservicio As Boolean
Public SC10_BorrarTotal As Boolean
Public SC10_Borrar_Ticket_dat As Boolean
Public SC10_Borrar_Eti_dat As Boolean
Public Iconificar As Boolean
Public IconificarInicio As Boolean

Public lAsigntec As Boolean

'****************
' Euro
Public UsaDual As Boolean
Public UsaEuro As Boolean
Public ValorEuro As Double
''''''''''''
'*******************************
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

'******************
' enlace Corte Inglés - Store Flow
'
Global StoreFlow As Boolean
Global CI_HOST As String
Global CI_TCP As Long
Global CI_UDP As Long
Global CI_Salir As Boolean
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
'**********************
' Modem
'**********************
Public PuertoModem As Integer
Public VelocidadModem As Long
Public TelefonoModem As String
Public CadModem1 As String
Public CadModem2 As String
Public CadModem3 As String
Public UsaModem As Boolean
Public ColgarModem As Boolean
'***********************
' Pasarela
'************************
Public Pasarela As Boolean
Public PasarelaHost As String
Public PasarelaPuerto As Long
'*************************
' Familias
'*************************
Global FamiliasFijas As Boolean
Global SubSeccionesFijas As Boolean
Global CommRetries As Long
'///////////////////////////////cas.v118
Global Const IDMAXPLU_C9COD As Integer = 6999
'\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\modulo idiomas
Public sIdioma() As String
Public nIdioma As Long
'/////////////////////////
Public nNumDecML100 As Integer
Public nNumBalML100 As Integer

'Public lastIpconn As String
'Public prevIpconn As String

'Leclerc...
Public lClR As Boolean
Public sPathMaj As String
Public lEsExport As Boolean
'c2f para 1.7.12 20/40 lineas de descripctivo
Public Type rLinTxt2040
    codigo As Long
    Plu As Long
    Mostrador As Long
    txt(19) As String
    tran_txt2 As String
    abm As Byte
End Type
''''''''
Public lintxt As rLinTxt2040
Public GrupoActual As Integer
Public sPathGigante As String
Public sUnidadGigante As String
Public lFinDia As Boolean
Public nContDiasGi As Integer
Public dArrGi(100) As Date
'1.9.0
Public lesUpdate As Boolean

Public sPathSecPlrm(60) As String
Public nContSecPlrm As Integer
Public nSecPLRM(60) As Integer
Public sFechaGesco As String
Public lTimeOut
Public sSpubS As String
Public lNoSi As Boolean
Public lastConnectionString As String
Public lAutoRecEnv As Boolean

'Public lUpperNW As Boolean 'atencionfamilia
Public Function sToUTF8(ByVal sS As String) As String
    Dim objStream As ADODB.Stream
    Dim data() As Byte
    Dim sDes As String
    Dim ncont As Integer
    If sS <> "" Then

        Set objStream = New ADODB.Stream
        objStream.Charset = "utf-8"
        objStream.mode = adModeReadWrite
        objStream.type = adTypeText
        objStream.Open

        objStream.WriteText sS
        objStream.Flush

        objStream.Position = 0
        objStream.type = adTypeBinary
        objStream.Read 3

        data = objStream.Read()
        objStream.Close

        On Error GoTo salir
        sDes = ""
        For ncont = 0 To 1000
            sDes = sDes & Chr(data(ncont))
        Next ncont

salir:
        On Error GoTo 0
    Else
        sDes = " "
    End If
    sToUTF8 = sDes

End Function


Attribute VB_Name = "var"
Option Explicit
'\\\\\\\\\\\\\\
'//////////////
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////////
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const DEBUGACTIVO = True
Public Const USAFICHASVACUNO = True
Public Const SC10_PROCESARTOTALVENTA = True
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'//////////////////////////////////////////////
'Public Const VERSION As String = "TIQAH v1.2.0"
Public version As String
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public lNoCoger As Boolean
Public lNoConnect As Boolean

'//////////////////////////////////////////////
Public cancelar As Integer
Public Const MyInterval = 2000 '1.4.1 2000

Public TimeAuto As Long
Public TimeAutoSeg As Long

Public NumMax As Long
Public PathSC10 As String
Public MiliTiquet As Long
Public PathExpor As String
Public modo As Integer
Public sAddIP As String
Public sPortBal As String
Public sInftxt1 As String
Public sInftxt2 As String
Public sInftxt3 As String
Public sInftxt4 As String
Public sInftxt5 As String
Public sInftxt6 As String
Public Base_General As String
Public lMarcaTime As Boolean
Public lTqRetail As Boolean
Public lNacho As Boolean

'\\\\\\\\\\\\\\\\\\\\\\\\\
'/////////////////////////
Global Taquion As Boolean
    Public TaquionFichero As String
    Public TaquionTiquets As Boolean
    Public TqStatus As Integer
    Public TqNombre As String
    Public TqDirecto As Boolean
    Public bunTicket As Boolean
    Public sunTicket As String
Global Sabeco As Boolean
Public UsaEuro As Boolean
Public decimales As Integer
Global TipoC9 As Boolean
'\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////
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
     ean13 As String
End Type
Public Type Tipo_NombreTiquet
    Balanza As Integer
    seccion As Integer
    tiquet As Long
    modo As Integer
    cliente As Long
    vendedor As Integer
End Type

Public sTiqGen As String
Public sHoraFin As String
Public dHoraFin As Date
Public nDelayPluAsc As Integer
Public lFinPluAsc As Boolean
Public TipoTQT(8) As String
Public nTocaLocal As Integer
Public nContLocal As Integer
Public lSerial As Boolean
Public sCfgSerial As String
Public nMaqType As Integer
Public sCOMMX As String


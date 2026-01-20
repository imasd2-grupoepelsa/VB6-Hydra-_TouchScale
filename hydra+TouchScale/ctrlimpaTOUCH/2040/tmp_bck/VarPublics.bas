Attribute VB_Name = "VarPublics"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global ImportVisible As Boolean
'*****************************************
' Espacio Reservado para los diferentes
' Supermercados
' Estas variables actúan como interruptores
' Para activar o desactivar funciones
'*****************************************

'*****************************************
' Fin de espacio Reservado
'*****************************************
' Tipo que contiene todos los datos relativos
' Al artículo
'************************************
' RESPONSE : activado de momento sólo
' para sabeco, indica línea por línea el
' resultado de la importación.
Global RESPONSE As Boolean
' /RESPONSE
'**********
Global FamiliasFijas As Boolean
Global SubseccionesFijas As Boolean
Global NoCaducidad As Boolean
Global MiProgreso As String
Global VacunoD As Boolean
Public Conf_Al As Boolean
Public LogActivo As Boolean
Public Success As Integer
Public Cancelar As Boolean
Public MostrarMensajes As Boolean
Public FamiliaVacuno As String
Public EtiquetaEsp As String
Public PathImport As String
Public PathLocal As String
Public FicheroImport As String
Public UsaEuro As Boolean
' Daniel Campos
' 24-7-2001
' Nuevas variables, para leer y configurar ficheros
' de importación
'
' Constante que determina el número de elementos
' que tiene un artículo (COD,FAM,SUB...)
Public Base_General As String
Public Const DimensionArticulo = 27
' Path que se utilizará para el fichero de importación
Public Propiedades(DimensionArticulo) As tipos
Public lineasconf() As lineaconfig
Public articulo As registroarticulos
Public Decimales As Integer
'
' Fin de modificación 24-7-2001
'
Public ValorEuro As Double
Public Type DB_Teclas
    Destino_NMostrador As Integer
    Destino_NBalanza As Integer
    Destino_Tabla As Integer
    Destino_CodiTec As Integer
    CodArticulo As Long
    PluArticulo As Integer
    SecArticulo As Integer
End Type
Public Type registroarticulos
    cod As Long
    plu As Long
    sec As Long
    sub As Long
    fam As Long
    prc As Double
    wgh As String
    cad As Long
    pre As Long
    tar As Long
    cba As String
    tlt(11) As Long
    txt(21) As String
    typ As String
    etq As Long
    gpc As Long
    trm1 As Long
    trm2 As Long
    prc1 As Double
    prc2 As Double
    iva As Integer
    tecla As Integer
    JuegoTecla As Integer
    Merma As Integer
    Presec As Integer
    prc100g As Boolean
    baja As Integer
    ' indica si vienen en el fichero
    ' ya que de no venir, en modificaciones
    ' no se introducirá ningún valor
    ' por defecto
    m_presec As Boolean
    m_merma As Boolean
    m_cod As Boolean
    m_plu As Boolean
    m_sec As Boolean
    m_sub As Boolean
    m_fam As Boolean
    m_prc As Boolean
    m_wgh As Boolean
    m_cad As Boolean
    m_pre As Boolean
    m_tar As Boolean
    m_cba As Boolean
    m_tlt(11) As Boolean
    m_txt(21) As Boolean
    m_typ As Boolean
    m_etq As Boolean
    m_gpc As Boolean
    m_trm1 As Boolean
    m_trm2 As Boolean
    m_prc1 As Boolean
    m_prc2 As Boolean
    m_iva As Boolean
    m_tecla As Boolean
    m_JuegoTecla As Boolean
    m_prc100g As Boolean
    m_baja As Boolean
End Type
' Para asignar las caracteríasticas de cada
' propiedad del artículo
Type tipos
    codigo As String
    descripcion As String
    max As Integer
End Type
' Para estudiar las líneas de un fichero
' de configuración a importar
Type lineaconfig
    codigo As String
    descripcion As String
    Inicio As Long
    fin As Long
    posicion As Integer
End Type
Global id As Integer
'///////////////////cas.v110
Global C9Cod As Integer
''''''''''''''''''''''''''''
'Ahorramás--> Se debe dar de baja un
'artículo cuando se envie a otra sección
'o exista el plu previamente.
Public lEsAho As Boolean

'Public Sub main()

    'C2f anular tras pruebas
    'frmimportar.Show
    'frmconfigurator.Show
    ''UsaEuro = False
    ''Decimales = 2
    ''PathLocal = "c:\trabajo\import_v141"
    ''PathImport = "c:\trabajo\import_v141"
    ''Base_General = "C:\trabajo\dbase.mdb"
    ''frmimportar.Show

'End Sub

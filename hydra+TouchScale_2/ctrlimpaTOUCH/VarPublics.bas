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
'Public lEsBajaMettler As Boolean
Public lEsFruteriaMettler As Boolean
Public lEsAdicional As Boolean

' Daniel Campos
' 24-7-2001
' Nuevas variables, para leer y configurar ficheros
' de importación
'
' Constante que determina el número de elementos
' que tiene un artículo (COD,FAM,SUB...)
Public Base_General As String
Public Const DimensionArticulo = 29
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
    onkey As Integer
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
    m_image As Boolean
    m_tarifa(10) As Boolean
    m_onkey As Boolean
    '17.5
    lmix As Boolean
    npeso As Integer
    imagen As String
    tarifa(10) As Double
    txt2040(100) As String   'cas.v->142
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
Public lEsBacK As Boolean
'C2F Gigante...
Public sPathGigante As String
Public sTipoImpGigante As String

Public TiendaActual As Integer
Public GrupoActual As Integer

'Public Sub main()

    'C2f anular tras pruebas
    'frmimportar.Show
    'frmconfigurator.Show
'    UsaEuro = True
'    Decimales = 2
'    PathLocal = "C:\hydra+174\import_v143"
'    PathImport = "C:\hydra+174\import_v143"
'    Base_General = "C:\hydra+174\dbasetouch.mdb"
'    frmimportar.Show ''

'End Sub
Public Sub Do_Events()
    'Sleep (1)
    DoEvents
    Sleep (1)
End Sub

Public Function V12Char(Sr As String) As String
    Dim buffer As String
    Dim Buffer2 As String
    Dim Bucle As Integer
    Buffer2 = ""
    buffer = StrConv(Sr, vbUpperCase)
    For Bucle = 1 To Len(buffer)
        Select Case Mid(buffer, Bucle, 1)
        'Case "Ñ"
        '    Buffer2 = Buffer2 & "¥"
        Case "@"
            Buffer2 = Buffer2 & "Ñ"
        Case "Á"
            Buffer2 = Buffer2 & "A"
        Case "É"
            Buffer2 = Buffer2 & "E"
        Case "Í"
            Buffer2 = Buffer2 & "I"
        Case "Ó"
            Buffer2 = Buffer2 & "O"
        Case "Ú"
            Buffer2 = Buffer2 & "U"
        Case "À"
            Buffer2 = Buffer2 & "A"
        Case "È"
            Buffer2 = Buffer2 & "E"
        Case "Ì"
            Buffer2 = Buffer2 & "I"
        Case "Ò"
            Buffer2 = Buffer2 & "O"
        Case "Ù"
            Buffer2 = Buffer2 & "U"
        Case "Ä"
            Buffer2 = Buffer2 & "A"
        Case "Ë"
            Buffer2 = Buffer2 & "E"
        Case "Ï"
            Buffer2 = Buffer2 & "I"
        Case "Ö"
            Buffer2 = Buffer2 & "O"
        Case "Ü"
            Buffer2 = Buffer2 & "U"
        Case "¥"
            Buffer2 = Buffer2 & "Ñ"
            'If Scaletype <> CLng(8) Then
            '    Buffer2 = Buffer2 & "@"
            'Else
            '    Buffer2 = Buffer2 & Chr(148)
            'End If
        'Case "Ñ"
        '    If Scaletype <> CLng(8) Then
        '        Buffer2 = Buffer2 & "@"
        '    Else
        '        Buffer2 = Buffer2 & Chr(148)
        '    End If
        Case "A" To "Z"
            Buffer2 = Buffer2 & Mid(buffer, Bucle, 1)
        Case "0" To "9"
            Buffer2 = Buffer2 & Mid(buffer, Bucle, 1)
        Case "."
            Buffer2 = Buffer2 & "."
        Case ","
            Buffer2 = Buffer2 & ","
        Case "<"
            Buffer2 = Buffer2 & "<"
        Case ">"
            Buffer2 = Buffer2 & ">"
        Case "_"
            Buffer2 = Buffer2 & "_"
        Case "-"
            Buffer2 = Buffer2 & "-"
        Case "/"
            Buffer2 = Buffer2 & "/"
        Case Else
            Buffer2 = Buffer2 & " "
        End Select
    Next Bucle
    V12Char = Buffer2
End Function
'Colores
'sColor = "#FFE0E0E0"
'sColor = "#FFC1C1C1"
'sColor = "#FFA2A2A2"
'sColor = "#FF6d6f81"
'sColor = "#FF838383"
'sColor = "#FF646464"
'sColor = "#FF454545"
'sColor = "#FF000000"
'sColor = "#FFF0FF00"
'sColor = "#FFc7de28"
'sColor = "#FFFFDA28"
'sColor = "#FFbba125"
'sColor = "#FFFFA500"
'sColor = "#FFFF7E34"
'sColor = "#FFFF7400"
'sColor = "#FF974500"
'sColor = "#FFC2FF49"
'sColor = "#FF4EFB22"
'sColor = "#FF16C700"
'sColor = "#FF00B761"
'sColor = "#FF00C7B4"
'sColor = "#FF00BAFF"
'sColor = "#FF00C7B4"
'sColor = "#FF242da3"
'sColor = "#FFC088FF"
'sColor = "#FFE082FF"
'sColor = "#FFD555FF"
'sColor = "#FFC518FF"
'sColor = "#FFFF00BE"
'sColor = "#FFFF005F"
'sColor = "#FFE20000"
'sColor = "#FFb70000"
Public Function APaleta(ByVal nColor As Integer) As String
Dim sColor As String
    Select Case nColor
    Case 0
        sColor = "#FFE0E0E0"
    Case 1
        sColor = "#FFC1C1C1"
    Case 2
        sColor = "#FFA2A2A2"
    Case 3
        sColor = "#FF6D6F81"
    Case 4
        sColor = "#FF838383"
    Case 5
        sColor = "#FF646464"
    Case 6
        sColor = "#FF454545"
    Case 7
        sColor = "#FF000000"
    Case 8
        sColor = "#FFF0FF00"
    Case 9
        sColor = "#FFC7DE28"
    Case 10
        sColor = "#FFFFDA28"
    Case 11
        sColor = "#FFBBA125"
    Case 12
        sColor = "#FFFFA500"
    Case 13
        sColor = "#FFFF7E34"
    Case 14
        sColor = "#FFFF7400"
    Case 15
        sColor = "#FF974500"
    Case 16
        sColor = "#FFC2FF49"
    Case 17
        sColor = "#FF4EFB22"
    Case 18
        sColor = "#FF16C700"
    Case 19
        sColor = "#FF00B761"
    Case 20
        sColor = "#FF00C7B4"
    Case 21
        sColor = "#FF00BAFF"
    Case 22
        sColor = "#FF00C7B4"
    Case 23
        sColor = "#FF242DA3"
    Case 24
        sColor = "#FFC088FF"
    Case 25
        sColor = "#FFE082FF"
    Case 26
        sColor = "#FFD555FF"
    Case 27
        sColor = "#FFC518FF"
    Case 28
        sColor = "#FFFF00BE"
    Case 29
        sColor = "#FFFF005F"
    Case 30
        sColor = "#FFE20000"
    Case 31
        sColor = "#FFB70000"
    Case Else
        sColor = "#FFE0E0E0"
    End Select
    APaleta = sColor
End Function

Public Function DigitoControl(ByVal Valor As String) As Integer
   '**************************
   ' Calcula el Checksum de
   ' un código de barras
   '***************************
   Dim Longitud As Integer
   Dim sumapar As Integer
   Dim sumaimpar As Integer
   Dim suma As Integer
   Dim digito, i As Integer
   sumapar = 0
   sumaimpar = 0
   Longitud = Len(Valor)
   For i = 1 To Longitud
      If i Mod 2 = 0 Then
         sumapar = sumapar + Val(Mid(Valor, i, 1))
      Else
         sumaimpar = sumaimpar + Val(Mid(Valor, i, 1))
      End If
   Next
   If Longitud Mod 2 = 0 Then
      suma = sumapar * 3 + sumaimpar
   Else
      suma = sumapar + sumaimpar * 3
   End If

   digito = 10 - suma Mod 10
   If digito = 10 Then digito = 0
   DigitoControl = digito
End Function


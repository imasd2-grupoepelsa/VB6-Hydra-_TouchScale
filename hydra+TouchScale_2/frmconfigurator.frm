VERSION 5.00
Begin VB.Form frmconfigurator 
   Caption         =   "Ficheros de Configuración"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdborrar 
      Caption         =   "Borrar Campo"
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   1200
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   5535
   End
   Begin VB.CommandButton cmdcancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Añadir"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Inserte aquí una descripción del contenido del campo"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox txtinicio 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Posición inicial, contando desde 1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtfin 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Posición final, contando desde 1 "
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cmbtipo 
      Height          =   315
      ItemData        =   "frmconfigurator.frx":0000
      Left            =   1680
      List            =   "frmconfigurator.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdeditar 
      Caption         =   "Editar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtfichero 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000018&
      Caption         =   "Elija el fichero de importación y pulse ""Editar"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Campos definidos hasta el momento :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Fin"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Inicio"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Campo"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del fichero  :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmconfigurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim contador As Long
Dim Position As Integer
' Constante que determina el número de elementos
' que tiene un artículo (COD,FAM,SUB...)
Private Const DimensionArticulo = 15
' Path que se utilizará para el fichero de importación
Private Propiedades(DimensionArticulo) As tipos
Private lineasconf() As lineaconfig
Private arTIculo As registroarticulos
Public PathImport As String

'
Private Sub inicia_articulo()
' *******************************************************************
' Daniel Campos Fernández 25-07-2001
' Esta función inicializa los valores de la variable "Propiedades"
' que contiene los datos sobre artículos :
' .codigo = codigo de la propiedad según el protocolo de comunicaciones
' .descripción = una descripción informativa
' .max = longitud máxima admisible del campo
' *******************************************************************
    Propiedades(0).codigo = "COD"
    Propiedades(0).descripcion = "Código"
    Propiedades(0).max = 6
    
    Propiedades(1).codigo = "PLU"
    Propiedades(1).descripcion = "Plu"
    Propiedades(1).max = 4
    
    Propiedades(2).codigo = "SEC"
    Propiedades(2).descripcion = "Sección"
    Propiedades(2).max = 2
    
    Propiedades(3).codigo = "SUB"
    Propiedades(3).descripcion = "Subsección"
    Propiedades(3).max = 3
    
    Propiedades(4).codigo = "FAM"
    Propiedades(4).descripcion = "Familia"
    Propiedades(4).max = 4
    
    Propiedades(5).codigo = "PRC"
    Propiedades(5).descripcion = "Precio"
    Propiedades(5).max = 7
    
    Propiedades(6).codigo = "WGH"
    Propiedades(6).descripcion = "Tipo de venta (pesado/no pesado)"
    Propiedades(6).max = 1
    
    Propiedades(7).codigo = "CAD"
    Propiedades(7).descripcion = "Caducidad"
    Propiedades(7).max = 3
    
    Propiedades(8).codigo = "PRE"
    Propiedades(8).descripcion = "Consumo preferente"
    
    Propiedades(9).codigo = "TAR"
    Propiedades(9).descripcion = "Tara Asociada"
    Propiedades(9).max = 5
    
    Propiedades(10).codigo = "ETQ"
    Propiedades(10).descripcion = "Número de etiqueta"
    Propiedades(10).max = 2
    
    Propiedades(11).codigo = "CBA"
    Propiedades(11).descripcion = "Código de barras del artículo"
    Propiedades(11).max = 12
    
    Propiedades(12).codigo = "TLT"
    Propiedades(12).descripcion = "Tipo de letra"
    Propiedades(12).max = 1
    
    Propiedades(13).codigo = "TXT"
    Propiedades(13).descripcion = "Texto"
    Propiedades(13).max = 25
    
    Propiedades(14).codigo = "TYP"
    Propiedades(14).descripcion = "Tipo de descriptivo"
    
    
    
       
    
End Sub





Private Sub cmbtipo_Click()
     If cmbtipo.ListIndex >= 0 Then
         txtdescripcion.Text = Propiedades(cmbtipo.ListIndex).descripcion
     End If
End Sub

Private Sub cmdaceptar_Click()
    Dim filenum As Integer
    Dim nombre As String
    Dim bucle As Long
    filenum = FreeFile
    
    nombre = PathImport & "\" & txtfichero.Text & ".cfg"
    Open nombre For Output As filenum
    For bucle = 0 To contador - 1
        Write #filenum, lineasconf(bucle).codigo, lineasconf(bucle).descripcion, lineasconf(bucle).Inicio, lineasconf(bucle).fin, lineasconf(bucle).posicion
        
    Next bucle
    Close filenum
    Unload Me
End Sub

Private Sub cmdborrar_Click()
    Dim bucle As Long
    If List1.Text = "" Then Exit Sub
    
    '
    ' Elimina el registro
    '
    contador = contador - 1
    For bucle = List1.ListIndex To contador
        lineasconf(bucle) = lineasconf(bucle + 1)
    Next bucle
    ReDim Preserve lineasconf(contador + 1)
    
    '
    ' Repone la lista de campos
    '
    List1.Clear
      For bucle = 0 To contador - 1
        If lineasconf(bucle).codigo = "TXT" Or lineasconf(bucle).codigo = "TLT" Then
         List1.AddItem lineasconf(bucle).codigo & "(" & lineasconf(bucle).posicion & ")" & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        Else
         List1.AddItem lineasconf(bucle).codigo & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        End If
    Next bucle
    
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdeditar_Click()
    Dim filenum As Integer
    Dim bucle As Long
    ' Habilita los diálogos
    
    List1.Enabled = True
    cmdaceptar.Enabled = True
    txtinicio.Enabled = True
    txtinicio.Enabled = True
    cmbtipo.Enabled = True
    cmdnuevo.Enabled = True
    txtfichero.Enabled = False
    cmdborrar.Enabled = True
    txtdescripcion.Enabled = True
    cmdeditar.Enabled = False
    Label6.Caption = "Edite el contenido del fichero, y pulse 'Aceptar' para guardar"
    ' Lee el fichero de configuración
    '
    filenum = FreeFile
    contador = 0
    On Error GoTo noexiste
    Open PathImport & "\" & txtfichero.Text & ".cfg" For Input As filenum
    On Error GoTo 0
   
    Do While Not EOF(filenum)
        Input #filenum, lineasconf(contador).codigo, lineasconf(contador).descripcion, lineasconf(contador).Inicio, lineasconf(contador).fin, lineasconf(contador).posicion
        contador = contador + 1
        ReDim Preserve lineasconf(contador + 1)
    Loop
    Close filenum
    '
    ' Carga la lista
    List1.Clear
    For bucle = 0 To contador - 1
        If lineasconf(bucle).codigo = "TXT" Or lineasconf(bucle).codigo = "TLT" Then
         List1.AddItem lineasconf(bucle).codigo & "(" & lineasconf(bucle).posicion & ")" & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        Else
         List1.AddItem lineasconf(bucle).codigo & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        End If
    Next bucle
    Exit Sub
noexiste:
End Sub

Private Sub cmdnuevo_Click()
    Dim bucle As Long, bucle2 As Long
    Dim buffer As String
    Dim BufferLineas As lineaconfig
    
    If cmbtipo.Text = "" Then Exit Sub
    ' Comprobación de que es correcto
      If (Not IsNumeric(txtinicio.Text)) Or (Not IsNumeric(txtfin.Text)) Then
        MsgBox "Inicio y Fin deben ser numéricos"
        Exit Sub
     End If
    
    If Val(txtinicio.Text) > Val(txtfin.Text) Then
        MsgBox "El punto de inicio no puede ser mayor que el final"
        Exit Sub
    End If
    ' Comprueba que no está entre medio de otro campo
    For bucle = 0 To contador
        If Val(txtinicio.Text) >= lineasconf(bucle).Inicio And Val(txtinicio.Text) <= lineasconf(bucle).fin Then
            MsgBox "Zona del fichero ya definida, compruebe el punto de inicio"
            Exit Sub
        End If
         If Val(txtfin.Text) >= lineasconf(bucle).Inicio And Val(txtfin.Text) <= lineasconf(bucle).fin Then
            MsgBox "Zona del fichero ya definida, compruebe el punto de inicio"
            Exit Sub
        End If
    Next bucle
    
    ' Casos especiales,txt, tlt, son arrays
    If cmbtipo.Text = "TXT" Then
        buffer = InputBox("¿Qué línea de descriptivo (0-20) ?", "Descriptivo")
        If Val(buffer) < 0 Or Val(buffer) > 20 Then
            MsgBox "Valor no válido"
            Exit Sub
        Else
            lineasconf(contador).posicion = Val(buffer)
        End If
    End If
    If cmbtipo.Text = "TLT" Then
        buffer = InputBox("¿Qué TLT corresponde (0-10) ?", "Tipo de Letra")
        If Val(buffer) < 0 Or Val(buffer) > 10 Then
            MsgBox "Valor no válido"
            Exit Sub
        Else
            lineasconf(contador).posicion = Val(buffer)
        End If
    End If
    lineasconf(contador).codigo = cmbtipo.Text
    lineasconf(contador).descripcion = txtdescripcion.Text
    lineasconf(contador).Inicio = Val(txtinicio.Text)
    lineasconf(contador).fin = Val(txtfin.Text)
    
    
    List1.Clear
    '
    ' Reordena la configuración
    For bucle = 0 To contador - 1
        For bucle2 = bucle + 1 To contador
            If lineasconf(bucle).Inicio > lineasconf(bucle2).Inicio Then
                
                BufferLineas.codigo = lineasconf(bucle).codigo
                BufferLineas.descripcion = lineasconf(bucle).descripcion
                BufferLineas.Inicio = lineasconf(bucle).Inicio
                BufferLineas.fin = lineasconf(bucle).fin
                BufferLineas.posicion = lineasconf(bucle).posicion
                
                lineasconf(bucle).codigo = lineasconf(bucle2).codigo
                lineasconf(bucle).descripcion = lineasconf(bucle2).descripcion
                lineasconf(bucle).Inicio = lineasconf(bucle2).Inicio
                lineasconf(bucle).fin = lineasconf(bucle2).fin
                lineasconf(bucle).posicion = lineasconf(bucle2).posicion
                
                lineasconf(bucle2).codigo = BufferLineas.codigo
                lineasconf(bucle2).descripcion = BufferLineas.descripcion
                lineasconf(bucle2).Inicio = BufferLineas.Inicio
                lineasconf(bucle2).fin = BufferLineas.fin
                lineasconf(bucle2).posicion = BufferLineas.posicion
            End If
        Next bucle2
    Next bucle
    '
    ' Actualiza el cuadro de lista
    '
    For bucle = 0 To contador
        If lineasconf(bucle).codigo = "TXT" Or lineasconf(bucle).codigo = "TLT" Then
         List1.AddItem lineasconf(bucle).codigo & "(" & lineasconf(bucle).posicion & ")" & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        Else
         List1.AddItem lineasconf(bucle).codigo & "----" & lineasconf(bucle).descripcion & "-Inicio : " & lineasconf(bucle).Inicio & "-Fin : " & lineasconf(bucle).fin
        End If
    Next bucle

    contador = contador + 1
    ReDim Preserve lineasconf(contador + 1)
End Sub

Private Sub Form_Load()
    Dim bucle As Integer
    ReDim lineasconf(1)
    inicia_articulo
    For bucle = 0 To DimensionArticulo - 1
        cmbtipo.AddItem Propiedades(bucle).codigo
    Next bucle
    
    ' Deshabilita todos los cuadros hasta que se pulse "Editar"
    '
    List1.Enabled = False
    cmdaceptar.Enabled = False
    txtinicio.Enabled = False
    txtinicio.Enabled = False
    cmbtipo.Enabled = False
    cmdnuevo.Enabled = False
    cmdborrar.Enabled = False
    txtdescripcion.Enabled = False
End Sub



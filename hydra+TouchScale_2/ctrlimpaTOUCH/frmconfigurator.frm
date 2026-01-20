VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconfigurator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficheros de Configuración"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "frmconfigurator.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lista 
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox ChkBalenv 
      Caption         =   "Enviar los artículos de este fichero a Ambos Sistemas"
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   4452
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "Borrar Campo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5760
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdcancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5760
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtdescripcion 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Inserte aquí una descripción del contenido del campo"
      Top             =   1440
      Width           =   4332
   End
   Begin VB.TextBox txtinicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Posición inicial, contando desde 1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtfin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Posición final, contando desde 1 "
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cmbtipo 
      Height          =   315
      ItemData        =   "frmconfigurator.frx":030A
      Left            =   1680
      List            =   "frmconfigurator.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdeditar 
      Caption         =   "Editar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtfichero 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1560
      Picture         =   "frmconfigurator.frx":030E
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4695
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
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label3 
      Caption         =   "Fin"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Inicio"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Campo"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del fichero  :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
Private Const DimensionArticulo = 30
' Path que se utilizará para el fichero de importación
Private Propiedades(DimensionArticulo) As tipos
Private lineasconf() As lineaconfig
Private articulo As registroarticulos
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////////////
Private Sub Form_Load()
Dim Bucle As Integer
ReDim lineasconf(1)
    Label1.Caption = LoadResString(id + 552)
    If id <> 0 Then Label2.Caption = ""
    ChkBalenv.Caption = LoadResString(id + 566)
    Label4.Caption = LoadResString(id + 780)
    Label3.Caption = LoadResString(id + 781)
    frmconfigurator.Caption = LoadResString(id + 258)
    If id <> 0 Then Label6.Visible = False
    cmdeditar.Caption = LoadResString(id + 774)
    cmdnuevo.Caption = LoadResString(id + 550)
    cmdborrar.Caption = LoadResString(id + 38)
    cmdaceptar.Caption = LoadResString(id + 287)
    cmdcancelar.Caption = LoadResString(id + 288)
    '
    Me.Lista.ColumnHeaders.Add , , LoadResString(id + 916)
    Me.Lista.ColumnHeaders.Add , , LoadResString(id + 48)
    Me.Lista.ColumnHeaders.Add , , LoadResString(id + 780)
    Me.Lista.ColumnHeaders.Add , , LoadResString(id + 781)
    '
    inicia_articulo
    For Bucle = 0 To DimensionArticulo - 1
        If Propiedades(Bucle).codigo <> "" Then
            cmbtipo.AddItem Propiedades(Bucle).codigo
        End If
    Next Bucle
    '
    ' Deshabilita todos los cuadros hasta que se pulse "Editar"
    '
    Lista.Enabled = False
    cmdaceptar.Enabled = False
    txtinicio.Enabled = False
    txtinicio.Enabled = False
    cmbtipo.Enabled = False
    cmdnuevo.Enabled = False
    cmdborrar.Enabled = False
    txtdescripcion.Enabled = False
    ChkBalenv.Enabled = False
End Sub


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
    ' *******************************************************************
' Daniel Campos Fernández 25-07-2001
' Esta función inicializa los valores de la variable "Propiedades"
' que contiene los datos sobre artículos :
' .codigo = codigo de la propiedad según el protocolo de comunicaciones
' .descripción = una descripción informativa
' .max = longitud máxima admisible del campo
' *******************************************************************
    Propiedades(0).codigo = "COD"
    Propiedades(0).descripcion = LoadResString(id + 89) ' "Código"
    Propiedades(0).max = 6
    
    Propiedades(1).codigo = "PLU"
    Propiedades(1).descripcion = "Plu"
    Propiedades(1).max = 4
    
    Propiedades(2).codigo = "MOS"
    Propiedades(2).descripcion = LoadResString(id + 302) '"Mostrador"
    Propiedades(2).max = 2
    
    Propiedades(3).codigo = "SUB"
    Propiedades(3).descripcion = LoadResString(id + 10) '"Subsección"
    Propiedades(3).max = 3
    
    Propiedades(4).codigo = "FAM"
    Propiedades(4).descripcion = LoadResString(id + 67) '"Familia"
    Propiedades(4).max = 4
    
    Propiedades(5).codigo = "PRC"
    Propiedades(5).descripcion = LoadResString(id + 71) '"Precio"
    Propiedades(5).max = 7
    
    Propiedades(6).codigo = "WGH"
    Propiedades(6).descripcion = LoadResString(id + 79) & "/" & LoadResString(id + 80) '"Tipo de venta (pesado/no pesado)"
    Propiedades(6).max = 1
    
    Propiedades(7).codigo = "CAD"
    Propiedades(7).descripcion = LoadResString(id + 72) '"Caducidad"
    Propiedades(7).max = 3
    
    Propiedades(8).codigo = "PRE"
    Propiedades(8).descripcion = LoadResString(id + 73) '"Consumo preferente"
    
    Propiedades(9).codigo = "TAR"
    Propiedades(9).descripcion = LoadResString(id + 74) '"Tara Asociada"
    Propiedades(9).max = 5
    
    Propiedades(10).codigo = "ETQ"
    Propiedades(10).descripcion = LoadResString(id + 623) '"Número de etiqueta"
    Propiedades(10).max = 2
    
    Propiedades(11).codigo = "CBA"
    Propiedades(11).descripcion = LoadResString(id + 23) '"Código de barras del artículo"
    Propiedades(11).max = 13
    
    Propiedades(12).codigo = "TLT"
    If id = 0 Then
        Propiedades(12).descripcion = "Tipo de letra"
    Else
        Propiedades(12).descripcion = LoadResString(id + 893)
    End If
    Propiedades(12).max = 1
    
    Propiedades(13).codigo = "TXT"
    Propiedades(13).descripcion = LoadResString(id + 539) '"Texto"
    Propiedades(13).max = 80
    
    Propiedades(14).codigo = "TYP"
    Propiedades(14).descripcion = "Tipo de descriptivo"
    
    Propiedades(15).codigo = "GPC"
    Propiedades(15).descripcion = LoadResString(id + 638) '"Grupo de Conservación"
    Propiedades(15).max = 2
    
        
    Propiedades(16).codigo = "TRM1"
    Propiedades(16).descripcion = LoadResString(id + 651) & " 1" '"Tramo 1"
    Propiedades(16).max = 9
       
    Propiedades(17).codigo = "PRC1"
    Propiedades(17).descripcion = LoadResString(id + 650) & " 1" '"Precio del Tramo 1"
    Propiedades(17).max = 9
       
    Propiedades(18).codigo = "TRM2"
    Propiedades(18).descripcion = LoadResString(id + 651) & " 2" '"Tramo 2"
    Propiedades(18).max = 9
    
    Propiedades(19).codigo = "PRC2"
    Propiedades(19).descripcion = LoadResString(id + 650) & " 2" '"Precio del Tramo 2"
    Propiedades(19).max = 9
       
    Propiedades(20).codigo = "IVA"
    Propiedades(20).descripcion = LoadResString(id + 553) ' "Tipo de IVA"
    Propiedades(20).max = 1
    
    Propiedades(21).codigo = "TEC"
    Propiedades(21).descripcion = LoadResString(id + 325) ' "Tecla en la sección"
    Propiedades(21).max = 3
    
    Propiedades(22).codigo = "JTEC"
    If id <> 0 Then
        Propiedades(22).descripcion = LoadResString(id + 223) ' "Juego de Tecla en la sección"
    Else
        Propiedades(22).descripcion = "Juego de Tecla en la sección"
    End If
    Propiedades(22).max = 2
    
    Propiedades(23).codigo = "NPR"
    Propiedades(23).descripcion = "Nº de Paquete Preseleccionado"
    Propiedades(23).max = 2
    
    Propiedades(24).codigo = "MER"
    Propiedades(24).descripcion = "Merma en tanto por ciento"
    Propiedades(24).max = 2
    
    Propiedades(25).codigo = "PRC100"
    Propiedades(25).descripcion = "Flag Precio por 100 gramos (EcoLabel)"
    Propiedades(25).max = 1
    
    Propiedades(26).codigo = "BAJA"
    Propiedades(26).descripcion = "Distinto de 0 Indica si baja de artículo"
    Propiedades(26).max = 2
    
    Propiedades(27).codigo = "IMAGE"
    Propiedades(27).descripcion = "Path + fichero Imagen"
    Propiedades(27).max = 250
    
    Propiedades(28).codigo = "TARIF."
    Propiedades(28).descripcion = "TARIFAS 1 a 10" 'LoadResString(id + 71) '"Precio"
    Propiedades(28).max = 7
    
    Propiedades(29).codigo = "ONKEY"
    Propiedades(29).descripcion = "Activar o no tecla en Balanza Touch"
    Propiedades(29).max = 1
    
End Sub





Private Sub cmbtipo_Click()
     If cmbtipo.ListIndex >= 0 Then
         txtdescripcion.Text = Propiedades(cmbtipo.ListIndex).descripcion
     End If
End Sub

Private Sub cmdaceptar_Click()
    Dim filenum As Integer
    Dim nombre As String
    Dim Bucle As Long
    filenum = FreeFile
    
    nombre = PathLocal & "\etc\" & txtfichero.Text & ".FIA"
    Open nombre For Output As filenum
    If ChkBalenv.Value = vbChecked Then
        Write #filenum, 1
    Else
        Write #filenum, 0
    End If
    For Bucle = 0 To contador - 1
        Write #filenum, lineasconf(Bucle).codigo, lineasconf(Bucle).descripcion, lineasconf(Bucle).Inicio, lineasconf(Bucle).fin, lineasconf(Bucle).posicion
        
    Next Bucle
    Close filenum
    Unload Me
End Sub

Private Sub cmdborrar_Click()
    Dim Bucle As Long
    If Lista.ListItems.Count = 0 Then Exit Sub
    
    '
    ' Elimina el registro
    '
    contador = contador - 1
    For Bucle = Lista.SelectedItem.Index - 1 To contador
        lineasconf(Bucle) = lineasconf(Bucle + 1)
    Next Bucle
    ReDim Preserve lineasconf(contador + 1)
    
    '
    ' Repone la lista de campos
    '
    Lista.ListItems.Clear
      For Bucle = 0 To contador - 1
        If lineasconf(Bucle).codigo = "TXT" Or lineasconf(Bucle).codigo = "TLT" Or lineasconf(Bucle).codigo = "TARIF." Then
         Lista.ListItems.Add , , lineasconf(Bucle).codigo & "(" & lineasconf(Bucle).posicion & ")"
        Else
         Lista.ListItems.Add , , lineasconf(Bucle).codigo
        End If
        Lista.ListItems(Lista.ListItems.Count).SubItems(1) = lineasconf(Bucle).descripcion
        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = lineasconf(Bucle).Inicio
        Lista.ListItems(Lista.ListItems.Count).SubItems(3) = lineasconf(Bucle).fin
    Next Bucle
    
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdeditar_Click()
    Dim filenum As Integer
    Dim EsBalenv As String
    Dim Bucle As Long
    ' Habilita los diálogos
    cmdborrar.Enabled = False
    Lista.Enabled = True
    ChkBalenv.Enabled = True
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
    Open PathLocal & "\etc\" & txtfichero.Text & ".fia" For Input As filenum
    On Error GoTo 0
    Input #filenum, EsBalenv
    If EsBalenv = "1" Then ChkBalenv.Value = vbChecked
    Do While Not EOF(filenum)
        Input #filenum, lineasconf(contador).codigo, lineasconf(contador).descripcion, lineasconf(contador).Inicio, lineasconf(contador).fin, lineasconf(contador).posicion
        contador = contador + 1
        ReDim Preserve lineasconf(contador + 1)
    Loop
    Close filenum
    '
    ' Carga la lista
    Lista.ListItems.Clear
    For Bucle = 0 To contador - 1
        If lineasconf(Bucle).codigo = "TXT" Or lineasconf(Bucle).codigo = "TLT" Or lineasconf(Bucle).codigo = "TARIF." Then
         Lista.ListItems.Add , , lineasconf(Bucle).codigo & "(" & lineasconf(Bucle).posicion & ")"
        Else
         Lista.ListItems.Add , , lineasconf(Bucle).codigo
        End If
        Lista.ListItems(Lista.ListItems.Count).SubItems(1) = lineasconf(Bucle).descripcion
        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = lineasconf(Bucle).Inicio
        Lista.ListItems(Lista.ListItems.Count).SubItems(3) = lineasconf(Bucle).fin
    Next Bucle
    Lista.ListItems(Lista.ListItems.Count).Selected = True
    Lista.SelectedItem.EnsureVisible
    Exit Sub
noexiste:
End Sub

Private Sub cmdnuevo_Click()
    Dim Bucle As Long, bucle2 As Long
    Dim buffer As String
    Dim BufferLineas As lineaconfig
    
    If cmbtipo.Text = "" Then Exit Sub
    ' Comprobación de que es correcto
      If (Not IsNumeric(txtinicio.Text)) Or (Not IsNumeric(txtfin.Text)) Then
        If MostrarMensajes Then MsgBox "Inicio y Fin deben ser numéricos"
        Exit Sub
     End If
    
    If Val(txtinicio.Text) > Val(txtfin.Text) Then
        If MostrarMensajes Then
            If id <> 0 Then
                MsgBox LoadResString(id + 795)
            Else
                MsgBox "El punto de inicio no puede ser mayor que el final"
            
            End If
        End If
        Exit Sub
    End If
    ' Comprueba que no está entre medio de otro campo
    For Bucle = 0 To contador
        If Val(txtinicio.Text) >= lineasconf(Bucle).Inicio And Val(txtinicio.Text) <= lineasconf(Bucle).fin Then
            If MostrarMensajes Then MsgBox LoadResString(id + 795)
            Exit Sub
        End If
         If Val(txtfin.Text) >= lineasconf(Bucle).Inicio And Val(txtfin.Text) <= lineasconf(Bucle).fin Then
            If MostrarMensajes Then MsgBox LoadResString(id + 795)
            Exit Sub
        End If
    Next Bucle
    
    ' Casos especiales,txt, tlt, son arrays
    If cmbtipo.Text = "TXT" Then
        If id <> 0 Then
            buffer = InputBox("¿TXT (0-40) ?", "Hydra Basic")
        Else
            buffer = InputBox("¿Qué línea de descriptivo (0-40) ?", "Descriptivo")
        End If
        If Val(buffer) < 0 Or Val(buffer) > 40 Then
            If MostrarMensajes Then
                If id <> 0 Then
                    MsgBox "Error."
                Else
                    MsgBox "Valor no válido"
                End If
            End If
            Exit Sub
        Else
            lineasconf(contador).posicion = Val(buffer)
        End If
    End If
    If cmbtipo.Text = "TLT" Then
        If id <> 0 Then
            buffer = InputBox("¿TLT (0-10) ?", "HydraBasic")
        Else
            buffer = InputBox("¿Qué TLT corresponde (0-10) ?", "Tipo de Letra")
        End If
        If Val(buffer) < 0 Or Val(buffer) > 10 Then
            If MostrarMensajes Then
                If id <> 0 Then
                    MsgBox "Error."
                Else
                    MsgBox "Valor no válido"
                End If
            End If
            Exit Sub
        Else
            lineasconf(contador).posicion = Val(buffer)
        End If
    End If
    If cmbtipo.Text = "TARIF." Then
        If id <> 0 Then
            buffer = InputBox("¿TARIFA (1-10) ?", "HydraBasic")
        Else
            buffer = InputBox("¿Qué Tarifa corresponde (1-10) ?", "Tarifa")
        End If
        If Val(buffer) < 1 Or Val(buffer) > 10 Then
            If MostrarMensajes Then
                If id <> 0 Then
                    MsgBox "Error."
                Else
                    MsgBox "Valor no válido"
                End If
            End If
            Exit Sub
        Else
            lineasconf(contador).posicion = Val(buffer)
        End If
    End If
    
    lineasconf(contador).codigo = cmbtipo.Text
    lineasconf(contador).descripcion = txtdescripcion.Text
    lineasconf(contador).Inicio = Val(txtinicio.Text)
    lineasconf(contador).fin = Val(txtfin.Text)
    
    
    Lista.ListItems.Clear
    '
    ' Reordena la configuración
    For Bucle = 0 To contador - 1
        For bucle2 = Bucle + 1 To contador
            If lineasconf(Bucle).Inicio > lineasconf(bucle2).Inicio Then
                
                BufferLineas.codigo = lineasconf(Bucle).codigo
                BufferLineas.descripcion = lineasconf(Bucle).descripcion
                BufferLineas.Inicio = lineasconf(Bucle).Inicio
                BufferLineas.fin = lineasconf(Bucle).fin
                BufferLineas.posicion = lineasconf(Bucle).posicion
                
                lineasconf(Bucle).codigo = lineasconf(bucle2).codigo
                lineasconf(Bucle).descripcion = lineasconf(bucle2).descripcion
                lineasconf(Bucle).Inicio = lineasconf(bucle2).Inicio
                lineasconf(Bucle).fin = lineasconf(bucle2).fin
                lineasconf(Bucle).posicion = lineasconf(bucle2).posicion
                
                lineasconf(bucle2).codigo = BufferLineas.codigo
                lineasconf(bucle2).descripcion = BufferLineas.descripcion
                lineasconf(bucle2).Inicio = BufferLineas.Inicio
                lineasconf(bucle2).fin = BufferLineas.fin
                lineasconf(bucle2).posicion = BufferLineas.posicion
            End If
        Next bucle2
    Next Bucle
    '
    ' Actualiza el cuadro de lista
    '
    For Bucle = 0 To contador
        If lineasconf(Bucle).codigo = "TXT" Or lineasconf(Bucle).codigo = "TLT" Or lineasconf(Bucle).codigo = "TARIF." Then
         Lista.ListItems.Add , , lineasconf(Bucle).codigo & "(" & lineasconf(Bucle).posicion & ")"
        Else
         Lista.ListItems.Add , , lineasconf(Bucle).codigo
        End If
        Lista.ListItems(Lista.ListItems.Count).SubItems(1) = lineasconf(Bucle).descripcion
        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = lineasconf(Bucle).Inicio
        Lista.ListItems(Lista.ListItems.Count).SubItems(3) = lineasconf(Bucle).fin
    Next Bucle

    contador = contador + 1
    ReDim Preserve lineasconf(contador + 1)
End Sub


Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdborrar.Enabled = True
End Sub

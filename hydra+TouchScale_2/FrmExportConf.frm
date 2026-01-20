VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Exportación"
   ClientHeight    =   5820
   ClientLeft      =   1455
   ClientTop       =   2370
   ClientWidth     =   7230
   Icon            =   "FrmExportConf.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4452
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6972
      Begin VB.CheckBox ChkNTienda 
         Caption         =   "Incluir dígitos de tienda en el nombre de fichero"
         Height          =   612
         Left            =   1920
         TabIndex        =   32
         Top             =   240
         Width           =   3012
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox TxtEAN 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtEAN 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Cmbcual 
         Height          =   288
         ItemData        =   "FrmExportConf.frx":000C
         Left            =   5160
         List            =   "FrmExportConf.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1080
         Width           =   1572
      End
      Begin VB.TextBox txtSC10 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   6240
         TabIndex        =   23
         Top             =   1680
         Width           =   492
      End
      Begin VB.TextBox TxtGA 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   6240
         TabIndex        =   22
         Top             =   1440
         Width           =   492
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Suprimir"
         Height          =   252
         Left            =   5280
         TabIndex        =   17
         Top             =   600
         Width           =   1452
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "&Insertar"
         Height          =   252
         Left            =   5280
         TabIndex        =   16
         Top             =   240
         Width           =   1452
      End
      Begin VB.CommandButton CmdAceptar 
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
         Height          =   492
         Left            =   240
         TabIndex        =   19
         Top             =   3840
         Width           =   1452
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   492
         Index           =   1
         Left            =   5280
         TabIndex        =   18
         Top             =   3840
         Width           =   1452
      End
      Begin VB.TextBox TxtFin 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   1092
      End
      Begin VB.TextBox TxtInicio 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1092
      End
      Begin VB.ComboBox CmbElemento 
         Height          =   288
         ItemData        =   "FrmExportConf.frx":0036
         Left            =   1680
         List            =   "FrmExportConf.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   1692
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1575
         Left            =   0
         TabIndex        =   9
         Top             =   2160
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label8 
         Caption         =   "Fichero de Líneas"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Fichero de Cabecera"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Exportar :"
         Height          =   252
         Left            =   4080
         TabIndex        =   24
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "Código SC10 :"
         Height          =   252
         Index           =   1
         Left            =   4560
         TabIndex        =   21
         Top             =   1800
         Width           =   1572
      End
      Begin VB.Label Label6 
         Caption         =   "Código Euroescale :"
         Height          =   252
         Index           =   0
         Left            =   4560
         TabIndex        =   20
         Top             =   1440
         Width           =   1572
      End
      Begin VB.Label Label5 
         Caption         =   "Fin :"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label4 
         Caption         =   "Inicio :"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Elemento :"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label lblinfo 
         BackColor       =   &H8000000E&
         Caption         =   "Nuevo Fichero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   612
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fichero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6972
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Width           =   1692
      End
      Begin VB.CommandButton CmdConfigurar 
         Caption         =   "&Configurar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1692
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   288
         ItemData        =   "FrmExportConf.frx":003A
         Left            =   1440
         List            =   "FrmExportConf.frx":0053
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   3132
      End
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo :"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
   End
End
Attribute VB_Name = "FrmExportConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TipoExport As Integer
' TipoExport
' 1= Tiquets sólo Euroscale
' 2= Tiquets sólo SC10
' 3= Tiquets de ambos sistemas
Private Tipo As Integer
Private Type TipoGtarti
    texto As String
    Min As Integer
End Type
Private DefLinea(30) As TipoGtarti 'c2f invicta eran 22 , pasa a 25 con numlote, pasa a 26 con cliente
Private DefGtarti(21) As TipoGtarti
Private DefGtsecs(12) As TipoGtarti
Private DefGtVend(16) As TipoGtarti
Private DefCabecera(37) As TipoGtarti 'c2f invicta eran 16 ... pasa a 19 con lote
Private DefGtHor(18) As TipoGtarti
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////////////
Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = Form2.Icon
    If Dir(App.Path & "\hydramulti.exe") <> "" Then
        ChkNTienda.Visible = True
    Else
        ChkNTienda.Visible = False
        ChkNTienda.Value = vbUnchecked
    End If
    Image1.Picture = Form2.Imagen.Picture
    Label6(0).Visible = False
    Label6(1).Visible = False
    Label7.Visible = False
    Cmbcual.Visible = False
    TxtGA.Visible = False
    txtSC10.Visible = False
    CmdEliminar.Enabled = False
    FrmExportConf.Height = 1628
    '******************
    ' GT Articulos
    DefGtarti(0).texto = "0   " & CargaCadena(215)  'Texto Libre"
    DefGtarti(1).texto = "1   " & CargaCadena(434)  ' "fecha"
    DefGtarti(2).texto = "2   " & CargaCadena(68)  '"Plu"
    DefGtarti(3).texto = "3   " & CargaCadena(70)  'Descriptivo"
    DefGtarti(4).texto = "4   " & CargaCadena(302)  'Mostrador"
    DefGtarti(5).texto = "5   " & CargaCadena(10)  'Subseccion"
    DefGtarti(6).texto = "6   " & CargaCadena(67)  'Familia"
    DefGtarti(7).texto = "7   " & CargaCadena(69)  'Código"
    DefGtarti(8).texto = "8   " & CargaCadena(71)  'Precio"
    DefGtarti(9).texto = "9   " & CargaCadena(132)  'Importe"
    DefGtarti(10).texto = "10   " & CargaCadena(131)  'Unidades"
    DefGtarti(11).texto = "11  " & CargaCadena(130)  'Peso"
    DefGtarti(12).texto = "12  " & CargaCadena(130) & "/" & CargaCadena(131)   'Peso/Unidades"
    DefGtarti(13).texto = "13  " & CargaCadena(74)  '"Tara"
    DefGtarti(14).texto = "14  " & CargaCadena(72)  'Caducidad"
    DefGtarti(15).texto = "15  " & CargaCadena(133)  'Operaciones"
    DefGtarti(16).texto = "16  " & CargaCadena(553)  'Tipo de Iva"
    DefGtarti(17).texto = "17  " & CargaCadena(436)  'Modificado"
    DefGtarti(18).texto = "18  " & CargaCadena(54)  'Sección Máquina"
    DefGtarti(19).texto = "19  " & CargaCadena(1261) '"Signo de importe"
    DefGtarti(20).texto = "20  " & CargaCadena(79) & " / " & CargaCadena(80) '"Pesado/No Pesado"
    DefGtarti(21).texto = "21  " & CargaCadena(1138) '"Número de Tienda"
    DefGtarti(0).Min = 1
    DefGtarti(1).Min = 8
    DefGtarti(2).Min = 4
    DefGtarti(3).Min = 1
    DefGtarti(4).Min = 3
    DefGtarti(5).Min = 3
    DefGtarti(6).Min = 4
    DefGtarti(7).Min = 6
    DefGtarti(8).Min = 7
    DefGtarti(9).Min = 7
    DefGtarti(10).Min = 5
    DefGtarti(11).Min = 5
    DefGtarti(12).Min = 5
    DefGtarti(13).Min = 3
    DefGtarti(14).Min = 3
    DefGtarti(15).Min = 5
    DefGtarti(16).Min = 3
    DefGtarti(17).Min = 1
    DefGtarti(18).Min = 2
    DefGtarti(19).Min = 1
    DefGtarti(20).Min = 1
    DefGtarti(21).Min = 1
    '**************************
    ' GT Secciones
    DefGtsecs(0).texto = "0  " & CargaCadena(215)  'Texto Libre"
    DefGtsecs(1).texto = "1  " & CargaCadena(434)  'Fecha"
    DefGtsecs(2).texto = "2  " & CargaCadena(302)  'Mostrador"
    DefGtsecs(3).texto = "3  " & CargaCadena(10)  'Subsección"
    DefGtsecs(4).texto = "4  " & CargaCadena(67)  'Familia"
    DefGtsecs(5).texto = "5  " & CargaCadena(134)  'Tiquets"
    DefGtsecs(6).texto = "6  " & CargaCadena(133)  'Operacion"
    DefGtsecs(7).texto = "7  " & CargaCadena(130)  'Peso"
    DefGtsecs(8).texto = "8  " & CargaCadena(131)  'Unidades"
    DefGtsecs(9).texto = "9  " & CargaCadena(132)  'Importe"
    DefGtsecs(10).texto = "10 " & CargaCadena(54)  'Sección Máquina"
    DefGtsecs(11).texto = "11  " & CargaCadena(1261) '"Signo de importe"
    DefGtsecs(12).texto = "12  " & CargaCadena(1138) '"Número de Tienda"
    DefGtsecs(0).Min = 1
    DefGtsecs(1).Min = 8
    DefGtsecs(2).Min = 3
    DefGtsecs(3).Min = 3
    DefGtsecs(4).Min = 4
    DefGtsecs(5).Min = 5
    DefGtsecs(6).Min = 4
    DefGtsecs(7).Min = 5
    DefGtsecs(8).Min = 5
    DefGtsecs(9).Min = 7
    DefGtsecs(10).Min = 2
    DefGtsecs(11).Min = 1
    DefGtsecs(12).Min = 1
   '************************
   ' GT Vendedores
   DefGtVend(0).texto = "0  " & CargaCadena(215)  'Texto Libre"
   DefGtVend(1).texto = "1  " & CargaCadena(434)  'Fecha"
   DefGtVend(2).texto = "2  " & CargaCadena(302)  'Mostrador"
   DefGtVend(3).texto = "3  " & CargaCadena(136)  'Nro. Vendedor"
   DefGtVend(4).texto = "4  " & CargaCadena(552)  'Nombre"
   DefGtVend(5).texto = "5  " & CargaCadena(786)  'Supertotal"
   DefGtVend(6).texto = "6  " & CargaCadena(122)  'Total Cancelado"
   DefGtVend(7).texto = "7  " & CargaCadena(132)  'Total"
   DefGtVend(8).texto = "8  " & CargaCadena(134)  'Tiquets"
   DefGtVend(9).texto = "9  " & CargaCadena(133)  'Operaciones"
   DefGtVend(10).texto = "10 " & CargaCadena(130)  'Peso"
   DefGtVend(11).texto = "11 " & CargaCadena(131)  'Unidades"
   DefGtVend(12).texto = "12 " & CargaCadena(787)  'Cancelaciones"
   DefGtVend(13).texto = "13 " & CargaCadena(54)  'Sección Máquina"
   DefGtVend(14).texto = "14  " & CargaCadena(1261) '"Signo de importe"
   DefGtVend(15).texto = "15  " & CargaCadena(1138) '"Número de Tienda"
   DefGtVend(0).Min = 1
   DefGtVend(1).Min = 8
   DefGtVend(2).Min = 3
   DefGtVend(3).Min = 4
   DefGtVend(4).Min = 10
   DefGtVend(5).Min = 7
   DefGtVend(6).Min = 7
   DefGtVend(7).Min = 7
   DefGtVend(8).Min = 4
   DefGtVend(9).Min = 4
   DefGtVend(10).Min = 5
   DefGtVend(11).Min = 5
   DefGtVend(12).Min = 4
   DefGtVend(13).Min = 2
   DefGtVend(14).Min = 1
   DefGtVend(15).Min = 1
   
   '************************
   ' Cabecera de Tiquet
   DefCabecera(0).texto = "0  " & CargaCadena(215)  'Texto Libre"
   DefCabecera(1).texto = "1  " & CargaCadena(123)  'Número de Tiquet"
   DefCabecera(2).texto = "2  " & CargaCadena(329)  'Sección"
   DefCabecera(3).texto = "3  " & CargaCadena(330)  'Balanza"
   DefCabecera(4).texto = "4  " & CargaCadena(136)  'Vendedor"
   DefCabecera(5).texto = "5  " & CargaCadena(791)  'Número de Líneas"
   DefCabecera(6).texto = "6  " & CargaCadena(792)  'Líneas Canceladas"
   DefCabecera(7).texto = "7  " & CargaCadena(132)  'Importe"
   DefCabecera(8).texto = "8  " & CargaCadena(435)  'Hora"
   DefCabecera(9).texto = "9  " & CargaCadena(789)  'Estado Tiquet"
   DefCabecera(10).texto = "10 " & CargaCadena(790)  'Código Cliente"
   DefCabecera(11).texto = "11 " & CargaCadena(434)  'Fecha"
   DefCabecera(12).texto = "12 " & CargaCadena(788)  'Tipo Tiquet"
   DefCabecera(13).texto = "13 " & "EAN13"
   DefCabecera(14).texto = "14  " & CargaCadena(1261) '"Signo de importe"
   DefCabecera(15).texto = "15  " & CargaCadena(1138) '"Número de Tienda"
   DefCabecera(16).texto = "16  " & "P.M." 'c2f invicta forma de pago
   DefCabecera(17).texto = "17  " & "B." & CargaCadena(132) 'c2f invicta importe bruto
   DefCabecera(18).texto = "18  " & CargaCadena(1040) '"LOTE" 'c2f lote
   
   DefCabecera(19).texto = "19  " & "F.Simplificada" 'c2f factura
   DefCabecera(20).texto = "20 " & "%Imp.0"
   DefCabecera(21).texto = "21 " & "%Imp.1"
   DefCabecera(22).texto = "22 " & "%Imp.2"
   DefCabecera(23).texto = "23 " & "%Imp.3"
   DefCabecera(24).texto = "24 " & "%Imp.4"
   DefCabecera(25).texto = "25 " & "Base.0"
   DefCabecera(26).texto = "26 " & "Base.1"
   DefCabecera(27).texto = "27 " & "Base.2"
   DefCabecera(28).texto = "28 " & "Base.3"
   DefCabecera(29).texto = "29 " & "Base.4"
   DefCabecera(30).texto = "30 " & "Imp.0"
   DefCabecera(31).texto = "31 " & "Imp.1"
   DefCabecera(32).texto = "32 " & "Imp.2"
   DefCabecera(33).texto = "33 " & "Imp.3"
   DefCabecera(34).texto = "34 " & "Imp.4"
   DefCabecera(35).texto = "35  " & "% Disc."
   DefCabecera(36).texto = "36  " & "T.Uni."
   DefCabecera(37).texto = "37  " & "T.Wei."
   
   DefCabecera(0).Min = 1
   DefCabecera(1).Min = 5
   DefCabecera(2).Min = 3
   DefCabecera(3).Min = 2
   DefCabecera(4).Min = 4
   DefCabecera(5).Min = 3
   DefCabecera(6).Min = 3
   DefCabecera(7).Min = 7
   DefCabecera(8).Min = 8
   DefCabecera(9).Min = 3
   DefCabecera(10).Min = 4
   DefCabecera(11).Min = 8
   DefCabecera(12).Min = 2
   DefCabecera(13).Min = 13
   DefCabecera(14).Min = 1
   DefCabecera(15).Min = 1
   DefCabecera(16).Min = 1 'c2f invicta forma de pago
   DefCabecera(17).Min = 7 'c2f invicta importe bruto
   DefCabecera(18).Min = 1 'c2f lote
   
   DefCabecera(19).Min = 16 'c2f factura
   DefCabecera(20).Min = 4
   DefCabecera(21).Min = 4
   DefCabecera(22).Min = 4
   DefCabecera(23).Min = 4
   DefCabecera(24).Min = 4
   DefCabecera(25).Min = 4
   DefCabecera(26).Min = 4
   DefCabecera(27).Min = 4
   DefCabecera(28).Min = 4
   DefCabecera(29).Min = 4
   DefCabecera(30).Min = 7
   DefCabecera(31).Min = 7
   DefCabecera(32).Min = 7
   DefCabecera(33).Min = 7
   DefCabecera(34).Min = 7
   DefCabecera(35).Min = 2
   DefCabecera(36).Min = 2 'total unidades
   DefCabecera(37).Min = 3 'total peso
   

   
   '*********************************
   ' Líneas de Tiquet
   DefLinea(0).texto = "0  " & CargaCadena(215)  'Texto Libre"
   DefLinea(1).texto = "1  " & CargaCadena(434)  'Fecha"
   DefLinea(2).texto = "2  " & CargaCadena(435)  'Hora"
   DefLinea(3).texto = "3  " & CargaCadena(136)  'Vendedor"
   DefLinea(4).texto = "4  " & CargaCadena(123)  'Numero Tiquet"
   DefLinea(5).texto = "5  " & CargaCadena(793)  'Numero Linea"
   DefLinea(6).texto = "6  " & CargaCadena(54)  'Sección Máquina"
   DefLinea(7).texto = "7  " & CargaCadena(69)  'Código"
   DefLinea(8).texto = "8  " & CargaCadena(71)  'Precio"
   DefLinea(9).texto = "9  " & CargaCadena(132)  'Importe"
   DefLinea(10).texto = "10  " & CargaCadena(130)  'Peso"
   DefLinea(11).texto = "11  " & CargaCadena(131)  'Unidades"
   DefLinea(12).texto = "12  " & CargaCadena(130) & "/" & CargaCadena(131)    'Peso/unidades"
   DefLinea(13).texto = "13  " & CargaCadena(788)  'Tipo Linea"
   DefLinea(14).texto = "14  " & CargaCadena(789)  'Estado Tiquet"
   DefLinea(15).texto = "15  " & CargaCadena(1261) '"Signo de importe"
   DefLinea(16).texto = "16  " & CargaCadena(1262) '"Fragmento de Cabecera"
   DefLinea(17).texto = "17  " & CargaCadena(1263) '"Descriptivo de artículo"
   DefLinea(18).texto = "18  " & CargaCadena(79) & " / " & CargaCadena(80) '"Pesado/No pesado"
   DefLinea(19).texto = "19  " & CargaCadena(792) '"Línea Cancelada"
   DefLinea(20).texto = "20  " & "PLU"
   DefLinea(21).texto = "21  " & CargaCadena(11) '"Equipo"
   DefLinea(22).texto = "22  " & CargaCadena(1138) '"Número de Tienda"
   DefLinea(23).texto = "23  " & "DPT." 'c2f invicta familia
   DefLinea(24).texto = "24  " & "B." & CargaCadena(132) 'c2f invicta importe bruto
   DefLinea(25).texto = "25  " & CargaCadena(1040) '"LOTE" 'c2f numlote
   DefLinea(26).texto = "26  " & CargaCadena(790) '"Cliente" 'c2f cliente
   DefLinea(27).texto = "27  " & "%Imp."
   DefLinea(28).texto = "28  " & "Factura"
   DefLinea(29).texto = "29  " & "% Disc."
   DefLinea(30).texto = "30  " & "VAT t."
   
   DefLinea(0).Min = 1
   DefLinea(1).Min = 8
   DefLinea(2).Min = 8
   DefLinea(3).Min = 4
   DefLinea(4).Min = 5
   DefLinea(5).Min = 4
   DefLinea(6).Min = 3
   DefLinea(7).Min = 6
   DefLinea(8).Min = 7
   DefLinea(9).Min = 7
   DefLinea(10).Min = 5
   DefLinea(11).Min = 5
   DefLinea(12).Min = 5
   DefLinea(13).Min = 1
   DefLinea(14).Min = 2
   DefLinea(15).Min = 1
   DefLinea(23).Min = 1
   DefLinea(24).Min = 7
   DefLinea(25).Min = 1
   DefLinea(26).Min = 1
   'DefLinea(27).Min = 3
   'DefLinea(28).Min = 1
   'DefLinea(29).Min = 4
   DefLinea(27).Min = 4
   DefLinea(28).Min = 16
   DefLinea(29).Min = 2
   DefLinea(30).Min = 2
   
    '******************
    ' GT Horario
    DefGtHor(0).texto = "0   " & "Texto Libre"
    DefGtHor(1).texto = "1   " & "Fecha"
    DefGtHor(2).texto = "2   " & "Hora"
    DefGtHor(3).texto = "3   " & "NºTiques"
    DefGtHor(4).texto = "4   " & "Operaciones"
    DefGtHor(5).texto = "5   " & "Oper. Cancel"
    DefGtHor(6).texto = "6   " & "Oper. Negati"
    DefGtHor(7).texto = "7   " & "Importe"
    DefGtHor(8).texto = "8   " & "Impor.Cancel"
    DefGtHor(9).texto = "9   " & "Impor.Negati"
    DefGtHor(10).texto = "10  " & "Peso"
    DefGtHor(11).texto = "11  " & "Peso Cancel"
    DefGtHor(12).texto = "12  " & "Peso Negati"
    DefGtHor(13).texto = "13  " & "Unidades"
    DefGtHor(14).texto = "14  " & "Unid.Cancel"
    DefGtHor(15).texto = "15  " & "Unid.Negati"
    DefGtHor(0).Min = 1
    DefGtHor(1).Min = 6
    DefGtHor(2).Min = 2
    DefGtHor(3).Min = 5
    DefGtHor(4).Min = 5
    DefGtHor(5).Min = 5
    DefGtHor(6).Min = 5
    DefGtHor(7).Min = 7
    DefGtHor(8).Min = 7
    DefGtHor(9).Min = 7
    DefGtHor(10).Min = 5
    DefGtHor(11).Min = 5
    DefGtHor(12).Min = 5
    DefGtHor(13).Min = 5
    DefGtHor(14).Min = 5
    DefGtHor(15).Min = 5
   
End Sub

Private Sub Refresca_Datos()
    Dim Fichero As Integer
    Dim Buffer As String
    Dim Elinicio As String
    Dim Elfin As String
    Dim eltipo As String
    Dim ElTexto As String
    Fichero = FreeFile
    Open Miruta & "\etc\" & TxtNombre.TexT & ".fet" For Input As #Fichero
    Line Input #Fichero, Buffer
    If InStr(1, Buffer, "+") Then
        If HayMulti Then ChkNTienda.Value = vbChecked
        Buffer = left(Buffer, InStr(1, Buffer, "+") - 1)
    End If
    If Not IsNumeric(Buffer) Then
        Close #Fichero
        MsgBox CargaCadena(797)  '"Fichero de Configuración de Exportación Corrupto"
        Unload Me
        Exit Sub
    End If
    CmbElemento.Clear
    Select Case Val(Buffer)
        Case 1
            Tipo = 1 ' fichero de Gtarti
            CmbTipo.TexT = CmbTipo.List(0)
            Do Until EOF(Fichero)
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                Lista.ListItems.Add , , Elinicio
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
            Loop
        Case 2
            Tipo = 2 ' fichero de Gtsecs
            CmbTipo.TexT = CmbTipo.List(1)
            Do Until EOF(Fichero)
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                Lista.ListItems.Add , , Elinicio
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
            Loop
        Case 3
            Tipo = 3 ' fichero de Gtvend
            CmbTipo.TexT = CmbTipo.List(2)
            Do Until EOF(Fichero)
                On Error Resume Next
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                If Err.Number = 0 Then
                    Lista.ListItems.Add , , Elinicio
                    Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                    Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                    Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
                End If
                On Error GoTo 0
            Loop
        Case 4
            Tipo = 4 ' fichero de Cabeceras
            CmbTipo.TexT = CmbTipo.List(3)
            Input #Fichero, Elinicio, Elfin, eltipo
            Select Case Val(Elinicio)
                Case 1
                    Cmbcual.TexT = Cmbcual.List(0)
                    TipoExport = 1
                Case 2
                    Cmbcual.TexT = Cmbcual.List(1)
                    TipoExport = 2
                Case 3
                    Cmbcual.TexT = Cmbcual.List(2)
                    TipoExport = 3
                    Label6(0).Visible = True
                    Label6(1).Visible = True
                    TxtGA.Visible = True
                    TxtGA.TexT = Elfin
                    txtSC10.Visible = True
                    txtSC10.TexT = eltipo
            End Select
            Do Until EOF(Fichero)
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                Lista.ListItems.Add , , Elinicio
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
                
            Loop
        Case 5
            Tipo = 5 ' fichero de Lineas
            CmbTipo.TexT = CmbTipo.List(4)
            Input #Fichero, Elinicio, Elfin, eltipo
            Select Case Val(Elinicio)
                Case 1
                    Cmbcual.TexT = Cmbcual.List(0)
                    TipoExport = 1
                Case 2
                    Cmbcual.TexT = Cmbcual.List(1)
                    TipoExport = 2
                Case 3
                    Cmbcual.TexT = Cmbcual.List(2)
                    TipoExport = 3
                    Label6(0).Visible = True
                    Label6(1).Visible = True
                    TxtGA.Visible = True
                    TxtGA.TexT = Elfin
                    txtSC10.Visible = True
                    txtSC10.TexT = eltipo
            End Select
            Do Until EOF(Fichero)
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                Lista.ListItems.Add , , Elinicio
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
            Loop
        Case 6 ' cabeceras + líneas
            CmbTipo.TexT = CmbTipo.List(5)
            Line Input #Fichero, Buffer
            Text1(0).TexT = Buffer
            Line Input #Fichero, Buffer
            Text1(1).TexT = Buffer
            LblInfo.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            CmbElemento.Visible = False
            TxtInicio.Visible = False
            TxtFin.Visible = False
            CmdIncluir.Visible = False
            CmdEliminar.Visible = False
            Lista.Visible = False
            Text1(0).Visible = True
            Text1(1).Visible = True
            Label8(0).Visible = True
            Label8(1).Visible = True
        Case 7
            Tipo = 7 ' fichero de GtHor
            CmbTipo.TexT = CmbTipo.List(6)
            Do Until EOF(Fichero)
                Input #Fichero, Elinicio, Elfin, eltipo, ElTexto
                Lista.ListItems.Add , , Elinicio
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Elfin
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = eltipo
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = ElTexto
            Loop
            
        Case Else
            MsgBox CargaCadena(797)  '"Fichero de Configuración de Exportación Corrupto"
            Close #Fichero
            Unload Me
            Exit Sub
    End Select
    Close #Fichero
    If CmbTipo.TexT <> CmbTipo.List(5) Then refresca_Elementos
End Sub
Private Sub refresca_Elementos()
    Dim bucle As Integer
    Dim Elmaximo As Integer
    Select Case CmbTipo.ListIndex
        Case 0
            Elmaximo = 21
        Case 1
            Elmaximo = 12
        Case 2
            Elmaximo = 15
        Case 3
            Elmaximo = 37 'c2f invicta eran 15 ... 18 con el lote
        Case 4
            Elmaximo = 30 'c2f invicta 22, 25 con número de lote(numlote), 26 con cliente
        Case 6
            Elmaximo = 15
    End Select
    CmbElemento.Clear
    For bucle = 0 To Elmaximo
        Select Case CmbTipo.ListIndex
            Case 0
                CmbElemento.AddItem DefGtarti(bucle).texto
            Case 1
                CmbElemento.AddItem DefGtsecs(bucle).texto
            Case 2
                CmbElemento.AddItem DefGtVend(bucle).texto
            Case 3
                CmbElemento.AddItem DefCabecera(bucle).texto
            Case 4
                CmbElemento.AddItem DefLinea(bucle).texto
            Case 6
                CmbElemento.AddItem DefGtHor(bucle).texto
                
        End Select
    Next bucle
End Sub

Private Sub Ordenar_Lista()
    Dim bucle As Long
    Dim Bucle2 As Long
    Dim Buffer1 As String
    Dim Buffer2 As String
    Dim buffer3 As String
    Dim buffer4 As String
    For bucle = 1 To Lista.ListItems.Count - 1
        For Bucle2 = bucle To Lista.ListItems.Count
            If Val(Lista.ListItems.Item(bucle)) > Val(Lista.ListItems.Item(Bucle2)) Then
                Buffer1 = Lista.ListItems.Item(bucle).TexT
                Buffer2 = Lista.ListItems.Item(bucle).SubItems(1)
                buffer3 = Lista.ListItems.Item(bucle).SubItems(2)
                buffer4 = Lista.ListItems.Item(bucle).SubItems(3)
                Lista.ListItems.Item(bucle).TexT = Lista.ListItems.Item(Bucle2).TexT
                Lista.ListItems.Item(bucle).SubItems(1) = Lista.ListItems.Item(Bucle2).SubItems(1)
                Lista.ListItems.Item(bucle).SubItems(2) = Lista.ListItems.Item(Bucle2).SubItems(2)
                Lista.ListItems.Item(bucle).SubItems(3) = Lista.ListItems.Item(Bucle2).SubItems(3)
                Lista.ListItems.Item(Bucle2).TexT = Buffer1
                Lista.ListItems.Item(Bucle2).SubItems(1) = Buffer2
                Lista.ListItems.Item(Bucle2).SubItems(2) = buffer3
                Lista.ListItems.Item(Bucle2).SubItems(3) = buffer4
            End If
        Next Bucle2
    Next bucle
End Sub





Private Sub Cmbcual_Click()
    If Cmbcual.ListIndex <> 2 Then
        Label6(0).Visible = False
        Label6(1).Visible = False
        TxtGA.Visible = False
        txtSC10.Visible = False
    Else
        Label6(0).Visible = True
        Label6(1).Visible = True
        TxtGA.Visible = True
        txtSC10.Visible = True
    End If
End Sub

Private Sub CmbElemento_Click()
    If (CmbTipo.ListIndex = 3 And CmbElemento.ListIndex = 13) _
    Or (CmbTipo.ListIndex = 4 And CmbElemento.ListIndex = 16) Then
        TxtEAN(0).Visible = True
        TxtEAN(1).Visible = True
        If (CmbTipo.ListIndex = 3 And CmbElemento.ListIndex = 13) Then
            TxtEAN(0).TexT = 1
            TxtEAN(1).TexT = 13
        Else
            TxtEAN(0).TexT = 1
            TxtEAN(1).TexT = 1
        End If
    Else
        TxtEAN(1).Visible = False
        TxtEAN(0).Visible = False
    End If
End Sub

Private Sub CmbTipo_Click()
   If Not HayMulti Then
   If CmbTipo.ListIndex > 2 Then
        If (Not UsaGamaAlta) And (Not usaSC10) And (Not (UsaEcoPrint And TipoEcoPrint = 2)) Then
            MsgBox CargaCadena(1180), vbExclamation
        End If
    End If
    End If
End Sub

Private Sub Cmdaceptar_Click()
    Dim Fichero As Integer
    Dim bucle As Long
    Dim MyFlag As String
    MyFlag = ""
    If HayMulti And ChkNTienda.Value = vbChecked Then MyFlag = "+"
    If HayMulti Then Cmbcual.ListIndex = 2
    If CmbTipo.ListIndex = 5 Then
        If Trim(Text1(0).TexT) = "" Or Trim(Text1(1).TexT) = "" Then
            MsgBox CargaCadena(801), vbCritical + vbOKOnly
            Exit Sub
        End If
    End If
    If CmbTipo.ListIndex = 3 Or CmbTipo.ListIndex = 4 Then
        If Cmbcual.ListIndex = 2 Then
            'If TxtGA.Text = "" Then
            '    MsgBox CargaCadena( 798), vbCritical + vbOKOnly
            '    Exit Sub
            'End If
            'If TxtSC10.Text = "" Then
            '    MsgBox CargaCadena( 798), vbCritical + vbOKOnly
            '    Exit Sub
            'End If
            'If TxtSC10.Text = TxtGA.Text Then
            '    MsgBox CargaCadena( 799), vbCritical + vbOKOnly
            '    Exit Sub
            'End If
        End If
    End If
    
    If Lista.ListItems.Count > 0 Or CmbTipo.ListIndex = 5 Then
        Fichero = FreeFile
        Open Miruta & "\etc\" & TxtNombre.TexT & ".fet" For Output As #Fichero
        MyFlag = Format((CmbTipo.ListIndex + 1), "0") & MyFlag
        Print #Fichero, MyFlag
        If CmbTipo.ListIndex = 3 Or CmbTipo.ListIndex = 4 Then
            Write #Fichero, (Cmbcual.ListIndex + 1), TxtGA.TexT, txtSC10.TexT
        End If
        If CmbTipo.ListIndex <> 5 Then
            For bucle = 1 To Lista.ListItems.Count
                Write #Fichero, Lista.ListItems(bucle).TexT, Lista.ListItems(bucle).SubItems(1), Lista.ListItems(bucle).SubItems(2), Lista.ListItems(bucle).SubItems(3)
            Next bucle
        Else
            Print #Fichero, Text1(0).TexT
            Print #Fichero, Text1(1).TexT
        End If
        Close #Fichero
        MsgBox CargaCadena(800), vbInformation + vbOKOnly
    End If
    
    Unload Me
End Sub

Private Sub CmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub CmdConfigurar_Click()
    Dim bucle As Long
    If Trim(TxtNombre.TexT) = "" Then
        MsgBox CargaCadena(801)  '"Introduzca el nombre del fichero"
        TxtNombre.SetFocus
        Exit Sub
    End If

    For bucle = 1 To Len(TxtNombre.TexT)
        If Mid(TxtNombre.TexT, bucle, 1) = ":" Or Mid(TxtNombre.TexT, bucle, 1) = "\" Then
            MsgBox CargaCadena(802)  '"Nombre de fichero incorrecto"
            TxtNombre.SetFocus
            Exit Sub
        End If
    Next bucle
    
    If Dir(Miruta & "\etc\" & TxtNombre.TexT & ".fet") <> "" Then
        Lista.ColumnHeaders.Add 1, , CargaCadena(780), 1000
        Lista.ColumnHeaders.Add 2, , CargaCadena(781), 1000
        Lista.ColumnHeaders.Add 3, , CargaCadena(55), 2000
        Lista.ColumnHeaders.Add 4, , CargaCadena(796), 2900
        Frame1.Visible = False
        Frame2.top = Frame1.top
        FrmExportConf.Height = 5100
        LblInfo.Caption = CargaCadena(525)
        Refresca_Datos
    Else
        If CmbTipo.TexT = "" Then
            MsgBox CargaCadena(803)  '"Introduzca el tipo de fichero"
            Exit Sub
        End If
        
        If CmbTipo.ListIndex = 3 Or CmbTipo.ListIndex = 4 Then
            If UsaGamaAlta And usaSC10 Then
                Label7.Visible = True
                Cmbcual.Visible = True
                Cmbcual.TexT = Cmbcual.List(0)
            End If
            If UsaGamaAlta And Not usaSC10 Then
               Label7.Visible = False
               Cmbcual.Visible = False
               Cmbcual.TexT = Cmbcual.List(0)
            End If
            If Not UsaGamaAlta And usaSC10 Then
               Label7.Visible = False
               Cmbcual.Visible = False
               Cmbcual.TexT = Cmbcual.List(1)
            End If
        End If
        Lista.ColumnHeaders.Add 1, , CargaCadena(780), 1000
        Lista.ColumnHeaders.Add 2, , CargaCadena(781), 1000
        Lista.ColumnHeaders.Add 3, , CargaCadena(55), 2000
        Lista.ColumnHeaders.Add 4, , CargaCadena(796), 2900
        Frame1.Visible = False
        Frame2.top = Frame1.top
        FrmExportConf.Height = 5100
        refresca_Elementos
        If CmbTipo.ListIndex = 5 Then
            LblInfo.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            CmbElemento.Visible = False
            TxtInicio.Visible = False
            TxtFin.Visible = False
            CmdIncluir.Visible = False
            CmdEliminar.Visible = False
            Lista.Visible = False
            Text1(0).Visible = True
            Text1(1).Visible = True
            Label8(0).Visible = True
            Label8(1).Visible = True
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Lista.ListItems.Count > 0 Then
        Lista.ListItems.Remove (Lista.SelectedItem.Index)
    End If
    CmdEliminar.Enabled = False
End Sub

Private Sub CmdIncluir_Click()
    Dim ElTexto As String
    Dim bucle As Long
    Dim Resp
    If CmbElemento.TexT = "" Then
        MsgBox CargaCadena(804)  '"Seleccione elemento"
        CmbElemento.SetFocus
        Exit Sub
    End If
    If Trim(TxtInicio.TexT) = "" Then
        MsgBox CargaCadena(805)  '"Introduzca el Inicio"
        TxtInicio.SetFocus
        Exit Sub
    End If
    If Trim(TxtFin.TexT) = "" Then
        MsgBox CargaCadena(806)  '"Introduzca el Inicio"
        TxtFin.SetFocus
        Exit Sub
    End If
   ' If (Val(TxtFin.Text) - Val(TxtInicio.Text) + 1) < DefGtarti(CmbElemento.ListIndex).Min Then
   '     Resp = MsgBox(CargaCadena( 807) & " " & CmbElemento.Text & " : " & DefGtarti(CmbElemento.ListIndex).Min & ")", vbExclamation + vbOKCancel)
   '     If Resp = vbNo Then
   '         Exit Sub
   '     End If
   ' End If
    If Val(TxtInicio.TexT) = 0 Or Val(TxtFin.TexT) = 0 Then
        MsgBox CargaCadena(808), vbCritical + vbOKOnly
        Exit Sub
    End If
    For bucle = 1 To Lista.ListItems.Count
        If (Val(TxtInicio.TexT) >= Val(Lista.ListItems.Item(bucle).TexT) And Val(TxtInicio.TexT) <= Val(Lista.ListItems.Item(bucle).SubItems(1))) Or _
            (Val(TxtFin.TexT) >= Val(Lista.ListItems.Item(bucle).TexT) And Val(TxtFin.TexT) <= Val(Lista.ListItems.Item(bucle).SubItems(1))) Then
            MsgBox CargaCadena(795), vbCritical + vbOKOnly
            Exit Sub
        End If
    Next bucle
    For bucle = 1 To Lista.ListItems.Count - 1
        If (Val(TxtInicio.TexT) >= Val(Lista.ListItems.Item(bucle).SubItems(1))) And _
        ((Val(TxtInicio.TexT) - Val(TxtFin.TexT)) >= (Val(Lista.ListItems.Item(bucle + 1).TexT) - Val(Lista.ListItems.Item(bucle).SubItems(1)))) Then
            MsgBox CargaCadena(795), vbCritical + vbOKOnly
            Exit Sub
        End If
    Next bucle

    If CmbElemento.ListIndex = 0 Then
        ElTexto = InputBox(CargaCadena(1035) & "(max " & Val(TxtFin.TexT) - Val(TxtInicio.TexT) + 1 & "(caracteres)", "Texto Libre")
        If Len(ElTexto) > Val(TxtFin.TexT) - Val(TxtInicio.TexT) + 1 Then
            MsgBox CargaCadena(794), vbCritical + vbOKOnly
            Exit Sub
        End If
    End If
    '
    If (CmbTipo.ListIndex = 3 And CmbElemento.ListIndex = 13) Or _
    (CmbTipo.ListIndex = 4 And CmbElemento.ListIndex = 16) Then
        If TxtEAN(0).TexT = "" Or TxtEAN(1).TexT = "" Then
            MsgBox CargaCadena(1033), vbCritical + vbOKOnly
            TxtEAN(0).SetFocus
            Exit Sub
        End If
        If Val(TxtEAN(0).TexT) > Val(TxtEAN(1).TexT) Then
            MsgBox CargaCadena(1034), vbCritical + vbOKOnly
            TxtEAN(0).SetFocus
            Exit Sub
        End If
        ElTexto = Format(Val(TxtEAN(0).TexT), "00") & Format(Val(TxtEAN(1).TexT), "00")
    End If
        
               
    Lista.ListItems.Add , , TxtInicio.TexT
    Lista.ListItems.Item(Lista.ListItems.Count).SubItems(1) = TxtFin.TexT
    Lista.ListItems.Item(Lista.ListItems.Count).SubItems(2) = CmbElemento.TexT
    Lista.ListItems.Item(Lista.ListItems.Count).SubItems(3) = ElTexto
    Ordenar_Lista
    TxtInicio.TexT = ""
    TxtFin.TexT = ""
End Sub

Private Sub CambiarIdioma()
    With Me
        .Caption = CargaCadena(680)
        .CmbTipo.List(0) = CargaCadena(165)
        .CmbTipo.List(1) = CargaCadena(153)
        .CmbTipo.List(2) = CargaCadena(166)
        .CmbTipo.List(3) = CargaCadena(784)
        .CmbTipo.List(4) = CargaCadena(785)
        .CmbTipo.List(5) = CargaCadena(1260)
        .Cmbcual.List(2) = CargaCadena(566)
        .Frame1.Caption = CargaCadena(745)
        .Frame2.Caption = CargaCadena(778)
        .Label1.Caption = CargaCadena(552)
        .Label2.Caption = CargaCadena(55)
        .Label3.Caption = CargaCadena(779)
        .Label4.Caption = CargaCadena(780)
        .Label5.Caption = CargaCadena(781)
        .Label6(0).Caption = CargaCadena(782)
        .Label6(1).Caption = CargaCadena(783)
        .Label7.Caption = CargaCadena(692)
        .LblInfo.Caption = CargaCadena(777)
        .CmdConfigurar.Caption = CargaCadena(774)
        .CmdCancelar(0).Caption = CargaCadena(288)
        .CmdCancelar(1).Caption = CargaCadena(288)
        .CmdAceptar.Caption = CargaCadena(287)
        .CmdIncluir.Caption = CargaCadena(775)
        .CmdEliminar.Caption = CargaCadena(776)
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub Lista_Click()
    CmdEliminar.Enabled = True
End Sub

Private Sub TxtEAN_KeyPress(Index As Integer, KeyAscii As Integer)
    If (CmbTipo.ListIndex = 3 And CmbElemento.ListIndex = 13) Then
        TxtEAN(Index).Locked = Checktexto(KeyAscii, 1, TxtEAN(Index).TexT, True, False)
    Else
        TxtEAN(Index).Locked = Checktexto(KeyAscii, 4, TxtEAN(Index).TexT, True, False)
    End If
End Sub

Private Sub TxtFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdIncluir.SetFocus
    TxtFin.Locked = Checktexto(KeyAscii, 5, TxtFin.TexT, True, False)
End Sub



Private Sub TxtGA_KeyPress(KeyAscii As Integer)
    TxtGA.Locked = Checktexto(KeyAscii, 1, TxtGA.TexT, False, False)
End Sub

Private Sub TxtInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtFin.SetFocus
    TxtInicio.Locked = Checktexto(KeyAscii, 5, TxtInicio.TexT, True, False)
End Sub



Private Sub txtSC10_KeyPress(KeyAscii As Integer)
txtSC10.Locked = Checktexto(KeyAscii, 1, txtSC10.TexT, False, False)
End Sub

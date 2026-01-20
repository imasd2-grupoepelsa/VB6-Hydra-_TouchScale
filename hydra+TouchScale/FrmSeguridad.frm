VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSeguridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguridad"
   ClientHeight    =   5655
   ClientLeft      =   2775
   ClientTop       =   2955
   ClientWidth     =   8865
   Icon            =   "FrmSeguridad.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8865
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   1200
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguridad.frx":000C
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguridad.frx":0166
            Key             =   "cmdline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguridad.frx":02C0
            Key             =   "normal"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Salir"
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
      Left            =   7320
      TabIndex        =   28
      Top             =   5160
      Width           =   1455
   End
   Begin TabDlg.SSTab Tabla 
      Height          =   5055
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSeguridad.frx":041A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtTienda"
      Tab(0).Control(1)=   "ChkTienda"
      Tab(0).Control(2)=   "TxtMostrador"
      Tab(0).Control(3)=   "ChkAsociar"
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(5)=   "TxtNombre"
      Tab(0).Control(6)=   "CmdCancelar"
      Tab(0).Control(7)=   "cmdAceptar"
      Tab(0).Control(8)=   "TxtPassword(1)"
      Tab(0).Control(9)=   "TxtPassword(0)"
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(11)=   "Label2(1)"
      Tab(0).Control(12)=   "Label2(0)"
      Tab(0).Control(13)=   "Label1"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Tienda"
      TabPicture(1)   =   "FrmSeguridad.frx":0436
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1(25)"
      Tab(1).Control(1)=   "Check1(24)"
      Tab(1).Control(2)=   "Check1(23)"
      Tab(1).Control(3)=   "Check1(18)"
      Tab(1).Control(4)=   "Check1(10)"
      Tab(1).Control(5)=   "Check1(9)"
      Tab(1).Control(6)=   "Check1(8)"
      Tab(1).Control(7)=   "Check1(7)"
      Tab(1).Control(8)=   "Check1(6)"
      Tab(1).Control(9)=   "Check1(5)"
      Tab(1).Control(10)=   "Check1(4)"
      Tab(1).Control(11)=   "Check1(3)"
      Tab(1).Control(12)=   "Check1(2)"
      Tab(1).Control(13)=   "Check1(1)"
      Tab(1).Control(14)=   "Check1(0)"
      Tab(1).Control(15)=   "LblInfo(0)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Comunicaciones"
      TabPicture(2)   =   "FrmSeguridad.frx":0452
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1(27)"
      Tab(2).Control(1)=   "Check1(26)"
      Tab(2).Control(2)=   "Check1(19)"
      Tab(2).Control(3)=   "Check1(22)"
      Tab(2).Control(4)=   "Check1(14)"
      Tab(2).Control(5)=   "Check1(13)"
      Tab(2).Control(6)=   "Check1(12)"
      Tab(2).Control(7)=   "Check1(11)"
      Tab(2).Control(8)=   "LblInfo(1)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Ficheros"
      TabPicture(3)   =   "FrmSeguridad.frx":046E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check1(33)"
      Tab(3).Control(1)=   "Check1(32)"
      Tab(3).Control(2)=   "Check1(31)"
      Tab(3).Control(3)=   "Check1(30)"
      Tab(3).Control(4)=   "Check1(29)"
      Tab(3).Control(5)=   "Check1(28)"
      Tab(3).Control(6)=   "Check1(17)"
      Tab(3).Control(7)=   "Check1(21)"
      Tab(3).Control(8)=   "Check1(20)"
      Tab(3).Control(9)=   "Check1(16)"
      Tab(3).Control(10)=   "Check1(15)"
      Tab(3).Control(11)=   "LblInfo(2)"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FrmSeguridad.frx":048A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Check1(34)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Check1(35)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Check1(36)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Check1(37)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Check1(38)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Check1(39)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Check1(40)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Check1(41)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Check1(42)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Check1(43)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72120
         MaxLength       =   3
         TabIndex        =   64
         Top             =   1440
         Width           =   612
      End
      Begin VB.CheckBox ChkTienda 
         Appearance      =   0  'Flat
         Caption         =   "Asociar a una Tienda"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74760
         TabIndex        =   63
         Top             =   1440
         Width           =   2172
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 10"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   43
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   62
         Top             =   3360
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 9"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   42
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   61
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 8"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   41
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   60
         Top             =   2880
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 7"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   59
         Top             =   2640
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 6"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   39
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   58
         Top             =   2400
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   38
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   57
         Top             =   2160
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   37
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   56
         Top             =   1920
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   36
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   55
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   54
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a menú 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   53
         Top             =   1200
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Gestión de Copias de Seguridad"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   52
         Top             =   2460
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Comando Libre 5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   51
         Top             =   3660
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Comando Libre 4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   50
         Top             =   3420
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Comando Libre 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   49
         Top             =   3180
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Comando Libre 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   48
         Top             =   2940
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Comando Libre 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   47
         Top             =   2700
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Programar StandBy de Maquinas Euroscale"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   46
         Top             =   2940
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Programar Fecha y Hora"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   45
         Top             =   2700
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a formularios de Totales de Ventas, envasado..."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   44
         Top             =   4620
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Artículos (consulta y programación)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   -74880
         TabIndex        =   43
         Top             =   1740
         Width           =   5055
      End
      Begin VB.TextBox TxtMostrador 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72120
         TabIndex        =   42
         Top             =   1800
         Width           =   612
      End
      Begin VB.CheckBox ChkAsociar 
         Appearance      =   0  'Flat
         Caption         =   "Asociar a un mostrador"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74760
         TabIndex        =   41
         Top             =   1860
         Width           =   2172
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Configuración General de NTQ"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   40
         Top             =   1740
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Borrado de Totales de Ventas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   39
         Top             =   4380
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Regenerar Bases de Datos Vacías"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   38
         Top             =   4140
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Borrados de Bases de Concentrador y Euroscale"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   37
         Top             =   2460
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Conectar/Desconectar ventas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   36
         Top             =   2220
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Exportación de Ficheros"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   35
         Top             =   2220
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Importación de Ficheros"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   34
         Top             =   1980
         Width           =   5055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cambiar"
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
         Left            =   -70440
         TabIndex        =   30
         Top             =   900
         Width           =   1335
      End
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72840
         MaxLength       =   8
         TabIndex        =   29
         Top             =   972
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Configuración de Ficheros de Exportación"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   27
         Top             =   1500
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a Configuración de Ficheros de Importación"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   26
         Top             =   1260
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Recogida y/o Borrado de Grandes Totales"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   25
         Top             =   1980
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Envío de Modificaciones y Modificaciones por separado"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   24
         Top             =   1740
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Envío de Datos por separado"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   23
         Top             =   1260
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Envío de Base Maestra "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   -74880
         MaskColor       =   &H00000000&
         TabIndex        =   22
         Top             =   1500
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Teclas Rápidas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   21
         Top             =   3900
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Ventas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   20
         Top             =   3660
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Publicidad y Ofertas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   -74880
         TabIndex        =   19
         Top             =   3432
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Códigos de Barras"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   -74880
         TabIndex        =   18
         Top             =   3192
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Cabereras y Leyendas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   -74880
         TabIndex        =   17
         Top             =   2952
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Tipos de IVA"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   -74880
         TabIndex        =   16
         Top             =   2712
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Países y Fichas de Vacuno"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   15
         Top             =   2472
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Grupos de Conservación"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   14
         Top             =   2232
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Cambios de Precio"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   13
         Top             =   1992
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Artículos (sólo consulta)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   12
         Top             =   1512
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Acceso a configuración de Mostradores y Equipos"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   11
         Top             =   1272
         Width           =   5055
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   -70560
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
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
         Left            =   -74760
         TabIndex        =   9
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         Height          =   288
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   -72600
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   3420
         Width           =   2055
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         Height          =   288
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   -72600
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   3060
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar Contraseña"
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
         Left            =   -74040
         TabIndex        =   4
         Top             =   2340
         Width           =   3852
      End
      Begin VB.Label LblInfo 
         BackColor       =   &H80000018&
         Caption         =   "Label3"
         Height          =   252
         Index           =   2
         Left            =   -73440
         TabIndex        =   33
         Top             =   780
         Width           =   2412
      End
      Begin VB.Label LblInfo 
         BackColor       =   &H80000018&
         Caption         =   "Label3"
         Height          =   252
         Index           =   1
         Left            =   -73440
         TabIndex        =   32
         Top             =   780
         Width           =   2412
      End
      Begin VB.Label LblInfo 
         BackColor       =   &H80000018&
         Caption         =   "Label3"
         Height          =   252
         Index           =   0
         Left            =   -73440
         TabIndex        =   31
         Top             =   780
         Width           =   2412
      End
      Begin VB.Label Label2 
         Caption         =   "Escribala otra vez :"
         Height          =   252
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   3420
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Escriba la contraseña :"
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   3060
         Width           =   1932
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de entrada :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   1032
         Width           =   1815
      End
   End
   Begin VB.CheckBox ChkSeguridad 
      Appearance      =   0  'Flat
      Caption         =   "Trabajar con usuarios y contraseñas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   6732
   End
   Begin MSComctlLib.TreeView Arbol 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8916
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General"
      Begin VB.Menu CreaUsuario 
         Caption         =   "Crear usuario"
      End
      Begin VB.Menu BorraUsuario 
         Caption         =   "Borrar Usuario"
      End
   End
End
Attribute VB_Name = "FrmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HabiaSeguridad As Boolean
Private CantidadUsuarios As Integer
Private QuantumLeap  As Boolean
Private QL2 As Boolean
Private QL3 As Boolean
'\\\\\\\\\\\\\\\\\\\\\
'/////////////////////
Private Sub Form_Load()
Dim Arch As Integer
Dim Buf1 As String
Dim Buf2 As String
    HabiaSeguridad = HaySeguridad
    
    If HayMulti Then
        ChkTienda.Visible = True
        txtTienda.Visible = True
        txtTienda.Enabled = False
    Else
        ChkTienda.Visible = False
        txtTienda.Visible = False
    End If
    CambiarIdioma
    Me.Icon = Form2.Icon
    TxtNombre.Enabled = False
    Arbol.ImageList = Imagenes
    mnuGeneral.Visible = False
    Modo_password (False)
    If Dir(MiruTa & "\hydplugin.exe") <> "" Then
        Tabla.TabVisible(4) = True
        Tabla.TabCaption(4) = "Plugin"
    Else
        Tabla.TabVisible(4) = False
    End If
    Tabla.Tab = 0
    CargaUsuarios
    Arbol.Nodes(1).Selected = True
    Tabla.TabEnabled(1) = False
    Tabla.TabEnabled(2) = False
    Tabla.TabEnabled(3) = False
    Tabla.TabEnabled(4) = False
    '
    If HayMulti Then
        If Dir(App.Path & "\hydramulti.ini") <> "" Then
            Arch = FreeFile()
            Open App.Path & "\hydramulti.ini" For Input As #Arch
            Input #Arch, Buf1, Buf2
            Close #Arch
            Buf1 = DesEncripta(Buf1)
            If Not IsNumeric(left(Buf1, 8)) Then
                MsgBox CargaCadena(851), vbCritical
                End
            Else
                If (Val(left(Buf1, 4)) + 3) = Val(Mid(Buf1, 5, 4)) Then
                    HaySeguridad = False
                Else
                    If (Val(left(Buf1, 4)) + 1) = Val(Mid(Buf1, 5, 4)) Then
                        HaySeguridad = True
                    Else
                        MsgBox CargaCadena(851), vbCritical
                        End
                    End If
                End If
            End If
        Else
            HaySeguridad = False
        End If
    End If
    '
    If HaySeguridad Then
        QL3 = True
        ChkSeguridad.Value = vbChecked
        QL3 = False
    Else
        QL3 = True
        ChkSeguridad.Value = vbUnchecked
        QL3 = False
    End If
    Command1.Enabled = False
    ChkAsociar.Enabled = False
    TxtMostrador.Enabled = False
End Sub

Private Sub Arbol_Click()
    Dim Bucle As Integer
    If QuantumLeap Then Exit Sub
    Refresca_Datos
End Sub
Public Sub Desglosa_Permisos(Permisos() As Byte)
Dim Valor As Integer
Dim Bucle As Integer
    Valor = Permisos(0)
    QL2 = True
    For Bucle = 0 To 7
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle).Value = vbUnchecked
        Else
            Check1(Bucle).Value = vbChecked
        End If
    Next Bucle
    Valor = Permisos(1)
    For Bucle = 0 To 7
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle + 8).Value = vbUnchecked
        Else
            Check1(Bucle + 8).Value = vbChecked
        End If
    Next Bucle
    Valor = Permisos(2)
    For Bucle = 0 To 7
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle + 16).Value = vbUnchecked
        Else
            Check1(Bucle + 16).Value = vbChecked
        End If
    Next Bucle
    Valor = Permisos(3)
    
    For Bucle = 0 To 7
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle + 24).Value = vbUnchecked
        Else
            Check1(Bucle + 24).Value = vbChecked
        End If
    Next Bucle
    Valor = Permisos(4)
    For Bucle = 0 To 7
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle + 32).Value = vbUnchecked
        Else
            Check1(Bucle + 32).Value = vbChecked
        End If
    Next Bucle
    Valor = Permisos(5)
    For Bucle = 0 To 3
        If (Valor And (2 ^ Bucle)) = 0 Then
            Check1(Bucle + 40).Value = vbUnchecked
        Else
            Check1(Bucle + 40).Value = vbChecked
        End If
    Next Bucle
    QL2 = False
End Sub
Private Sub Arbol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        QuantumLeap = True
        If Arbol.SelectedItem.TexT = "admin" Or Arbol.SelectedItem.TexT = "cmdline" Then
            BorraUsuario.Enabled = False
        Else
            BorraUsuario.Enabled = True
        End If
        PopupMenu mnuGeneral, 0, x + 200, y + 200
    Else
        QuantumLeap = False
    End If
End Sub

Private Sub BorraUsuario_Click()
Dim Bucle As Integer
Dim B2 As Integer
Dim BufUser() As Tipo_User
Dim usrAct As Integer
Dim Inc As Integer
    Inc = 0
    ReDim BufUser(CantidadUsuarios)
    usrAct = Arbol.SelectedItem.Index
    For Bucle = 1 To CantidadUsuarios
        If Bucle <> usrAct Then
            BufUser(Bucle - Inc).Nombre = usuarios(Bucle).Nombre
            BufUser(Bucle - Inc).Password = usuarios(Bucle).Password
            For B2 = 0 To 7
                BufUser(Bucle - Inc).Permisos(B2) = usuarios(Bucle).Permisos(B2)
            Next B2
        Else
            Inc = 1
        End If
    Next Bucle
    CantidadUsuarios = CantidadUsuarios - 1
    ReDim usuarios(CantidadUsuarios)
    For Bucle = 1 To CantidadUsuarios
        usuarios(Bucle).Nombre = BufUser(Bucle).Nombre
        usuarios(Bucle).Password = BufUser(Bucle).Password
        For B2 = 0 To 7
            usuarios(Bucle).Permisos(B2) = BufUser(Bucle).Permisos(B2)
        Next B2
    Next Bucle
    EscribirUsuarios
    CargaUsuarios
    Arbol.Nodes(1).Selected = True
    Refresca_Datos
End Sub


Private Sub Check1_Click(Index As Integer)
    Dim UsAct As Integer
    Dim Bucle As Integer
    If QL2 Then Exit Sub
    UsAct = Arbol.SelectedItem.Index
    
    usuarios(UsAct).Permisos(0) = Check1(0).Value
    
    For Bucle = 1 To 7
        usuarios(UsAct).Permisos(0) = usuarios(UsAct).Permisos(0) + (2 * Check1(Bucle).Value) ^ Bucle
    Next Bucle
    
    usuarios(UsAct).Permisos(1) = Check1(8).Value
    For Bucle = 1 To 7
        usuarios(UsAct).Permisos(1) = usuarios(UsAct).Permisos(1) + (2 * Check1(Bucle + 8).Value) ^ Bucle
    Next Bucle
    usuarios(UsAct).Permisos(2) = Check1(16).Value
    For Bucle = 1 To 7
        usuarios(UsAct).Permisos(2) = usuarios(UsAct).Permisos(2) + (2 * Check1(Bucle + 16).Value) ^ Bucle
    Next Bucle
    usuarios(UsAct).Permisos(3) = Check1(24).Value
    For Bucle = 1 To 7
        usuarios(UsAct).Permisos(3) = usuarios(UsAct).Permisos(3) + (2 * Check1(Bucle + 24).Value) ^ Bucle
    Next Bucle
    usuarios(UsAct).Permisos(4) = Check1(32).Value
    For Bucle = 1 To 7
        usuarios(UsAct).Permisos(4) = usuarios(UsAct).Permisos(4) + (2 * Check1(Bucle + 32).Value) ^ Bucle
    Next Bucle
    usuarios(UsAct).Permisos(5) = Check1(40).Value
    For Bucle = 1 To 3
        usuarios(UsAct).Permisos(5) = usuarios(UsAct).Permisos(5) + (2 * Check1(Bucle + 40).Value) ^ Bucle
    Next Bucle
    EscribirUsuarios
    Arbol.Nodes(UsAct).Selected = True
End Sub

Private Sub ChkAsociar_Click()
    If ChkAsociar.Value = vbChecked Then
        TxtMostrador.Enabled = True
    Else
        TxtMostrador.Enabled = False
        TxtMostrador.TexT = "0"
        TxtMostrador_LostFocus
    End If
End Sub

Private Sub ChkSeguridad_Click()
    Dim Archivo As Integer
    If QL3 Then Exit Sub
    If ChkSeguridad.Value = vbChecked Then
        'If Dir(MiruTa & "\pass.sec") = "" Then
        If Dir(App.Path & "\pass.sec") = "" Then
            Archivo = FreeFile()
            'Open MiruTa & "\pass.sec" For Output As Archivo
            Open App.Path & "\pass.sec" For Output As Archivo
            Close #Archivo
        End If
        HaySeguridad = True
        Activa_Seguridad_Multi
        If Not HayMulti Then
            Load FrmConfiguracion2
            EscribirParametrosInicio
            Unload FrmConfiguracion2
            MsgBox CargaCadena(767)  '"Ha activado la Seguridad de Usuarios.Reinicie el programa", vbExclamation
        End If
    Else
        HaySeguridad = False
        Activa_Seguridad_Multi
        If Not HayMulti Then
            Load FrmConfiguracion2
            EscribirParametrosInicio
            Unload FrmConfiguracion2
            MsgBox CargaCadena(768)  '"Ha desactivado la Seguridad de Usuarios", vbExclamation
        End If
    End If
End Sub
Private Sub Activa_Seguridad_Multi()
    '
    Dim MiValor As Integer
    Dim MiValor2 As Integer
    Dim Arch As Integer
    Dim Buf() As String
    Dim BufCont As Integer
    Dim Bucle As Integer
    If HayMulti Then
        If Dir(App.Path & "\hydramulti.ini") <> "" Then
            BufCont = 0
            Arch = FreeFile()
            Open App.Path & "\hydramulti.ini" For Input As #Arch
            Do Until EOF(Arch)
                BufCont = BufCont + 1
                ReDim Preserve Buf(BufCont)
                Line Input #Arch, Buf(BufCont - 1)
            Loop
            Close #Arch
            Randomize Time
            MiValor = Int((4500 * Rnd) + 1)
            If HaySeguridad Then
                MiValor2 = MiValor + 1
            Else
                MiValor2 = MiValor + 3
            End If
            Arch = FreeFile()
            Open App.Path & "\hydramulti.ini" For Output As #Arch
            Write #Arch, Encripta(Format(MiValor, "0000") & Format(MiValor2, "0000")), "NEVER EDIT NOR REMOVE THIS LINE"
            For Bucle = 1 To BufCont - 1
                Print #Arch, Buf(Bucle)
            Next Bucle
            Close #Arch
        End If
    End If
    '
End Sub

Private Sub ChkTienda_Click()
    If ChkTienda.Value = vbChecked Then
        txtTienda.Enabled = True
    Else
        txtTienda.Enabled = False
        txtTienda.TexT = "0"
    End If
End Sub

Private Sub Cmdaceptar_Click()
    TxtPassword(0).TexT = StrConv(TxtPassword(0).TexT, vbLowerCase)
    TxtPassword(1).TexT = StrConv(TxtPassword(1).TexT, vbLowerCase)
    If Trim(TxtPassword(0).TexT) = "" Or Trim(TxtPassword(1).TexT) = "" Then
        MsgBox CargaCadena(769)  '"Contraseña Incorrecta", vbCritical
        Exit Sub
    End If
    If Trim(TxtPassword(0).TexT) <> Trim(TxtPassword(1).TexT) Then
        MsgBox CargaCadena(770)  '"Las contraseñas no coinciden", vbCritical
        Exit Sub
    End If
    usuarios(Arbol.SelectedItem.Index).Password = Trim(TxtPassword(0).TexT)
    EscribirUsuarios
    Modo_password (False)
End Sub

Private Sub CmdCancelar_Click()
    Command1.Enabled = True
    Modo_password (False)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Modo_password (True)
End Sub

Private Sub Command2_Click()
    Dim Bucle As Integer
    Dim Buffer As Integer
    If Trim(TxtNombre.TexT) = "" Then
        MsgBox CargaCadena(771)  '"Nombre no válido"
        Exit Sub
    End If
    If IsNumeric(TxtNombre.TexT) Then
        MsgBox CargaCadena(771)  '"Nombre no válido"
        Exit Sub
    End If
    For Bucle = 1 To CantidadUsuarios
        If Trim(Arbol.Nodes(Bucle).TexT) = Trim(TxtNombre.TexT) Then
            MsgBox CargaCadena(771)  '"Nombre no válido"
            Exit Sub
        End If
    Next Bucle
    usuarios(Arbol.SelectedItem.Index).Nombre = TxtNombre.TexT
    Buffer = Arbol.SelectedItem.Index
    lblinfo(0).Caption = TxtNombre.TexT
    lblinfo(1).Caption = TxtNombre.TexT
    lblinfo(2).Caption = TxtNombre.TexT
    EscribirUsuarios
    CargaUsuarios
    Arbol.Nodes(Buffer).Selected = True

End Sub

Private Sub CreaUsuario_Click()
    Dim Bucle As Integer
    Dim cadena As String
    Dim Contador As Integer
    Dim Encontrado As Boolean
    Dim MiNodo As Integer
    Contador = 1
    cadena = "user"
    Do Until Encontrado
        For Bucle = 1 To Arbol.Nodes.Count
            If Trim(Arbol.Nodes(Bucle).TexT) = (cadena & Contador) Then
                Contador = Contador + 1
                Exit For
            End If
        Next Bucle
        If Arbol.Nodes.Count < Bucle Then
            Encontrado = True
            cadena = cadena & Contador
        End If
    Loop
    QL2 = True
    For Bucle = 0 To 19
        Check1(Bucle).Value = vbChecked
    Next Bucle
    QL2 = False
    Arbol.Nodes.Add , tvwChild, cadena, cadena
    
    CantidadUsuarios = CantidadUsuarios + 1
    ReDim Preserve usuarios(CantidadUsuarios + 1)
    usuarios(CantidadUsuarios).Nombre = cadena
    usuarios(CantidadUsuarios).Password = "drake"
    usuarios(CantidadUsuarios).Permisos(0) = 255
    usuarios(CantidadUsuarios).Permisos(1) = 255
    usuarios(CantidadUsuarios).Permisos(2) = 255
    usuarios(CantidadUsuarios).Permisos(3) = 255
    usuarios(CantidadUsuarios).Permisos(4) = 255
    usuarios(CantidadUsuarios).Permisos(7) = 255
    usuarios(CantidadUsuarios).Asociado = 0
    EscribirUsuarios
    CargaUsuarios
    Arbol.Nodes(Arbol.Nodes.Count).Selected = True
    Refresca_Datos
End Sub
Public Sub EscribirUsuarios()
    Dim Archivo As Integer
    Dim MiCheck
    Dim Buffer As String
    Dim BufByte As Byte
    Dim B2 As Integer
    Dim MiPerm(4) As Byte
    Dim MiTienda As Byte
    Dim Bucle As Integer
    Dim Bucle2 As Integer
    Archivo = FreeFile
    'Open MiruTa & "\pass.sec" For Output As #Archivo
    Open App.Path & "\pass.sec" For Output As #Archivo
    Close #Archivo
    Archivo = FreeFile
    'Open MiruTa & "\pass.sec" For Binary As #Archivo
    Open App.Path & "\pass.sec" For Binary As #Archivo
    For Bucle2 = 1 To CantidadUsuarios
        Buffer = usuarios(Bucle2).Nombre & Space(8 - Len(usuarios(Bucle2).Nombre))
        Buffer = Buffer & usuarios(Bucle2).Password & Space(8 - Len(usuarios(Bucle2).Password))
        MiPerm(0) = usuarios(Bucle2).Permisos(0)
        MiPerm(1) = usuarios(Bucle2).Permisos(1)
        MiPerm(2) = usuarios(Bucle2).Permisos(2)
        MiPerm(3) = usuarios(Bucle2).Permisos(3)
        MiPerm(4) = usuarios(Bucle2).Permisos(4)
        MiTienda = usuarios(Bucle2).Permisos(7)
        
        
        Buffer = Buffer & Chr(MiPerm(0)) & Chr(MiPerm(1)) & Chr(MiPerm(2)) _
        & Chr(MiPerm(3)) & Chr(MiPerm(4)) & Chr(0) & Chr(0) & Chr(MiTienda)
        Buffer = Buffer & Format(usuarios(Bucle2).Asociado, "000")
        Buffer = Buffer & Chr(Bucle2)
        calcheck Buffer, MiCheck
        Buffer = Buffer & Chr(MiCheck)
        Buffer = Encripta(Buffer)
        For B2 = 1 To Len(Buffer)
            BufByte = Asc(Mid(Buffer, B2, 1))
            Put #Archivo, , BufByte
        Next B2
    Next Bucle2
    Close #Archivo
End Sub
Private Sub CambiarIdioma()
    FrmSeguridad.Caption = CargaCadena(523)
    Tabla.TabCaption(0) = CargaCadena(753)
    Tabla.TabCaption(1) = CargaCadena(46)
    Tabla.TabCaption(2) = CargaCadena(4)
    Tabla.TabCaption(3) = CargaCadena(745)
    Label1.Caption = CargaCadena(520)
    ChkAsociar.Caption = CargaCadena(754)
    Label2(0).Caption = CargaCadena(755)
    Label2(1).Caption = CargaCadena(756)
    Command2.Caption = CargaCadena(759)
    Command1.Caption = CargaCadena(758)
    cmdAceptar.Caption = CargaCadena(287)
    CmdCancelar.Caption = CargaCadena(288)
    CmdExit.Caption = CargaCadena(7)
    ChkSeguridad.Caption = CargaCadena(757)
    '
    Check1(0).Caption = CargaCadena(760) & " " & CargaCadena(580)
    Check1(1).Caption = CargaCadena(760) & " " & CargaCadena(526) & " " & CargaCadena(772)
    Check1(2).Caption = CargaCadena(760) & " " & CargaCadena(24)
    Check1(3).Caption = CargaCadena(760) & " " & CargaCadena(638)
    Check1(4).Caption = CargaCadena(760) & " " & CargaCadena(673)
    Check1(5).Caption = CargaCadena(760) & " " & CargaCadena(8)
    Check1(6).Caption = CargaCadena(760) & " " & CargaCadena(716)
    Check1(7).Caption = CargaCadena(760) & " " & CargaCadena(23)
    Check1(8).Caption = CargaCadena(760) & " " & CargaCadena(505)
    Check1(9).Caption = CargaCadena(760) & " " & CargaCadena(25)
    Check1(10).Caption = CargaCadena(760) & " " & CargaCadena(22)
    Check1(11).Caption = CargaCadena(764)
    Check1(12).Caption = CargaCadena(763)
    Check1(13).Caption = CargaCadena(765)
    Check1(14).Caption = CargaCadena(160)
    Check1(15).Caption = CargaCadena(760) & " " & CargaCadena(680)
    Check1(16).Caption = CargaCadena(760) & " " & CargaCadena(681)
    Check1(17).Caption = CargaCadena(760) & " " & CargaCadena(683)
    Check1(18).Caption = CargaCadena(760) & " " & CargaCadena(249)
    Check1(19).Caption = CargaCadena(760) & " " & CargaCadena(762)
    Check1(20).Caption = CargaCadena(685)
    Check1(21).Caption = CargaCadena(686)
    Check1(22).Caption = CargaCadena(766)
    Check1(23).Caption = CargaCadena(282)
    Check1(24).Caption = CargaCadena(760) & " " & CargaCadena(526) & " " & CargaCadena(773)
    Check1(25).Caption = CargaCadena(760) & " " & CargaCadena(113)
    Check1(26).Caption = CargaCadena(101)
    Check1(27).Caption = CargaCadena(103)
    Check1(28).Caption = CargaCadena(760) & " " & CargaCadena(761) & " 1"
    Check1(29).Caption = CargaCadena(760) & " " & CargaCadena(761) & " 2"
    Check1(30).Caption = CargaCadena(760) & " " & CargaCadena(761) & " 3"
    Check1(31).Caption = CargaCadena(760) & " " & CargaCadena(761) & " 4"
    Check1(32).Caption = CargaCadena(760) & " " & CargaCadena(761) & " 5"
    Check1(33).Caption = CargaCadena(32)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v140
    Me.CreaUsuario.Caption = CargaCadena(1272)
    Me.BorraUsuario.Caption = CargaCadena(1273)
    '//////////////////////////////////////////
End Sub
Private Sub Modo_password(Si_o_no As Boolean)
    Label2(0).Visible = Si_o_no
    Label2(1).Visible = Si_o_no
    TxtPassword(0).Visible = Si_o_no
    TxtPassword(1).Visible = Si_o_no
    cmdAceptar.Visible = Si_o_no
    CmdCancelar.Visible = Si_o_no
    Arbol.Enabled = Not (Si_o_no)
    ChkSeguridad.Enabled = Not (Si_o_no)
    Command1.Enabled = Not (Si_o_no)
    Tabla.TabEnabled(0) = Not (Si_o_no)
    If Trim(TxtNombre.TexT) <> "admin" Then
        Tabla.TabEnabled(1) = Not (Si_o_no)
        Tabla.TabEnabled(2) = Not (Si_o_no)
        Tabla.TabEnabled(3) = Not (Si_o_no)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
    If Not HayMulti And HabiaSeguridad <> HaySeguridad Then salir_programa
    Exit Sub
FIN:
End Sub



Private Sub Tabla_Click(PreviousTab As Integer)
    QL2 = False
End Sub


Private Sub TxtMostrador_KeyPress(KeyAscii As Integer)
    TxtMostrador.Locked = Checktexto(KeyAscii, 3, TxtMostrador.TexT, True, False)
End Sub



Private Sub TxtMostrador_LostFocus()
    Dim UsAct As Integer
    UsAct = Arbol.SelectedItem.Index
    If Not IsNumeric(TxtMostrador.TexT) Then
        TxtMostrador.TexT = "0"
        usuarios(UsAct).Asociado = 0
    Else
        usuarios(UsAct).Asociado = Val(TxtMostrador.TexT)
    End If
    EscribirUsuarios
End Sub


Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    TxtNombre.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) _
    Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
        TxtNombre.Locked = True
    End If
End Sub


Private Sub TxtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
TxtPassword(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) _
    Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
        TxtPassword(Index).Locked = True
    End If
End Sub


Private Sub TxtTienda_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        If Len(txtTienda.TexT) = 0 Then
            If KeyAscii = 48 Then KeyAscii = 0
        End If
        If Len(txtTienda.TexT) = 2 Then
            If Val((txtTienda.TexT) & Chr(KeyAscii)) <> 999 Then
            
            
                Select Case Val(left(txtTienda.TexT, 1))
                    Case 0
                    Case 1
                    Case 2
                        Select Case Val(Right(txtTienda.TexT, 1))
                            Case 0 To 4
                            Case 5
                                If Not (KeyAscii >= 48 And KeyAscii <= 53) Then KeyAscii = 0
                            Case Else
                                KeyAscii = 0
                        End Select
                    Case Else
                        KeyAscii = 0
                End Select
            End If
        End If
    End If
End Sub

Private Sub TxtTienda_LostFocus()
    Dim UsAct As Integer
    UsAct = Arbol.SelectedItem.Index
    If Not IsNumeric(txtTienda.TexT) Then
        txtTienda.TexT = "0"
        usuarios(UsAct).Permisos(7) = 255
    Else
        If Val(txtTienda.TexT) > 0 Then
            If Val(txtTienda.TexT) = 999 Then
                usuarios(UsAct).Permisos(7) = (255 - 1)
            Else
                usuarios(UsAct).Permisos(7) = (Val(txtTienda.TexT) - 1)
            End If
        Else
            usuarios(UsAct).Permisos(7) = 255
        End If
    End If
    EscribirUsuarios
End Sub
'//////
'//////
'//////
'////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////// METODOS
Private Sub Refresca_Datos()
    Dim Bucle As Integer
    Tabla.Tab = 0
    Select Case Arbol.SelectedItem.TexT
    Case "cmdline"
        Tabla.TabEnabled(0) = False
        Tabla.TabEnabled(1) = True
        Tabla.TabEnabled(2) = True
        Tabla.TabEnabled(3) = True
        Tabla.TabEnabled(4) = True
        TxtNombre.Enabled = False
        Command1.Enabled = False
        Command2.Enabled = False
        ChkAsociar.Value = False
        TxtMostrador.Enabled = False
        TxtMostrador.TexT = "0"
        ChkAsociar.Enabled = False
        ChkTienda.Enabled = False
        txtTienda.Enabled = False
        txtTienda.TexT = "0"
    Case "admin"
        Tabla.TabEnabled(0) = True
        Tabla.TabEnabled(1) = False
        Tabla.TabEnabled(2) = False
        Tabla.TabEnabled(3) = False
        Tabla.TabEnabled(4) = False
        TxtNombre.Enabled = False
        Command2.Enabled = False
        Command1.Enabled = True
        ChkAsociar.Value = False
        TxtMostrador.Enabled = False
        TxtMostrador.TexT = "0"
        ChkAsociar.Enabled = False
        ChkTienda.Enabled = False
        txtTienda.Enabled = False
        txtTienda.TexT = "0"
    Case Else
        Tabla.TabEnabled(0) = True
        Tabla.TabEnabled(1) = True
        Tabla.TabEnabled(2) = True
        Tabla.TabEnabled(3) = True
        Tabla.TabEnabled(4) = True
        TxtNombre.Enabled = True
        Command2.Enabled = True
        Command1.Enabled = True
        ChkAsociar.Enabled = True
        ChkTienda.Enabled = True
        txtTienda.Enabled = True
    End Select
    For Bucle = 1 To CantidadUsuarios
        If Trim(Arbol.SelectedItem.TexT) = Trim(usuarios(Bucle).Nombre) Then
            TxtNombre.TexT = usuarios(Bucle).Nombre
            lblinfo(0).Caption = TxtNombre.TexT
            lblinfo(1).Caption = TxtNombre.TexT
            lblinfo(2).Caption = TxtNombre.TexT
            TxtPassword(0).TexT = usuarios(Bucle).Password
            TxtPassword(1).TexT = usuarios(Bucle).Password
            If usuarios(Bucle).Asociado = 0 Then
                ChkAsociar.Value = vbUnchecked
                TxtMostrador.TexT = "0"
            Else
                ChkAsociar.Value = vbChecked
                TxtMostrador.TexT = usuarios(Bucle).Asociado
            End If
            If usuarios(Bucle).Permisos(7) = 255 Then
                ChkTienda.Value = vbUnchecked
                txtTienda.TexT = "0"
            Else
                ChkTienda.Value = vbChecked
                
                If usuarios(Bucle).Permisos(7) = 254 Then
                    txtTienda.TexT = "999"
                Else
                    txtTienda.TexT = (usuarios(Bucle).Permisos(7) + 1)
                End If
            End If
            Desglosa_Permisos usuarios(Bucle).Permisos
        End If
    Next Bucle
End Sub
'//////
'//////
'//////
Public Sub CargaUsuarios()
    Dim Archivo As Integer
    Dim Buffer As String * 29
    Dim Contador As Integer
    Dim Nulos As String * 29
    Arbol.Nodes.Clear
    Contador = 0
    CantidadUsuarios = 0
    ReDim usuarios(1)
    'If Dir(MiruTa & "\pass.sec") <> "" Then
    If Dir(App.Path & "\pass.sec") <> "" Then
        Archivo = FreeFile()
        'Open MiruTa & "\pass.sec" For Binary As Archivo
        Open App.Path & "\pass.sec" For Binary As Archivo
        Do Until EOF(Archivo)
            Get #Archivo, , Buffer
            If Buffer <> Nulos Then Buffer = DesEncripta(Buffer)
            If Trim(Buffer) <> "" And Buffer <> Nulos Then
                Contador = Contador + 1
                CantidadUsuarios = CantidadUsuarios + 1
                ReDim Preserve usuarios(Contador + 1)
                usuarios(Contador).Nombre = left(Buffer, 8)
                usuarios(Contador).Password = Mid(Buffer, 9, 8)
                usuarios(Contador).Asociado = Val(Mid(Buffer, 25, 3))
                CargaPermisos usuarios(Contador), Buffer
                If Trim(usuarios(Contador).Nombre) = "admin" Then
                    Arbol.Nodes.Add , tvwChild, Trim(usuarios(Contador).Nombre), Trim(usuarios(Contador).Nombre), "admin"
                Else
                    If Trim(usuarios(Contador).Nombre) = "cmdline" Then
                        Arbol.Nodes.Add , tvwChild, Trim(usuarios(Contador).Nombre), Trim(usuarios(Contador).Nombre), "cmdline"
                    Else
                        Arbol.Nodes.Add , tvwChild, Trim(usuarios(Contador).Nombre), Trim(usuarios(Contador).Nombre), "normal"
                    End If
                End If
            End If
        Loop
        Close #Archivo
    End If
End Sub

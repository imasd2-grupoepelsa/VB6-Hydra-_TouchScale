VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmEpelsa 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   7830
   ClientLeft      =   1620
   ClientTop       =   2685
   ClientWidth     =   15375
   FillStyle       =   0  'Solid
   ForeColor       =   &H0000FFFF&
   Icon            =   "epelsa.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   Picture         =   "epelsa.frx":030A
   ScaleHeight     =   7830
   ScaleWidth      =   15375
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   9
      Left            =   480
      Picture         =   "epelsa.frx":49257
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   1932
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Entrada Modo Inventario."
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
      Index           =   6
      Left            =   4080
      Picture         =   "epelsa.frx":4CE38
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
      Width           =   1572
   End
   Begin VB.CommandButton CommandMarcar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marcar Artículos Para Envio."
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
      Left            =   4080
      Picture         =   "epelsa.frx":50A19
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Width           =   1572
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir >>"
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
      Left            =   7680
      Picture         =   "epelsa.frx":545FA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton CmdTotales 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdTotales"
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
      Index           =   2
      Left            =   7680
      Picture         =   "epelsa.frx":581DB
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5520
      Width           =   2052
   End
   Begin VB.CommandButton CmdTotales 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdTotales"
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
      Index           =   1
      Left            =   7680
      Picture         =   "epelsa.frx":5BDBC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   2052
   End
   Begin VB.CommandButton CmdTotales 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdTotales"
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
      Index           =   0
      Left            =   7680
      Picture         =   "epelsa.frx":5F99D
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Width           =   2052
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C3"
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
      Index           =   5
      Left            =   4080
      Picture         =   "epelsa.frx":6357E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C2"
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
      Index           =   4
      Left            =   4080
      Picture         =   "epelsa.frx":6715F
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C1"
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
      Index           =   3
      Left            =   4080
      Picture         =   "epelsa.frx":6AD40
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1572
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdComunicaciones"
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
      Index           =   2
      Left            =   4080
      Picture         =   "epelsa.frx":6E921
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   1572
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdComunicaciones"
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
      Index           =   1
      Left            =   4080
      Picture         =   "epelsa.frx":72502
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Width           =   1572
   End
   Begin VB.CommandButton CmdComunicaciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CmdComunicaciones"
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
      Index           =   0
      Left            =   4080
      Picture         =   "epelsa.frx":760E3
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Width           =   1572
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   8
      Left            =   480
      Picture         =   "epelsa.frx":79CC4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   7
      Left            =   480
      Picture         =   "epelsa.frx":7D8A5
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   6
      Left            =   480
      Picture         =   "epelsa.frx":81486
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   5
      Left            =   480
      Picture         =   "epelsa.frx":85067
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   4
      Left            =   480
      Picture         =   "epelsa.frx":88C48
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   3
      Left            =   480
      Picture         =   "epelsa.frx":8C829
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   2
      Left            =   480
      Picture         =   "epelsa.frx":9040A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   1
      Left            =   480
      Picture         =   "epelsa.frx":93FEB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1932
   End
   Begin VB.CommandButton CmdTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
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
      Index           =   0
      Left            =   480
      MaskColor       =   &H0000C0C0&
      Picture         =   "epelsa.frx":97BCC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1932
   End
   Begin MSChart20Lib.MSChart gRafico 
      Height          =   3135
      Left            =   7440
      OleObjectBlob   =   "epelsa.frx":9B7AD
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4332
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   855
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label LblGA 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   4080
      TabIndex        =   5
      Top             =   3840
      Width           =   2892
   End
   Begin VB.Label LblSC10 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   4080
      TabIndex        =   4
      Top             =   3480
      Width           =   2892
   End
   Begin VB.Label LblTotales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label LblComunicaciones 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Top             =   2040
      Width           =   2892
   End
   Begin VB.Label LblTienda 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   2652
   End
   Begin VB.Menu mnuficheros 
      Caption         =   "1"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMostradoryEquipo 
         Caption         =   "2"
      End
      Begin VB.Menu mnuFamilias 
         Caption         =   "2_5"
      End
      Begin VB.Menu mnuVacunoGama 
         Caption         =   "6"
         Begin VB.Menu mnuPaises 
            Caption         =   "7"
         End
         Begin VB.Menu mnuFichaVacuno 
            Caption         =   "8"
         End
      End
      Begin VB.Menu mnu_iva 
         Caption         =   "9"
      End
      Begin VB.Menu mnuArticulos 
         Caption         =   "10"
      End
      Begin VB.Menu mnuTe 
         Caption         =   "Teclas Directas Equipos"
      End
      Begin VB.Menu mnuTeclas 
         Caption         =   "15"
      End
      Begin VB.Menu mnuCambioPrecio 
         Caption         =   "11"
      End
      Begin VB.Menu mnuVendedores 
         Caption         =   "12"
      End
      Begin VB.Menu mnuTrade 
         Caption         =   "Datos Genéricos Tique"
      End
      Begin VB.Menu mnuTextLibre 
         Caption         =   "Textos Libres"
      End
      Begin VB.Menu mnunada1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCbarras 
         Caption         =   "16"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu MnuNada556 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrarSesion 
         Caption         =   "20"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "21"
      End
   End
   Begin VB.Menu mnuBase_de_datos 
      Caption         =   "2"
      Begin VB.Menu mnutotalsec 
         Caption         =   "2"
      End
      Begin VB.Menu mnutotalarticulos 
         Caption         =   "3"
      End
      Begin VB.Menu mnutotalvendedores 
         Caption         =   "4"
      End
      Begin VB.Menu mnuInformeVen 
         Caption         =   "Informe total vendedores"
      End
      Begin VB.Menu mnuInformeDes 
         Caption         =   "Informe descuentos"
      End
      Begin VB.Menu MnuBaseVacuno 
         Caption         =   "5"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConsultaTiquets 
         Caption         =   "Consulta de Tiquets"
      End
      Begin VB.Menu mnuInformeHorario 
         Caption         =   "Informe total horario"
      End
      Begin VB.Menu mnuInforHistorico 
         Caption         =   "Informe historico"
      End
      Begin VB.Menu mnunadacomun 
         Caption         =   "-"
      End
      Begin VB.Menu mnumantener 
         Caption         =   "7"
      End
      Begin VB.Menu mnunada 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopiasSeguridad 
         Caption         =   "8"
      End
   End
   Begin VB.Menu mnuComunicaciones 
      Caption         =   "3"
      Begin VB.Menu mnuRecibeTiquet 
         Caption         =   "Recibir Tiquets"
      End
      Begin VB.Menu mnunada294 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrandes_totales 
         Caption         =   "2"
      End
      Begin VB.Menu mnuNadaComu 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEnvio_modificaciones 
         Caption         =   "3"
      End
      Begin VB.Menu mnubasemaestra 
         Caption         =   "4"
      End
      Begin VB.Menu mnu_envio_separado 
         Caption         =   "5"
         Begin VB.Menu mnuEnviarPaises 
            Caption         =   "5"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEnviarFichasVacuno 
            Caption         =   "6"
         End
         Begin VB.Menu mnu_Envio_IVA 
            Caption         =   "7"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_envio_articulos 
            Caption         =   "8"
         End
         Begin VB.Menu mnuEnvioCLientes 
            Caption         =   "Envío de Clientes"
         End
         Begin VB.Menu mnu_envio_codigobar 
            Caption         =   "11"
         End
         Begin VB.Menu mnu_envio_vendedor 
            Caption         =   "12"
         End
         Begin VB.Menu KillTec 
            Caption         =   "Borrado de Teclas"
         End
         Begin VB.Menu KILLART 
            Caption         =   "Borrado Artículos y Teclas."
         End
      End
      Begin VB.Menu mnuModSeparado 
         Caption         =   "14"
         Begin VB.Menu mnuModPais 
            Caption         =   "19"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuModFichas 
            Caption         =   "20"
         End
         Begin VB.Menu mnuModIVA 
            Caption         =   "21"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuModArt 
            Caption         =   "22"
         End
         Begin VB.Menu MnuModCli 
            Caption         =   "Mod. Cli."
         End
         Begin VB.Menu mnuModCodBar 
            Caption         =   "25"
         End
         Begin VB.Menu mnuModVend 
            Caption         =   "26"
         End
      End
      Begin VB.Menu mnuEstadoGA 
         Caption         =   "Touch Com. Veri."
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "4"
      Begin VB.Menu Importfichero 
         Caption         =   "2"
      End
      Begin VB.Menu mnuExportarFcihero 
         Caption         =   "3"
      End
      Begin VB.Menu mnunada1024 
         Caption         =   "-"
      End
      Begin VB.Menu globaldat1 
         Caption         =   "GLOBAL.DAT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGlobalDat 
         Caption         =   "Importar GLOBAL.DAT"
      End
      Begin VB.Menu mnuExportaGlobal 
         Caption         =   "Exportar GLOBAL.DAT"
      End
      Begin VB.Menu Exportartglobal 
         Caption         =   "Export.ART(Global.dat)(CHG.)"
      End
      Begin VB.Menu exportarttodos 
         Caption         =   "Export.ART(Global.dat)"
      End
      Begin VB.Menu ImportTeclasTXT 
         Caption         =   "Import. Teclas.txt"
      End
      Begin VB.Menu expmsb100 
         Caption         =   "6"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export..."
         Begin VB.Menu mnuhistorico 
            Caption         =   "7"
            Index           =   0
         End
         Begin VB.Menu mnuhistorico 
            Caption         =   "7-1"
            Index           =   1
         End
      End
      Begin VB.Menu mnunada1029 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSQL 
         Caption         =   "SQL"
      End
      Begin VB.Menu Mnunada456 
         Caption         =   "-"
      End
      Begin VB.Menu rechazos 
         Caption         =   "8"
      End
      Begin VB.Menu mnuListadoCambiosPrecio 
         Caption         =   "9"
      End
      Begin VB.Menu mnutemporales 
         Caption         =   "Reporte Cambios Temporales"
      End
      Begin VB.Menu mnnuada437 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorrado_totales 
         Caption         =   "12"
         Visible         =   0   'False
         Begin VB.Menu BorradoGeneral 
            Caption         =   "13"
            Index           =   0
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "14"
            Index           =   1
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "15"
            Index           =   2
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "16"
            Index           =   3
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "17"
            Index           =   4
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "18"
            Index           =   5
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "19"
            Index           =   7
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "20"
            Index           =   9
         End
         Begin VB.Menu BorradoGeneral 
            Caption         =   "21"
            Index           =   24
         End
         Begin VB.Menu mnunada458 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnugenerabasevacia 
         Caption         =   "21"
      End
      Begin VB.Menu mnuAbsorver 
         Caption         =   "Absorver Estructura de Mostradores y Equipos"
      End
   End
   Begin VB.Menu mnuopciones 
      Caption         =   "22"
      Begin VB.Menu mnuficherosexportacion 
         Caption         =   "23"
      End
      Begin VB.Menu configurator 
         Caption         =   "24"
      End
      Begin VB.Menu mnuParamInic 
         Caption         =   "25"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConfiguracion2 
         Caption         =   "26"
      End
      Begin VB.Menu mnuseuridad 
         Caption         =   "27"
      End
   End
   Begin VB.Menu mnucdmlibres 
      Caption         =   "1"
      Begin VB.Menu cmdlibre 
         Caption         =   "1"
         Index           =   0
      End
      Begin VB.Menu cmdlibre 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu cmdlibre 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu cmdlibre 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu cmdlibre 
         Caption         =   "5"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPlugin 
      Caption         =   "Plugin"
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   0
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   1
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   2
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   3
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   4
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   5
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   6
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   7
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   8
      End
      Begin VB.Menu mnuplg 
         Caption         =   "plg"
         Index           =   9
      End
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "1"
   End
End
Attribute VB_Name = "frmEpelsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************
'**************
'**************
'Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" _
'         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
'          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Const HWND_TOPMOST = -1
'Const SWP_NOMOVE = &H2
'Const SWP_NOSIZE = &H1
'**********************
'**********************
'**********************
Dim pathorigen
Dim pathdestino
Dim mostra(50)
Dim contasec
Dim myserver
Dim aviso As String
Dim ncont As Integer
Dim abier As Boolean
Dim Hiper As String
Dim dbase As dao.Database
Dim Maquina As dao.Recordset
Dim NDireccion As String
Private Type TipoImagen
    Height As Long
    Width As Long
End Type
Private Imagen As TipoImagen
Private QLResize As Boolean
#Const BALSAM = False

Private Sub CambiarIdioma()
    'mnuSubSec.Caption = CargaCadena(677)
    mnuFamilias.Caption = CargaCadena(678)
    mnuEnvioCLientes.Caption = CargaCadena(225) & " " & CargaCadena(1056)
    KillTec.Caption = CargaCadena(1454)
    KILLART.Caption = CargaCadena(1455)
    mnuConsultaTiquets.Caption = CargaCadena(112)
    mnuRecibeTiquet.Caption = CargaCadena(1030)
    'mnuTotalesGB.Caption = CargaCadena(1270)
    mnuGlobalDat.Caption = CargaCadena(1175)
    mnutemporales.Caption = CargaCadena(1176)
    CmdSalir.Caption = CargaCadena(7) & " >>"
    LblTienda.Caption = CargaCadena(46)
    LblComunicaciones.Caption = CargaCadena(4)
    LblTotales.Caption = CargaCadena(602)
    'mnuCambioDirecto.Caption = CargaCadena(916)
    mnuAbsorver.Caption = CargaCadena(1400)
    mnuClientes.Caption = CargaCadena(1056)
    CmdTienda(0).Caption = CargaCadena(580)
    CmdTienda(1).Caption = CargaCadena(678)
    CmdTienda(2).Caption = CargaCadena(13)
    CmdTienda(3).Caption = CargaCadena(12)
    CmdTienda(4).Caption = CargaCadena(81) 'CargaCadena(24)
    CmdTienda(5).Caption = CargaCadena(214)
    CmdTienda(6).Caption = CargaCadena(1056)
    CmdTienda(7).Caption = CargaCadena(23)
    
    CmdTienda(8).Caption = CargaCadena(1397)
    CmdTienda(9).Caption = CargaCadena(1398)
    
    CmdComunicaciones(0).Caption = CargaCadena(151)
    CmdComunicaciones(1).Caption = CargaCadena(1030) 'CargaCadena( 161)
    CmdComunicaciones(2).Caption = CargaCadena(204)
    CmdComunicaciones(3).Caption = CargaCadena(704)
    CmdTotales(0).Caption = CargaCadena(153)
    CmdTotales(1).Caption = CargaCadena(156)
    CmdTotales(2).Caption = CargaCadena(157)
    mnuficheros.Caption = CargaCadena(46)
    mnuBase_de_datos.Caption = CargaCadena(3)
    mnuComunicaciones.Caption = CargaCadena(4)
    mnuUtil.Caption = CargaCadena(5)
    mnuopciones.Caption = CargaCadena(334)
    mnucdmlibres.Caption = CargaCadena(605)
    mnuHelpAbout.Caption = CargaCadena(296)
    '
    mnuficheros.Caption = CargaCadena(46)  ' "&Tienda"
    mnuMostradoryEquipo.Caption = CargaCadena(580)  ' "Mostradores y Equipos"
    mnuVacunoGama.Caption = CargaCadena(676)  ' "Vacuno..."
    mnuPaises.Caption = CargaCadena(212)  ' "Países"
    mnuFichaVacuno.Caption = CargaCadena(619)  ' "Fichas de Vacuno"
    mnu_iva.Caption = CargaCadena(8)  ' "Iva"
    mnuArticulos.Caption = CargaCadena(12)  ' "Articulos"
    mnuCambioPrecio.Caption = CargaCadena(24)  ' "Cambios de Precio"
    mnuVendedores.Caption = CargaCadena(25)  ' "Vendedores"
    mnuTeclas.Caption = CargaCadena(1397) 'CargaCadena(22)  ' "Teclas de balanza" 'revisar texto
    mnuCbarras.Caption = CargaCadena(23)  ' "Codigos de barras"
    mnuCerrarSesion.Caption = CargaCadena(26)  ' "Cerrar Sesión"
    mnuSalir.Caption = CargaCadena(7)  ' "Salir"
    mnuBase_de_datos.Caption = CargaCadena(3)  ' "&Base de Datos"
    mnutotalsec.Caption = CargaCadena(153)  ' "Total por Sección"
    mnutotalarticulos.Caption = CargaCadena(156)  ' "Total por Articulos"
    mnutotalvendedores.Caption = CargaCadena(157)  ' "Total por Vendedores"
    mnuInformeVen.Caption = CargaCadena(1239) ' "Informe total por Vendedores"
    MnuBaseVacuno.Caption = CargaCadena(642)  ' "Total Fichas de Vacuno"
    'mnutotalgeneral.Caption = CargaCadena(115)  ' "Total Diario/Acumulado"
    mnumantener.Caption = CargaCadena(282)  ' "Mantenimiento"
    mnuCopiasSeguridad.Caption = CargaCadena(32)  ' "Copias de Seguridad"
    mnuComunicaciones.Caption = CargaCadena(4)  ' "&Comunicaciones"
    mnuGrandes_totales.Caption = CargaCadena(161)  ' "Grandes Totales"
    mnuEnvio_modificaciones.Caption = CargaCadena(203)  ' "Envio de Modificaciones"
    mnubasemaestra.Caption = CargaCadena(204)  ' "Base Maestra"
    mnu_envio_separado.Caption = CargaCadena(229)  ' "Envio datos separados"
    'mnuEnviarSecciones.Caption = CargaCadena(225) & CargaCadena(9)   ' "Envío Secciones"
    'mnu_enviarEquipos.Caption = CargaCadena(225) & CargaCadena(11)   ' "Envío Equipos"
    mnuEnviarPaises.Caption = CargaCadena(225) & CargaCadena(212)   ' "Envío Países"
    mnuEnviarFichasVacuno.Caption = CargaCadena(225) & CargaCadena(619)   ' "Envío Fichas de Vacuno"
    mnu_Envio_IVA.Caption = CargaCadena(225) & CargaCadena(8)   ' "Envío Tipos de IVA"
    mnu_envio_articulos.Caption = CargaCadena(230)  ' "Envio Artículos"
    'mnu_envio_publicidad.Caption = CargaCadena(231)  ' "Envio de Publicidad"
    'mnu_envio_cabley.Caption = CargaCadena(232)  ' "Envio Cabeceras y Leyendas"
    mnu_envio_codigobar.Caption = CargaCadena(233)  ' "Envio Código de barras"
    mnu_envio_vendedor.Caption = CargaCadena(225) & CargaCadena(13)   ' "Envio de Vendedores"
    'mnu_envio_teclas_rapidas.Caption = CargaCadena(234)  ' "Envio de Teclas rápidas"
    mnuModSeparado.Caption = CargaCadena(809)  ' "Envío de Modificaciones por separado"
    'mnuModSec.Caption = CargaCadena(679) & CargaCadena(9)   ' "Modificaciones de Secciones"
    'mnuModSub.Caption = CargaCadena(679) & CargaCadena(677)   ' "Modificaciones de Subsecciones"
'    mnuModFam.Caption = CargaCadena(679) & CargaCadena(678)   ' "Modificaciones de Familias"
'    mnuModEqp.Caption = CargaCadena(679) & CargaCadena(11)   ' "Modificaciones de Equipos"
    mnuModPais.Caption = CargaCadena(679) & CargaCadena(212)   ' "Modificaciones de Países"
    mnuModFichas.Caption = CargaCadena(679) & CargaCadena(619)   ' "Modificaciones de Fichas de Vacuno"
    mnuModIVA.Caption = CargaCadena(679) & CargaCadena(8)   ' "Modificaciones de tipos de IVA"
    mnuModArt.Caption = CargaCadena(679) & CargaCadena(12)   ' "Modificaciones de Artículos"
'    mnuModPubli.Caption = CargaCadena(679) & CargaCadena(14)   ' "Modificaciones de Publicidad/Ofertas"
    'menuModCabley.Caption = CargaCadena(679) & CargaCadena(21)   ' "Modificaciones de Cabeceras/Leyendas"
    mnuModCodBar.Caption = CargaCadena(679) & CargaCadena(23)   ' "Modificaciones de Códigos de Barras"
    mnuModVend.Caption = CargaCadena(679) & CargaCadena(13)   ' "Modificaciones de Vendedores"
    mnuUtil.Caption = CargaCadena(5)  ' "&Utilidades"
    Importfichero.Caption = CargaCadena(685)  ' "Importar desde Fichero..."
    mnuExportarFcihero.Caption = CargaCadena(686)  ' "Exportar a Fichero..."
    mnuExportaGlobal.Caption = CargaCadena(692) & " Global.dat"
    mnuhistorico(0).Caption = CargaCadena(690)  ' "Histórico desde ficheros"
    mnuhistorico(1).Caption = CargaCadena(691)  ' "Historico desde base de datos"
    rechazos.Caption = CargaCadena(688) & " " & CargaCadena(689)   ' "Listado rechazos"
    mnuListadoCambiosPrecio.Caption = CargaCadena(688) & " " & CargaCadena(24)   ' "Listado Cambios de Precio"
'    mnuconectar_ventas.Caption = CargaCadena(247)  ' "Conectar Ventas"
'    mnudesconectar_ventas.Caption = CargaCadena(248)  ' "Desconectar Ventas"
    mnuBorrado_totales.Caption = CargaCadena(237)  ' "Borrados"
    BorradoGeneral(0).Caption = CargaCadena(238)  ' "Borrar Todo"
    BorradoGeneral(1).Caption = CargaCadena(38) & " " & CargaCadena(9)   ' "Borrar Mostradores"
    BorradoGeneral(2).Caption = CargaCadena(38) & " " & CargaCadena(12)   ' "Borrar Articulos"
    BorradoGeneral(3).Caption = CargaCadena(38) & " " & CargaCadena(13)   ' "Borrar Vendedores"
    BorradoGeneral(4).Caption = CargaCadena(38) & " " & CargaCadena(11)   ' "SC10 - Borrar Equipos"
    BorradoGeneral(5).Caption = CargaCadena(242)  ' "Borrar Publicidad + Cabeceras y Leyendas"
    BorradoGeneral(7).Caption = CargaCadena(243)  ' "Borrar Tiquets y Líneas de Tiquets"
    BorradoGeneral(9).Caption = CargaCadena(245)  ' "Borrar Códigos de Barras"
    BorradoGeneral(24).Caption = CargaCadena(246)  ' "Borrar Teclas"
    mnugenerabasevacia.Caption = CargaCadena(249)  ' "&Regenerar base de datos vacia"
    mnuopciones.Caption = CargaCadena(334)  ' "Opciones"
    mnuficherosexportacion.Caption = CargaCadena(680)  ' "Configurar ficheros de exportación"
    configurator.Caption = CargaCadena(681)  ' "Configurar ficheros de importación"
    mnuParamInic.Caption = CargaCadena(682)  ' "Configurar parámetros de Inicio"
    mnuConfiguracion2.Caption = CargaCadena(683)  ' "Configuración de General de HydraTouch"
    mnuseuridad.Caption = CargaCadena(684)  ' "Configuración de Seguridad"
    mnucdmlibres.Caption = CargaCadena(605)  ' "Comandos Libres"
    mnuHelpAbout.Caption = CargaCadena(296)  ' "Acerca de ..."
    Me.mnuInformeDes.Caption = CargaCadena(1281) ' "Acerca de ..."
    mnuTe.Caption = CargaCadena(1398)
    mnuTrade.Caption = CargaCadena(1399)
    mnuTextLibre.Caption = CargaCadena(215)
    mnuInformeHorario.Caption = CargaCadena(1401)
    mnuInforHistorico.Caption = CargaCadena(1402)
End Sub

Private Sub BorradoGeneral_Click(Index As Integer)
        Dim Resp1
        Dim Resp2
        If HaySeguridad Then
            If ConsultarPermiso(19) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        End If
        Resp1 = MsgBox(CargaCadena(694), vbYesNoCancel + vbQuestion)
        If Resp1 = vbCancel Then Exit Sub
        If Resp1 = vbYes Then
            Resp2 = MsgBox(CargaCadena(695), vbOKCancel + vbExclamation)
            If Resp2 = vbOK Then AN_BorrarBases Index, True
        Else
            Resp2 = MsgBox(CargaCadena(696), vbOKCancel + vbExclamation)
            If Resp2 = vbOK Then AN_BorrarBases Index, False
        End If
End Sub
'******
'******
'******
Private Sub CmdComunicaciones_Click(Index As Integer)
Dim Arch As Integer
Dim Resp As Long

    lCogeTiquet = False
    Select Case Index
        Case 0
            If HaySeguridad Then
                If ConsultarPermiso(13) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            mnuEnvio_modificaciones_Click
        Case 1
            If HaySeguridad Then
                If ConsultarPermiso(14) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            If Not Supeco Then
                mnuRecibeTiquet_Click
            Else
                Resp = MsgBox(CargaCadena(1013), vbYesNo + vbDefaultButton2 + vbExclamation)
                If Resp = vbYes Then mnuGrandes_totales_Click
            End If
        Case 2
            If HaySeguridad Then
                If ConsultarPermiso(12) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            mnuBaseMaestra_Click
        Case 3
            If PausaTiquets = False Then
                PausaTiquets = True
                CmdComunicaciones(3).Enabled = False
                CmdComunicaciones(3).Caption = CargaCadena(273)
                canCelar = 1
                On Error Resume Next
                Arch = FreeFile()
                Open App.Path & "\cancel.tqt" For Output As #Arch
                Close #Arch
                On Error GoTo 0
                'c2f 1.7.2
                frmEpelsa.HabilitarMenues
                '''''''''''''''''''''''''
            Else
                PausaTiquets = False
                BloquearMenues
                CmdComunicaciones(3).Caption = CargaCadena(704)
                'c2f 1.7.2
                CmdComunicaciones(3).Enabled = True
                frmControl.timetqt.Enabled = True
                '''''''''''''''''''''''''''''''''
            End If
        Case 4
            Select Case TqStatus
                Case 1
                    frmControl.ControlImport1.CancelarImportacion
                Case 2
                    canCelar = 1
            End Select
        Case 5
            frmEpelsa.Enabled = False

            CmdComunicaciones(5).Visible = False
            frmEpelsa.Enabled = True
            
        Case 6 'digitonuba entrada modo inventario
            
    End Select
    
    'c2f/caspiunza **
    On Error Resume Next
    If Trim(Command) = "" Then
        Unload frmMonitor
    End If
    On Error GoTo 0
    '''''''''''''''
    
    lCogeTiquet = True
End Sub
Public Sub BloquearMenues()
    If StoreFlow Then Exit Sub
    mnuBorrado_totales.Enabled = False
    mnuConfiguracion2.Enabled = False
    mnuComunicaciones.Enabled = False
    'mnutotalgeneral.Enabled = False
    mnuCopiasSeguridad.Enabled = False
    frmEpelsa.Importfichero.Enabled = False
    frmEpelsa.CmdComunicaciones(0).Enabled = False
    frmEpelsa.CmdComunicaciones(1).Enabled = False
    frmEpelsa.CmdComunicaciones(2).Enabled = False
    'frmEpelsa.mnuCambioDirecto.Enabled = False
    mnuPlugin.Enabled = False
    CmdTienda(4).Enabled = False
    
End Sub
Public Sub HabilitarMenues()
    If StoreFlow Then Exit Sub
    mnuBorrado_totales.Enabled = True
    mnuConfiguracion2.Enabled = True
    mnuComunicaciones.Enabled = True
    'mnutotalgeneral.Enabled = True
    mnuCopiasSeguridad.Enabled = True
    frmEpelsa.Importfichero.Enabled = True
    frmEpelsa.CmdComunicaciones(0).Enabled = True
    frmEpelsa.CmdComunicaciones(1).Enabled = True
    frmEpelsa.CmdComunicaciones(2).Enabled = True
    frmEpelsa.CmdComunicaciones(3).Enabled = True
    'frmEpelsa.mnuCambioDirecto.Enabled = True
    CmdTienda(4).Enabled = True
    CmdTienda(5).Enabled = True
    mnuPlugin.Enabled = True
    If PausaTiquets = True Then frmEpelsa.CmdComunicaciones(3).Caption = CargaCadena(705)
End Sub

Private Sub cmdlibre_Click(Index As Integer)
    Dim MyResul As Long
    Dim MyProc As Long
    If HaySeguridad Then
        If ConsultarPermiso(28 + Index) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    On Error GoTo errorcomando
    'Shell EjecutarLibre(Index), vbMaximizedFocus -> cas.v143
    'Shell EjecutarLibre(Index), vbNormalFocus
    PlugActivo = True
    MyResul = STILL_ACTIVE
    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(EjecutarLibre(Index), vbNormalFocus))
    Do While MyResul = STILL_ACTIVE
        GetExitCodeProcess MyProc, MyResul
        Do_Events
    Loop
    PlugActivo = False
Exit Sub
errorcomando:
    MsgBox CargaCadena(383), vbCritical
End Sub

Private Sub CmdSalir_Click()
    salir_programa
End Sub

Private Sub CmdTienda_Click(Index As Integer)
    Select Case Index
        Case 0
            If HaySeguridad Then
                If ConsultarPermiso(0) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmSeccionyMaquina
        Case 1
            If HaySeguridad Then
                If ConsultarPermiso(0) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmFamilias ' 'FrmFamilias
        Case 2
            If HaySeguridad Then
                If ConsultarPermiso(9) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmVendedor2
        Case 3
            'artículos...
            If HaySeguridad Then
                If ConsultarPermiso(1) = False Then
                    If ConsultarPermiso(24) = False Then
                        MsgBox CargaCadena(693), vbCritical
                        Exit Sub
                    End If
                End If
            End If
            MostrarModal FrmArticulo2
            ''''''''''''''''
            'If HaySeguridad Then
            '    If ConsultarPermiso(2) = False Then
            '        MsgBox CargaCadena(693), vbCritical
            '        Exit Sub
            '    End If
            'End If
            'lCogeTiquet = False
            'MostrarModal FrmCambioRapido
        Case 4
            'cambios de precio...
            If HaySeguridad Then
                If ConsultarPermiso(2) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            'textart MostrarModal FrmCambiosPrecio
            MostrarModal FrmArtText
            '''''''''''''''''''''
        Case 5
            If HaySeguridad Then
                If ConsultarPermiso(5) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmIVA
        
        '    If HaySeguridad Then
        '        If ConsultarPermiso(6) = False Then
        '            MsgBox CargaCadena(693), vbCritical
        '            Exit Sub
        '        End If
        '    End If
        '    MostrarModal FrmCabley2
        Case 6
            If HaySeguridad Then
                If ConsultarPermiso(6) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmClientes
        '    If HaySeguridad Then
        '        If ConsultarPermiso(8) = False Then
        '            MsgBox CargaCadena(693), vbCritical
        '            Exit Sub
        '        End If
        '    End If
        '    MostrarModal FrmPublicidad2
        Case 7
            'códigos de barras...
            If HaySeguridad Then
                If ConsultarPermiso(7) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
#If BALSAM = True Then
            Dim message, title, defaultValue As String
            Dim myValue As String
            ' Set prompt.
            message = "Selecionar:" & vbCrLf & " (1) Ean13" & vbCrLf & " (2) Ean14" & vbCrLf & " (3) Ean128" & vbCrLf & " (4) Code128" & vbCrLf & " (5) GS1."
            ' Set title.
            title = "Tipo Código de Barras"
            defaultValue = "1"   ' Set default value.
            
            ' Display message, title, and default value.
            myValue = InputBox(message, title, defaultValue)
            ' If user has clicked Cancel, set myValue to defaultValue
            If myValue = "" Then myValue = defaultValue
            
            Select Case myValue
            Case "1"
            MostrarModal FrmCodBar2
            Case "2"
            MostrarModal FrmCodBar14
            Case "3"
            MostrarModal FrmEan128
            Case "4"
            MostrarModal FrmEanCod
            Case "5"
            MostrarModal FrmEanGs1
            
            End Select
            
#Else
            MostrarModal FrmCodBar2
            '''''''''''''''''''''
#End If
        Case 8
            If HaySeguridad Then
                If ConsultarPermiso(10) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            If Dir(App.Path & "\test.tst") <> "" Then
                MostrarModal FrmOrdTec
            Else
                MostrarModal FrmTeclas2T
            End If
        Case 9
            If HaySeguridad Then
                If ConsultarPermiso(10) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            If Dir(App.Path & "\test.tst") <> "" Then
                MsgBox "Se ha seleccionado un modo de trabajo que no permite esta opción...", vbExclamation
            Else
                MostrarModal FrmTeclas2TE
            End If
        Case 10
            If HaySeguridad Then
                If ConsultarPermiso(9) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal FrmDatosTique
            
    End Select
End Sub


Private Sub CmdTotales_Click(Index As Integer)
    
'prueba leclerc quitar
'Dim n As Integer
'n = fVerLeclerc()
''''''''''''''''''''''
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    Select Case Index
        Case 0
            mnutotalsec_Click
        Case 1
            mnutotalarticulos_Click
        Case 2
            mnutotalvendedores_Click
    End Select
    lCogeTiquet = True
End Sub




Public Sub configurator_Click()
    If HaySeguridad Then
        If ConsultarPermiso(15) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmControl.ini_import
    frmControl.ControlImport1.FamFijas = FamiliasFijas
    frmControl.ControlImport1.SubFijas = SubSeccionesFijas
    frmControl.ControlImport1.PathFile = Miruta
    frmControl.ControlImport1.PathNTQ = Miruta
    frmControl.ControlImport1.PathBase = Base_General
    frmControl.ControlImport1.configurar
    Unload frmControl
End Sub


Private Sub expmsb100_Click()
    If HaySeguridad Then
        If ConsultarPermiso(21) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    ExportaTotalesDat False
    frmEpelsa.Enabled = True
End Sub

Private Sub Exportartglobal_Click()
    Dim dB1 As dao.Database
    Set dB1 = OpenDatabase(Base_General)
    Call Exporta_GlobalDat_Art(dB1)
    dB1.Close
End Sub

Private Sub exportarttodos_Click()
    Dim dB1 As dao.Database
    Set dB1 = OpenDatabase(Base_General)
    Call Exporta_GlobalDat_Art_todos(dB1)
    dB1.Close

End Sub

Private Sub Form_Activate()
Dim a
a = a
End Sub

Private Sub Form_Click()
Dim a
a = a
End Sub

Private Sub Form_DblClick()
    If Iconificar Then
        frmEpelsa.Visible = False
    End If
End Sub
Public Sub mnuTotalFichaVAcuno_Click()
    frmMonitor.List1.Clear
    lCogeTiquet = False
    'MUE_LanzaDatosComunes 5, False
    lCogeTiquet = True
End Sub


Private Sub ImportTeclasTXT_Click()
Dim nf As Integer
Dim sS As String
    If Dir(App.Path & "\tecimp.cfg") <> "" Then
        On Error GoTo finTECIMP
        nf = FreeFile()
        Open App.Path & "\tecimp.cfg" For Input As #nf
        Line Input #nf, sS
        sS = Trim(sS)
        Close #nf
    Else
        sS = App.Path & "\teclas.txt"
    End If
    If Dir(sS) <> "" Then
        Call ImportTeclas(sS)
    Else
        MsgBox "No existe fichero:" & sS, vbExclamation
    End If
finTECIMP:
If Err.Number <> 0 Then
    MsgBox "tecimp.cfg Erroneo.", vbCritical
End If
On Error GoTo 0
End Sub

Private Sub KILLART_Click()
    Call KillARTICLES
End Sub

Private Sub KillTec_Click()
    Call KillTECLAS
End Sub

Private Sub mnuEstadoGA_Click()
    Do_EstadoGA
End Sub

Public Sub mnuTeclas_Click()
    'CmdTienda_Click (8)
    If HaySeguridad Then
        If ConsultarPermiso(10) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    If Dir(App.Path & "\test.tst") <> "" Then
        MostrarModal FrmOrdTec
    Else
        MostrarModal FrmTeclas2T
    End If
    'MostrarModal FrmTeclas2T
End Sub
Public Sub mnuTE_Click()
    'CmdTienda_Click (8)
    If HaySeguridad Then
        If ConsultarPermiso(10) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    If Dir(App.Path & "\test.tst") <> "" Then
        MsgBox "Se ha seleccionado un modo de trabajo que no permite esta opción...", vbExclamation
    Else
        MostrarModal FrmTeclas2TE
    End If
End Sub
Private Sub Form_GotFocus()
   lCogeTiquet = True
End Sub

Public Sub Activar_Menues()
Dim bucle As Integer

CmdTienda(0).Visible = True
CmdTienda(1).Visible = True
CmdTienda(2).Visible = True
CmdTienda(3).Visible = True
CmdTienda(4).Visible = True
CmdTienda(8).Visible = True
CmdTienda(7).Visible = True
CmdTienda(6).Visible = True
CmdTienda(5).Visible = True
mnuVendedores.Visible = True
mnuTrade.Visible = True
'mnuCambioDirecto.Visible = True
mnuConsultaTiquets.Visible = True
'mnuEnviarSecciones.Visible = True
'mnu_envio_publicidad.Visible = True
mnu_envio_vendedor.Visible = True
'mnu_envio_teclas_rapidas.Visible = True
mnuRecibeTiquet.Visible = True
CmdComunicaciones(1).Visible = True
mnuModVend.Visible = True
mnuTeclas.Visible = True
'*********************************************

'Sam
#If BALSAM = True Then
    mnuTextLibre.Visible = True
#Else
    mnuTextLibre.Visible = False
#End If

If HaySeguridad Then
    frmEpelsa.mnuCerrarSesion.Visible = True
Else
    frmEpelsa.mnuCerrarSesion.Visible = False
End If
If ComandoLibre(0) <> "" Or ComandoLibre(1) <> "" Or ComandoLibre(2) <> "" _
Or ComandoLibre(3) <> "" Or ComandoLibre(4) <> "" Then
    frmEpelsa.mnucdmlibres.Visible = True
Else
    frmEpelsa.mnucdmlibres.Visible = False
End If
For bucle = 0 To 4
    If ComandoLibre(bucle) <> "" Then
        frmEpelsa.cmdlibre(bucle).Visible = True
        frmEpelsa.cmdlibre(bucle).Caption = ComandoLibre(bucle)
    Else
        If frmEpelsa.mnucdmlibres.Visible = True Then
            frmEpelsa.cmdlibre(bucle).Visible = False
        End If
    End If
Next bucle
'************
' /comandos libres
'************
'frmEpelsa.mnuExportarFcihero.Visible = True
If lBoka = False And lFornes = False Then
    frmEpelsa.mnuExportarFcihero.Visible = True
Else
    frmEpelsa.mnuExportarFcihero.Visible = False
End If
frmEpelsa.expmsb100.Caption = CargaCadena(706) & " TOTALES.DAT"

If (descAuto And Right(descautotime, 3) = "sec") And (Not StoreFlow) Then
    mnuBorrado_totales.Enabled = False
    CmdComunicaciones(3).Visible = True
    mnuConfiguracion2.Enabled = False
Else
    mnuBorrado_totales.Enabled = True
    CmdComunicaciones(3).Visible = False
    mnuConfiguracion2.Enabled = True
End If


    frmEpelsa.mnu_Envio_IVA.Visible = True
    frmEpelsa.mnuModIVA.Visible = True

    mnuClientes.Visible = True
    mnuEnvioCLientes.Visible = True
    frmEpelsa.mnuVacunoGama.Visible = True
    frmEpelsa.MnuBaseVacuno.Visible = True
    frmEpelsa.mnuEnviarFichasVacuno.Visible = True
    frmEpelsa.mnuEnviarPaises.Visible = True
    frmEpelsa.mnuModPais.Visible = True
    frmEpelsa.mnuModFichas.Visible = True
If (Not UsaGamaAlta And Not usaSC10) Or (Not UsaGamaAlta And (usaSC10 And TipoC9)) Then
    'frmEpelsa.mnuconectar_ventas.Visible = False
    'frmEpelsa.mnudesconectar_ventas.Visible = False
    frmEpelsa.BorradoGeneral(0).Visible = False
    frmEpelsa.BorradoGeneral(1).Visible = False
    frmEpelsa.BorradoGeneral(2).Visible = False
    frmEpelsa.BorradoGeneral(3).Visible = False
    frmEpelsa.BorradoGeneral(4).Visible = False
    frmEpelsa.BorradoGeneral(5).Visible = False
    frmEpelsa.BorradoGeneral(7).Visible = False
    frmEpelsa.BorradoGeneral(9).Visible = False
    frmEpelsa.BorradoGeneral(24).Visible = False
Else
    frmEpelsa.BorradoGeneral(0).Visible = True
    frmEpelsa.BorradoGeneral(1).Visible = True
    frmEpelsa.BorradoGeneral(2).Visible = True
    frmEpelsa.BorradoGeneral(3).Visible = True
    frmEpelsa.BorradoGeneral(4).Visible = True
    frmEpelsa.BorradoGeneral(5).Visible = True
    frmEpelsa.BorradoGeneral(7).Visible = True
    frmEpelsa.BorradoGeneral(9).Visible = True
    frmEpelsa.BorradoGeneral(24).Visible = True
End If

mnuAbsorver.Visible = False
If UsaGamaAlta Or (usaSC10 And (Not TipoC9) And SC10Rs232) Then
    mnuAbsorver.Visible = True
End If
'If (Not UsaECO2) And ((Not UsaECO1) Or (UsaECO1 And TipoV8)) And (Not UsaEcoPrint) And usaSC10 And (Not UsaGamaAlta) And (Not TipoC9) And (Not SC10Rs232) Then
'    frmEpelsa.mnuConsultaArt.Visible = False
'Else
'
'    frmEpelsa.mnuConsultaArt.Visible = True'
'
'End If
If usaSC10 And (Not UsaGamaAlta) And TipoC9 Then
    frmEpelsa.CmdComunicaciones(1).Visible = False
    frmEpelsa.mnuRecibeTiquet.Visible = False
Else
    If usaSC10 Or UsaGamaAlta Then
        frmEpelsa.CmdComunicaciones(1).Visible = True
        frmEpelsa.mnuRecibeTiquet.Visible = True
    End If
End If
If HayMulti Then
    CmdSalir.Visible = True
    configurator.Visible = False
    mnuficherosexportacion.Visible = False
    mnuParamInic.Visible = False
    mnuseuridad.Visible = False
    mnuCopiasSeguridad.Visible = False
Else
    mnuParamInic.Visible = True
    mnuParamInic.Enabled = True
End If
'If (Not UsaECO1) And (Not UsaECO2) And (Not UsaEcoPrint) Then
'    mnuTotalesGB.Visible = False
'Else
    'mnuTotalesGB.Visible = True
'End If
mnuFamilias.Visible = False
'mnuSubSec.Visible = False
'If FamiliasFijas Then
mnuFamilias.Visible = True
'If SubSeccionesFijas Then mnuSubSec.Visible = True
Form_Resize
End Sub

Public Sub Form_Load()
    Dim b As Integer
    'Me.Icon = Form2.Icon
    '*******************
    ' logo
    '*******************
    'If Dir(App.Path & "\etc\ekobal.dat.fia") <> "" Then
    '    frmEpelsa.WindowState = 2
    '    frmEpelsa.BorderStyle = 0
    'End If
    
    CommandMarcar.Visible = False
    CommandMarcar.Enabled = False
    CmdComunicaciones(6).Visible = False
    CmdComunicaciones(6).Enabled = False
    'If Dir(App.Path & "\digitonuba.txt") <> "" Then
    '    CmdComunicaciones(6).Visible = True
    '    CmdComunicaciones(6).Enabled = True
    'End If
    Carga_Logo
    If NombreTienda <> "" Then
        frmEpelsa.Label2.ForeColor = vbWhite
        frmEpelsa.Label2.Caption = TiendaActual & "-" & NombreTienda
    End If
    CambiarIdioma
    Imagen.Width = 10800
    Imagen.Height = 5940
    On Error Resume Next
    QLResize = True
    Me.Width = GetSetting("hydra", "interface", "principal_width")
    Me.Height = GetSetting("hydra", "interface", "principal_height")
    QLResize = False
    If GetSetting("hydra", "interface", "principal_max") = "1" Then
        Me.WindowState = vbMaximized
    End If
    If Err.Number <> 0 Then
        If Screen.Width <= 9600 Then
            Width = Screen.Width * 0.9
            Height = Screen.Height * 0.9
        Else
            Width = Screen.Width * 0.85
            Height = Screen.Height * 0.8
        End If
    End If
    On Error GoTo 0
    '
    Activar_Menues
    '
    left = (Screen.Width - Width) / 2
    top = (Screen.Height - Height) / 2
    '************************
    ' botón de parar recogida
    '************************
    
    If descAuto = False Or Right(descautotime, 3) <> "sec" Or StoreFlow Then
        CmdComunicaciones(3).Visible = False
    Else
        mnuComunicaciones.Enabled = False
    End If


    frmEpelsa.Caption = "HydraTouchLite " & miver

    mnuPlugin.Visible = False

End Sub



Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Buf1 As String
    Dim Buf2 As String
    Dim bucle As Integer
    If data.GetFormat(vbCFFiles) Then
        If HaySeguridad Then
            If ConsultarPermiso(20) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        End If
        Buf1 = data.Files(1)
        If Dir(Buf1) <> "" Then
            For bucle = Len(Buf1) To 1 Step -1
                If Mid(Buf1, bucle, 1) <> "\" Then
                    Buf2 = Mid(Buf1, bucle, 1) & Buf2
                Else
                    Exit For
                End If
            Next bucle
            If Trim(Buf1) <> Trim(Miruta & "\" & Buf2) Then FileCopy Buf1, Miruta & "\" & Buf2
            'If Sabeco And UCase(Buf2) = "_ARTICUS.DAT" Then Buf2 = "sabeco.dat"
            If Dir(Miruta & "\etc\" & Buf2 & ".fia") <> "" Then
                frmControl.ini_import
                lCogeTiquet = False
                frmEpelsa.Enabled = False
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.FactorEuro = 1
                frmControl.ControlImport1.NumeroDecimales = decimales
                frmControl.ControlImport1.Euros = True
                If UsaEuro Then
                    frmControl.ControlImport1.Euros = True
                Else
                    frmControl.ControlImport1.Euros = False
                End If
                frmControl.ControlImport1.Fichero = FicheroImportacion
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                '****************************
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.MuestraMensajes = True
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.ActivarLog = LogArticulos
                ' /especial
                '*****************
                frmControl.ControlImport1.Importar_Background Buf2
                Corregir_Nulos
                frmEpelsa.Enabled = True
                
            Else
                MsgBox CargaCadena(1031), vbCritical
            End If
        End If
    End If
    If Dir(App.Path & "\asign.tst") <> "" Then
        Call AsignTec
    End If
End Sub

Private Sub Form_Resize()
    Dim bucle As Integer
    
    'If Dir(App.Path & "\etc\ekobal.dat.fia") <> "" Then
    '    Me.WindowState = vbMaximized
    'End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If EstiloInterfaz = 1 Then Exit Sub
    If Me.Width < (0.74 * Screen.Width) Then Me.Width = (0.74 * Screen.Width)
    If Me.Height < (0.7 * Screen.Height) Then Me.Height = (0.7 * Screen.Height)
    
    Imagen.Width = frmEpelsa.Width
    Imagen.Height = frmEpelsa.Height * 5940 / 6996
    gRafico.top = (Imagen.Height * 50) / 5940
    gRafico.Height = Imagen.Height * 1800 / 5940
    CmdTienda(0).top = (Imagen.Height * 2400) / 5940
    CmdTienda(0).Height = Imagen.Height * 372 / 5940
    CmdComunicaciones(0).top = (Imagen.Height * 2400) / 5940
    CmdComunicaciones(0).Height = Imagen.Height * 372 / 5940 - 5
    CmdComunicaciones(3).Height = Imagen.Height * 372 / 5940 - 5
    CmdComunicaciones(3).top = (Imagen.Height * 4200) / 5940
    CmdComunicaciones(4).Height = Imagen.Height * 372 / 5940 - 5
    CmdComunicaciones(5).Height = Imagen.Height * 372 / 5940 - 5
    CmdComunicaciones(4).top = (Imagen.Height * 4560) / 5940
    CmdComunicaciones(5).top = (Imagen.Height * 4920) / 5940
    'If Dir(App.Path & "\digitonuba.txt") <> "" Then
        CmdComunicaciones(6).top = (Imagen.Height * 5280) / 5940
        CmdComunicaciones(6).Height = Imagen.Height * 372 / 5940 - 5
    'End If
    If AhorraMas Then
        CommandMarcar.Height = Imagen.Height * 372 / 5940 - 5
        CommandMarcar.top = (Imagen.Height * 5280) / 5940
    End If
    
    LblGA.top = (Imagen.Height * 3480) / 5940 + 10
    LblGA.Height = Imagen.Height * 372 / 5940 - 2
    LblSC10.top = (Imagen.Height * 3840) / 5940 + 10
    LblSC10.Height = Imagen.Height * 372 / 5940 - 2
    'CmdTotales(0).Top = (Imagen.Height * 3960) / 5940
    'CmdTotales(0).Height = Imagen.Height * 372 / 5940
    CmdTotales(0).top = (Imagen.Height * 2400) / 5940
    CmdTotales(0).Height = Imagen.Height * 372 / 5940
    LblTienda.top = CmdTienda(0).top - CmdTienda(0).Height
    LblTotales.top = CmdTotales(0).top - CmdTotales(0).Height
    LblComunicaciones.top = CmdComunicaciones(0).top - CmdComunicaciones(0).Height
    'LblVersion.Top = (Imagen.Height * 240) / 5940
    'LblVersion.Height = Imagen.Height * 872 / 5940
    LblTienda.Height = CmdTienda(0).Height
    LblComunicaciones.Height = CmdComunicaciones(0).Height
    LblTotales.Height = CmdTotales(0).Height
    For bucle = 1 To 8
        If bucle = 1 Or bucle = 2 Then
            CmdTotales(bucle).top = CmdTotales(0).top + (CmdTotales(0).Height * bucle)
            CmdTotales(bucle).Height = CmdTotales(0).Height
        Else
            If bucle = 8 Then
                CmdSalir.top = CmdTotales(0).top + (CmdTotales(0).Height * bucle)
                CmdSalir.Height = CmdTotales(0).Height
            End If
        End If
    Next bucle
    For bucle = 1 To 2
        If CmdComunicaciones(1).Visible = True Then
            CmdComunicaciones(bucle).top = CmdComunicaciones(0).top + (CmdComunicaciones(0).Height * bucle)
            CmdComunicaciones(bucle).Height = CmdComunicaciones(0).Height
        Else
            CmdComunicaciones(bucle).top = CmdComunicaciones(0).top + (CmdComunicaciones(0).Height * (bucle - 1))
            CmdComunicaciones(bucle).Height = CmdComunicaciones(0).Height
        End If
    Next bucle
    For bucle = 1 To 9
        
        'If Bucle = 8 Then
        'Else
            CmdTienda(bucle).Height = CmdTienda(0).Height
            'If (CmdTienda(3).Visible = True) Or (Bucle <> 4 And Bucle <> 5) Then
                CmdTienda(bucle).top = CmdTienda(0).top + (CmdTienda(0).Height * bucle)
            'Else
            '    CmdTienda(Bucle).top = CmdTienda(0).top + (CmdTienda(0).Height * (Bucle - 1))
            'End If
        'End If
    Next bucle
    gRafico.left = Imagen.Width * 7000 / 10800
    gRafico.Width = Imagen.Width * 3600 / 10800
    CmdTienda(0).left = Imagen.Width * 240 / 10800
    CmdTienda(0).Width = Imagen.Width * 2892 / 10800
    CmdComunicaciones(0).left = Imagen.Width * 4000 / 10800
    CmdComunicaciones(0).Width = Imagen.Width * 2870 / 10800
    CmdComunicaciones(3).left = Imagen.Width * 4000 / 10800
    CmdComunicaciones(3).Width = Imagen.Width * 2870 / 10800
    CmdComunicaciones(5).left = Imagen.Width * 4000 / 10800
    CmdComunicaciones(5).Width = Imagen.Width * 2870 / 10800
    If AhorraMas Then
        CommandMarcar.left = Imagen.Width * 4000 / 10800
        CommandMarcar.Width = Imagen.Width * 2870 / 10800
    End If
    If Dir(App.Path & "\digitonuba.txt") <> "" Then
        CmdComunicaciones(6).left = Imagen.Width * 4000 / 10800
        CmdComunicaciones(6).Width = Imagen.Width * 2870 / 10800
    End If
    
    LblGA.left = Imagen.Width * 4000 / 10800
    LblSC10.left = Imagen.Width * 4000 / 10800
    LblGA.Width = Imagen.Width * 2892 / 10800
    LblSC10.Width = Imagen.Width * 2892 / 10800
    CmdTotales(0).left = Imagen.Width * 7560 / 10800
    CmdTotales(0).Width = Imagen.Width * 2892 / 10800
    LblTienda.left = CmdTienda(0).left
    LblTienda.Width = CmdTienda(0).Width
    LblComunicaciones.left = CmdComunicaciones(0).left
    LblComunicaciones.Width = CmdComunicaciones(0).Width
    LblTotales.left = CmdTotales(0).left
    LblTotales.Width = CmdTotales(0).Width
    'LblVersion.Left = Imagen.Width * 4560 / 10800
    'LblVersion.Width = Imagen.Width * 2052 / 10800
    For bucle = 1 To 2
        CmdTotales(bucle).left = CmdTotales(0).left
        CmdTotales(bucle).Width = CmdTotales(0).Width
    Next bucle
    CmdSalir.left = CmdTotales(0).left
    CmdSalir.Width = CmdTotales(0).Width
    For bucle = 1 To 4
        CmdComunicaciones(bucle).left = CmdComunicaciones(0).left
        CmdComunicaciones(bucle).Width = CmdComunicaciones(0).Width
    Next bucle
    For bucle = 1 To 9
        'If Bucle = 8 Then
        'Else
            CmdTienda(bucle).left = CmdTienda(0).left
            CmdTienda(bucle).Width = CmdTienda(0).Width
            'If Bucle = 8 Then  '2.0.9 Dir(App.Path & "\marte.new") <> ""  Then
            '    CmdTienda(Bucle).left = CmdTienda(4).left
            '    CmdTienda(Bucle).Width = CmdTienda(4).Width
            '    CmdTienda(Bucle).top = CmdTienda(4).top + CmdTienda(4).Height
            '
            'End If
        'End If
    Next bucle
    'If Dir(App.Path & "\nover1.txt") = "" Then
    '    grventas
    'End If
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1 '3 Color (Green-Blue-White)
    If Dir(App.Path & "\backgrounde.jpg") <> "" Then
        Me.Picture = LoadPicture(App.Path & "\backgrounde.jpg")
    End If
    If Dir(App.Path & "\etc\modifica.bal.fia") <> "" And Dir(App.Path & "\etc\articulo.bal.fia") <> "" Then
        If Dir(App.Path & "\hydra-background-1024x768_1_belros.jpg") <> "" Then
            Me.Picture = LoadPicture(App.Path & "\hydra-background-1024x768_1_belros.jpg")
        End If
    Else
        If Dir(App.Path & "\hydra-background-1024x768_1.jpg") <> "" Then
            Me.Picture = LoadPicture(App.Path & "\hydra-background-1024x768_1.jpg")
        End If
    End If
    
    frmEpelsa.Refresh
    On Error Resume Next
    If Not QLResize Then
        If Me.WindowState <> vbMaximized Then
            SaveSetting "hydra", "interface", "principal_width", CStr(Me.Width)
            SaveSetting "hydra", "interface", "principal_height", CStr(Me.Height)
        End If
        If Me.WindowState = vbMaximized Then
            SaveSetting "hydra", "interface", "principal_max", "1"
        Else
            SaveSetting "hydra", "interface", "principal_max", "0"
        End If
    End If
    
    'If Dir(App.Path & "\nover1.txt") <> "" Then
        gRafico.Visible = False
    'End If
    
    On Error GoTo 0
End Sub
Private Sub Carga_Logo()
    Dim MyImag As String
    MyImag = ""
    If Dir(App.Path & "\logo.gif") <> "" Then MyImag = "logo.gif"
    If Dir(App.Path & "\logo.jpg") <> "" Then MyImag = "logo.jpg"
    If Dir(App.Path & "\logo.bmp") <> "" Then MyImag = "logo.bmp"
    If MyImag <> "" Then
        On Error Resume Next
        Image1.Picture = LoadPicture(App.Path & "\" & MyImag)
        On Error GoTo 0
    End If
End Sub


Public Sub Form_Unload(Cancel As Integer)
    'If PlugActivo = True Then Exit Sub
    
    If cgdtiquet = True Then
        Cancel = 1
    Else
        If Dir(App.Path & "\etc\ekobal.dat.fia") = "" Then
            salir_programa
        Else
            Cancel = 1
        End If
    End If
End Sub



Private Sub Label1_Click()
    mnuArticulos_Click
End Sub



Private Sub Importfichero_Click()
Dim AL_Exito As Integer
Dim nLcLr As Integer
Dim sPath As String
Dim MyFich As Integer

    If HaySeguridad Then
        If ConsultarPermiso(20) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    
    If lFornes Then
        lCogeTiquet = False
        If Dir(App.Path & "\pluasc.dir") <> "" Then
            MyFich = FreeFile()
            Open App.Path & "\pluasc.dir" For Input As #MyFich
            Line Input #MyFich, sPath
            Close #MyFich
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        Else
            sPath = App.Path & "\"
            MyFich = FreeFile()
            Open App.Path & "\pluasc.dir" For Output As #MyFich
            Print #MyFich, sPath
            Close #MyFich
        End If
        If Dir(sPath & "Bal_Epelsa.txt") <> "" Then
            Load frmWAIT
            frmWAIT.Label1.Caption = "Tratar fichero: Bal_Epelsa.txt"
            frmWAIT.Label2.Caption = "Espere unos instantes ..."
            
            frmWAIT.Visible = True
            frmWAIT.Refresh
            Call transforma_Fornes
            frmWAIT.Visible = False
            Unload frmWAIT
            If Dir(App.Path & "\global.dat") <> "" Then
                Me.Enabled = False
                Importa_Global_Dat
                Me.Enabled = True
            End If
        Else
            MsgBox "No existe fichero: " & sPath & "Bal_Epelsa.txt", vbInformation
        End If
    Else
    
        If DebugActivo Then
            If id <> 0 Then
                CadenadeLog "Accesing to Import Menu..."
            Else
                CadenadeLog "Acceso a Formulario de Importación..."
            End If
        End If
        
        'leclerc....
        CadenadeLog "ImportFichero_click.:" & CStr(lClR) & " :" & sPathMaj
        If lClR And sPathMaj <> "" Then
            nLcLr = fVerLeclerc()
            CadenadeLog "Verificar si se ha Generado BAL.DAT"
            If Dir(App.Path & "\bal.dat") <> "" Then
                CadenadeLog "Generado bal.dat con longitud:" & CStr(FileLen(App.Path & "\bal.dat"))
            End If
        End If
        '...........
    
        frmControl.ini_import
        lCogeTiquet = False
        frmEpelsa.Enabled = False
        frmControl.ControlImport1.FamFijas = FamiliasFijas
        frmControl.ControlImport1.SubFijas = SubSeccionesFijas
        frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
        frmControl.ControlImport1.PathBase = Base_General
        frmControl.ControlImport1.FactorEuro = 1
        frmControl.ControlImport1.NumeroDecimales = decimales
        frmControl.ControlImport1.Euros = UsaEuro
        frmControl.ControlImport1.Fichero = FicheroImportacion
        frmControl.ControlImport1.FamVacuno = FamiliaEspecial
        frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
        frmControl.ControlImport1.PathNTQ = Miruta
        frmControl.ControlImport1.PathFile = Miruta
        frmControl.ControlImport1.MuestraMensajes = True
        frmControl.ControlImport1.MostrarFormulario = True
        frmControl.ControlImport1.ActivarLog = LogArticulos
        frmControl.ControlImport1.activarC9COD = var.C9Cod
        '**********
        ' /especial
        '**********
        frmControl.ControlImport1.importar
    
        On Error GoTo errorexito
        Sleep (200)
        Do While frmControl.ControlImport1.Exito = 4
            Sleep (10)
            Do_Events
errorexito:
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Do
            End If
            
        Loop
        Corregir_Nulos
        On Error GoTo 0
    'ERROREXITO:
    '    If Err.Number <> 0 Then
    '        If Err.Number <> 0 Then
    '
    '        End If
    '        On Error GoTo 0
    '    End If
        If AHOLD Or lClR Then
            If Dir(App.Path & "\bal.dat") <> "" Then
                If lClR Then
                    FileCopy App.Path & "\bal.dat", App.Path & "\bal.bck"
                    Sleep (3000)
                End If
                On Error Resume Next
                Kill App.Path & "\bal.dat"
                On Error GoTo 0
            End If
        End If
        'If Sabeco Then
        '    If frmControl.ControlImport1.Exito = 0 Or frmControl.ControlImport1.Exito = 1 _
        '    Or frmControl.ControlImport1.Exito = 10 Then
        '        ChDir MiruTa
        '        On Error Resume Next
        '        Shell ("XMSB200 /RETURN")
        '        If Err.Number <> 0 Then
        '            MsgBox "XMSB200.EXE no pudo ser ejecutado", vbCritical
        '        Else
        '            Sleep (1000)
        '            If Dir(MiruTa & "\balanza.res") <> "" Then
        '                If Dir(MiruTa & "\control", vbDirectory) = "" Then MkDir MiruTa & "\control"
        '                If Dir(MiruTa & "\balanza.res") <> "" Then FileCopy MiruTa & "\balanza.res", MiruTa & "\control\balanza.res"
        '            End If
        '        End If
        '        On Error GoTo 0
        '    End If
        'End If
        On Error GoTo salSinExito
            Select Case frmControl.ControlImport1.Exito
                Case 0
                    MsgBox CargaCadena(697)
                Case 1
                    MsgBox CargaCadena(698)
                Case 10
                    MsgBox CargaCadena(699)
                Case 11
                    MsgBox CargaCadena(700)
                    If DebugActivo Then
                        If id <> 0 Then
                            CadenadeLog "Import File not Found..."
                        Else
                            CadenadeLog "Fichero de importación no encontrado..."
                        End If
                    End If
            End Select
        'Else
            AL_Exito = frmControl.ControlImport1.Exito
        'End If
salSinExito:
        On Error GoTo 0
        If Dir(App.Path & "\asign.tst") <> "" Then
            Call AsignTec
        End If
        
        frmEpelsa.Enabled = True
        Unload frmControl
        Set frmControl = Nothing
        frmEpelsa.Enabled = True
    End If
    lCogeTiquet = True
End Sub







Public Sub mnu_envio_articulos_Click()
    Dim Resp
    If HaySeguridad Then
        If ConsultarPermiso(12) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 7
    'FrmExportar.AN_Articulos False
End Sub

Public Sub mnu_envio_codigobar_Click()
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 4
    'FrmExportar.AN_CodigosBarras False
End Sub

Public Sub mnu_Envio_IVA_Click()
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 3
    'FrmExportar.AN_IVAS False
End Sub

Public Sub mnu_envio_vendedor_Click()
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 1
    'FrmExportar.AN_Vendedores False
End Sub

Public Sub mnu_exp_bal_Click()
'atr_conti
End Sub

Public Sub mnu_iva_Click()
    If HaySeguridad Then
        If ConsultarPermiso(5) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    FrmIVA.Show
End Sub


Private Sub mnuAbsorver_Click()
Dim Resp As Integer
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703), vbCritical
        Exit Sub
    End If
    Resp = MsgBox(CargaCadena(1005), vbYesNo + vbDefaultButton2 + vbExclamation)
    If Resp = vbYes Then
        Resp = MsgBox(CargaCadena(1006), vbOKCancel + vbDefaultButton2 + vbExclamation)
        If Resp = vbOK Then
            Call Absorver_Estructura
        End If
    End If
End Sub

Public Sub mnuBaseMaestra_Click()
    Dim Resp
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'If (usaSC10 And (SC10Rs232 = False) And (Not TipoC9)) And (Not HayComandos) Then
    '    FrmNewLogin.ElUsuario = "FABRICA"
    '    FrmNewLogin.Form_Load
    '    FrmNewLogin.Show vbModal
    '    If LoginSucceeded Then
    '        Resp = MsgBox(CargaCadena(1000) & vbCrLf & _
    '        CargaCadena(1001) & vbCrLf & _
    '        CargaCadena(1002) & vbCrLf & _
    '        CargaCadena(1003) & vbCrLf & vbCrLf & _
    '        CargaCadena(1004), vbExclamation + vbYesNo + vbDefaultButton2)
    '        If Resp = vbNo Then
    '            Exit Sub
    '        Else
    '            If HaySeguridad Then
    '                CadenadeLog "Recibida orden de envío de base maestra por parte del usuario : " & UsuarioActual.Nombre
    '            Else
    '                CadenadeLog "Recibida orden de envío de Base Maestra por Parte de un usuario"
    '            End If
    '        End If
    '    Else
    '        Exit Sub
    '    End If
    'End If
    'MostrarMonitor
    AN_BaseMaestra False
End Sub

Private Sub MnuBaseVacuno_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'frmEpelsa.Enabled = False
    'FrmBaseVacuno.Show
End Sub

Public Sub mnuArticulos_Click()
    CmdTienda_Click (3)
End Sub


Private Sub mnuBorrarSecBal_Click()
    Dim Res As String
    Dim Conf
    If HaySeguridad Then
            If ConsultarPermiso(19) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        End If
    Res = InputBox(CargaCadena(1007))
    If Trim(Res) = "" Then Exit Sub
    If IsNumeric(Res) Then
        If Val(Res) > 0 And Val(Res) < 99 Then
            Conf = MsgBox(CargaCadena(1008), vbExclamation + vbYesNo)
            If Conf = vbYes Then
                Conf = AN_BorrarSeccion(Val(Res))
            End If
        Else
            MsgBox CargaCadena(1009), vbCritical
        End If
    Else
        MsgBox CargaCadena(1009), vbCritical
    End If
End Sub

Private Sub mnuBorrarTotAcum_Click()
    Dim Resp
    Dim bucle As Integer
    If HaySeguridad Then
        If ConsultarPermiso(19) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    Resp = MsgBox(CargaCadena(1010), vbYesNo + vbExclamation + vbDefaultButton2)
    If Resp = vbYes Then
        canCelar = 0
        MostrarMonitor
        
        
        If UsaGamaAlta Then
            MUE_Sistemas
            If SisEur.NSistemas > 0 Then
                For bucle = 0 To SisEur.NSistemas - 1
                    'CG_BorraTotales (700 + Bucle)
                Next bucle
            End If
        End If
        canCelar = 0
        
        frmMonitor.cmdmonitorstop.Enabled = False
        HabilitarSalir
        frmMonitor.Label1.Caption = CargaCadena(445) '"Comunicación Finalizada"
    Else
        MsgBox CargaCadena(1011), vbInformation
    End If
End Sub

'Public Sub mnuCabeceras_Click()
'    CmdTienda_Click (5)
'End Sub

'Private Sub mnuCambioDirecto_Click()
'    CmdTienda_Click (3)
'End Sub

Public Sub mnuCambioPrecio_Click()
    'CmdTienda_Click (4)
    'cambios de precio...
    If HaySeguridad Then
        If ConsultarPermiso(2) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    MostrarModal FrmCambiosPrecio
End Sub

Public Sub mnuCbarras_Click()
   CmdTienda_Click (7)
End Sub





Private Sub mnuCerrarSesion_Click()
    frmEpelsa.Hide
    UsuarioActual.Nombre = ""
    UsuarioActual.password = ""
    If Not Iconificar Then
        If HaySeguridad And Not IconificarInicio Then
            If EncontrarUsuario = False Then
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    On Error Resume Next
                    Kill App.Path & "\hydrarun.txt"
                    On Error GoTo 0
                    
                End If
                salir_programa
            Else
                MostrarInterfaz
            End If
        End If
    End If
End Sub

Private Sub mnuClientes_Click()
    If HaySeguridad Then
        If ConsultarPermiso(6) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    MostrarModal FrmClientes
End Sub


Private Sub mnuConfiguracion2_Click()
    If HaySeguridad Then
        If ConsultarPermiso(17) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    
    lCogeTiquet = False
    Me.Enabled = False
    FrmConfiguracion2.Form_Load
    FrmConfiguracion2.Show
    Do While FrmConfiguracion2.Visible
        Do_Events
    Loop
    Me.Enabled = True
    CambiarIdioma
    Activar_Menues
    
    'If frmControl.Enlazado = True Then
    '    frmControl.Balanza.RemoteUnlink
    '    EstadoBal = 1000
    'End If
    
    If descAuto Then frmControl.timetqt.Enabled = True
    lCogeTiquet = True
End Sub




Private Sub mnuConsultaTiquets_Click()
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703), vbCritical
        Exit Sub
    End If
    MostrarModal FrmConsultaTiquets
End Sub


Private Sub mnuCopiasSeguridad_Click()
    If HaySeguridad Then
        If ConsultarPermiso(33) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    Me.Enabled = False
    FrmResturarBackup.Show
    Do While FrmResturarBackup.Visible = True
        Do_Events
    Loop
    Me.Enabled = True
    Set FrmResturarBackup = Nothing
    lCogeTiquet = True
End Sub


Public Sub mnuEnviarFichasVacuno_Click()
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 5
    'FrmExportar.AN_FichasVacuno False
End Sub

Public Sub mnuEnviarPaises_Click()
    If HaySeguridad Then
        If ConsultarPermiso(11) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'MostrarMonitor
    AN_BaseMaestra False, , 5
    'FrmExportar.AN_Paises False
End Sub

Public Sub mnuEnvio_modificaciones_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    MostrarMonitor
    If Trim(Command) <> "" Then
        CadenadeLog "mnuEnvio_modificaciones_Clik()"
    End If
    AN_BaseMaestra True
End Sub
Public Sub mnuEnvioCLientes_Click()
    'MostrarMonitor
    
    AN_BaseMaestra False, , 2
    
    'FrmExportar.AN_CLientes False
End Sub
Public Sub mnumodCLi_Click()
    'MostrarMonitor
    
    AN_BaseMaestra True, , 2
    
    'FrmExportar.AN_CLientes False
End Sub


Private Sub mnuExportaGlobal_Click()
    Dim Resp
    Resp = MsgBox(CargaCadena(1408), vbYesNo + vbExclamation + vbDefaultButton2)
    If Resp = vbYes Then
        Me.Enabled = False
        Exporta_GlobalDat
        Me.Enabled = True
    End If
End Sub

'Private Sub mnuExportarFcihero_Click()
'    If HaySeguridad Then
'        If ConsultarPermiso(21) = False Then
'            MsgBox CargaCadena(693), vbCritical
'            Exit Sub
'        End If
'    End If
'    lCogeTiquet = False
'    Me.Enabled = False
'    FrmExportar.Show
'    Do While FrmExportar.Visible = True
'        Do_Events
'    Loop
'    Me.Enabled = True
'    lCogeTiquet = True
'End Sub
Private Sub mnuExportarFcihero_Click()
    If lBoka = False And lFornes = False And lBelRos = False Then
        If HaySeguridad Then
            If ConsultarPermiso(21) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        End If
        lCogeTiquet = False
        Me.Enabled = False
        FrmExportar.Show
        Do While FrmExportar.Visible = True
            Do_Events
        Loop
        Me.Enabled = True
        lCogeTiquet = True
    End If
End Sub


Private Sub mnuFamilias_Click()
    If HaySeguridad Then
        If UsuarioActual.Asociado <> 0 Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
        If ConsultarPermiso(0) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    MostrarModal FrmFamilias ' 'FrmFamilias
End Sub

Public Sub mnuFichaVacuno_Click()
    If HaySeguridad Then
        If ConsultarPermiso(4) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    frmFichasVacuno.Show
    
End Sub
Public Sub mnuficherosexportacion_Click()
    If HaySeguridad Then
        If ConsultarPermiso(16) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    Me.Enabled = False
    FrmExportConf.Show
    Do While FrmExportConf.Visible = True
        Do_Events
    Loop
    Me.Enabled = True
End Sub


Private Sub mnugenerabasevacia_Click()
Dim Resp
If HaySeguridad Then
    If ConsultarPermiso(18) = False Then
        MsgBox CargaCadena(693), vbCritical
        Exit Sub
    End If
End If
FrmNewLogin.ElUsuario = "FABRICA"
FrmNewLogin.Form_Load
FrmNewLogin.Show vbModal
If LoginSucceeded Then
    If HaySeguridad Then
        CadenadeLog "Recibida orden de borrar base de datos por parte del usuario : " & UsuarioActual.Nombre
    Else
        CadenadeLog "Recibida orden de borrar base de datos por parte de un usuario"
    End If
Else
    Exit Sub
End If
Resp = MsgBox(CargaCadena(701), vbOKCancel + vbExclamation + vbDefaultButton2)
If Resp = vbCancel Then Exit Sub
Resp = MsgBox(CargaCadena(1012), vbOKCancel + vbExclamation + vbDefaultButton2)
If Resp = vbCancel Then Exit Sub
On Error GoTo errorborrar
Kill Base_General
On Error GoTo 0
Crear_Base_Maestra
grventas
MsgBox CargaCadena(362)
Exit Sub
errorborrar:
    MsgBox CargaCadena(702), vbCritical
End Sub


Private Sub mnuGlobalDat_Click()
 Me.Enabled = False
 Importa_Global_Dat
 Me.Enabled = True
End Sub

Public Sub mnuGrandes_totales_Click()
Dim Resp As Variant
    If HaySeguridad Then
        If ConsultarPermiso(14) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    frmControl.timetqt.Enabled = False
    If Trim(Command) = "" Then
        Resp = MsgBox(CargaCadena(1013), vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbYes Then
            MostrarMonitor
            'TouchScale -->
            FrmExportar.AN_FinDeDia True
        End If
    Else
        MostrarMonitor
        'TouchScale -->
        FrmExportar.AN_FinDeDia True
    End If
    lCogeTiquet = True

End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show
    Me.Enabled = False
    Do While frmSplash.Visible = True
        Do_Events
    Loop
    Me.Enabled = True
End Sub

'Private Sub mnuhistorico_Click(Index As Integer)
'    Load frmhistorico
'    frmhistorico.Caption = "TOTALES.DAT " & CargaCadena(708)
'    If Index = 0 Then
'        frmhistorico.Tag = 0
'        frmhistorico.Label2.Caption = CargaCadena(690)
'    Else
'        frmhistorico.Tag = 1
'        frmhistorico.Label2.Caption = CargaCadena(691)
'    End If
'    lCogeTiquet = False
'    Me.Enabled = False
'    frmhistorico.Show
'    Do While frmhistorico.Visible = True
'        Do_Events
'    Loop
'    Me.Enabled = True
'    lCogeTiquet = True
'    Set frmhistorico = Nothing
'End Sub
Private Sub mnuhistorico_Click(Index As Integer)
    
    If (lBoka Or lFornes Or lBelRos) And Index = 0 Then
        MsgBox "No es posible desde fichero...", vbExclamation
        Exit Sub
    Else
        If lBelRos Then
            MsgBox "No es posible desde base de datos...", vbExclamation
            Exit Sub
        End If
        Load frmhistorico
        If lAlba Then
            frmhistorico.Caption = "TQGEN HIST."
        Else
        
        If lBoka Then
            frmhistorico.Caption = "BOKA.txt " & CargaCadena(708)
        Else
            If lFornes Then
                frmhistorico.Caption = "RX / BL " & CargaCadena(708)
            Else
                frmhistorico.Caption = "TOTALES.DAT " & CargaCadena(708)
            End If
        End If
        
        End If
        
        If Index = 0 Then
            frmhistorico.Tag = 0
            frmhistorico.Label2.Caption = CargaCadena(690)
        Else
            frmhistorico.Tag = 1
            frmhistorico.Label2.Caption = CargaCadena(691)
        End If
        lCogeTiquet = False
        Me.Enabled = False
        frmhistorico.Show
        Do While frmhistorico.Visible = True
            Do_Events
        Loop
        Me.Enabled = True
        lCogeTiquet = True
        Set frmhistorico = Nothing
    End If
End Sub


Private Sub mnuIV4COM_Click()

End Sub

Private Sub mnuListadoCambiosPrecio_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    Aux_Ejecuta "rpcprecios", ""
    Me.Enabled = True
End Sub

Private Sub mnumantener_Click()
    If HaySeguridad Then
        If ConsultarPermiso(23) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    Me.Enabled = False
    frmMantenimiento.Show
    Do While frmMantenimiento.Visible = True
        Do_Events
    Loop
    Me.Enabled = True
    lCogeTiquet = True
End Sub

Private Sub mnuModArt_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 7
    
    'FrmExportar.AN_Articulos True
End Sub

Private Sub mnuModCodBar_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 4
    
    'FrmExportar.AN_CodigosBarras True
End Sub


Private Sub mnuModFichas_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 5
End Sub

Private Sub mnuModIVA_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 3
    'FrmExportar.AN_IVAS True
End Sub

Private Sub mnuModPais_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 5
    'FrmExportar.AN_Paises True
End Sub

Private Sub mnuModVend_Click()
    If HaySeguridad Then
        If ConsultarPermiso(13) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    AN_BaseMaestra True, , 1

    'FrmExportar.AN_Vendedores True
End Sub

Public Sub mnuMostradoryEquipo_Click()
    CmdTienda_Click (0)
End Sub

Public Sub mnuPaises_Click()
    If HaySeguridad Then
        If ConsultarPermiso(4) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    FrmPaises.Show
End Sub


Public Sub mnuParamInic_Click()
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703)
        Exit Sub
    End If
    frmEpelsa.Enabled = False
    FrmLineaComandos.Show
    'frmEpelsa.Enabled = True
End Sub



Private Sub mnuplg_Click(Index As Integer)
    If HaySeguridad Then
        If ConsultarPermiso(Index + 34) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    frmEpelsa.Enabled = False
    PlugActivo = True
    FrmHlink.Start "000"
    FrmHlink.WPlugin DataPlugin.accion(Index)
    If PlugOculta Then Me.Hide
    If Iconificar Or IconificarInicio Then frmControl.cSysTray1.InTray = False
    Do While PlugActivo = True
        lCogeTiquet = False
        Do_Events
    Loop
    If Iconificar Or IconificarInicio Then frmControl.cSysTray1.InTray = True
    If PlugOculta Then Me.Show
    frmEpelsa.Enabled = True
    lCogeTiquet = True
End Sub


Private Sub mnuRecibeTiquet_Click()
    If HaySeguridad Then
        If ConsultarPermiso(14) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    'lCogeTiquet = False
    'Me.Enabled = False
    'MostrarMonitor
    'FrmExportar.AN_RecogerTotales False
    'grventas
    'frmMonitor.Enabled = False
    'lCogeTiquet = True
    lCogeTiquet = False
    frmControl.timetqt.Enabled = False
    MostrarMonitor
    'TouchScale -->
    FrmExportar.AN_FinDeDia False
    lCogeTiquet = True
    
End Sub

Private Sub mnuSalir_Click()
    If cgdtiquet = False Then salir_programa
End Sub

Public Sub mnuseuridad_Click()
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703), vbCritical
        Exit Sub
    End If
    Me.Enabled = False
    FrmSeguridad.Show
    Do While FrmSeguridad.Visible = True
        Do_Events
    Loop
    Me.Enabled = True
End Sub

Private Sub mnuSQL_Click()
    Dim Resp
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703), vbCritical
        Exit Sub
    Else
        Resp = MsgBox(CargaCadena(995), vbYesNo + vbExclamation + vbDefaultButton2)
        If Resp <> vbYes Then Exit Sub
    End If
    Resp = MsgBox(CargaCadena(1006), vbExclamation + vbDefaultButton2 + vbYesNo)
    If Resp <> vbYes Then Exit Sub
    frmEpelsa.Enabled = False
    FrmSQL.Show
    Do While FrmSQL.Visible = True
        Do_Events
    Loop
    frmEpelsa.Enabled = True
End Sub

Private Sub mnutemporales_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    Aux_Ejecuta "rpcptemporal"
                 
    Me.Enabled = True
End Sub

Private Sub mnutotalarticulos_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtarti", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtarti", "0"
    End If
    frmEpelsa.Enabled = True
End Sub



'Private Sub CommandMarcar_Click()
'    If HaySeguridad Then
'        If ConsultarPermiso(13) = False Then
'            MsgBox CargaCadena(693), vbCritical
'            Exit Sub
'        End If
'    End If
'
'    MostrarModal frmMARCAR
'
'End Sub


Private Sub mnuTotalesGB_Click()
    If HaySeguridad Then
        If ConsultarPermiso(14) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    lCogeTiquet = False
    frmControl.timetqt.Enabled = False
    MostrarMonitor
    'TouchScale -->
    FrmExportar.AN_FinDeDia False
    lCogeTiquet = True

End Sub

Private Sub mnutotalsec_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtsecs", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtsecs", "0"
    End If
    frmEpelsa.Enabled = True
End Sub

Private Sub mnutotalvendedores_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtvend", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtvend", "0"
    End If
    frmEpelsa.Enabled = True
End Sub

Private Sub mnuInformeVen_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtInfVen", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtInfVen", "0"
    End If
    frmEpelsa.Enabled = True
End Sub


Private Sub mnuInformeDes_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtInfDes", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtInfDes", "0"
    End If
    frmEpelsa.Enabled = True
End Sub
Private Sub mnuInformeHorario_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        'Aux_Ejecuta "gtInfDes", CStr(TiendaActual)
        Aux_Ejecuta "gtInfHora"
    Else
        'Aux_Ejecuta "gtInfDes", "0"
        Aux_Ejecuta "gtInfHora"
    End If
    frmEpelsa.Enabled = True
End Sub

Public Sub mnuTrade_Click()
    CmdTienda_Click (10)
End Sub

Public Sub mnuVendedores_Click()
    CmdTienda_Click (2)
End Sub
'''''''
'''''''
Private Sub rechazos_Click()
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    
    frmEpelsa.Enabled = False
    Aux_Ejecuta "rrechazos", ""
    frmEpelsa.Enabled = True
End Sub
'''''''
'''''''
Private Sub mnuTextLibre_Click()
    frmEpelsa.Enabled = False
    Load FrmTextoLibre
    FrmTextoLibre.Show vbModal
    frmEpelsa.Enabled = True
End Sub
'''''''
'''''''
Private Sub mnuInforHistorico_Click()
    Dim Salida As Byte
    Dim sruta As String
    On Error GoTo Ret
    Salida = 0
    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    If HayMulti Then
        'c2f 1.7.2
        If TiendaActual < 100 Then
            sruta = Dir(App.Path & "\T" & Format(TiendaActual, "00") & "\S", vbDirectory)
        Else
            sruta = Dir(App.Path & "\T" & Format(TiendaActual, "000") & "\S", vbDirectory)
        End If
    Else
        sruta = Dir(App.Path & "\S", vbDirectory)
    End If
    If (sruta = "") Then Salida = 1
    If (Salida) Then Err.Raise 1000
    frmEpelsa.Enabled = False
    Load fechaHistorico
    fechaHistorico.Show vbModal
    frmEpelsa.Enabled = True
Ret:
    If Err.Number <> 0 Then
        If (Err.Number = 1000) Then
            MsgBox CargaCadena(1419), vbInformation
        Else
            MsgBox Err.Description, vbCritical
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

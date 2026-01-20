VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConfiguracion2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   6000
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9084
   ControlBox      =   0   'False
   Icon            =   "FrmConfiguracion2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9084
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7680
      TabIndex        =   1
      Top             =   5400
      Width           =   1092
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1092
   End
   Begin TabDlg.SSTab Tabla 
      Height          =   5172
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   9123
      _Version        =   393216
      Tabs            =   12
      Tab             =   5
      TabsPerRow      =   4
      TabHeight       =   420
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmConfiguracion2.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Image1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Image1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label46"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame27"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Euroscale"
      TabPicture(1)   =   "FrmConfiguracion2.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "Frame16"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "SC10"
      TabPicture(2)   =   "FrmConfiguracion2.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Totales"
      TabPicture(3)   =   "FrmConfiguracion2.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame28"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "Frame8"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Moneda"
      TabPicture(4)   =   "FrmConfiguracion2.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Especiales"
      TabPicture(5)   =   "FrmConfiguracion2.frx":04CE
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame14"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame13"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame23"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Cliente/Servidor"
      TabPicture(6)   =   "FrmConfiguracion2.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ChkStoreFlow"
      Tab(6).Control(1)=   "Frame17"
      Tab(6).Control(2)=   "ChkTaquion"
      Tab(6).Control(3)=   "Frame21"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Automatización"
      TabPicture(7)   =   "FrmConfiguracion2.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame20"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame19"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame18"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Comandos Libres"
      TabPicture(8)   =   "FrmConfiguracion2.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label29(3)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Label29(2)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Label29(1)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "Label28(4)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "Label28(3)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "Label28(2)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).Control(6)=   "Label28(1)"
      Tab(8).Control(6).Enabled=   0   'False
      Tab(8).Control(7)=   "Label29(0)"
      Tab(8).Control(7).Enabled=   0   'False
      Tab(8).Control(8)=   "Label28(0)"
      Tab(8).Control(8).Enabled=   0   'False
      Tab(8).Control(9)=   "Label29(4)"
      Tab(8).Control(9).Enabled=   0   'False
      Tab(8).Control(10)=   "TxtEjecutar(4)"
      Tab(8).Control(10).Enabled=   0   'False
      Tab(8).Control(11)=   "TxtComando(4)"
      Tab(8).Control(11).Enabled=   0   'False
      Tab(8).Control(12)=   "TxtEjecutar(3)"
      Tab(8).Control(12).Enabled=   0   'False
      Tab(8).Control(13)=   "TxtComando(3)"
      Tab(8).Control(13).Enabled=   0   'False
      Tab(8).Control(14)=   "TxtEjecutar(2)"
      Tab(8).Control(14).Enabled=   0   'False
      Tab(8).Control(15)=   "TxtComando(2)"
      Tab(8).Control(15).Enabled=   0   'False
      Tab(8).Control(16)=   "TxtEjecutar(1)"
      Tab(8).Control(16).Enabled=   0   'False
      Tab(8).Control(17)=   "TxtComando(1)"
      Tab(8).Control(17).Enabled=   0   'False
      Tab(8).Control(18)=   "TxtEjecutar(0)"
      Tab(8).Control(18).Enabled=   0   'False
      Tab(8).Control(19)=   "TxtComando(0)"
      Tab(8).Control(19).Enabled=   0   'False
      Tab(8).ControlCount=   20
      TabCaption(9)   =   "Master/Slave"
      TabPicture(9)   =   "FrmConfiguracion2.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "ChkMaestro"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Frame22"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Exportación"
      TabPicture(10)  =   "FrmConfiguracion2.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame24"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "Frame25"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Modem"
      TabPicture(11)  =   "FrmConfiguracion2.frx":0576
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame26"
      Tab(11).ControlCount=   1
      Begin VB.Frame Frame28 
         Appearance      =   0  'Flat
         Caption         =   "Cambios de Precio Realizados en Balanzas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   612
         Left            =   -74880
         TabIndex        =   219
         Top             =   1560
         Width           =   8172
         Begin VB.CheckBox ChkAbsorverCambios 
            Appearance      =   0  'Flat
            Caption         =   "Absorver en la base de datos, tras cada recogida de tiquets"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   220
            Top             =   240
            Width           =   7332
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Modem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -70680
         TabIndex        =   209
         Top             =   1440
         Width           =   3972
         Begin VB.CheckBox ChkModem 
            Appearance      =   0  'Flat
            Caption         =   "Comunicar a través de Modem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   192
            Left            =   120
            TabIndex        =   210
            Top             =   360
            Width           =   3612
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Parámetros de Configuración de la Conexión por Modem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3972
         Left            =   -74760
         TabIndex        =   199
         Top             =   960
         Width           =   8172
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   2
            Left            =   120
            TabIndex        =   213
            Top             =   3360
            Visible         =   0   'False
            Width           =   7812
         End
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   120
            TabIndex        =   212
            Top             =   3120
            Visible         =   0   'False
            Width           =   7812
         End
         Begin VB.CheckBox ChkColgarModem 
            Caption         =   "Colgar el Modem Automáticamente al terminar una transmisión"
            Height          =   252
            Left            =   120
            TabIndex        =   211
            Top             =   1800
            Width           =   7812
         End
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   120
            TabIndex        =   208
            Top             =   2880
            Width           =   7812
         End
         Begin VB.TextBox TxtModemTel 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4320
            TabIndex        =   206
            Top             =   1440
            Width           =   3612
         End
         Begin VB.ComboBox CmbModemPuerto 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0592
            Left            =   4320
            List            =   "FrmConfiguracion2.frx":05A2
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   840
            Width           =   1212
         End
         Begin VB.ComboBox CmbModemVelocidad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":05BE
            Left            =   4320
            List            =   "FrmConfiguracion2.frx":05CE
            Style           =   2  'Dropdown List
            TabIndex        =   201
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label Label43 
            Caption         =   $"FrmConfiguracion2.frx":05EB
            Height          =   492
            Left            =   120
            TabIndex        =   207
            Top             =   2160
            Width           =   7812
         End
         Begin VB.Label Label42 
            Caption         =   "Numero de Teléfono a Marcar"
            Height          =   252
            Left            =   120
            TabIndex        =   205
            Top             =   1440
            Width           =   4092
         End
         Begin VB.Label Label41 
            Caption         =   "Puerto de Comunicaciones en el que se halla el Modem :"
            Height          =   492
            Left            =   120
            TabIndex        =   203
            Top             =   840
            Width           =   3372
         End
         Begin VB.Label Label40 
            Caption         =   "Baudios"
            Height          =   252
            Left            =   5760
            TabIndex        =   202
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label39 
            Caption         =   "Velocidad de transmisión de Datos"
            Height          =   252
            Left            =   120
            TabIndex        =   200
            Top             =   360
            Width           =   3972
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Sistemas a utilizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   -74880
         TabIndex        =   188
         Top             =   2280
         Width           =   3492
         Begin VB.OptionButton OptSC10 
            Caption         =   "C10 / C9"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   192
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox ChkSC10 
            Appearance      =   0  'Flat
            Caption         =   "Equipos Conectados a concentrador SC10 (V-12)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   191
            Top             =   840
            Width           =   3132
         End
         Begin VB.CheckBox ChkGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   "Equipos Euroscale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   492
            Left            =   240
            TabIndex        =   190
            Top             =   480
            Width           =   2892
         End
         Begin VB.OptionButton OptSC10 
            Caption         =   "SC10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   189
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.Frame Frame25 
         Appearance      =   0  'Flat
         Caption         =   "Exportación Configurable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   165
         Top             =   2040
         Width           =   8055
         Begin VB.OptionButton Optdiario 
            Appearance      =   0  'Flat
            Caption         =   "Acumulado"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   177
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton Optdiario 
            Appearance      =   0  'Flat
            Caption         =   "Diario"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   176
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox CmbFormato 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmConfiguracion2.frx":0685
            Left            =   2520
            List            =   "FrmConfiguracion2.frx":068F
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox CmbFormato 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmConfiguracion2.frx":06A0
            Left            =   2520
            List            =   "FrmConfiguracion2.frx":06BC
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtPositivo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2520
            MaxLength       =   1
            TabIndex        =   172
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TxtPositivo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2520
            MaxLength       =   1
            TabIndex        =   171
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox ChkAbsoluto 
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Valores Absolutos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   168
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label37 
            Caption         =   "Exportar :"
            Height          =   255
            Left            =   5160
            TabIndex        =   175
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Formato de Hora"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   170
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label36 
            Caption         =   "Formato de Fecha"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   169
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Signo de Precio Negativo"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   167
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Signo de Precio Positivo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   166
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.Frame Frame24 
         Appearance      =   0  'Flat
         Caption         =   "Ruta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74760
         TabIndex        =   162
         Top             =   1080
         Width           =   8055
         Begin VB.TextBox TxtPathExport 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   164
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label35 
            Caption         =   "Path donde depositar los ficheros : "
            Height          =   255
            Left            =   240
            TabIndex        =   163
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame23 
         Height          =   615
         Left            =   240
         TabIndex        =   155
         Top             =   2040
         Width           =   7935
         Begin VB.TextBox Txtdias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   157
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label33 
            Caption         =   "días"
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   158
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "Borrar totales acumulados más antiguos de ..."
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Parámetros de Configuración Maestro/Esclavo"
         Height          =   2532
         Left            =   -74760
         TabIndex        =   144
         Top             =   1440
         Width           =   7932
         Begin MSMask.MaskEdBox TxtMasterRHost 
            Height          =   252
            Left            =   2400
            TabIndex        =   152
            Top             =   1200
            Width           =   1812
            _ExtentX        =   3196
            _ExtentY        =   445
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   15
            Mask            =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtMasterPath 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   288
            Left            =   2400
            MaxLength       =   255
            TabIndex        =   151
            Top             =   1440
            Width           =   5172
         End
         Begin VB.TextBox TxtMasterLPort 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   288
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   150
            Top             =   960
            Width           =   1812
         End
         Begin VB.OptionButton OptMaestro 
            Appearance      =   0  'Flat
            Caption         =   "Hydra Basic Slave"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   146
            Top             =   600
            Width           =   3612
         End
         Begin VB.OptionButton OptMaestro 
            Appearance      =   0  'Flat
            Caption         =   "Hydra Basic Master"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   145
            Top             =   360
            Width           =   3492
         End
         Begin VB.Label Label31 
            Caption         =   "Path de Red Remota"
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   149
            Top             =   1440
            Width           =   1932
         End
         Begin VB.Label Label31 
            Caption         =   "Host Remoto"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   148
            Top             =   1200
            Width           =   2052
         End
         Begin VB.Label Label31 
            Caption         =   "Puerto Local"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   147
            Top             =   960
            Width           =   2052
         End
      End
      Begin VB.CheckBox ChkMaestro 
         Appearance      =   0  'Flat
         Caption         =   "Activar Configuración Master/Slave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74760
         TabIndex        =   143
         Top             =   1080
         Width           =   7212
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Conexión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   -74880
         TabIndex        =   128
         Top             =   1080
         Width           =   8052
         Begin VB.OptionButton OptTCPIP 
            Appearance      =   0  'Flat
            Caption         =   "Ethernet (TCP/IP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   240
            TabIndex        =   130
            Top             =   240
            Width           =   2172
         End
         Begin VB.OptionButton OPTRS232 
            Appearance      =   0  'Flat
            Caption         =   "Serial (RS-232)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   5160
            TabIndex        =   129
            Top             =   240
            Width           =   2772
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "TCP/IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   -74880
         TabIndex        =   125
         Top             =   2040
         Width           =   2292
         Begin VB.TextBox TxtPuerto 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1200
            TabIndex        =   126
            Top             =   444
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Puerto"
            Height          =   252
            Left            =   120
            TabIndex        =   127
            Top             =   480
            Width           =   852
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "RS-232"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   -72480
         TabIndex        =   116
         Top             =   2040
         Width           =   2532
         Begin VB.ComboBox CmbCOM 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0710
            Left            =   1320
            List            =   "FrmConfiguracion2.frx":0723
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   360
            Width           =   1092
         End
         Begin VB.ComboBox CmbVelocidad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0745
            Left            =   1320
            List            =   "FrmConfiguracion2.frx":0758
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   1080
            Width           =   1092
         End
         Begin VB.ComboBox CmbBits 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":077C
            Left            =   1320
            List            =   "FrmConfiguracion2.frx":0786
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   720
            Width           =   1092
         End
         Begin VB.ComboBox CmbParidad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0790
            Left            =   1320
            List            =   "FrmConfiguracion2.frx":079D
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   1440
            Width           =   1092
         End
         Begin VB.Label Label2 
            Caption         =   "Puerto "
            Height          =   252
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label3 
            Caption         =   "Baudios"
            Height          =   252
            Left            =   120
            TabIndex        =   123
            Top             =   1080
            Width           =   852
         End
         Begin VB.Label Label4 
            Caption         =   "Bits de Datos"
            Height          =   252
            Left            =   120
            TabIndex        =   122
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label Label5 
            Caption         =   "Paridad"
            Height          =   252
            Left            =   120
            TabIndex        =   121
            Top             =   1440
            Width           =   972
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Método de Comunicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2652
         Left            =   -74760
         TabIndex        =   102
         Top             =   1200
         Width           =   7932
         Begin VB.ComboBox CmbTipoC9 
            Height          =   315
            ItemData        =   "FrmConfiguracion2.frx":07AA
            Left            =   3360
            List            =   "FrmConfiguracion2.frx":07B7
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   960
            Width           =   4215
         End
         Begin VB.ComboBox CmbProtocolo 
            Height          =   315
            ItemData        =   "FrmConfiguracion2.frx":081B
            Left            =   3360
            List            =   "FrmConfiguracion2.frx":082E
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox TxtPathSC10 
            Appearance      =   0  'Flat
            Height          =   408
            Left            =   3000
            TabIndex        =   110
            Text            =   "Text1"
            Top             =   720
            Width           =   2652
         End
         Begin VB.CommandButton CmdElegirPAthSC10 
            Caption         =   "Seleccionar"
            Height          =   372
            Left            =   5880
            TabIndex        =   109
            Top             =   720
            Width           =   1092
         End
         Begin VB.OptionButton OptSC10Comu 
            Appearance      =   0  'Flat
            Caption         =   "Path de Red"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   108
            Top             =   360
            Width           =   3612
         End
         Begin VB.OptionButton OptSC10Comu 
            Appearance      =   0  'Flat
            Caption         =   "Serial RS-232"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   107
            Top             =   1320
            Width           =   3612
         End
         Begin VB.ComboBox CmbSC10bits 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":085F
            Left            =   1440
            List            =   "FrmConfiguracion2.frx":0869
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   1800
            Width           =   1092
         End
         Begin VB.ComboBox CmbSC10Port 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0873
            Left            =   1440
            List            =   "FrmConfiguracion2.frx":0886
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   2160
            Width           =   1092
         End
         Begin VB.ComboBox CmbSC10Paridad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":08A8
            Left            =   4200
            List            =   "FrmConfiguracion2.frx":08B5
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   2160
            Width           =   1092
         End
         Begin VB.ComboBox CMBSC10Baudios 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":08C2
            Left            =   4200
            List            =   "FrmConfiguracion2.frx":08D5
            Style           =   2  'Dropdown List
            TabIndex        =   103
            Top             =   1800
            Width           =   1092
         End
         Begin VB.Label LblTipoC9 
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionamiento del Concentrador :"
            Height          =   375
            Left            =   120
            TabIndex        =   195
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label lblProtocolo 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Protocolo"
            Height          =   375
            Left            =   120
            TabIndex        =   193
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Ruta del directorio del concentrador"
            Height          =   252
            Left            =   120
            TabIndex        =   115
            Top             =   840
            Width           =   2772
         End
         Begin VB.Label Label14 
            Caption         =   "Bits de Datos"
            Height          =   252
            Left            =   120
            TabIndex        =   114
            Top             =   1800
            Width           =   1212
         End
         Begin VB.Label Label15 
            Caption         =   "Puerto "
            Height          =   252
            Left            =   120
            TabIndex        =   113
            Top             =   2160
            Width           =   972
         End
         Begin VB.Label Label16 
            Caption         =   "Paridad"
            Height          =   252
            Left            =   3000
            TabIndex        =   112
            Top             =   2160
            Width           =   972
         End
         Begin VB.Label Label17 
            Caption         =   "Baudios"
            Height          =   252
            Left            =   3000
            TabIndex        =   111
            Top             =   1800
            Width           =   852
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "Recepción Contínua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   612
         Left            =   -74880
         TabIndex        =   98
         Top             =   840
         Width           =   8172
         Begin VB.CheckBox ChkContinua 
            Appearance      =   0  'Flat
            Caption         =   "Recepción Contínua de Tiquets"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   100
            Top             =   240
            Width           =   2772
         End
         Begin VB.ComboBox TxtIntervalo 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":08F9
            Left            =   6120
            List            =   "FrmConfiguracion2.frx":0921
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   240
            Width           =   1452
         End
         Begin VB.Label LblIntervalo 
            Caption         =   "Intervalo de recogida "
            Height          =   252
            Left            =   3600
            TabIndex        =   101
            Top             =   240
            Width           =   2292
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Euroscale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   -74880
         TabIndex        =   80
         Top             =   2280
         Width           =   3972
         Begin VB.CheckBox ChkNoTiquet 
            Appearance      =   0  'Flat
            Caption         =   "Chequeo rígido de Tiquets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   197
            Top             =   2400
            Width           =   3495
         End
         Begin VB.CheckBox ChkGaReinit 
            Appearance      =   0  'Flat
            Caption         =   "Reiniciar numeración de Tiquets tras borrar"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   2160
            Width           =   3735
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   2880
            TabIndex        =   91
            Top             =   360
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Fichas de Vacuno"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   90
            Top             =   1440
            Width           =   3252
         End
         Begin VB.CheckBox ChkBorrarVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Borrar Totales de Fichas de Vacuno"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   89
            Top             =   1920
            Width           =   2892
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   2880
            TabIndex        =   88
            Top             =   600
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   2880
            TabIndex        =   87
            Top             =   840
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   3
            Left            =   2880
            TabIndex        =   86
            Top             =   1080
            Width           =   252
         End
         Begin VB.CheckBox ChkBorrarTotalGA 
            Appearance      =   0  'Flat
            Caption         =   "Borrar totales de Euroscale"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   85
            Top             =   1680
            Width           =   2772
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   4
            Left            =   3480
            TabIndex        =   84
            Top             =   360
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   5
            Left            =   3480
            TabIndex        =   83
            Top             =   600
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   6
            Left            =   3480
            TabIndex        =   82
            Top             =   840
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   7
            Left            =   3480
            TabIndex        =   81
            Top             =   1080
            Width           =   252
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Venta"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   2652
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Super"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   96
            Top             =   600
            Width           =   2652
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Envasado"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   95
            Top             =   840
            Width           =   2652
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Autoservicio"
            Height          =   372
            Index           =   3
            Left            =   120
            TabIndex        =   94
            Top             =   1080
            Width           =   2652
         End
         Begin VB.Label Label25 
            Caption         =   "Red"
            Height          =   252
            Index           =   0
            Left            =   2760
            TabIndex        =   93
            Top             =   120
            Width           =   492
         End
         Begin VB.Label Label25 
            Caption         =   "Local"
            Height          =   252
            Index           =   1
            Left            =   3480
            TabIndex        =   92
            Top             =   120
            Width           =   492
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         Caption         =   "SC10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   -70680
         TabIndex        =   72
         Top             =   2280
         Width           =   3972
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Venta"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   3252
         End
         Begin VB.CheckBox ChkBorrarTotalSC10 
            Appearance      =   0  'Flat
            Caption         =   "Borrar totales del Concentrador"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   78
            Top             =   1920
            Width           =   2772
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de modo Super"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   600
            Width           =   3252
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Envasado"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   3252
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Autoservicio"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   1080
            Width           =   3252
         End
         Begin VB.CheckBox ChkBorrarTiquet 
            Appearance      =   0  'Flat
            Caption         =   "Borrar TICKET:DAT"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   2772
         End
         Begin VB.CheckBox ChkBorrarEti 
            Appearance      =   0  'Flat
            Caption         =   "Borrar ETI.DAT"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   73
            Top             =   1680
            Width           =   2772
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         Caption         =   "Opciones de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   -71280
         TabIndex        =   67
         Top             =   2280
         Width           =   4572
         Begin VB.TextBox TxtPasarelaPort 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   218
            Top             =   2040
            Width           =   732
         End
         Begin MSMask.MaskEdBox TxtPasarelaHost 
            Height          =   252
            Left            =   960
            TabIndex        =   216
            Top             =   2040
            Width           =   1332
            _ExtentX        =   2350
            _ExtentY        =   445
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   15
            Mask            =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox ChkPasarela 
            Appearance      =   0  'Flat
            Caption         =   "Comunicar a través de Pasarela Remota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   192
            Left            =   120
            TabIndex        =   214
            Top             =   1680
            Width           =   4212
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Version D o posterior"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   183
            Top             =   720
            Width           =   3612
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Versión C"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   182
            Top             =   480
            Width           =   3612
         End
         Begin VB.TextBox TxtPuertoPantallas 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3480
            TabIndex        =   154
            Top             =   1200
            Width           =   972
         End
         Begin VB.CheckBox ChkPantallas 
            Appearance      =   0  'Flat
            Caption         =   "Pantallas Remotas.     Puerto :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   153
            Top             =   1200
            Width           =   3252
         End
         Begin VB.CheckBox ChkFichaVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Hacer uso de Fichas de Vacuno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   3612
         End
         Begin VB.CheckBox ChkGruposConservacion 
            Appearance      =   0  'Flat
            Caption         =   "Hacer uso de Grupos de Conservación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   3732
         End
         Begin VB.CheckBox ChkGedit 
            Appearance      =   0  'Flat
            Caption         =   "Emplear el Editor de Etiquetas Gedit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.CheckBox ChkEnvioDirecto 
            Appearance      =   0  'Flat
            Caption         =   "Envío Directo de Modificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Visible         =   0   'False
            Width           =   3732
         End
         Begin VB.Label Label45 
            Caption         =   "Puerto"
            Height          =   252
            Left            =   2400
            TabIndex        =   217
            Top             =   2040
            Width           =   852
         End
         Begin VB.Label Label44 
            Caption         =   "Host :"
            Height          =   252
            Left            =   120
            TabIndex        =   215
            Top             =   2040
            Width           =   732
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         Caption         =   "Moneda Nacional (No Euros)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   -74880
         TabIndex        =   62
         Top             =   1200
         Width           =   7932
         Begin VB.TextBox TxtDecimales 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   2280
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   324
            Width           =   492
         End
         Begin VB.TextBox TxtSimbolo 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5040
            TabIndex        =   63
            Top             =   324
            Width           =   1332
         End
         Begin VB.Label Label7 
            Caption         =   "Número de Decimales :"
            Height          =   252
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1932
         End
         Begin VB.Label Label10 
            Caption         =   "Símbolo Monetario :"
            Height          =   252
            Left            =   3000
            TabIndex        =   65
            Top             =   360
            Width           =   1812
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         Caption         =   "Euro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1812
         Left            =   -74880
         TabIndex        =   57
         Top             =   2040
         Width           =   7932
         Begin VB.TextBox TxtFactorEuro 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1680
            TabIndex        =   59
            Text            =   "Text2"
            Top             =   840
            Width           =   1332
         End
         Begin VB.ComboBox CmbMoneda 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":0998
            Left            =   3960
            List            =   "FrmConfiguracion2.frx":09A2
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   1692
         End
         Begin VB.Label Label8 
            Caption         =   "Factor Euro :"
            Height          =   372
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   1212
         End
         Begin VB.Label Label9 
            Caption         =   "Moneda a Emplear en las comunicaciones "
            Height          =   252
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   3492
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         Caption         =   "Nombre de la Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   -74880
         TabIndex        =   55
         Top             =   1440
         Width           =   3972
         Begin VB.TextBox TxtEmpresa 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   3612
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         Caption         =   "Backup de la Base de Datos"
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   7932
         Begin VB.TextBox TxtBackupErase 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   160
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox ChkBackup 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Realizar Backup de la Base tras la recogida de totales"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4800
            TabIndex        =   53
            Top             =   240
            Width           =   3012
         End
         Begin VB.TextBox TxtBackup 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   288
            Left            =   1080
            TabIndex        =   52
            Top             =   240
            Width           =   3372
         End
         Begin VB.Label Label33 
            Caption         =   "días"
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   161
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label34 
            Caption         =   "Borrar copias más antíguas de"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label11 
            Caption         =   "Path"
            Height          =   252
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   732
         End
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  'Flat
         Caption         =   "Etiqueta Especial"
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   240
         TabIndex        =   46
         Top             =   2760
         Width           =   7932
         Begin VB.TextBox TxtFamiliaVacuno 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1920
            TabIndex        =   48
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox TxtEtiquetaVacuno 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   6120
            TabIndex        =   47
            Top             =   360
            Width           =   732
         End
         Begin VB.Label Label12 
            Caption         =   "Familia Especial"
            Height          =   252
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label13 
            Caption         =   "Etiqueta"
            Height          =   252
            Left            =   5040
            TabIndex        =   49
            Top             =   360
            Width           =   972
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         Caption         =   "Ubicación de dbteclas.dbf"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   3600
         Visible         =   0   'False
         Width           =   7932
         Begin VB.TextBox TxtDbteclas 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   2892
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         Caption         =   "Versión 9 y posteriores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   -69840
         TabIndex        =   42
         Top             =   2040
         Width           =   3012
         Begin VB.CheckBox ChkClientes 
            Appearance      =   0  'Flat
            Caption         =   "Versión D o posterior : Programación de Clientes y Datos de Merma en artículo"
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            TabIndex        =   198
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox Chk20Lineas 
            Appearance      =   0  'Flat
            Caption         =   "20 Líneas de descriptivo y precios por tramos"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CheckBox ChkStoreFlow 
         Appearance      =   0  'Flat
         Caption         =   "Enlace Mediante Socket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   41
         Top             =   960
         Width           =   7092
      End
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         Caption         =   "Datos de Comunicaciones con el Host"
         ForeColor       =   &H80000008&
         Height          =   1092
         Left            =   -74880
         TabIndex        =   34
         Top             =   1320
         Width           =   8052
         Begin VB.TextBox txtCI_TCP 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4200
            TabIndex        =   36
            Top             =   360
            Width           =   732
         End
         Begin VB.TextBox txtCI_UDP 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   6240
            TabIndex        =   35
            Top             =   360
            Width           =   732
         End
         Begin MSMask.MaskEdBox txtCI_HOST 
            Height          =   252
            Left            =   1320
            TabIndex        =   37
            Top             =   360
            Width           =   1572
            _ExtentX        =   2773
            _ExtentY        =   445
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   15
            Mask            =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "IP del Host : "
            Height          =   492
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label19 
            Caption         =   "Puerto TCP"
            Height          =   492
            Left            =   3120
            TabIndex        =   39
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label20 
            Caption         =   "Puerto UDP"
            Height          =   492
            Left            =   5040
            TabIndex        =   38
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame Frame18 
         Appearance      =   0  'Flat
         Caption         =   "Proceso Automático"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1332
         Left            =   -74880
         TabIndex        =   22
         Top             =   960
         Width           =   8052
         Begin VB.TextBox txtCI_MINUTO 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4680
            TabIndex        =   26
            Top             =   960
            Width           =   732
         End
         Begin VB.TextBox TxtCI_HORA 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4680
            TabIndex        =   25
            Top             =   600
            Width           =   732
         End
         Begin VB.TextBox TxtRE_HORA 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1440
            TabIndex        =   24
            Top             =   600
            Width           =   732
         End
         Begin VB.TextBox TxtRE_MINUTO 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1440
            TabIndex        =   23
            Top             =   960
            Width           =   732
         End
         Begin VB.Label Label23 
            Caption         =   "Dejando los valores en blanco no provoca Finalización del programa y/o recogida de Grandes Totales."
            Height          =   972
            Left            =   5520
            TabIndex        =   33
            Top             =   240
            Width           =   2412
         End
         Begin VB.Label Label22 
            Caption         =   "Minuto (0 -59)"
            Height          =   252
            Index           =   0
            Left            =   3360
            TabIndex        =   32
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label21 
            Caption         =   "Hora (0 -23)"
            Height          =   252
            Left            =   3360
            TabIndex        =   31
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label26 
            Caption         =   "Cerrar el Programa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   3360
            TabIndex        =   30
            Top             =   240
            Width           =   2052
         End
         Begin VB.Label Label27 
            Caption         =   "Hora (0 -23)"
            Height          =   252
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label22 
            Caption         =   "Minuto (0 -59)"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label26 
            Caption         =   "Recoger/Borrar totales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2772
         End
      End
      Begin VB.Frame Frame19 
         Appearance      =   0  'Flat
         Caption         =   "Trabajo Cotidiano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   972
         Left            =   -74880
         TabIndex        =   19
         Top             =   2400
         Width           =   8052
         Begin VB.CheckBox ChkIconificar 
            Appearance      =   0  'Flat
            Caption         =   "Iconificar el programa en la barra de Estado al minimizar o doble click sobre la interfaz"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   7572
         End
         Begin VB.CheckBox ChkIconificarInicio 
            Appearance      =   0  'Flat
            Caption         =   "Arrancar el programa en modo Iconificado"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   7692
         End
      End
      Begin VB.CheckBox ChkTaquion 
         Appearance      =   0  'Flat
         Caption         =   "Enlace ASCII"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   18
         Top             =   2520
         Width           =   2412
      End
      Begin VB.Frame Frame20 
         Appearance      =   0  'Flat
         Caption         =   "LOGS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   972
         Left            =   -74880
         TabIndex        =   15
         Top             =   3480
         Width           =   8052
         Begin VB.CheckBox ChkLogArticulos 
            Appearance      =   0  'Flat
            Caption         =   "Activar LOG de Cambios de Precio y Rechazos"
            ForeColor       =   &H80000008&
            Height          =   492
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   3852
         End
         Begin VB.CheckBox ChklogSistema 
            Appearance      =   0  'Flat
            Caption         =   "Activar Traza de Procesos (logntq.txt)"
            ForeColor       =   &H80000008&
            Height          =   612
            Left            =   4320
            TabIndex        =   16
            Top             =   240
            Width           =   3132
         End
      End
      Begin VB.Frame Frame21 
         Appearance      =   0  'Flat
         Caption         =   "Opciones"
         ForeColor       =   &H80000008&
         Height          =   2172
         Left            =   -74880
         TabIndex        =   13
         Top             =   2880
         Width           =   8052
         Begin VB.TextBox TxtAllTicket 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   186
            Top             =   1680
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "Un solo fichero para todos"
            ForeColor       =   &H80000008&
            Height          =   372
            Index           =   1
            Left            =   4800
            TabIndex        =   185
            Top             =   1320
            Width           =   2895
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "Un fichero por cada tiquet"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   0
            Left            =   4800
            TabIndex        =   184
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox ChkTqDirecto 
            Caption         =   "Enviar Modificaciones conforme se recogen los ficheros"
            Height          =   615
            Left            =   120
            TabIndex        =   181
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox TxtNombreTicket 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   179
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox ChkTaquionTiquets 
            Caption         =   "Crear ficheros de tiquets conforme se recogen de las balanzas."
            Height          =   732
            Left            =   4800
            TabIndex        =   142
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox TxtTaquionFichero 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   2172
         End
         Begin VB.Label Label38 
            Caption         =   ".ddm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   6840
            TabIndex        =   180
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label Label38 
            Caption         =   "Nombre del Fichero de Tiquets"
            Height          =   252
            Index           =   0
            Left            =   2280
            TabIndex        =   178
            Top             =   1680
            Width           =   2412
         End
         Begin VB.Label Label30 
            Caption         =   "Fichero de Importación"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   12
         Top             =   1320
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   -70680
         MaxLength       =   255
         TabIndex        =   11
         Top             =   1320
         Width           =   3852
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   10
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   -70680
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1800
         Width           =   3852
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   8
         Top             =   2280
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   -70680
         MaxLength       =   255
         TabIndex        =   7
         Top             =   2280
         Width           =   3852
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2760
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   -70680
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2760
         Width           =   3852
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   4
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   4
         Top             =   3240
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   4
         Left            =   -70680
         MaxLength       =   255
         TabIndex        =   3
         Top             =   3240
         Width           =   3852
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Configuración General de la Tienda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   372
         Left            =   -74880
         TabIndex        =   221
         Top             =   960
         Width           =   8412
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   9
         Left            =   -69840
         Picture         =   "FrmConfiguracion2.frx":09B6
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   8
         Left            =   -70320
         Picture         =   "FrmConfiguracion2.frx":0DF8
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   7
         Left            =   -70800
         Picture         =   "FrmConfiguracion2.frx":123A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   6
         Left            =   -71280
         Picture         =   "FrmConfiguracion2.frx":167C
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   5
         Left            =   -71760
         Picture         =   "FrmConfiguracion2.frx":1ABE
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   4
         Left            =   -72240
         Picture         =   "FrmConfiguracion2.frx":1F00
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   255
         Index           =   4
         Left            =   -71880
         TabIndex        =   140
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   -74160
         Picture         =   "FrmConfiguracion2.frx":2342
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   -73680
         Picture         =   "FrmConfiguracion2.frx":2784
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   -73200
         Picture         =   "FrmConfiguracion2.frx":2BC6
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   3
         Left            =   -72720
         Picture         =   "FrmConfiguracion2.frx":3490
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   139
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   255
         Index           =   0
         Left            =   -71880
         TabIndex        =   138
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   137
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   136
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   135
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   134
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   255
         Index           =   1
         Left            =   -71880
         TabIndex        =   133
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   255
         Index           =   2
         Left            =   -71880
         TabIndex        =   132
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   255
         Index           =   3
         Left            =   -71880
         TabIndex        =   131
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Image Image2 
      Height          =   492
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6252
   End
End
Attribute VB_Name = "FrmConfiguracion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TRS232
    puerto As Integer
    baudios As Integer
    paridad As String
    bits As Integer
    
End Type
Private Sub CambiarIdioma()
    Tabla.TabCaption(3) = CargaCadena(602)
    Tabla.TabCaption(4) = CargaCadena(601)
    Tabla.TabCaption(5) = CargaCadena(265)
    Tabla.TabCaption(6) = CargaCadena(603)
    Tabla.TabCaption(7) = CargaCadena(604)
    Tabla.TabCaption(8) = CargaCadena(605)
    Frame26.Caption = CargaCadena(937)
    Label39.Caption = CargaCadena(938)
    Label40.Caption = CargaCadena(939)
    Label41.Caption = CargaCadena(1065)
    Label42.Caption = CargaCadena(1066)
    ChkColgarModem.Caption = CargaCadena(1067)
    Label43.Caption = CargaCadena(1068)
    Option1(0).Caption = CargaCadena(931)
    Option1(1).Caption = CargaCadena(932)
    ChkModem.Caption = CargaCadena(933)
    ChkPasarela.Caption = CargaCadena(934)
    ChkClientes.Caption = CargaCadena(935)
    ChkNoTiquet.Caption = CargaCadena(936)
    Frame1.Caption = CargaCadena(606)
    Frame2.Caption = CargaCadena(272)
    Frame5.Caption = CargaCadena(272)
    Frame6.Caption = CargaCadena(508)
    Frame9.Caption = CargaCadena(607)
    Frame10.Caption = CargaCadena(609) & " (No Euro)"
    Frame12.Caption = CargaCadena(509)
    Frame13.Caption = CargaCadena(610)
    Frame14.Caption = CargaCadena(611)
    'Frame15.Caption = CargaCadena( 509)
    Frame16.Caption = CargaCadena(608)
    Frame17.Caption = CargaCadena(615)
    Frame18.Caption = CargaCadena(613)
    Frame19.Caption = CargaCadena(614)
    Frame21.Caption = CargaCadena(334)
    Frame22.Caption = CargaCadena(612)
    CmdAceptar.Caption = CargaCadena(287)
    CmdCancelar.Caption = CargaCadena(288)
    Label1.Caption = CargaCadena(616)
    Label2.Caption = CargaCadena(616)
    Label3.Caption = CargaCadena(261)
    Label4.Caption = CargaCadena(263)
    Label5.Caption = CargaCadena(262)
    Label7.Caption = CargaCadena(292)
    Label45.Caption = CargaCadena(616)
    Label8.Caption = CargaCadena(284)
    Label9.Caption = CargaCadena(621)
    Label10.Caption = CargaCadena(620)
    Label12.Caption = CargaCadena(622)
    Label13.Caption = CargaCadena(623)
    Label14.Caption = CargaCadena(263)
    Label15.Caption = CargaCadena(616)
    Label16.Caption = CargaCadena(262)
    Label17.Caption = CargaCadena(261)
    Label18.Caption = CargaCadena(624)
    Label19.Caption = CargaCadena(625)
    Label20.Caption = CargaCadena(626)
    Label21.Caption = CargaCadena(435) & "(0-23)"
    Label22(0).Caption = CargaCadena(627) & "(0-59)"
    Label22(1).Caption = CargaCadena(627) & "(0-59)"
    Label23.Caption = CargaCadena(630)
    Label24(0).Caption = CargaCadena(514) & " " & CargaCadena(336)
    Label24(1).Caption = CargaCadena(514) & " " & "Super"
    Label24(2).Caption = CargaCadena(514) & " " & CargaCadena(618)
    Label24(3).Caption = CargaCadena(514) & " " & CargaCadena(196)
    Label25(0).Caption = CargaCadena(332)
    Label26(0).Caption = CargaCadena(629)
    Label26(1).Caption = CargaCadena(628)
    Label27.Caption = CargaCadena(435) & "(0-23)"
    Label28(0).Caption = CargaCadena(552)
    Label28(1).Caption = CargaCadena(552)
    Label28(2).Caption = CargaCadena(552)
    Label28(3).Caption = CargaCadena(552)
    Label28(4).Caption = CargaCadena(552)
    Label29(0).Caption = CargaCadena(289)
    Label29(1).Caption = CargaCadena(289)
    Label29(2).Caption = CargaCadena(289)
    Label29(3).Caption = CargaCadena(289)
    Label29(4).Caption = CargaCadena(289)
    Label30.Caption = CargaCadena(633)
    Label31(0).Caption = CargaCadena(631)
    Label31(1).Caption = CargaCadena(632)
    Label31(3).Caption = CargaCadena(266)
    LblIntervalo.Caption = CargaCadena(617)
    ChkBorrarTotalGA.Caption = CargaCadena(38) & " " & CargaCadena(641)
    ChkBorrarTotalSC10.Caption = CargaCadena(38) & " " & CargaCadena(643)
    ChkBorrarVacuno.Caption = CargaCadena(38) & " " & CargaCadena(642)
    ChkBorrarEti.Caption = CargaCadena(38) & " " & " ETI.DAT"
    ChkBorrarTiquet.Caption = CargaCadena(38) & " " & " TICKET.DAT"
    ChkGamaAlta.Caption = CargaCadena(634)
    ChkGruposConservacion.Caption = CargaCadena(637) & " " & CargaCadena(638)
    ChkFichaVacuno.Caption = CargaCadena(637) & " " & CargaCadena(636)
    ChkTotalesVacuno.Caption = CargaCadena(514) & " " & CargaCadena(636)
    ChkGedit.Caption = CargaCadena(637) & " " & CargaCadena(639)
    ChkSC10.Caption = CargaCadena(635)
    ChkPantallas.Caption = CargaCadena(640) & ".   " & CargaCadena(616)
    ChkBackup.Caption = CargaCadena(644)
    ChkTotalesSC10(0).Caption = CargaCadena(514) & " " & CargaCadena(336)
    ChkTotalesSC10(1).Caption = CargaCadena(514) & " " & "Super"
    ChkTotalesSC10(2).Caption = CargaCadena(514) & " " & CargaCadena(618)
    ChkTotalesSC10(3).Caption = CargaCadena(514) & " " & CargaCadena(196)
    Chk20Lineas.Caption = CargaCadena(530)
    ChkStoreFlow.Caption = CargaCadena(185)
    ChkTaquion.Caption = CargaCadena(184)
    ChkMaestro.Caption = CargaCadena(297)
    ChkTaquionTiquets.Caption = CargaCadena(480)
    ChkIconificar.Caption = CargaCadena(481)
    ChkIconificarInicio.Caption = CargaCadena(482)
    ChkLogArticulos.Caption = CargaCadena(483)
    ChklogSistema.Caption = CargaCadena(484)
    ChkContinua.Caption = CargaCadena(508)
    ChkBorrarTotalSC10.Caption = CargaCadena(942)
    ChkGaReinit.Caption = CargaCadena(943)
    Frame28.Caption = CargaCadena(941)
    ChkAbsorverCambios.Caption = CargaCadena(940)
    Label35.Caption = CargaCadena(944)
    Frame25.Caption = CargaCadena(945)
    Label36(2).Caption = CargaCadena(946)
    Label36(3).Caption = CargaCadena(947)
    Label36(0).Caption = CargaCadena(948)
    Label36(1).Caption = CargaCadena(949)
    ChkAbsoluto.Caption = CargaCadena(950)
    lblProtocolo.Caption = CargaCadena(951)
    LblTipoC9.Caption = CargaCadena(952)
    ChkTqDirecto.Caption = CargaCadena(953)
    Label38(0).Caption = CargaCadena(954)
    Option2(0).Caption = CargaCadena(955)
    Option2(1).Caption = CargaCadena(956)
    '***************
    ' temporalmente
    If id <> 0 Then
        Label34.Caption = "Erase backups older than"
        Label32.Caption = "Erase Totals older than"
        Label33(1).Caption = " Days"
        Label33(0).Caption = " Days"
    End If
    Label37.Caption = CargaCadena(1052)
    Optdiario(0).Caption = CargaCadena(1049)
    Optdiario(1).Caption = CargaCadena(1050)
    ChkIconificar.Caption = CargaCadena(1051)
    Tabla.TabCaption(10) = CargaCadena(1052)
    CmbTipoC9.List(0) = "3 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
    CmbTipoC9.List(1) = "5 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
    CmbTipoC9.List(2) = "7 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
End Sub


Private Sub ChkContinua_Click()
    If ChkContinua.Value = vbChecked Then
        LblIntervalo.Enabled = True
        TxtIntervalo.Enabled = True
        TxtIntervalo.Text = TxtIntervalo.List(5)
    Else
       
        LblIntervalo.Enabled = False
        TxtIntervalo.Enabled = False
    End If
End Sub

Private Sub ChkFichaVacuno_Click()
    If ChkFichaVacuno.Value = vbChecked Then
        ChkBorrarVacuno.Enabled = True
        ChkTotalesVacuno.Enabled = True
        Option1(0).Enabled = True
        Option1(1).Enabled = True
        If Option1(0).Value = False And Option1(1).Value = False Then Option1(1).Value = True
        'Frame14.Enabled = True
    Else
        ChkBorrarVacuno.Enabled = False
        ChkTotalesVacuno.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Option1(0).Value = False
        Option1(0).Value = False
        'Frame14.Enabled = False
    End If
End Sub

Private Sub ChkGamaAlta_Click()
If ChkGamaAlta.Value = vbChecked Then
    Tabla.TabVisible(1) = True
    ChkFichaVacuno.Enabled = True
    Frame7.Enabled = True
    Frame14.Enabled = True
    ChkGruposConservacion.Enabled = True
    If ChkModem.Value = vbChecked Then
        ChkSC10.Enabled = False
        ChkSC10.Value = vbUnchecked
        OPTRS232.Value = vbChecked
    End If
Else
    ChkSC10.Enabled = True
    Tabla.TabVisible(1) = False
    ChkFichaVacuno.Enabled = False
    Frame7.Enabled = False
    Frame14.Enabled = False
    If OptSC10(1).Value Then ChkGruposConservacion.Enabled = False
End If
End Sub

Private Sub ChkIconificar_Click()
    If ChkIconificar.Value = vbChecked Then
        ChkIconificarInicio.Enabled = True
    Else
        ChkIconificarInicio.Enabled = False
        ChkIconificarInicio.Value = vbUnchecked
    End If
End Sub

Private Sub ChkMaestro_Click()
    If ChkMaestro.Value = vbChecked Then
        TxtMasterLPort.Text = "32335"
        TxtMasterRHost.Text = "127.000.000.001"
        TxtMasterPath.Text = "y:\hydra"
        OptMaestro(0).Value = True
        Frame22.Enabled = True
    Else
        Frame22.Enabled = False
    End If
End Sub

Private Sub ChkModem_Click()
    Select Case ChkModem.Value
        Case vbChecked
            If ChkPasarela.Value = vbChecked Then
                MsgBox CargaCadena(959), vbCritical  '"Este modo de Trabajo es incompatible con el uso de Pasarela remota", vbCritical
                ChkModem.Value = vbUnchecked
            Else
                If ChkGamaAlta.Value = vbChecked And ChkSC10.Value = vbChecked Then
                    MsgBox CargaCadena(960), vbCritical  '"Si emplea Modem, no puede comunicar con dos sistemas diferentes a la Vez", vbCritical
                    ChkModem.Value = vbUnchecked
                Else
                    If ChkGamaAlta.Value = vbChecked Then ChkSC10.Enabled = False
                    If ChkSC10.Value = vbChecked Then ChkGamaAlta.Enabled = False
                    Tabla.TabVisible(11) = True
                    OptTCPIP.Enabled = False
                    OPTRS232.Enabled = False
                    CmbCOM.Enabled = False
                    TxtPuerto.Enabled = False
                    CmbBits.Enabled = False
                    CmbVelocidad.Enabled = False
                    CmbParidad.Enabled = False
                    CmbSC10bits.Enabled = False
                    CmbSC10Port.Enabled = False
                    CMBSC10Baudios.Enabled = False
                    CmbSC10Paridad.Enabled = False
                    CmbProtocolo.Enabled = False
                    OPTRS232.Value = True
                    OptSC10Comu(1).Value = True
                End If
            End If
        Case vbUnchecked
            Tabla.TabVisible(11) = False
            ChkGamaAlta.Enabled = True
            ChkSC10.Enabled = True
            OptTCPIP.Enabled = True
            OPTRS232.Enabled = True
            CmbCOM.Enabled = True
            TxtPuerto.Enabled = True
            CmbBits.Enabled = True
            CmbVelocidad.Enabled = True
            CmbParidad.Enabled = True
            CmbSC10Port.Enabled = True
            If Not TipoC9 Then
                CmbSC10bits.Enabled = True
                
                CMBSC10Baudios.Enabled = True
                CmbSC10Paridad.Enabled = True
            End If
            CmbProtocolo.Enabled = True
    End Select
End Sub

Private Sub ChkPantallas_Click()
    If ChkPantallas.Value = vbUnchecked Then
        TxtPuertoPantallas.Enabled = False
    Else
        TxtPuertoPantallas.Enabled = True
    End If
End Sub

Private Sub ChkPasarela_Click()
    If ChkPasarela.Value = vbChecked Then
        If ChkModem.Value = vbChecked Then
            MsgBox CargaCadena(961)  '"Este modo de trabajo es incompatible con el trabajo por Modem", vbCritical
            ChkPasarela.Value = vbUnchecked
        Else
            TxtPasarelaHost.Enabled = True
            Label44.Enabled = True
            Label45.Enabled = True
            TxtPasarelaPort.Enabled = True
            OptSC10Comu(0).Enabled = False
            OptSC10Comu(1).Value = True
        End If
    Else
        TxtPasarelaHost.Enabled = False
        Label44.Enabled = False
        Label45.Enabled = False
        TxtPasarelaPort.Enabled = False
        If ChkModem.Value = vbUnchecked Then OptSC10Comu(0).Enabled = True
    End If
End Sub

Private Sub chkSC10_Click()
    If ChkSC10.Value = vbChecked Then
        Tabla.TabVisible(2) = True
        Frame8.Enabled = True
        OptSC10(0).Enabled = True
        OptSC10(1).Enabled = True
        If ChkModem.Value = vbChecked Then
            ChkGamaAlta.Value = vbUnchecked
            ChkGamaAlta.Enabled = False
        End If
    Else
        Tabla.TabVisible(2) = False
        Frame8.Enabled = False
        OptSC10(0).Enabled = False
        OptSC10(1).Enabled = False
        ChkGamaAlta.Enabled = True
    End If
    
End Sub







Private Sub ChkStoreFlow_Click()
    If ChkStoreFlow.Value = vbChecked Then
        ChkTaquion.Enabled = False
        ChkTaquion.Value = vbUnchecked
        Frame17.Enabled = True
        Frame19.Enabled = False
        If txtCI_HOST.Text = "___.___.___.___" Then txtCI_HOST.Text = "127.000.000.001"
        If txtCI_TCP.Text = "" Then txtCI_TCP.Text = 6000
        If txtCI_UDP.Text = "" Then txtCI_UDP.Text = 6580
        
    Else
        Frame17.Enabled = False
        Frame19.Enabled = True
        ChkTaquion.Enabled = True
    End If
End Sub







Private Sub ChkTaquion_Click()
    If ChkTaquion.Value = vbChecked Then
        ChkStoreFlow.Enabled = False
        ChkStoreFlow.Value = vbUnchecked
        Frame21.Enabled = True
    Else
        ChkStoreFlow.Enabled = True
        Frame21.Enabled = False
    End If
End Sub



Private Sub ChkTotalesGamaAlta_Click(Index As Integer)
    If Index < 4 Then
        If ChkTotalesGamaAlta(Index).Value = vbChecked Then
            ChkTotalesGamaAlta(Index + 4).Value = vbChecked
        Else
            'ChkTotalesGamaAlta(Index + 4).Value = vbUnchecked
        End If
    End If
End Sub

Private Sub ChkTotalesVacuno_Click()
    If ChkTotalesVacuno.Value = vbChecked And UsaFichasVacuno Then
        ChkBorrarVacuno.Value = vbChecked
    End If
End Sub

Private Sub CmbMoneda_Click()
    If UsaEuro = False And CmbMoneda.Text = CmbMoneda.List(1) Then
        'msgbox "Recuerde Verificar el Factor Euro", vbInformation
        MsgBox CargaCadena(209), vbInformation
    End If
    If UsaEuro = True And CmbMoneda.Text = CmbMoneda.List(0) Then
        'MsgBox "Recuerde Verificar el número de decimales y el símbolo Monetario", vbInformation
        MsgBox CargaCadena(210), vbInformation
    End If
    
End Sub



Private Sub CmbProtocolo_Click()
    Dim MyParam(3) As String
    Select Case CmbProtocolo.ListIndex
        Case 0
            MyParam(0) = "1200"
            MyParam(1) = "E"
            MyParam(2) = "7"
        Case 1
            MyParam(0) = "1200"
            MyParam(1) = "N"
            MyParam(2) = "8"
        Case 2
            MyParam(0) = "9600"
            MyParam(1) = "N"
            MyParam(2) = "8"
        Case 3
            MyParam(0) = "9600"
            MyParam(1) = "E"
            MyParam(2) = "8"
        Case 4
            MyParam(0) = "19200"
            MyParam(1) = "N"
            MyParam(2) = "8"
    End Select
    CmbSC10bits.Text = MyParam(2)
    CMBSC10Baudios.Text = MyParam(0)
    CmbSC10Paridad.Text = MyParam(1)
End Sub

Private Sub Cmdaceptar_Click()
    Dim LaRespuesta As Variant
    Dim BufferSimbolo As String
    Dim BufferDecimales As String
    Dim Devuelve As Boolean
    Dim Bucle As Integer
    Dim MasterFile As Integer
    If Trim(TxtNombreTicket.Text) = "" Then TxtNombreTicket.Text = "BBTTTT"
    If Trim(Txtdias.Text) = "" Then Txtdias.Text = "0"
    If Trim(TxtBackupErase.Text) = "" Then TxtBackupErase.Text = "0"
    TxtPathExport.Text = Trim(TxtPathExport.Text)
    If TxtPathExport.Text = "" Then TxtPathExport.Text = MiruTa
    If Right(TxtPathExport.Text, 1) = "\" Then TxtPathExport.Text = Left(TxtPathExport.Text, Len(TxtPathExport.Text) - 1)
    '****************
    ' pantallas remotas
    '****************
    If Not IsNumeric(TxtPuertoPantallas.Text) Then TxtPuertoPantallas.Text = 32337
    If Val(TxtPuertoPantallas.Text) < 1024 Or Val(TxtPuertoPantallas.Text) > 65535 Then
        'MsgBox "Puerto para pantallas remotas no válido", vbCritical
        MsgBox CargaCadena(211), vbCritical
        Tabla.Tab = 1
        Exit Sub
    End If
    '********************
    ' Master/Slave
    '********************
    If ChkMaestro.Value = vbUnchecked Then
        If Dir(MiruTa & "\master.cfg") <> "" Then Kill MiruTa & "\master.cfg"
        If Dir(MiruTa & "\slave.cfg") <> "" Then Kill MiruTa & "\slave.cfg"
    Else
        MasterFile = FreeFile()
        If OptMaestro(0).Value = True Then
            If Dir(MiruTa & "\slave.cfg") <> "" Then Kill MiruTa & "\slave.cfg"
            Open MiruTa & "\master.cfg" For Output As MasterFile
        Else
            If Dir(MiruTa & "\master.cfg") <> "" Then Kill MiruTa & "\master.cfg"
            Open MiruTa & "\slave.cfg" For Output As MasterFile
        End If
        Print #MasterFile, TxtMasterLPort.Text
        Print #MasterFile, TxtMasterRHost.Text
        Print #MasterFile, TxtMasterPath.Text
        Close #MasterFile
    End If
    '*********************
    ' puertos serie
    '*********************
    If OPTRS232.Value = False And OptSC10Comu(1).Value = True Then
        If CmbCOM.ListIndex = CmbSC10Port.ListIndex Then
            If CmbSC10Port.ListIndex < 3 Then
                CmbCOM.Text = CmbCOM.List(CmbCOM.ListIndex + 1)
            Else
                CmbCOM.Text = CmbCOM.List(0)
            End If
        End If
    Else
        If OPTRS232.Value = True And OptSC10Comu(1).Value = False Then
            If CmbCOM.ListIndex = CmbSC10Port.ListIndex Then
                If CmbCOM.ListIndex < 3 Then
                    CmbSC10Port.Text = CmbSC10Port.List(CmbSC10Port.ListIndex + 1)
                Else
                    CmbSC10Port.Text = CmbSC10Port.List(0)
                End If
            End If
        Else
            If OPTRS232.Value = True And OptSC10Comu(1).Value = True Then
                If CmbCOM.ListIndex = CmbSC10Port.ListIndex Then
                    'MsgBox "Los puertos serie de SC10 y Euroscale han de ser distintos", vbCritical
                    If ChkSC10.Value = vbChecked And ChkGamaAlta.Value = vbChecked Then
                        MsgBox CargaCadena(217), vbCritical
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    '********************
    ' Convesión Euros
    '********************
    BufferSimbolo = TxtSimbolo.Text
    BufferDecimales = TxtDecimales.Text
    If Dir(MiruTa & "\dbase.mdb") <> "" Then
        If UsaEuro = False And CmbMoneda.Text = CmbMoneda.List(1) Then
            FrmConfiguracion2.Hide
            FrmNewLogin.ElUsuario = "Euros"
            FrmNewLogin.Form_Load
            FrmNewLogin.Show vbModal, Me
            Unload FrmNewLogin
            If LoginSucceeded Then
                FrmConfiguracion2.Show
                Devuelve = ConvertirAEuros
                If Devuelve = False Then
                        TxtSimbolo.Text = BufferSimbolo
                        TxtDecimales.Text = BufferDecimales
                        CmbMoneda.Text = CmbMoneda.List(0)
                                  
                End If
                
            Else
                FrmConfiguracion2.Show
                CmbMoneda.Text = CmbMoneda.List(0)
                Unload Me
                Exit Sub
            End If
        End If
        If UsaEuro = True And CmbMoneda.Text = CmbMoneda.List(0) Then
            FrmConfiguracion2.Hide
            FrmNewLogin.ElUsuario = "Euros"
            FrmNewLogin.Form_Load
            FrmNewLogin.Show vbModal, Me
            Unload FrmNewLogin
            If LoginSucceeded Then
                FrmConfiguracion2.Show
                LaRespuesta = MsgBox("National Currency Symbol : " & TxtSimbolo.Text & Chr(13) & _
                Chr(10) & "Decimals number is : " & TxtDecimales.Text _
                & Chr(13) & Chr(10) & "¿Do you want to Continue?", vbQuestion + vbYesNo)
                If LaRespuesta = vbYes Then
                    Devuelve = ConvertirAPtas
                    If Devuelve = False Then
                        
                        
                        CmbMoneda.Text = CmbMoneda.List(1)
                    End If
                Else
                    CmbMoneda.Text = CmbMoneda.List(1)
                    
                    
                    
                End If
            Else
                FrmConfiguracion2.Show
                CmbMoneda.Text = CmbMoneda.List(1)
                
                
                
                Unload Me
                Exit Sub
            End If
        End If
    End If
    '*************************
    '*************************
    ' condiciones relativas a storeflow
    If ChkStoreFlow.Value = vbUnchecked And StoreFlow = True Then
        'MsgBox "El sistema pasará a modo Normal la próxima vez que arranque el programa", vbInformation
        MsgBox CargaCadena(269), vbInformation
    End If
    If ChkStoreFlow.Value = vbChecked And StoreFlow = False Then
        'MsgBox "El sistema pasará a modo Enlace StoreFlow la próxima vez que arranque el programa", vbInformation
        MsgBox CargaCadena(316), vbInformation
    End If
    If ChkStoreFlow.Value = vbChecked Then
        If comprueba_formato_ip(txtCI_HOST.Text) = False Then
            MsgBox "Enlace StoreFlow / Número IP de Host no válido", vbExclamation
            Tabla.Tab = 6
            txtCI_HOST.SetFocus
            Exit Sub
        End If
        If txtCI_TCP < 0 Or txtCI_TCP > 65535 Then
            MsgBox "Enlace StoreFlow / Puerto no válido", vbExclamation
            Tabla.Tab = 6
            txtCI_TCP.SetFocus
            Exit Sub
        End If
        If txtCI_UDP < 0 Or txtCI_UDP > 65535 Then
            MsgBox "Enlace StoreFlow / Puerto no válido", vbExclamation
            Tabla.Tab = 6
            txtCI_UDP.SetFocus
            Exit Sub
        End If
        '*************************
        ' Cierre y recogida automática
        '*************************
        If IsNumeric(TxtCI_HORA) Or IsNumeric(txtCI_MINUTO) Then
            If TxtCI_HORA < 0 Or TxtCI_HORA > 23 Then
                MsgBox "Enlace StoreFlow / Hora no válida", vbExclamation
                Tabla.Tab = 6
                TxtCI_HORA.SetFocus
                Exit Sub
            End If
            If txtCI_MINUTO < 0 Or txtCI_MINUTO > 59 Then
                MsgBox "Enlace StoreFlow / Minuto no válido", vbExclamation
                Tabla.Tab = 6
                txtCI_MINUTO.SetFocus
                Exit Sub
            End If
        End If
        If IsNumeric(TxtRE_HORA) Or IsNumeric(TxtRE_MINUTO) Then
            If TxtRE_HORA < 0 Or TxtRE_HORA > 23 Then
                MsgBox "Enlace StoreFlow / Hora no válida", vbExclamation
                Tabla.Tab = 6
                TxtRE_HORA.SetFocus
                Exit Sub
            End If
            If TxtRE_MINUTO < 0 Or TxtRE_MINUTO > 59 Then
                MsgBox "Enlace StoreFlow / Minuto no válido", vbExclamation
                Tabla.Tab = 6
                TxtRE_MINUTO.SetFocus
                Exit Sub
            End If
        End If
        '*************************
    End If
    ' fin storeflow
    '*************************
    If Trim(TxtFactorEuro.Text) = "" Then
        TxtFactorEuro.Text = "1"
    End If
    If ChkGamaAlta = vbUnchecked And ChkSC10 = vbUnchecked Then
        'MsgBox "Debe Elegir al Menos un tipo de Comunicación (Euroscale o SC-10)"
        MsgBox CargaCadena(270)
        Tabla.Tab = 0
        Exit Sub
    End If
    If (Trim(TxtFamiliaVacuno.Text) <> "" And Trim(TxtEtiquetaVacuno.Text) = "") _
    Or (Trim(TxtFamiliaVacuno.Text) = "" And Trim(TxtEtiquetaVacuno.Text) <> "") Then
        'MsgBox "Para definir la Familia Especial debe definir tanto la familia como la etiqueta"
        MsgBox CargaCadena(271)
        Tabla.Tab = 5
        
        Exit Sub
    End If
    '************************
    ' comandos libres
    '************************
    For Bucle = 0 To 4
        TxtComando(Bucle).Text = Trim(TxtComando(Bucle).Text)
        TxtEjecutar(Bucle).Text = Trim(TxtEjecutar(Bucle).Text)
    Next Bucle
    EscribirParametrosInicio
    LeerParametrosInicio
    If frmEpelsa.Visible = True Then
        frmEpelsa.Enabled = True
        
    End If
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    If UsaEuro Then
        decimales = 2
    End If
    Unload Me
End Sub

Private Sub CmdElegirPAthSC10_Click()
    frmSC10path.Show vbModal, Me
End Sub
Public Sub Antonio_Banderas()
    '*****************************
    ' Para cubrir las peticiones
    ' del S.A.C. sobre idiomas
    ' solicitadas por Antonio,
    ' incluímos aquí la carga
    ' de banderas para los idiomas,
    ' de modo dinámico
    '******************************
    Dim Bucle As Integer
    For Bucle = 1 To 10
        If Dir(App.Path & "\lang\" & Bucle & ".ico") <> "" Then
            Image1(Bucle - 1).Visible = True
            Image1(Bucle - 1).Picture = LoadPicture(App.Path & "\lang\" & Bucle & ".ico")
        Else
            Image1(Bucle - 1).Visible = False
        End If
    Next Bucle
End Sub
Private Sub Form_Load()
    Dim Buffer As String
    Dim Contador As Long
    Dim PSerie As TRS232
    Dim Bucle As Integer
    Dim MasterFile As Integer
    If HayMulti Then
        Label46.Visible = True
    Else
        Label46.Visible = False
    End If
    CambiarIdioma
    Antonio_Banderas
    If AbsorverCambios Then
        ChkAbsorverCambios.Value = vbChecked
    Else
        ChkAbsorverCambios.Value = vbUnchecked
    End If
    If ColgarModem Then
        ChkColgarModem.Value = vbChecked
    Else
        ChkColgarModem.Value = vbUnchecked
    End If
    If UsaModem Then
        ChkModem.Value = vbChecked
    Else
        ChkModem.Value = vbUnchecked
    End If
    ChkModem_Click
    If AhorraMas Then Frame15.Visible = True
    
    '************************
    ' modem
    '************************
    CmbModemPuerto.ListIndex = PuertoModem - 1
    CmbModemVelocidad.ListIndex = VelocidadModem - 1
    TxtModemTel.Text = TelefonoModem
    TxtModemCadena(0).Text = CadModem1
    TxtModemCadena(1).Text = CadModem2
    TxtModemCadena(2).Text = CadModem3
    '***************************
    ' Pasarela
    '***************************
    If Pasarela Then
        ChkPasarela.Value = vbChecked
    Else
        ChkPasarela.Value = vbUnchecked
    End If
    ChkPasarela_Click
    TxtPasarelaHost.Text = PasarelaHost
    TxtPasarelaPort.Text = PasarelaPuerto
    ChkClientes.Value = BoolToCheck(ClienteD)
    ChkNoTiquet.Value = BoolToCheck(CheckNoTiquet)
    If Not usaSC10 Then
        OptSC10(0).Enabled = False
        OptSC10(1).Enabled = False
    End If
    If TipoC9 Then
        OptSC10(1).Value = True
    Else
        OptSC10(0).Value = True
    End If
    CmbTipoC9.Text = CmbTipoC9.List(ModoC9)
    ChkGaReinit.Value = BoolToCheck(GA_Reinit)
    If bunTicket Then
        Option2(1).Value = True
    Else
        Option2(0).Value = True
    End If
    TxtAllTicket.Text = sunTicket
    If VacunoD = True Then
        Option1(1).Value = True
    Else
        Option1(0).Value = True
    End If
    ChkTqDirecto.Value = BoolToCheck(TqDirecto)
    TxtNombreTicket.Text = TqNombre
    If Exp_Diario = True Then
        Optdiario(0).Value = True
    Else
        Optdiario(1).Value = True
    End If
    If gedit = "1" Then
        ChkGedit.Value = vbChecked
    Else
        ChkGedit.Value = vbUnchecked
    End If
    If DiasBorrado > 0 Then
        Txtdias.Text = DiasBorrado
    Else
        Txtdias.Text = ""
    End If
    If BorrarBackup > 0 Then
        TxtBackupErase.Text = BorrarBackup
    Else
        TxtBackupErase.Text = ""
    End If
    If Not UsaDual Then
        CmbMoneda.Enabled = False
    End If
    '****************
    ' pantallas
    '****************
    If UsarPantallas Then
        ChkPantallas.Value = vbChecked
        TxtPuertoPantallas.Enabled = True
    Else
        ChkPantallas.Value = vbUnchecked
        TxtPuertoPantallas.Enabled = False
    End If
    TxtPuertoPantallas.Text = PuertoPantallas
    '********************
    ' master/slave
    '********************
    If Dir(MiruTa & "\master.cfg") <> "" Then
        ChkMaestro.Value = vbChecked
        OptMaestro(0).Value = True
        MasterFile = FreeFile()
        Open MiruTa & "\master.cfg" For Input As #MasterFile
        Line Input #MasterFile, Buffer
        TxtMasterLPort.Text = Buffer
        Line Input #MasterFile, Buffer
        TxtMasterRHost.Text = Buffer
        Line Input #MasterFile, Buffer
        TxtMasterPath.Text = Buffer
        Close #MasterFile
    Else
        If Dir(MiruTa & "\slave.cfg") <> "" Then
            ChkMaestro.Value = vbChecked
            OptMaestro(1).Value = True
            MasterFile = FreeFile()
            Open MiruTa & "\slave.cfg" For Input As #MasterFile
            Line Input #MasterFile, Buffer
            TxtMasterLPort.Text = Buffer
            Line Input #MasterFile, Buffer
            TxtMasterRHost.Text = Buffer
            Line Input #MasterFile, Buffer
            TxtMasterPath.Text = Buffer
            Close #MasterFile
        Else
            ChkMaestro.Value = False
            Frame22.Enabled = False
        End If
    End If
    '*********************
    ' comandos libres
    '*********************
    For Bucle = 0 To 4
        TxtComando(Bucle).Text = Trim(ComandoLibre(Bucle))
        TxtEjecutar(Bucle).Text = Trim(EjecutarLibre(Bucle))
    Next Bucle
    '*********************
    ' /comandos libres
    '*********************
    Image2.Picture = Form2.Imagen.Picture
    If LogArticulos Then FrmConfiguracion2.ChkLogArticulos.Value = vbChecked
    If DebugActivo Then FrmConfiguracion2.ChklogSistema.Value = vbChecked
    '**************
    ' Taquion
    '**************
    If Taquion = True Then
        ChkTaquion.Value = vbChecked
        Frame21.Enabled = True
    Else
        Frame21.Enabled = False
    End If
    If TaquionTiquets Then FrmConfiguracion2.ChkTaquionTiquets = vbChecked
    FrmConfiguracion2.TxtTaquionFichero = TaquionFichero
    ' *******
    ' El Corte Inglés / StoreFlow
    ' *******
    
    If StoreFlow Then
        ChkStoreFlow.Value = vbChecked
        Frame17.Enabled = True
        Frame19.Enabled = False
        
    Else
        ChkStoreFlow.Value = vbUnchecked
        Frame17.Enabled = False
       
    End If
    txtCI_HOST.Text = CI_HOST
    txtCI_TCP.Text = CI_TCP
    txtCI_UDP.Text = CI_UDP
    '********
    ' automatización
    '*********
    If CI_HORA < 24 And CI_MINUTO < 60 Then
        TxtCI_HORA.Text = CI_HORA
        txtCI_MINUTO.Text = CI_MINUTO
    Else
        TxtCI_HORA.Text = ""
        txtCI_MINUTO = ""
    End If
    If RE_HORA < 24 And RE_MINUTO < 60 Then
        TxtRE_HORA.Text = RE_HORA
        TxtRE_MINUTO.Text = RE_MINUTO
    Else
        TxtRE_HORA.Text = ""
        TxtRE_MINUTO = ""
    End If
    ' ********
    ' General
    ' ********
    
    TxtEmpresa.Text = Empresa
    '******************
    ' Puerto Serie Euroscale
    '******************
    PSerie = ConvierteSerie(cngvelocidad)
    Select Case PSerie.baudios
        Case 19200
            CmbVelocidad.Text = CmbVelocidad.List(4)
        Case 14400
            CmbVelocidad.Text = CmbVelocidad.List(3)
        Case 9600
            CmbVelocidad.Text = CmbVelocidad.List(2)
        Case 2400
            CmbVelocidad.Text = CmbVelocidad.List(1)
        Case 1200
            CmbVelocidad.Text = CmbVelocidad.List(0)
    End Select
    Select Case PSerie.bits
        Case 8
            CmbBits.Text = CmbBits.List(0)
        Case 7
            CmbBits.Text = CmbBits.List(1)
    End Select
    Select Case PSerie.paridad
        Case "E"
            CmbParidad.Text = CmbParidad.List(0)
        Case "O"
            CmbParidad.Text = CmbParidad.List(1)
        Case "N"
            CmbParidad.Text = CmbParidad.List(2)
    End Select
    Select Case PSerie.puerto
        Case 1
            CmbCOM.Text = CmbCOM.List(0)
        Case 2
            CmbCOM.Text = CmbCOM.List(1)
        Case 3
            CmbCOM.Text = CmbCOM.List(2)
        Case 4
            CmbCOM.Text = CmbCOM.List(3)
        Case 5
            CmbCOM.Text = CmbCOM.List(4)
    End Select
    '****************
    ' Puerto Serie SC10
    '****************
    PSerie = ConvierteSerie(SC10ParametrosSerie)
    Select Case PSerie.baudios
        Case 19200
            CMBSC10Baudios.Text = CMBSC10Baudios.List(4)
        Case 14400
            CMBSC10Baudios.Text = CMBSC10Baudios.List(3)
        Case 9600
            CMBSC10Baudios.Text = CMBSC10Baudios.List(2)
        Case 2400
            CMBSC10Baudios.Text = CMBSC10Baudios.List(1)
        Case 1200
            CMBSC10Baudios.Text = CMBSC10Baudios.List(0)
    End Select
    Select Case PSerie.bits
        Case 8
            CmbSC10bits.Text = CmbSC10bits.List(0)
        Case 7
            CmbSC10bits.Text = CmbSC10bits.List(1)
    End Select
    Select Case PSerie.paridad
        Case "E"
            CmbSC10Paridad.Text = CmbSC10Paridad.List(0)
        Case "O"
            CmbSC10Paridad.Text = CmbSC10Paridad.List(1)
        Case "N"
            CmbSC10Paridad.Text = CmbSC10Paridad.List(2)
    End Select
    Select Case PSerie.puerto
        Case 1
            CmbSC10Port.Text = CmbSC10Port.List(0)
        Case 2
            CmbSC10Port.Text = CmbSC10Port.List(1)
        Case 3
            CmbSC10Port.Text = CmbSC10Port.List(2)
        Case 4
            CmbSC10Port.Text = CmbSC10Port.List(3)
        Case 5
            CmbSC10Port.Text = CmbSC10Port.List(4)
    End Select
    If TipoC9 Then
        If CmbSC10bits.Text = "7" And CmbSC10Paridad.Text = "E" And CMBSC10Baudios.Text = "1200" Then
            CmbProtocolo.Text = CmbProtocolo.List(0)
        Else
            If CmbSC10bits.Text = "8" And CmbSC10Paridad.Text = "N" And CMBSC10Baudios.Text = "1200" Then
                CmbProtocolo.Text = CmbProtocolo.List(1)
            Else
                If CmbSC10bits.Text = "8" And CmbSC10Paridad.Text = "N" And CMBSC10Baudios.Text = "9600" Then
                    CmbProtocolo.Text = CmbProtocolo.List(2)
                Else
                    If CmbSC10bits.Text = "8" And CmbSC10Paridad.Text = "E" And CMBSC10Baudios.Text = "9600" Then
                        CmbProtocolo.Text = CmbProtocolo.List(3)
                    Else
                        If CmbSC10bits.Text = "8" And CmbSC10Paridad.Text = "N" And CMBSC10Baudios.Text = "19200" Then
                            CmbProtocolo.Text = CmbProtocolo.List(4)
                        Else
                            CmbProtocolo.Text = CmbProtocolo.List(0)
                            CmbSC10bits.Text = "7"
                            CmbSC10Paridad.Text = "E"
                            CMBSC10Baudios.Text = "1200"
                        End If
                    End If
                End If
            End If
        End If
        CmbSC10bits.Enabled = False
        CmbSC10Paridad.Enabled = False
        CMBSC10Baudios.Enabled = False
    End If
    '***************
    If UsaGamaAlta Then
        ChkGamaAlta.Value = vbChecked
        Tabla.TabVisible(1) = True
        ChkFichaVacuno.Enabled = True
        Frame7.Enabled = True
        Frame14.Enabled = True
    Else
        ChkGamaAlta.Value = vbUnchecked
        Tabla.TabVisible(1) = False
        ChkFichaVacuno.Enabled = False
        Frame7.Enabled = False
        Frame14.Enabled = False
    End If
    If UsaFichasVacuno Then
        ChkFichaVacuno.Value = vbChecked
        ChkBorrarVacuno.Enabled = True
        ChkTotalesVacuno.Enabled = True
        
    Else
        ChkFichaVacuno.Value = vbUnchecked
        ChkBorrarVacuno.Enabled = False
        ChkTotalesVacuno.Enabled = False
        
    End If
    ChkGruposConservacion.Value = BoolToCheck(UsaGruposdeConservacion)
    If usaSC10 Then
        ChkSC10.Value = vbChecked
        Tabla.TabVisible(2) = True
        Frame8.Enabled = True
    Else
        ChkSC10.Value = vbUnchecked
        Tabla.TabVisible(2) = False
        Frame8.Enabled = False
    End If
    ' ************
    ' Euroscale
    ' ************
    TxtPuerto.Text = sPortBal
    '
    ' Desglosa configuración serie
    '
    
    If tcpip Then
        OptTCPIP.Value = True
        Frame3.Enabled = True
        Frame4.Enabled = False
    Else
        OPTRS232.Value = True
        Frame4.Enabled = True
        Frame3.Enabled = False
    End If
    Chk20Lineas.Value = BoolToCheck(Usa20Lineas)
      
    
    '********
    ' Concentrador
    TxtPathSC10.Text = pathSC10
    If SC10Rs232 = False Then
        OptSC10Comu(0).Value = True
        Label6.Enabled = True
        TxtPathSC10.Enabled = True
        CmdElegirPAthSC10.Enabled = True
        CmbSC10Port.Enabled = False
        CmbSC10bits.Enabled = False
        CMBSC10Baudios.Enabled = False
        CmbSC10Paridad.Enabled = False
        Label14.Enabled = False
        Label15.Enabled = False
        Label16.Enabled = False
        Label17.Enabled = False
    Else
        OptSC10Comu(1).Value = True
        Label6.Enabled = False
        TxtPathSC10.Enabled = False
        CmdElegirPAthSC10.Enabled = False
        CmbSC10Port.Enabled = True
        If Not TipoC9 Then
            CmbSC10bits.Enabled = True
            CMBSC10Baudios.Enabled = True
            CmbSC10Paridad.Enabled = True
        End If
        Label14.Enabled = True
        Label15.Enabled = True
        Label16.Enabled = True
        Label17.Enabled = True
    End If
    
    '********
    ' Totales
    If descAuto Then
        ChkContinua.Value = vbChecked
        LblIntervalo.Enabled = True
        TxtIntervalo.Enabled = True
    Else
        ChkContinua.Value = vbUnchecked
        LblIntervalo.Enabled = False
        TxtIntervalo.Enabled = False
    End If
    If descautotime <> "" Then
        TxtIntervalo.Text = descautotime
    Else
        TxtIntervalo.Text = TxtIntervalo.List(5)
    End If
    ChkBorrarTotalGA.Value = BoolToCheck(GA_BorrarTotal)
    ChkBorrarTotalSC10.Value = BoolToCheck(SC10_BorrarTotal)
    ChkTotalesSC10(0).Value = BoolToCheck(SC10_ProcesarTotalVenta)
    ChkTotalesSC10(1).Value = BoolToCheck(SC10_ProcesarTotalSuper)
    ChkTotalesSC10(2).Value = BoolToCheck(SC10_ProcesarTotalEnvasado)
    ChkTotalesSC10(3).Value = BoolToCheck(SC10_ProcesarTotalAutoservicio)
    ChkTotalesGamaAlta(0) = BoolToCheck(GA_RecibirTotalVenta)
    ChkTotalesGamaAlta(1) = BoolToCheck(GA_RecibirTotalSuper)
    ChkTotalesGamaAlta(2) = BoolToCheck(GA_RecibirTotalEnvasado)
    ChkTotalesGamaAlta(3) = BoolToCheck(GA_RecibirTotalAutoservicio)
    ChkTotalesGamaAlta(4) = BoolToCheck(GA_RecibirTotalVentaL)
    ChkTotalesGamaAlta(5) = BoolToCheck(BoolToCheck(GA_RecibirTotalSuperL))
    ChkTotalesGamaAlta(6) = BoolToCheck(GA_RecibirTotalEnvasadoL)
    ChkTotalesGamaAlta(7) = BoolToCheck(GA_RecibirTotalAutoservicioL)
    ChkTotalesVacuno = BoolToCheck(RecibirTotalVacuno)
    ChkBorrarTiquet.Value = BoolToCheck(SC10_Borrar_Ticket_dat)
    ChkBorrarEti.Value = BoolToCheck(SC10_Borrar_Eti_dat)
    '*********************
    ' Iconificar
    '*********************
    ChkIconificarInicio.Value = BoolToCheck(IconificarInicio)
    If Iconificar Then
        ChkIconificar.Value = vbChecked
    Else
        ChkIconificar.Value = vbUnchecked
        ChkIconificarInicio.Value = vbUnchecked
        ChkIconificarInicio.Enabled = False
    End If
    '******************
    ' Moneda
    TxtSimbolo.Text = SimboloMonetario
    If IsNumeric(ValorEuro) Then
        TxtFactorEuro.Text = ValorEuro
    Else
        TxtFactorEuro.Text = 0
    End If
    If IsNumeric(decimales) Then
        TxtDecimales.Text = decimales
    Else
        decimales = 0
    End If
    If UsaEuro Then
        CmbMoneda.Text = CmbMoneda.List(1)
    Else
        CmbMoneda.Text = CmbMoneda.List(0)
    End If
    ChkEnvioDirecto.Value = BoolToCheck(EnvioDirecto)
    
    ' ***********************
    ' Especiales
    If UsaBackup Then
        ChkBackup.Value = vbChecked
    Else
        ChkBackup.Value = vbUnchecked
    End If
    TxtBackup.Text = pathred
    Tabla.Tab = 0
    TxtFamiliaVacuno.Text = FamiliaEspecial
    TxtEtiquetaVacuno.Text = EtiquetaEspecial
    TxtDbteclas = Trim(PathDBF)
    '***************+
    ' export
    CmbFormato(0).Text = CmbFormato(0).List(Exp_Fecha)
    CmbFormato(1).Text = CmbFormato(1).List(Exp_hora)
    TxtPositivo(0).Text = Exp_Mas
    TxtPositivo(1).Text = Exp_Menos
    ChkAbsoluto.Value = BoolToCheck(Exp_Absoluto)
    
    TxtPathExport.Text = Exp_Path
    '
    Opciones_Multi
End Sub
Private Sub Opciones_Multi()
    If HayMulti Then
        Tabla.TabVisible(9) = False
        Tabla.TabVisible(10) = False
        ChkPantallas.Visible = False
        TxtPuertoPantallas.Visible = False
        Frame6.Visible = False
        Frame18.Visible = False
        Frame19.Visible = False
        Frame13.Visible = False
        Tabla.TabVisible(6) = False
        Label46.Caption = "Configuración General de la Tienda " & TiendaActual
    End If
End Sub
Private Function ConvierteSerie(ByVal CadenaSerie As String) As TRS232
    Dim Contador As Long
    Dim Buffer As String
    Dim Devuelve As TRS232
    Contador = 1
    Buffer = ""
    Do Until Mid(CadenaSerie, Contador, 1) = ","
        Buffer = Buffer & Mid(CadenaSerie, Contador, 1)
        Contador = Contador + 1
    Loop
    Devuelve.baudios = Trim(Buffer)
    
    Buffer = ""
    Contador = Contador + 1
    Do Until Mid(CadenaSerie, Contador, 1) = ","
        Buffer = Buffer & Mid(CadenaSerie, Contador, 1)
        Contador = Contador + 1
    Loop
    Devuelve.paridad = Trim(Buffer)
    '
    Buffer = ""
    Contador = Contador + 1
    Do Until Mid(CadenaSerie, Contador, 1) = ","
        Buffer = Buffer & Mid(CadenaSerie, Contador, 1)
        Contador = Contador + 1
    Loop
    Devuelve.bits = Trim(Buffer)
    Buffer = ""
    Contador = Contador + 1
    Do Until Mid(CadenaSerie, Contador, 1) = ","
        Buffer = Buffer & Mid(CadenaSerie, Contador, 1)
        Contador = Contador + 1
    Loop
    Buffer = ""
    Contador = Contador + 1
    Do Until Contador > Len(CadenaSerie)
        Buffer = Buffer & Mid(CadenaSerie, Contador, 1)
        Contador = Contador + 1
    Loop
    Devuelve.puerto = Trim(Buffer)
    ConvierteSerie = Devuelve
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not StoreFlow Then frmEpelsa.Enabled = True
End Sub

Private Sub Image1_Click(Index As Integer)
    id = (Index) * 1000
    CargaIdiomas
    CambiarIdioma
End Sub



Private Sub Option2_Click(Index As Integer)
    If Option2(0).Value = True Then
        Label38(1).Caption = ".ddm"
        TxtAllTicket.Visible = False
    End If
    If Option2(1).Value = True Then
        Label38(1).Caption = ""
        TxtAllTicket.Visible = True
    End If
End Sub

Private Sub OPTRS232_Click()
    If OPTRS232.Value = True Then
        Frame4.Enabled = True
        Frame3.Enabled = False
    End If
End Sub

Private Sub OptSC10_Click(Index As Integer)
    Select Case OptSC10(0).Value
        Case True
            Tabla.TabCaption(2) = "SC10"
            OptSC10Comu(0).Visible = True
            OptSC10Comu(1).Visible = True
            Label6.Visible = True
            TxtPathSC10.Visible = True
            CmdElegirPAthSC10.Visible = True
            lblProtocolo.Visible = False
            CmbProtocolo.Visible = False
            LblTipoC9.Visible = False
            CmbTipoC9.Visible = False
            CmbSC10bits.Enabled = True
            CmbSC10Paridad.Enabled = True
            CMBSC10Baudios.Enabled = True
            ChkGruposConservacion.Enabled = True
            ChkTotalesSC10(1).Visible = True
            ChkTotalesSC10(2).Visible = True
            ChkTotalesSC10(3).Visible = True
            ChkBorrarTiquet.Visible = True
            ChkBorrarEti.Visible = True
            Frame8.Caption = "SC10"
            ChkBorrarTotalSC10.Caption = CargaCadena(942) & " SC10"
        Case False
            ChkTotalesSC10(1).Visible = False
            ChkTotalesSC10(2).Visible = False
            ChkTotalesSC10(3).Visible = False
            ChkBorrarTiquet.Visible = False
            ChkBorrarEti.Visible = False
            CmbProtocolo.Text = CmbProtocolo.List(0)
            CmbSC10bits.Text = "7"
            CmbSC10Paridad.Text = "E"
            CMBSC10Baudios.Text = "1200"
            Tabla.TabCaption(2) = "C10/C9"
            OptSC10Comu(0).Visible = False
            OptSC10Comu(1).Visible = False
            OptSC10Comu(1).Value = True
            Label6.Visible = False
            TxtPathSC10.Visible = False
            CmdElegirPAthSC10.Visible = False
            lblProtocolo.Visible = True
            CmbProtocolo.Visible = True
            CmbSC10bits.Enabled = False
            CmbSC10Paridad.Enabled = False
            CMBSC10Baudios.Enabled = False
            LblTipoC9.Visible = True
            CmbTipoC9.Visible = True
            If CmbTipoC9.Text = "" Then CmbTipoC9.Text = CmbTipoC9.List(0)
            If ChkGamaAlta.Value = vbUnchecked Then
                ChkGruposConservacion.Enabled = False
                ChkGruposConservacion.Value = vbUnchecked
            End If
            Frame8.Caption = "C9/C10"
            ChkBorrarTotalSC10.Caption = CargaCadena(942) & " C9/C10"
    End Select
End Sub

Private Sub OptSC10Comu_Click(Index As Integer)
    If OptSC10Comu(0).Value = True Then
        Label6.Enabled = True
        TxtPathSC10.Enabled = True
        CmdElegirPAthSC10.Enabled = True
        CmbSC10Port.Enabled = False
        CmbSC10bits.Enabled = False
        CMBSC10Baudios.Enabled = False
        CmbSC10Paridad.Enabled = False
        Label14.Enabled = False
        Label15.Enabled = False
        Label16.Enabled = False
        Label17.Enabled = False
    Else
        Label6.Enabled = False
        TxtPathSC10.Enabled = False
        CmdElegirPAthSC10.Enabled = False
        CmbSC10Port.Enabled = True
        CmbSC10bits.Enabled = True
        CMBSC10Baudios.Enabled = True
        CmbSC10Paridad.Enabled = True
        Label14.Enabled = True
        Label15.Enabled = True
        Label16.Enabled = True
        Label17.Enabled = True
    End If
        
End Sub

Private Sub OptTCPIP_Click()
    If OptTCPIP.Value = True Then
        Frame3.Enabled = True
        Frame4.Enabled = False
    End If
End Sub







Private Sub TxtBackup_LostFocus()
    TxtBackup.Text = Trim(TxtBackup.Text)
    If Right(TxtBackup.Text, 1) = "\" Then
        TxtBackup.Text = Left(TxtBackup.Text, Len(TxtBackup.Text) - 1)
    End If
End Sub












Private Sub TxtBackupErase_KeyPress(KeyAscii As Integer)
    TxtBackupErase.Locked = Checktexto(KeyAscii, 2, TxtBackupErase.Text, True, False)
End Sub

Private Sub TxtCI_HORA_KeyPress(KeyAscii As Integer)
    TxtCI_HORA.Locked = Checktexto(KeyAscii, 2, TxtCI_HORA.Text, True, False)
End Sub



Private Sub txtCI_MINUTO_KeyPress(KeyAscii As Integer)
    txtCI_MINUTO.Locked = Checktexto(KeyAscii, 2, txtCI_MINUTO.Text, True, False)
End Sub

Private Sub txtCI_TCP_KeyPress(KeyAscii As Integer)
    txtCI_TCP.Locked = Checktexto(KeyAscii, 5, txtCI_TCP.Text, True, False)
End Sub

Private Sub txtCI_UDP_keypress(KeyAscii As Integer)
    txtCI_UDP.Locked = Checktexto(KeyAscii, 5, txtCI_UDP.Text, True, False)
End Sub

Private Sub TxtDbteclas_LostFocus()
    If Right(TxtDbteclas.Text, 1) = "\" Then
        TxtDbteclas.Text = Left(TxtDbteclas.Text, Len(TxtDbteclas.Text) - 1)
    End If
End Sub



Private Sub TxtDecimales_KeyPress(KeyAscii As Integer)
    TxtDecimales.Locked = Checktexto(KeyAscii, 2, TxtDecimales.Text, True, False)
End Sub





Private Sub Txtdias_KeyPress(KeyAscii As Integer)
    Txtdias.Locked = Checktexto(KeyAscii, 2, Txtdias.Text, True, False)
End Sub

Private Sub TxtEjecutar_KeyPress(Index As Integer, KeyAscii As Integer)
    TxtEjecutar(Index).Locked = False
    If KeyAscii = 34 Then TxtEjecutar(Index).Locked = True
End Sub

Private Sub TxtEtiquetaVacuno_keypress(KeyAscii As Integer)
    TxtEtiquetaVacuno.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtEtiquetaVacuno.Locked = True
    If Len(TxtEtiquetaVacuno.Text) >= 2 Then TxtEtiquetaVacuno.Locked = True
    
End Sub

Private Sub TxtEtiquetaVacuno_LostFocus()
    If TxtEtiquetaVacuno.Text <> "" Then
        If Val(TxtEtiquetaVacuno.Text) > 16 Then
            MsgBox CargaCadena(485)  '"El valor de la etiqueta debe estar entre 0 y 16"
            Tabla.Tab = 5
            TxtEtiquetaVacuno.SetFocus
        End If
    End If
End Sub



Private Sub TxtFactorEuro_KeyPress(KeyAscii As Integer)
    TxtFactorEuro.Locked = Checktexto(KeyAscii, 255, TxtFactorEuro.Text, True, True)
    
End Sub



Private Sub TxtFamiliaVacuno_KeyPress(KeyAscii As Integer)
    TxtFamiliaVacuno.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtFamiliaVacuno.Locked = True
    If Len(TxtFamiliaVacuno.Text) >= 4 Then TxtFamiliaVacuno.Locked = True

End Sub

Private Sub TxtMasterLPort_keypress(KeyAscii As Integer)
    TxtMasterLPort.Locked = Checktexto(KeyAscii, 5, TxtMasterLPort.Text, True, False)
End Sub





Private Sub TxtNombreTicket_KeyPress(KeyAscii As Integer)
    TxtNombreTicket.Locked = False
    Select Case KeyAscii
        Case 98
            KeyAscii = 66
        Case 116
            KeyAscii = 84
        Case 115
            KeyAscii = 83
        Case 120
            KeyAscii = 88
        Case 99
            KeyAscii = 67
        Case 118
            KeyAscii = 86
    End Select
    If KeyAscii <> 86 And KeyAscii <> 67 And KeyAscii <> 66 And KeyAscii <> 84 And KeyAscii <> 83 And KeyAscii <> 88 And KeyAscii <> 8 Then
        TxtNombreTicket.Locked = True
    End If
End Sub



Private Sub TxtPasarelaHost_LostFocus()
    If (Not IsNumeric(Left(TxtPasarelaHost.Text, 3))) Or _
    (Not IsNumeric(Mid(TxtPasarelaHost.Text, 5, 3))) Or _
    (Not IsNumeric(Mid(TxtPasarelaHost.Text, 9, 3))) Or _
    (Not IsNumeric(Right(TxtPasarelaHost.Text, 3))) Then
        MsgBox CargaCadena(957), vbExclamation  '"Dirección IP incorrecta", vbExclamation
        Tabla.Tab = 0
        TxtPasarelaHost.SetFocus
    Else
        If (Val(Left(TxtPasarelaHost.Text, 3)) > 255) Or _
           (Val(Mid(TxtPasarelaHost.Text, 5, 3)) > 255) Or _
           (Val(Mid(TxtPasarelaHost.Text, 9, 3)) > 255) Or _
           (Val(Right(TxtPasarelaHost.Text, 3)) > 255) Then
                MsgBox CargaCadena(957), vbExclamation  '"Dirección IP incorrecta", vbExclamation
                Tabla.Tab = 0
                TxtPasarelaHost.SetFocus
        End If
    End If
End Sub



Private Sub TxtPasarelaPort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyDelete
        Case vbKeyBack
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPasarelaPort_LostFocus()
    If Val(TxtPasarelaPort.Text) < 1024 Or Val(TxtPasarelaPort.Text) > 65535 Then
        MsgBox CargaCadena(958)  '"Valor de Puerto Incorrecto", vbExclamation
        Tabla.Tab = 0
        TxtPasarelaPort.SetFocus
    End If
End Sub

Private Sub TxtPathSC10_LostFocus()
    TxtPathSC10.Text = Trim(TxtPathSC10.Text)
    If Right(TxtPathSC10.Text, 1) = "\" Then
        TxtPathSC10.Text = Left(TxtPathSC10.Text, Len(TxtPathSC10.Text) - 1)
    End If
End Sub

Private Sub TxtPuerto_KeyPress(KeyAscii As Integer)
    TxtPuerto.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtPuerto.Enabled = True
    If Len(KeyAscii) >= 5 Then TxtPuerto.Enabled = True
End Sub

Private Sub TxtPuerto_LostFocus()
    If IsNumeric(TxtPuerto.Text) Then
        If Val(TxtPuerto.Text) < 1 Or Val(TxtPuerto.Text) > 65535 Then
            MsgBox CargaCadena(486)  '"El puerto debe valer entre 1 y 65535"
            TxtPuerto.SetFocus
        End If
    End If
End Sub

Private Sub TxtPuertoPantallas_keypress(KeyAscii As Integer)
    TxtPuertoPantallas.Enabled = Checktexto(KeyAscii, 5, TxtPuertoPantallas.Text, True, False)
End Sub

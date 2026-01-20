VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmConfiguracion2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   6675
   ClientLeft      =   4785
   ClientTop       =   2520
   ClientWidth     =   10185
   ControlBox      =   0   'False
   Icon            =   "FrmConfiguracion2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10185
   Begin MSComctlLib.ListView Lista 
      Height          =   4572
      Left            =   0
      TabIndex        =   173
      Top             =   120
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12648447
      BackColor       =   8388608
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3196
      EndProperty
   End
   Begin VB.CommandButton CmdCancelar 
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
      Height          =   732
      Left            =   120
      Picture         =   "FrmConfiguracion2.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1692
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
      Height          =   732
      Left            =   120
      Picture         =   "FrmConfiguracion2.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1692
   End
   Begin TabDlg.SSTab Tabla 
      Height          =   6495
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      TabsPerRow      =   12
      TabHeight       =   176
      TabMaxWidth     =   88
      WordWrap        =   0   'False
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmConfiguracion2.frx":11A0
      Tab(0).ControlEnabled=   -1  'True
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
      Tab(0).Control(14)=   "ChkPosFam"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ChkAsignTec"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Option5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Option7"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Option8"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ChkSHname"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmConfiguracion2.frx":11BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkInventarioFinal"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ChkBorrarInventario"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ChkSolicitarInventario"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ChkSolicitarPedido"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame16"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame29"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmConfiguracion2.frx":11D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FrmConfiguracion2.frx":11F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ChkParked"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "ChkAcuET"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "ChkAcuSS"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "ChkNoPre"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "ChkNoPack"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "ChkNoLabel"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "ChkBorrarEti"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "ChkBorrarTotalSC10"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "ChkBorrarTotalGA"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "ChkGaReinit"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Frame28"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Frame6"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Frame7"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Frame8"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FrmConfiguracion2.frx":1210
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame11"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "FrmConfiguracion2.frx":122C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame30"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame23"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame3"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "FrmConfiguracion2.frx":1248
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame17"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "ChkStoreFlow"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "ChkTaquion"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Frame21"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "FrmConfiguracion2.frx":1264
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Txtdias"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame15"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "ChkLogArticulos"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "ChklogSistema"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Frame13"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Frame19"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Frame18"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Label32"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Label33(0)"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).ControlCount=   9
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "FrmConfiguracion2.frx":1280
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Dialogo"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "CmdPath(4)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "CmdPath(3)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "CmdPath(2)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "CmdPath(1)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "CmdPath(0)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).Control(6)=   "TxtComando(0)"
      Tab(8).Control(6).Enabled=   0   'False
      Tab(8).Control(7)=   "TxtEjecutar(0)"
      Tab(8).Control(7).Enabled=   0   'False
      Tab(8).Control(8)=   "TxtComando(1)"
      Tab(8).Control(8).Enabled=   0   'False
      Tab(8).Control(9)=   "TxtEjecutar(1)"
      Tab(8).Control(9).Enabled=   0   'False
      Tab(8).Control(10)=   "TxtComando(2)"
      Tab(8).Control(10).Enabled=   0   'False
      Tab(8).Control(11)=   "TxtEjecutar(2)"
      Tab(8).Control(11).Enabled=   0   'False
      Tab(8).Control(12)=   "TxtComando(3)"
      Tab(8).Control(12).Enabled=   0   'False
      Tab(8).Control(13)=   "TxtEjecutar(3)"
      Tab(8).Control(13).Enabled=   0   'False
      Tab(8).Control(14)=   "TxtComando(4)"
      Tab(8).Control(14).Enabled=   0   'False
      Tab(8).Control(15)=   "TxtEjecutar(4)"
      Tab(8).Control(15).Enabled=   0   'False
      Tab(8).Control(16)=   "Label29(4)"
      Tab(8).Control(16).Enabled=   0   'False
      Tab(8).Control(17)=   "Label28(0)"
      Tab(8).Control(17).Enabled=   0   'False
      Tab(8).Control(18)=   "Label29(0)"
      Tab(8).Control(18).Enabled=   0   'False
      Tab(8).Control(19)=   "Label28(1)"
      Tab(8).Control(19).Enabled=   0   'False
      Tab(8).Control(20)=   "Label28(2)"
      Tab(8).Control(20).Enabled=   0   'False
      Tab(8).Control(21)=   "Label28(3)"
      Tab(8).Control(21).Enabled=   0   'False
      Tab(8).Control(22)=   "Label28(4)"
      Tab(8).Control(22).Enabled=   0   'False
      Tab(8).Control(23)=   "Label29(1)"
      Tab(8).Control(23).Enabled=   0   'False
      Tab(8).Control(24)=   "Label29(2)"
      Tab(8).Control(24).Enabled=   0   'False
      Tab(8).Control(25)=   "Label29(3)"
      Tab(8).Control(25).Enabled=   0   'False
      Tab(8).ControlCount=   26
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "FrmConfiguracion2.frx":129C
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame22"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "ChkMaestro"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Tab 10"
      TabPicture(10)  =   "FrmConfiguracion2.frx":12B8
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame24"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "Frame25"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Tab 11"
      TabPicture(11)  =   "FrmConfiguracion2.frx":12D4
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame26"
      Tab(11).ControlCount=   1
      Begin VB.CheckBox ChkSHname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "shortname=name (modif.)"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   294
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox ChkParked 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Clear ticket Parked"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   293
         Top             =   1920
         Width           =   3015
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Orden Descriptivo ascendente."
         Height          =   195
         Left            =   360
         TabIndex        =   292
         Top             =   2880
         Width           =   2895
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Orden = PLU"
         Height          =   255
         Left            =   360
         TabIndex        =   291
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Orden Correlativo por PLU."
         Height          =   255
         Left            =   360
         TabIndex        =   290
         Top             =   2400
         Width           =   2895
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Orden correlativo por Código."
         Height          =   195
         Left            =   360
         TabIndex        =   289
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox ChkAcuET 
         Appearance      =   0  'Flat
         Caption         =   "Acumular Envasado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70920
         TabIndex        =   288
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox ChkAcuSS 
         Appearance      =   0  'Flat
         Caption         =   "Acumular SS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70920
         TabIndex        =   287
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox ChkNoPre 
         Appearance      =   0  'Flat
         Caption         =   "No recoger etiquetas Envasado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70920
         TabIndex        =   286
         Top             =   2280
         Width           =   3855
      End
      Begin VB.CheckBox ChkNoPack 
         Appearance      =   0  'Flat
         Caption         =   "No recoger etiquetas AutoServicio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70920
         TabIndex        =   285
         Top             =   1920
         Width           =   3735
      End
      Begin VB.CheckBox ChkNoLabel 
         Appearance      =   0  'Flat
         Caption         =   "No Recoger Etiquetas (Envasado, AutoServicio)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70920
         TabIndex        =   284
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CheckBox ChkBorrarEti 
         Appearance      =   0  'Flat
         Caption         =   "Borrar ETI.DAT"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   281
         Top             =   3000
         Width           =   3252
      End
      Begin VB.CheckBox ChkBorrarTotalSC10 
         Appearance      =   0  'Flat
         Caption         =   "Borrar totales del Concentrador"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   280
         Top             =   2640
         Width           =   3252
      End
      Begin VB.CheckBox ChkBorrarTotalGA 
         Appearance      =   0  'Flat
         Caption         =   "Borrar totales de Euroscale"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   279
         Top             =   1560
         Width           =   5775
      End
      Begin VB.CheckBox ChkGaReinit 
         Appearance      =   0  'Flat
         Caption         =   "Reiniciar numeración de Tiquets tras borrar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   278
         Top             =   2280
         Width           =   3855
      End
      Begin VB.CheckBox ChkInventarioFinal 
         Appearance      =   0  'Flat
         Caption         =   "Borrar Inventario al final del día"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   277
         Top             =   3480
         Width           =   4092
      End
      Begin VB.CheckBox ChkBorrarInventario 
         Appearance      =   0  'Flat
         Caption         =   "Borrar Inventario tras cada recogida"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   276
         Top             =   3240
         Width           =   4092
      End
      Begin VB.CheckBox ChkSolicitarInventario 
         Appearance      =   0  'Flat
         Caption         =   "Solicitar Inventario en cada recogida de tickets"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   275
         Top             =   3000
         Width           =   4092
      End
      Begin VB.CheckBox ChkSolicitarPedido 
         Appearance      =   0  'Flat
         Caption         =   "Solicitar pedido  en cada recogida de tickets"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   274
         Top             =   2760
         Width           =   3612
      End
      Begin VB.CheckBox ChkAsignTec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Asignar Posición Artículos en Familias"
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
         Height          =   375
         Left            =   120
         TabIndex        =   273
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox ChkPosFam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Trabajo con posición Familias."
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
         Height          =   255
         Left            =   120
         TabIndex        =   272
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Txtdias 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70800
         MaxLength       =   2
         TabIndex        =   252
         Top             =   3840
         Width           =   975
      End
      Begin MSComDlg.CommonDialog Dialogo 
         Left            =   -67680
         Top             =   1800
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   -69600
         TabIndex        =   244
         Top             =   4200
         Width           =   372
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -69600
         TabIndex        =   243
         Top             =   3240
         Width           =   372
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69600
         TabIndex        =   242
         Top             =   2400
         Width           =   372
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69600
         TabIndex        =   241
         Top             =   1560
         Width           =   372
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69600
         TabIndex        =   240
         Top             =   720
         Width           =   372
      End
      Begin VB.Frame Frame30 
         Caption         =   "EcoPrint / IV4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5652
         Left            =   -74880
         TabIndex        =   233
         Top             =   360
         Width           =   7692
         Begin VB.CheckBox ChkPedidoIV4 
            Appearance      =   0  'Flat
            Caption         =   "Solicitar pedido/inventario al recoger totales"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   3000
            TabIndex        =   238
            Top             =   960
            Visible         =   0   'False
            Width           =   4572
         End
         Begin VB.CheckBox ChkEcoSoloUna 
            Appearance      =   0  'Flat
            Caption         =   "Comunicar sólo con la máquina que se solicite"
            ForeColor       =   &H80000008&
            Height          =   492
            Left            =   3000
            TabIndex        =   237
            Top             =   360
            Width           =   4452
         End
         Begin VB.ComboBox CmbPortPrint 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":12F0
            Left            =   1200
            List            =   "FrmConfiguracion2.frx":1306
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   480
            Width           =   1452
         End
         Begin VB.Label Label51 
            Caption         =   "Puerto"
            Height          =   252
            Left            =   240
            TabIndex        =   234
            Top             =   480
            Width           =   852
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Red Eco2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1332
         Left            =   -74880
         TabIndex        =   222
         Top             =   1800
         Visible         =   0   'False
         Width           =   7692
         Begin VB.ComboBox CmbPortLabel 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":132E
            Left            =   1200
            List            =   "FrmConfiguracion2.frx":1344
            Style           =   2  'Dropdown List
            TabIndex        =   224
            Top             =   480
            Width           =   1332
         End
         Begin VB.TextBox TxtEquipoLabel 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   223
            Top             =   480
            Width           =   972
         End
         Begin VB.Label Label50 
            Caption         =   "Puerto"
            Height          =   252
            Left            =   240
            TabIndex        =   226
            Top             =   480
            Width           =   2292
         End
         Begin VB.Label Label49 
            Caption         =   "Número de la Balanza Conectada"
            Height          =   252
            Left            =   3240
            TabIndex        =   225
            Top             =   480
            Width           =   3012
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Red Eco1 / V8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1332
         Left            =   -74880
         TabIndex        =   214
         Top             =   360
         Width           =   7692
         Begin VB.TextBox TxtEquipoV8 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   221
            Top             =   480
            Width           =   972
         End
         Begin VB.ComboBox CmbPortV8 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":136C
            Left            =   1200
            List            =   "FrmConfiguracion2.frx":1382
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label48 
            Caption         =   "Número de la Balanza Conectada"
            Height          =   252
            Left            =   3240
            TabIndex        =   220
            Top             =   480
            Width           =   3012
         End
         Begin VB.Label Label47 
            Caption         =   "Puerto"
            Height          =   252
            Left            =   240
            TabIndex        =   218
            Top             =   480
            Width           =   2292
         End
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  'Flat
         Caption         =   "Etiqueta Especial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   612
         Left            =   -74880
         TabIndex        =   207
         Top             =   5400
         Width           =   7692
         Begin VB.TextBox TxtFamiliaVacuno 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1920
            TabIndex        =   209
            Top             =   240
            Width           =   612
         End
         Begin VB.TextBox TxtEtiquetaVacuno 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4200
            TabIndex        =   208
            Top             =   240
            Width           =   732
         End
         Begin VB.Label Label12 
            Caption         =   "Familia Especial"
            Height          =   252
            Left            =   240
            TabIndex        =   211
            Top             =   240
            Width           =   1572
         End
         Begin VB.Label Label13 
            Caption         =   "Etiqueta"
            Height          =   252
            Left            =   3120
            TabIndex        =   210
            Top             =   240
            Width           =   972
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         Caption         =   "Ubicación de dbteclas.dbf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   -74880
         TabIndex        =   205
         Top             =   5400
         Visible         =   0   'False
         Width           =   7812
         Begin VB.TextBox TxtDbteclas 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   240
            TabIndex        =   206
            Top             =   240
            Width           =   5052
         End
      End
      Begin VB.CheckBox ChkLogArticulos 
         Appearance      =   0  'Flat
         Caption         =   "Activar LOG de Cambios de Precio y Rechazos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   196
         Top             =   4320
         Width           =   7332
      End
      Begin VB.CheckBox ChklogSistema 
         Appearance      =   0  'Flat
         Caption         =   "Activar Traza de Procesos (logntq.txt)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74880
         TabIndex        =   195
         Top             =   4560
         Width           =   7572
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         Caption         =   "Versión 9 y posteriores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5055
         Left            =   -69960
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   2772
         Begin VB.CheckBox ChkL000 
            Appearance      =   0  'Flat
            Caption         =   "Version >= 00L-000."
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   271
            Top             =   4320
            Width           =   2055
         End
         Begin VB.CheckBox chkLin2040 
            Appearance      =   0  'Flat
            Caption         =   "Versión F.00C o posterior: Programación de textos adicionales de artículo líneas 20 - 40"
            ForeColor       =   &H80000008&
            Height          =   972
            Left            =   120
            TabIndex        =   262
            Top             =   3240
            Width           =   2292
         End
         Begin VB.CheckBox ChkClientes 
            Appearance      =   0  'Flat
            Caption         =   "Versión D o posterior : Programación de Clientes y Datos de Merma en artículo"
            ForeColor       =   &H80000008&
            Height          =   972
            Left            =   120
            TabIndex        =   204
            Top             =   2280
            Width           =   2292
         End
         Begin VB.CheckBox Chk20Lineas 
            Appearance      =   0  'Flat
            Caption         =   "20 Líneas de descriptivo y precios por tramos"
            ForeColor       =   &H80000008&
            Height          =   732
            Left            =   120
            TabIndex        =   203
            Top             =   1680
            Width           =   2292
         End
         Begin VB.CheckBox ChkFichaVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Hacer uso de Fichas de Vacuno"
            ForeColor       =   &H80000008&
            Height          =   852
            Left            =   120
            TabIndex        =   194
            Top             =   360
            Width           =   2052
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Versión C"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   193
            Top             =   1200
            Width           =   2292
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Version D o posterior"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   192
            Top             =   1440
            Width           =   2292
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         Caption         =   "Backup de la Base de Datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1092
         Left            =   -74880
         TabIndex        =   174
         Top             =   2640
         Width           =   7812
         Begin VB.CommandButton CmdCarpeta 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   4560
            TabIndex        =   245
            Top             =   240
            Width           =   372
         End
         Begin VB.TextBox TxtBackup 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   288
            Left            =   1080
            TabIndex        =   177
            Top             =   240
            Width           =   3372
         End
         Begin VB.CheckBox ChkBackup 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Realizar Backup de la Base tras la recogida de totales"
            ForeColor       =   &H80000008&
            Height          =   732
            Left            =   5400
            TabIndex        =   176
            Top             =   240
            Width           =   2292
         End
         Begin VB.TextBox TxtBackupErase 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   175
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Path"
            Height          =   252
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Width           =   732
         End
         Begin VB.Label Label34 
            Caption         =   "Borrar copias más antíguas de"
            Height          =   255
            Left            =   120
            TabIndex        =   179
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label33 
            Caption         =   "días"
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   178
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame29 
         Appearance      =   0  'Flat
         Caption         =   "Pedido / Inventario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   -74880
         TabIndex        =   168
         Top             =   3960
         Visible         =   0   'False
         Width           =   4812
         Begin VB.CheckBox ChkSolicitarInventarioFin 
            Appearance      =   0  'Flat
            Caption         =   "Solicitar inventario al final del día"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   172
            Top             =   1560
            Width           =   4092
         End
         Begin VB.CheckBox ChkSolicitarPedidoFin 
            Appearance      =   0  'Flat
            Caption         =   "Solicitar pedido al final del día"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   171
            Top             =   720
            Width           =   3612
         End
         Begin VB.CheckBox ChkPedidoFinalDia 
            Appearance      =   0  'Flat
            Caption         =   "Borrar pedido al final del día"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   170
            Top             =   960
            Width           =   3612
         End
         Begin VB.CheckBox ChkBorrarPedido 
            Appearance      =   0  'Flat
            Caption         =   "Borrar pedido tras cada recogida"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   169
            Top             =   480
            Width           =   3612
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3720
            Y1              =   1320
            Y2              =   1320
         End
      End
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         Caption         =   "Datos de Comunicaciones con el Host"
         ForeColor       =   &H80000008&
         Height          =   1332
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         Begin VB.TextBox txtCI_TCP 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1320
            TabIndex        =   33
            Top             =   600
            Width           =   732
         End
         Begin VB.TextBox txtCI_UDP 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1320
            TabIndex        =   32
            Top             =   840
            Width           =   732
         End
         Begin MSMask.MaskEdBox txtCI_HOST 
            Height          =   252
            Left            =   1320
            TabIndex        =   34
            Top             =   360
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   15
            Mask            =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "IP del Host : "
            Height          =   252
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label19 
            Caption         =   "Puerto TCP"
            Height          =   252
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label20 
            Caption         =   "Puerto UDP"
            Height          =   252
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1092
         End
      End
      Begin VB.Frame Frame28 
         Appearance      =   0  'Flat
         Caption         =   "Cambios de Precio Realizados en Balanzas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   612
         Left            =   -74880
         TabIndex        =   164
         Top             =   3240
         Visible         =   0   'False
         Width           =   7572
         Begin VB.CheckBox ChkAbsorverCambios 
            Appearance      =   0  'Flat
            Caption         =   "Absorver en la base de datos, tras cada recogida de tiquets"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   165
            Top             =   240
            Width           =   7092
         End
      End
      Begin VB.Frame Frame26 
         Appearance      =   0  'Flat
         Caption         =   "Parámetros de Configuración de la Conexión por Modem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5772
         Left            =   -74880
         TabIndex        =   151
         Top             =   240
         Width           =   7812
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   2
            Left            =   120
            TabIndex        =   163
            Top             =   3960
            Visible         =   0   'False
            Width           =   7332
         End
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   120
            TabIndex        =   162
            Top             =   3720
            Visible         =   0   'False
            Width           =   7332
         End
         Begin VB.CheckBox ChkColgarModem 
            Appearance      =   0  'Flat
            Caption         =   "Colgar el Modem Automáticamente al terminar una transmisión"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   161
            Top             =   2280
            Width           =   7212
         End
         Begin VB.TextBox TxtModemCadena 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   120
            TabIndex        =   160
            Top             =   3480
            Width           =   7332
         End
         Begin VB.TextBox TxtModemTel 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   120
            TabIndex        =   158
            Top             =   1800
            Width           =   3612
         End
         Begin VB.ComboBox CmbModemPuerto 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":13AA
            Left            =   4320
            List            =   "FrmConfiguracion2.frx":13BA
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   840
            Width           =   1212
         End
         Begin VB.ComboBox CmbModemVelocidad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":13D6
            Left            =   4320
            List            =   "FrmConfiguracion2.frx":13E6
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label Label43 
            Caption         =   $"FrmConfiguracion2.frx":1403
            Height          =   732
            Left            =   120
            TabIndex        =   159
            Top             =   2640
            Width           =   7212
         End
         Begin VB.Label Label42 
            Caption         =   "Numero de Teléfono a Marcar"
            Height          =   252
            Left            =   120
            TabIndex        =   157
            Top             =   1440
            Width           =   4092
         End
         Begin VB.Label Label41 
            Caption         =   "Puerto de Comunicaciones en el que se halla el Modem :"
            Height          =   492
            Left            =   120
            TabIndex        =   155
            Top             =   840
            Width           =   3372
         End
         Begin VB.Label Label40 
            Caption         =   "Baudios"
            Height          =   252
            Left            =   5760
            TabIndex        =   154
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label39 
            Caption         =   "Velocidad de transmisión de Datos"
            Height          =   252
            Left            =   120
            TabIndex        =   152
            Top             =   360
            Width           =   3972
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Sistemas a utilizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3015
         Left            =   120
         TabIndex        =   141
         Top             =   3360
         Visible         =   0   'False
         Width           =   3132
         Begin VB.CheckBox Check66T 
            Caption         =   "66T."
            Height          =   255
            Left            =   2280
            TabIndex        =   270
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox Check49T 
            Caption         =   "49T."
            Height          =   255
            Left            =   1560
            TabIndex        =   269
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox CheckMarteAlfa 
            Caption         =   "Marte V4 Alfa."
            Height          =   255
            Left            =   120
            TabIndex        =   268
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CheckBox CheckMarteNew 
            Caption         =   "Marte V:1-X-YL0"
            Height          =   255
            Left            =   120
            TabIndex        =   267
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Frame Frame27 
            BorderStyle     =   0  'None
            Caption         =   "Frame27"
            Height          =   615
            Left            =   120
            TabIndex        =   228
            Top             =   4200
            Width           =   2892
            Begin VB.OptionButton OptIV4 
               Appearance      =   0  'Flat
               Caption         =   "IV4 LCD"
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   3
               Left            =   1560
               TabIndex        =   232
               Top             =   360
               Width           =   1092
            End
            Begin VB.OptionButton OptIV4 
               Appearance      =   0  'Flat
               Caption         =   "IV4 TQ"
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   2
               Left            =   1560
               TabIndex        =   231
               Top             =   0
               Width           =   1092
            End
            Begin VB.OptionButton OptIV4 
               Appearance      =   0  'Flat
               Caption         =   "IV4 CE"
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   1
               Left            =   0
               TabIndex        =   230
               Top             =   360
               Width           =   1092
            End
            Begin VB.OptionButton OptIV4 
               Appearance      =   0  'Flat
               Caption         =   "IV4 STD"
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   0
               Left            =   0
               TabIndex        =   229
               Top             =   0
               Width           =   1092
            End
         End
         Begin VB.CheckBox option4 
            Caption         =   "19200"
            Height          =   195
            Left            =   1680
            TabIndex        =   266
            Top             =   1920
            Width           =   855
         End
         Begin VB.CheckBox option3 
            Caption         =   "9600"
            Height          =   195
            Left            =   120
            TabIndex        =   265
            Top             =   1920
            Width           =   975
         End
         Begin VB.CheckBox chkC9Cod 
            Caption         =   "COD"
            Height          =   255
            Left            =   1800
            TabIndex        =   261
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox ChkEcoPrint 
            Appearance      =   0  'Flat
            Caption         =   "IV4 / IV4 LCD"
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
            Left            =   600
            TabIndex        =   227
            Top             =   2640
            Width           =   2412
         End
         Begin VB.Frame Frame20 
            BorderStyle     =   0  'None
            Caption         =   "Frame20"
            Height          =   372
            Left            =   120
            TabIndex        =   215
            Top             =   2400
            Width           =   2892
            Begin VB.OptionButton OptECO 
               Appearance      =   0  'Flat
               Caption         =   "V8"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   1
               Left            =   1560
               TabIndex        =   217
               Top             =   0
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.OptionButton OptECO 
               Appearance      =   0  'Flat
               Caption         =   "Eco"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   0
               Left            =   0
               TabIndex        =   216
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1212
            End
         End
         Begin VB.CheckBox ChkEcoLabel 
            Appearance      =   0  'Flat
            Caption         =   "Red EcoLabel / EcoNet"
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
            Height          =   255
            Left            =   120
            TabIndex        =   213
            Top             =   1560
            Width           =   2652
         End
         Begin VB.CheckBox ChkEcoNet 
            Appearance      =   0  'Flat
            Caption         =   "Red V-8 Abierta"
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
            Height          =   375
            Left            =   120
            TabIndex        =   212
            Top             =   1320
            Width           =   2892
         End
         Begin VB.OptionButton OptSC10 
            Appearance      =   0  'Flat
            Caption         =   "C10 / C9"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   145
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox ChkSC10 
            Appearance      =   0  'Flat
            Caption         =   "Equipos Conectados a concentrador SC10 (V-12)"
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
            Height          =   495
            Left            =   120
            TabIndex        =   144
            Top             =   480
            Width           =   2772
         End
         Begin VB.CheckBox ChkGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   "Equipos Euroscale"
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
            Height          =   375
            Left            =   120
            TabIndex        =   143
            Top             =   240
            Width           =   2532
         End
         Begin VB.OptionButton OptSC10 
            Appearance      =   0  'Flat
            Caption         =   "SC10"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   142
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame25 
         Appearance      =   0  'Flat
         Caption         =   "Exportación Configurable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4812
         Left            =   -74880
         TabIndex        =   121
         Top             =   1200
         Width           =   7812
         Begin VB.OptionButton Optdiario 
            Appearance      =   0  'Flat
            Caption         =   "Acumulado"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   133
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
            TabIndex        =   132
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox CmbFormato 
            Height          =   288
            Index           =   1
            ItemData        =   "FrmConfiguracion2.frx":149D
            Left            =   2520
            List            =   "FrmConfiguracion2.frx":14AD
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox CmbFormato 
            Height          =   288
            Index           =   0
            ItemData        =   "FrmConfiguracion2.frx":14C9
            Left            =   2520
            List            =   "FrmConfiguracion2.frx":14E8
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtPositivo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2520
            MaxLength       =   1
            TabIndex        =   128
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TxtPositivo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2520
            MaxLength       =   1
            TabIndex        =   127
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox ChkAbsoluto 
            Appearance      =   0  'Flat
            Caption         =   "Utilizar Valores Absolutos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   124
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label37 
            Caption         =   "Exportar :"
            Height          =   255
            Left            =   5160
            TabIndex        =   131
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Formato de Hora"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   126
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label36 
            Caption         =   "Formato de Fecha"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   125
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Signo de Precio Negativo"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   123
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Signo de Precio Positivo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.Frame Frame24 
         Appearance      =   0  'Flat
         Caption         =   "Ruta "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   -74880
         TabIndex        =   118
         Top             =   240
         Width           =   7812
         Begin VB.CommandButton CmdCarpeta 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   7200
            TabIndex        =   246
            Top             =   360
            Width           =   372
         End
         Begin VB.TextBox TxtPathExport 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3240
            TabIndex        =   120
            Top             =   360
            Width           =   3732
         End
         Begin VB.Label Label35 
            Caption         =   "Path donde depositar los ficheros : "
            Height          =   252
            Left            =   120
            TabIndex        =   119
            Top             =   360
            Width           =   2892
         End
      End
      Begin VB.Frame Frame22 
         Appearance      =   0  'Flat
         Caption         =   "Parámetros de Configuración Maestro/Esclavo"
         ForeColor       =   &H00800000&
         Height          =   5292
         Left            =   -74880
         TabIndex        =   107
         Top             =   720
         Visible         =   0   'False
         Width           =   7812
         Begin MSMask.MaskEdBox TxtMasterRHost 
            Height          =   252
            Left            =   2400
            TabIndex        =   115
            Top             =   1200
            Width           =   1812
            _ExtentX        =   3201
            _ExtentY        =   450
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
            TabIndex        =   114
            Top             =   1440
            Width           =   4932
         End
         Begin VB.TextBox TxtMasterLPort 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   288
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   113
            Top             =   960
            Width           =   1812
         End
         Begin VB.OptionButton OptMaestro 
            Appearance      =   0  'Flat
            Caption         =   "Hydra Basic Slave"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   109
            Top             =   600
            Width           =   3612
         End
         Begin VB.OptionButton OptMaestro 
            Appearance      =   0  'Flat
            Caption         =   "Hydra Basic Master"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   108
            Top             =   360
            Width           =   3492
         End
         Begin VB.Label Label31 
            Caption         =   "Path de Red Remota"
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   112
            Top             =   1440
            Width           =   1932
         End
         Begin VB.Label Label31 
            Caption         =   "Host Remoto"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   111
            Top             =   1200
            Width           =   2052
         End
         Begin VB.Label Label31 
            Caption         =   "Puerto Local"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   110
            Top             =   960
            Width           =   2052
         End
      End
      Begin VB.CheckBox ChkMaestro 
         Appearance      =   0  'Flat
         Caption         =   "Activar Configuración Master/Slave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   -74880
         TabIndex        =   106
         Top             =   360
         Visible         =   0   'False
         Width           =   7212
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Conexión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2412
         Left            =   -74880
         TabIndex        =   91
         Top             =   240
         Width           =   4812
         Begin VB.CheckBox ChkMultiEuroscale 
            Appearance      =   0  'Flat
            Caption         =   " Trabajar con Servidora por Sección (Ned ID)"
            ForeColor       =   &H80000008&
            Height          =   852
            Left            =   120
            TabIndex        =   247
            Top             =   1440
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Caption         =   "RS-232"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1692
            Left            =   2160
            TabIndex        =   183
            Top             =   600
            Width           =   2532
            Begin VB.ComboBox CmbParidad 
               Height          =   288
               ItemData        =   "FrmConfiguracion2.frx":1546
               Left            =   1320
               List            =   "FrmConfiguracion2.frx":1553
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   1320
               Visible         =   0   'False
               Width           =   1092
            End
            Begin VB.ComboBox CmbBits 
               Height          =   288
               ItemData        =   "FrmConfiguracion2.frx":1560
               Left            =   1320
               List            =   "FrmConfiguracion2.frx":156A
               Style           =   2  'Dropdown List
               TabIndex        =   186
               Top             =   600
               Visible         =   0   'False
               Width           =   1092
            End
            Begin VB.ComboBox CmbVelocidad 
               Height          =   288
               ItemData        =   "FrmConfiguracion2.frx":1574
               Left            =   1320
               List            =   "FrmConfiguracion2.frx":1587
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   960
               Visible         =   0   'False
               Width           =   1092
            End
            Begin VB.ComboBox CmbCOM 
               Height          =   288
               ItemData        =   "FrmConfiguracion2.frx":15AB
               Left            =   1320
               List            =   "FrmConfiguracion2.frx":15BE
               Style           =   2  'Dropdown List
               TabIndex        =   184
               Top             =   240
               Visible         =   0   'False
               Width           =   1092
            End
            Begin VB.Label Label5 
               Caption         =   "Paridad"
               Height          =   252
               Left            =   0
               TabIndex        =   191
               Top             =   1320
               Visible         =   0   'False
               Width           =   972
            End
            Begin VB.Label Label4 
               Caption         =   "Bits de Datos"
               Height          =   252
               Left            =   0
               TabIndex        =   190
               Top             =   600
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.Label Label3 
               Caption         =   "Baudios"
               Height          =   252
               Left            =   0
               TabIndex        =   189
               Top             =   960
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label Label2 
               Caption         =   "Puerto "
               Height          =   252
               Left            =   0
               TabIndex        =   188
               Top             =   240
               Visible         =   0   'False
               Width           =   972
            End
         End
         Begin VB.TextBox TxtPuerto 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   120
            TabIndex        =   182
            Top             =   1080
            Width           =   852
         End
         Begin VB.OptionButton OptTCPIP 
            Appearance      =   0  'Flat
            Caption         =   "Ethernet (TCP/IP)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   1932
         End
         Begin VB.OptionButton OPTRS232 
            Appearance      =   0  'Flat
            Caption         =   "Serial (RS-232)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   2160
            TabIndex        =   92
            Top             =   240
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Left            =   120
            TabIndex        =   181
            Top             =   720
            Width           =   852
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Método de Comunicación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5772
         Left            =   -74760
         TabIndex        =   77
         Top             =   240
         Width           =   7692
         Begin VB.TextBox TxtTCPSC10 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            TabIndex        =   258
            Top             =   3840
            Width           =   852
         End
         Begin VB.TextBox TxtIPSC10 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            TabIndex        =   257
            Top             =   3480
            Width           =   2172
         End
         Begin VB.OptionButton OptSC10Comu 
            Appearance      =   0  'Flat
            Caption         =   "TCP/IP"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   256
            Top             =   3000
            Width           =   3612
         End
         Begin VB.CheckBox Chk4Lineas 
            Appearance      =   0  'Flat
            Caption         =   "4 lineas"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   248
            Top             =   4440
            Width           =   6372
         End
         Begin VB.TextBox TxtPathSC10 
            Appearance      =   0  'Flat
            Height          =   408
            Left            =   3000
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   720
            Width           =   2652
         End
         Begin VB.CommandButton CmdElegirPAthSC10 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5880
            TabIndex        =   84
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton OptSC10Comu 
            Appearance      =   0  'Flat
            Caption         =   "Path de Red"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   83
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
            TabIndex        =   82
            Top             =   1320
            Width           =   3612
         End
         Begin VB.ComboBox CmbSC10bits 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":15E0
            Left            =   1440
            List            =   "FrmConfiguracion2.frx":15EA
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1800
            Width           =   1092
         End
         Begin VB.ComboBox CmbSC10Port 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":15F4
            Left            =   1440
            List            =   "FrmConfiguracion2.frx":1607
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   2160
            Width           =   1092
         End
         Begin VB.ComboBox CmbSC10Paridad 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":1629
            Left            =   4200
            List            =   "FrmConfiguracion2.frx":1636
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   2160
            Width           =   1092
         End
         Begin VB.ComboBox CMBSC10Baudios 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":1643
            Left            =   4200
            List            =   "FrmConfiguracion2.frx":1656
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1800
            Width           =   1092
         End
         Begin VB.ComboBox CmbProtocolo 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":167A
            Left            =   3360
            List            =   "FrmConfiguracion2.frx":168D
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   600
            Width           =   2295
         End
         Begin VB.ComboBox CmbTipoC9 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":16BE
            Left            =   3360
            List            =   "FrmConfiguracion2.frx":16CB
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   960
            Width           =   4092
         End
         Begin VB.Label LblSC10TCP 
            Caption         =   "Port:"
            Height          =   252
            Left            =   240
            TabIndex        =   260
            Top             =   3840
            Width           =   1212
         End
         Begin VB.Label LblSC10IP 
            Caption         =   "IP:"
            Height          =   252
            Left            =   240
            TabIndex        =   259
            Top             =   3480
            Width           =   1212
         End
         Begin VB.Label LblTipoC9 
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionamiento del Concentrador :"
            Height          =   255
            Left            =   360
            TabIndex        =   148
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label lblProtocolo 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Protocolo"
            Height          =   255
            Left            =   360
            TabIndex        =   146
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Ruta del directorio del concentrador"
            Height          =   255
            Left            =   360
            TabIndex        =   90
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label14 
            Caption         =   "Bits de Datos"
            Height          =   252
            Left            =   120
            TabIndex        =   89
            Top             =   1800
            Width           =   1212
         End
         Begin VB.Label Label15 
            Caption         =   "Puerto "
            Height          =   252
            Left            =   120
            TabIndex        =   88
            Top             =   2160
            Width           =   972
         End
         Begin VB.Label Label16 
            Caption         =   "Paridad"
            Height          =   252
            Left            =   3000
            TabIndex        =   87
            Top             =   2160
            Width           =   972
         End
         Begin VB.Label Label17 
            Caption         =   "Baudios"
            Height          =   252
            Left            =   3000
            TabIndex        =   86
            Top             =   1800
            Width           =   852
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "Recepción Contínua"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   -74880
         TabIndex        =   73
         Top             =   240
         Width           =   7572
         Begin VB.CheckBox ChkEnvioContinuo 
            Appearance      =   0  'Flat
            Caption         =   "Envío de modificaciones en el mismo intervalo"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   239
            Top             =   720
            Width           =   5532
         End
         Begin VB.CheckBox ChkContinua 
            Appearance      =   0  'Flat
            Caption         =   "Recepción Contínua de Tiquets"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   75
            Top             =   240
            Width           =   2772
         End
         Begin VB.ComboBox TxtIntervalo 
            Height          =   315
            ItemData        =   "FrmConfiguracion2.frx":172F
            Left            =   5880
            List            =   "FrmConfiguracion2.frx":175A
            TabIndex        =   74
            Text            =   "TxtIntervalo"
            Top             =   240
            Width           =   1452
         End
         Begin VB.Label LblIntervalo 
            Caption         =   "Intervalo de recogida "
            Height          =   255
            Left            =   3720
            TabIndex        =   76
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Euroscale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   -74880
         TabIndex        =   56
         Top             =   3840
         Visible         =   0   'False
         Width           =   3972
         Begin VB.CheckBox CheckAHO 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   263
            Top             =   2280
            Width           =   255
         End
         Begin VB.CheckBox ChkNoTiquet 
            Appearance      =   0  'Flat
            Caption         =   "Chequeo rígido de Tiquets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   2160
            Width           =   3495
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   2880
            TabIndex        =   66
            Top             =   600
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Fichas de Vacuno"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   3252
         End
         Begin VB.CheckBox ChkBorrarVacuno 
            Appearance      =   0  'Flat
            Caption         =   "Borrar Totales de Fichas de Vacuno"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   64
            Top             =   1920
            Width           =   2892
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   2880
            TabIndex        =   63
            Top             =   840
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   2880
            TabIndex        =   62
            Top             =   1080
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   3
            Left            =   2880
            TabIndex        =   61
            Top             =   1320
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   4
            Left            =   3480
            TabIndex        =   60
            Top             =   600
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   5
            Left            =   3480
            TabIndex        =   59
            Top             =   840
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   6
            Left            =   3480
            TabIndex        =   58
            Top             =   1080
            Width           =   252
         End
         Begin VB.CheckBox ChkTotalesGamaAlta 
            Appearance      =   0  'Flat
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   7
            Left            =   3480
            TabIndex        =   57
            Top             =   1320
            Width           =   252
         End
         Begin VB.Label LabelAHO 
            Caption         =   "Eliminar Vendedores tras Borrado de totales."
            Height          =   255
            Left            =   480
            TabIndex        =   264
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Venta"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Super"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Envasado"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   70
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label24 
            Caption         =   "Recibir Totales Autoservicio"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label25 
            Caption         =   "Red"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   68
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label25 
            Caption         =   "Local"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   67
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         Caption         =   "SC10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   -70800
         TabIndex        =   52
         Top             =   3960
         Visible         =   0   'False
         Width           =   3492
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Autoservicio"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   283
            Top             =   1440
            Width           =   3252
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Envasado"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   282
            Top             =   1080
            Width           =   3252
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de Venta"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   3252
         End
         Begin VB.CheckBox ChkTotalesSC10 
            Appearance      =   0  'Flat
            Caption         =   "Recibir Totales de modo Super"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   54
            Top             =   720
            Width           =   3252
         End
         Begin VB.CheckBox ChkBorrarTiquet 
            Appearance      =   0  'Flat
            Caption         =   "Borrar TICKET:DAT"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   240
            TabIndex        =   53
            Top             =   1800
            Width           =   3252
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         Caption         =   "Opciones de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3855
         Left            =   3360
         TabIndex        =   49
         Top             =   2520
         Visible         =   0   'False
         Width           =   4452
         Begin VB.CheckBox ChkSubSec 
            Appearance      =   0  'Flat
            Caption         =   "Trabajo con Subsecciones Fijas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   255
            Top             =   2760
            Width           =   4092
         End
         Begin VB.TextBox TxtRetries 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   250
            Text            =   "TxtRetries"
            Top             =   3240
            Width           =   852
         End
         Begin VB.CheckBox ChkFamilias 
            Appearance      =   0  'Flat
            Caption         =   "Trabajo con Familias Fijas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   249
            Top             =   2520
            Width           =   4092
         End
         Begin VB.CheckBox ChkEstiloInterfaz 
            Appearance      =   0  'Flat
            Caption         =   "Interfaz Estilo Barra de Tareas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   0
            TabIndex        =   236
            Top             =   3480
            Visible         =   0   'False
            Width           =   3732
         End
         Begin VB.CheckBox ChkModem 
            Appearance      =   0  'Flat
            Caption         =   "Comunicar a través de Modem"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   120
            TabIndex        =   199
            Top             =   1200
            Width           =   2892
         End
         Begin VB.CheckBox ChkPasarela 
            Appearance      =   0  'Flat
            Caption         =   "Comunicar a través de Pasarela Remota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   120
            TabIndex        =   198
            Top             =   1440
            Width           =   2652
         End
         Begin VB.TextBox TxtPasarelaPort 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   197
            Top             =   2160
            Width           =   732
         End
         Begin VB.CheckBox ChkPlu0 
            Appearance      =   0  'Flat
            Caption         =   "Rechazar artículos con precio cero"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   120
            TabIndex        =   167
            Top             =   720
            Width           =   3252
         End
         Begin VB.TextBox TxtPuertoPantallas 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3360
            TabIndex        =   117
            Top             =   480
            Width           =   972
         End
         Begin VB.CheckBox ChkPantallas 
            Appearance      =   0  'Flat
            Caption         =   "Pantallas Remotas.     Puerto :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   116
            Top             =   480
            Width           =   3012
         End
         Begin VB.CheckBox ChkGruposConservacion 
            Appearance      =   0  'Flat
            Caption         =   "Hacer uso de Grupos de Conservación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   4092
         End
         Begin VB.CheckBox ChkGedit 
            Appearance      =   0  'Flat
            Caption         =   "Emplear el Editor de Etiquetas Gedit"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Visible         =   0   'False
            Width           =   4215
         End
         Begin MSMask.MaskEdBox TxtPasarelaHost 
            Height          =   255
            Left            =   1320
            TabIndex        =   200
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   15
            Mask            =   "###.###.###.###"
            PromptChar      =   "_"
         End
         Begin VB.Line Line2 
            X1              =   2640
            X2              =   120
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label8 
            Caption         =   "Reintentos de comunicación :"
            Height          =   495
            Left            =   0
            TabIndex        =   251
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Host :"
            Height          =   255
            Left            =   360
            TabIndex        =   202
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "Puerto"
            Height          =   255
            Left            =   360
            TabIndex        =   201
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         Caption         =   "Moneda Nacional (No Euros)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   732
         Left            =   -74880
         TabIndex        =   44
         Top             =   240
         Width           =   7812
         Begin VB.TextBox TxtDecimales 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   2280
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   324
            Width           =   492
         End
         Begin VB.TextBox TxtSimbolo 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   5040
            TabIndex        =   45
            Top             =   324
            Width           =   1332
         End
         Begin VB.Label Label7 
            Caption         =   "Número de Decimales :"
            Height          =   252
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   1932
         End
         Begin VB.Label Label10 
            Caption         =   "Símbolo Monetario :"
            Height          =   252
            Left            =   3000
            TabIndex        =   47
            Top             =   360
            Width           =   1812
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         Caption         =   "Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1812
         Left            =   -74880
         TabIndex        =   41
         Top             =   1080
         Width           =   7812
         Begin VB.ComboBox CmbMoneda 
            Height          =   288
            ItemData        =   "FrmConfiguracion2.frx":17DA
            Left            =   3960
            List            =   "FrmConfiguracion2.frx":17E4
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   360
            Width           =   1692
         End
         Begin VB.Label Label9 
            Caption         =   "Moneda a Emplear en las comunicaciones "
            Height          =   252
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   3492
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         Caption         =   "Nombre de la Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   8055
         Begin VB.TextBox TxtEmpresa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   295
            Top             =   240
            Width           =   8055
         End
      End
      Begin VB.CheckBox ChkStoreFlow 
         Appearance      =   0  'Flat
         Caption         =   "Enlace Mediante Socket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   -74880
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   6252
      End
      Begin VB.Frame Frame19 
         Appearance      =   0  'Flat
         Caption         =   "Trabajo Cotidiano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   852
         Left            =   -74880
         TabIndex        =   16
         Top             =   1680
         Width           =   7812
         Begin VB.CheckBox ChkIconificar 
            Appearance      =   0  'Flat
            Caption         =   "Iconificar el programa en la barra de Estado al minimizar o doble click sobre la interfaz"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   7572
         End
         Begin VB.CheckBox ChkIconificarInicio 
            Appearance      =   0  'Flat
            Caption         =   "Arrancar el programa en modo Iconificado"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   7452
         End
      End
      Begin VB.CheckBox ChkTaquion 
         Appearance      =   0  'Flat
         Caption         =   "Enlace ASCII"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   -74880
         TabIndex        =   15
         Top             =   1920
         Width           =   2412
      End
      Begin VB.Frame Frame21 
         Appearance      =   0  'Flat
         Caption         =   "Opciones"
         ForeColor       =   &H80000008&
         Height          =   3612
         Left            =   -74880
         TabIndex        =   13
         Top             =   2400
         Width           =   7812
         Begin VB.TextBox TxtAllTicket 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   140
            Top             =   2880
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "Un solo fichero para todos"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   139
            Top             =   2400
            Width           =   3972
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            Caption         =   "Un fichero por cada tiquet"
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   0
            Left            =   120
            TabIndex        =   138
            Top             =   2040
            Width           =   4332
         End
         Begin VB.CheckBox ChkTqDirecto 
            Appearance      =   0  'Flat
            Caption         =   "Enviar Modificaciones conforme se recogen los ficheros"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   137
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox TxtNombreTicket 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   135
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CheckBox ChkTaquionTiquets 
            Appearance      =   0  'Flat
            Caption         =   "Crear ficheros de tiquets conforme se recogen de las balanzas."
            ForeColor       =   &H80000008&
            Height          =   732
            Left            =   120
            TabIndex        =   105
            Top             =   1320
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   4800
            TabIndex        =   136
            Top             =   2880
            Width           =   612
         End
         Begin VB.Label Label38 
            Caption         =   "Nombre del Fichero de Tiquets"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   134
            Top             =   2880
            Width           =   2412
         End
         Begin VB.Label Label30 
            Caption         =   "Fichero de Importación"
            Height          =   255
            Left            =   120
            TabIndex        =   104
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
         Top             =   360
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   -73560
         MaxLength       =   255
         TabIndex        =   11
         Top             =   720
         Width           =   3732
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   10
         Top             =   1200
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   -73560
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1560
         Width           =   3732
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   8
         Top             =   2040
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   -73560
         MaxLength       =   255
         TabIndex        =   7
         Top             =   2400
         Width           =   3732
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2880
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   -73560
         MaxLength       =   255
         TabIndex        =   5
         Top             =   3240
         Width           =   3732
      End
      Begin VB.TextBox TxtComando 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   4
         Left            =   -73560
         MaxLength       =   25
         TabIndex        =   4
         Top             =   3840
         Width           =   1572
      End
      Begin VB.TextBox TxtEjecutar 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   4
         Left            =   -73560
         MaxLength       =   255
         TabIndex        =   3
         Top             =   4200
         Width           =   3732
      End
      Begin VB.Frame Frame18 
         Appearance      =   0  'Flat
         Caption         =   "Proceso Automático"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1332
         Left            =   -74880
         TabIndex        =   19
         Top             =   240
         Width           =   7812
         Begin VB.TextBox txtCI_MINUTO 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4320
            TabIndex        =   23
            Top             =   960
            Width           =   732
         End
         Begin VB.TextBox TxtCI_HORA 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4320
            TabIndex        =   22
            Top             =   600
            Width           =   732
         End
         Begin VB.TextBox TxtRE_HORA 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1440
            TabIndex        =   21
            Top             =   600
            Width           =   732
         End
         Begin VB.TextBox TxtRE_MINUTO 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1440
            TabIndex        =   20
            Top             =   960
            Width           =   732
         End
         Begin VB.Label Label23 
            Caption         =   "Dejando los valores en blanco no provoca Finalización del programa y/o recogida de Grandes Totales."
            Height          =   972
            Left            =   5160
            TabIndex        =   30
            Top             =   240
            Width           =   2412
         End
         Begin VB.Label Label22 
            Caption         =   "Minuto (0 -59)"
            Height          =   252
            Index           =   0
            Left            =   3000
            TabIndex        =   29
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label21 
            Caption         =   "Hora (0 -23)"
            Height          =   252
            Left            =   3000
            TabIndex        =   28
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label26 
            Caption         =   "Cerrar el Programa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   3000
            TabIndex        =   27
            Top             =   240
            Width           =   2052
         End
         Begin VB.Label Label27 
            Caption         =   "Hora (0 -23)"
            Height          =   252
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label22 
            Caption         =   "Minuto (0 -59)"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Label26 
            Caption         =   "Recoger/Borrar totales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2772
         End
      End
      Begin VB.Label Label32 
         Caption         =   "Borrar totales acumulados más antiguos de ..."
         Height          =   252
         Left            =   -74880
         TabIndex        =   254
         Top             =   3840
         Width           =   4092
      End
      Begin VB.Label Label33 
         Caption         =   "días"
         Height          =   252
         Index           =   0
         Left            =   -69480
         TabIndex        =   253
         Top             =   3840
         Width           =   852
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Configuración General de la Tienda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   166
         Top             =   240
         Width           =   7575
      End
      Begin VB.Image Image1 
         Height          =   315
         Index           =   9
         Left            =   5160
         Picture         =   "FrmConfiguracion2.frx":17F8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   8
         Left            =   4680
         Picture         =   "FrmConfiguracion2.frx":21A6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   7
         Left            =   4200
         Picture         =   "FrmConfiguracion2.frx":25E8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   6
         Left            =   3720
         Picture         =   "FrmConfiguracion2.frx":2A2A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   5
         Left            =   3240
         Picture         =   "FrmConfiguracion2.frx":2E6C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   4
         Left            =   2760
         Picture         =   "FrmConfiguracion2.frx":32AE
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   252
         Index           =   4
         Left            =   -74760
         TabIndex        =   103
         Top             =   4200
         Width           =   1092
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   840
         Picture         =   "FrmConfiguracion2.frx":36F0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   1320
         Picture         =   "FrmConfiguracion2.frx":3B32
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   1800
         Picture         =   "FrmConfiguracion2.frx":3F74
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   348
         Index           =   3
         Left            =   2280
         Picture         =   "FrmConfiguracion2.frx":483E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   102
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   101
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   252
         Index           =   1
         Left            =   -74760
         TabIndex        =   100
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   252
         Index           =   2
         Left            =   -74760
         TabIndex        =   99
         Top             =   2040
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   252
         Index           =   3
         Left            =   -74760
         TabIndex        =   98
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "Nombre : "
         Height          =   252
         Index           =   4
         Left            =   -74760
         TabIndex        =   97
         Top             =   3840
         Width           =   1092
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   252
         Index           =   1
         Left            =   -74760
         TabIndex        =   96
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   252
         Index           =   2
         Left            =   -74760
         TabIndex        =   95
         Top             =   2400
         Width           =   1092
      End
      Begin VB.Label Label29 
         Caption         =   "Ejecutar :"
         Height          =   252
         Index           =   3
         Left            =   -74760
         TabIndex        =   94
         Top             =   3240
         Width           =   1092
      End
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   10
      Left            =   7560
      Picture         =   "FrmConfiguracion2.frx":4C80
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FrmConfiguracion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private enLoad As Boolean

Private QLMONEDA As Boolean
Private Type TRS232
    puerto As Integer
    baudios As Integer
    paridad As String
    bits As Integer
End Type

Private Sub CheckMarteAlfa_Click()
    If (ChkEcoLabel.Value = vbChecked) Then 'And (CheckMarteNew.Value = vbChecked) Then
        If CheckMarteAlfa.Value = vbChecked Then
            lMarteAlfa = True
            l49T = True
            Check49T.Value = vbChecked
            lMarteNew = True
            CheckMarteNew.Value = vbChecked
        Else
            lMarteAlfa = False
            l49T = False
            l66T = False
            Check49T.Value = 0
            Check66T = 0
        End If
    
    End If
End Sub

Private Sub Check49T_Click()
    If (ChkEcoLabel.Value = vbChecked) And (CheckMarteAlfa.Value = vbChecked) Then
        If Check49T.Value = vbChecked Then
            l49T = True
            l66T = False
            Check66T.Value = vbUnchecked
        Else
            l49T = False
            l66T = True
            Check66T.Value = vbChecked
        End If
    Else
        Check49T.Value = vbUnchecked
    End If
End Sub

Private Sub Check66T_Click()
    If (ChkEcoLabel.Value = vbChecked) And (CheckMarteAlfa.Value = vbChecked) Then
        If Check66T.Value = vbChecked Then
            l66T = True
            l49T = False
            Check49T.Value = vbUnchecked
        Else
            l66T = False
            l49T = True
            Check49T.Value = vbChecked
        End If
    Else
        Check66T.Value = vbUnchecked
    End If
End Sub

Private Sub CheckAHO_Click()
    Dim nf As Integer
    If CheckAHO.Value = vbUnchecked Then
        If Dir(App.Path & "\killvnd.aho") <> "" Then
            Kill App.Path & "\killvnd.aho"
        End If
    Else
        If Dir(App.Path & "\killvnd.aho") = "" Then
            nf = FreeFile()
            Open App.Path & "\killvnd.aho" For Output As #nf
            Close #nf
        End If
    End If
End Sub


Private Sub CheckMarteNew_Click()
'Dim n As Integer
    If (ChkEcoLabel.Value = vbChecked) Then
        If CheckMarteNew.Value Then
            
            'n = FreeFile()
            'Open App.Path & "\marte.new" For Output As #n
            'Close #n
            lMarteNew = True
        Else
            'If Dir(App.Path & "\marte.new") <> "" Then
            '    Kill App.Path & "\marte.new"
            'End If
            lMarteNew = False
            lMarteAlfa = False
            l49T = False
            l66T = False
            CheckMarteAlfa.Value = 0
            Check49T.Value = 0
            Check66T.Value = 0
            
        End If
    End If
End Sub


Private Sub ChkBorrarTotalGA_Click()
    If ChkBorrarTotalGA.Value = False Then
        GA_Reinit = False
        ChkGaReinit.Value = False
        ChkGaReinit.Enabled = False
        
        SC10_BorrarTotal = False
        ChkBorrarTotalSC10.Value = False
        ChkBorrarTotalSC10.Enabled = False
        
        SC10_Borrar_Eti_dat = False
        ChkBorrarEti.Value = False
        ChkBorrarEti.Enabled = False
        
        ChkParked.Value = False
        ChkParked.Enabled = False
        If Dir(App.Path & "\nopark") <> "" Then
            Kill App.Path & "\nopark"
        End If
    Else
        ChkGaReinit.Enabled = True
        ChkBorrarTotalSC10.Enabled = True
        ChkBorrarEti.Enabled = True
        ChkParked.Enabled = True
    End If
End Sub

Private Sub ChkL000_Click()
    If ChkL000.Value = vbChecked Then
        ChkFichaVacuno.Value = vbChecked
        Option1(1).Value = True
        Chk20Lineas.Value = vbChecked
    End If
End Sub

Private Sub ChkNoLabel_Click()
Dim nf As Integer
    If ChkNoLabel.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\nolabel.txt" For Output As #nf
        Close #nf
        ChkNoPack.Value = vbChecked
        ChkNoPre.Value = vbChecked
        ChkNoPack.Enabled = False
        ChkNoPre.Enabled = False
        ChkAcuSS.Value = vbUnchecked
        ChkAcuET.Value = vbUnchecked
        ChkAcuSS.Enabled = False
        ChkAcuET.Enabled = False
    Else
        If Dir(App.Path & "\nolabel.txt") <> "" Then
            Kill App.Path & "\nolabel.txt"
        End If
        ChkNoPack.Enabled = True
        ChkNoPre.Enabled = True
    End If
End Sub

Private Sub ChkNopack_Click()
Dim nf As Integer
    If ChkNoPack.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\nopack.txt" For Output As #nf
        Close #nf
        ChkAcuSS.Value = vbUnchecked
        ChkAcuSS.Enabled = False
    Else
        If Dir(App.Path & "\nopack.txt") <> "" Then
            Kill App.Path & "\nopack.txt"
        End If
        ChkAcuSS.Enabled = True
    End If

End Sub
Private Sub Chknopre_Click()
Dim nf As Integer
    If ChkNoPre.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\nopre.txt" For Output As #nf
        Close #nf
        ChkAcuET.Value = vbUnchecked
        ChkAcuET.Enabled = False
    Else
        If Dir(App.Path & "\nopre.txt") <> "" Then
            Kill App.Path & "\nopre.txt"
        End If
        ChkAcuET.Enabled = True
    End If

End Sub

Private Sub Chkacuss_Click()
Dim nf As Integer
    If ChkAcuSS.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\acumulass" For Output As #nf
        Close #nf
        lAcumulaSS = True
    Else
        If Dir(App.Path & "\acumulass") <> "" Then
            Kill App.Path & "\acumulass"
        End If
        lAcumulaSS = False
    End If

End Sub

Private Sub Chkacuet_Click()
Dim nf As Integer
    If ChkAcuET.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\acumulaet" For Output As #nf
        Close #nf
        lAcumulaET = True
    Else
        If Dir(App.Path & "\acumulaet") <> "" Then
            Kill App.Path & "\acumulaet"
        End If
        lAcumulaET = False
    End If

End Sub


Private Sub ChkParked_Click()
Dim nf As Integer
If ChkParked.Value = vbChecked Then
    nf = FreeFile()
    Open App.Path & "\nopark" For Output As #nf
    Close #nf
Else
    If Dir(App.Path & "\nopark") <> "" Then
        Kill App.Path & "\nopark"
    End If
End If
End Sub

Private Sub ChkPosFam_Click()
Dim nf As Integer
    If ChkPosFam.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\test.tst" For Output As #nf
        Close #nf
        ChkAsignTec.Enabled = True
    Else
        If Dir(App.Path & "\test.tst") <> "" Then
            Kill App.Path & "\test.tst"
        End If
        ChkAsignTec.Value = False
        ChkAsignTec.Enabled = False
    End If
    'atencionfamilia
    'If Dir(App.Path & "\upper.txt") <> "" And Dir(App.Path & "\test.tst") <> "" Then
    '    lUpperNW = True
    'Else
    '    lUpperNW = False
    'End If
    
End Sub
Private Sub disableOrden()
    Option5.Value = False
    Option6.Value = False
    Option7.Value = False
    Option8.Value = False
    Option5.Enabled = False
    Option6.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
    
    If Dir(App.Path & "\asign1.tst") <> "" Then
        Kill App.Path & "\asign1.tst"
    End If
    If Dir(App.Path & "\asign2.tst") <> "" Then
        Kill App.Path & "\asign2.tst"
    End If
    If Dir(App.Path & "\asign3.tst") <> "" Then
        Kill App.Path & "\asign3.tst"
    End If
    If Dir(App.Path & "\asign4.tst") <> "" Then
        Kill App.Path & "\asign4.tst"
    End If
    
End Sub

Private Sub ChkSHname_Click()
Dim nf As Integer
    If ChkSHname.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\shname.txt" For Output As #nf
        Print #nf, "shname"
        Close #nf
    Else
        If Dir(App.Path & "\shname.txt") <> "" Then
            Kill App.Path & "\shname.txt"
        End If
    End If
End Sub

Private Sub Option5_Click()
Dim nf As Integer
    If Option5.Value = True Then
        nf = FreeFile()
        Open App.Path & "\asign1.tst" For Output As #nf
        Close #nf
        Option6.Value = False
        Option7.Value = False
        Option8.Value = False
        If Dir(App.Path & "\asign2.tst") <> "" Then
            Kill App.Path & "\asign2.tst"
        End If
        If Dir(App.Path & "\asign3.tst") <> "" Then
            Kill App.Path & "\asign3.tst"
        End If
        If Dir(App.Path & "\asign4.tst") <> "" Then
            Kill App.Path & "\asign4.tst"
        End If
        If enLoad = False Then
            If Dir(Base_General) <> "" Then
                AsignTec
            End If
        End If
    Else
        If Dir(App.Path & "\asign1.tst") <> "" Then
            Kill App.Path & "\asign1.tst"
        End If
    End If
End Sub

Private Sub Option6_Click()
Dim nf As Integer
    If Option6.Value = True Then
        nf = FreeFile()
        Open App.Path & "\asign2.tst" For Output As #nf
        Close #nf
        Option5.Value = False
        Option7.Value = False
        Option8.Value = False
        If Dir(App.Path & "\asign1.tst") <> "" Then
            Kill App.Path & "\asign1.tst"
        End If
        If Dir(App.Path & "\asign3.tst") <> "" Then
            Kill App.Path & "\asign3.tst"
        End If
        If Dir(App.Path & "\asign4.tst") <> "" Then
            Kill App.Path & "\asign4.tst"
        End If
        
        If enLoad = False Then
            If Dir(Base_General) <> "" Then
                AsignTec
            End If
        End If
    Else
        If Dir(App.Path & "\asign2.tst") <> "" Then
            Kill App.Path & "\asign2.tst"
        End If
    End If
End Sub

Private Sub Option7_Click()
Dim nf As Integer
    If Option7.Value = True Then
        nf = FreeFile()
        Open App.Path & "\asign3.tst" For Output As #nf
        Close #nf
        Option5.Value = False
        Option6.Value = False
        Option8.Value = False
        If Dir(App.Path & "\asign1.tst") <> "" Then
            Kill App.Path & "\asign1.tst"
        End If
        If Dir(App.Path & "\asign2.tst") <> "" Then
            Kill App.Path & "\asign2.tst"
        End If
        If Dir(App.Path & "\asign4.tst") <> "" Then
            Kill App.Path & "\asign4.tst"
        End If
        
        If enLoad = False Then
            If Dir(Base_General) <> "" Then
                AsignTec
            End If
        End If
    Else
        If Dir(App.Path & "\asign3.tst") <> "" Then
            Kill App.Path & "\asign3.tst"
        End If
    End If
End Sub

Private Sub Option8_Click()
Dim nf As Integer
    If Option8.Value = True Then
        nf = FreeFile()
        Open App.Path & "\asign4.tst" For Output As #nf
        Close #nf
        Option5.Value = False
        Option6.Value = False
        Option7.Value = False
        If Dir(App.Path & "\asign1.tst") <> "" Then
            Kill App.Path & "\asign1.tst"
        End If
        If Dir(App.Path & "\asign2.tst") <> "" Then
            Kill App.Path & "\asign2.tst"
        End If
        If Dir(App.Path & "\asign3.tst") <> "" Then
            Kill App.Path & "\asign3.tst"
        End If
        
        If enLoad = False Then
            If Dir(Base_General) <> "" Then
                AsignTec
            End If
        End If
    Else
        If Dir(App.Path & "\asign4.tst") <> "" Then
            Kill App.Path & "\asign4.tst"
        End If
    End If
End Sub


Private Sub chkAsignTec_Click()
Dim nf As Integer


    If ChkAsignTec.Value = vbChecked Then
        nf = FreeFile()
        Open App.Path & "\asign.tst" For Output As #nf
        Close #nf
        'nf = FreeFile()
        'Open App.Path & "\asign1.tst" For Output As #nf
        'Close #nf
        Option5.Enabled = True
        Option6.Enabled = True
        Option7.Enabled = True
        Option8.Enabled = True
        'Option5.Value = True
        If Dir(App.Path & "\asign1.tst") = "" And Dir(App.Path & "\asign2.tst") = "" And Dir(App.Path & "\asign3.tst") = "" And Dir(App.Path & "\asign4.tst") = "" Then
            nf = FreeFile()
            Open App.Path & "\asign1.tst" For Output As #nf
            Close #nf
            Option5.Value = True
        End If
        If enLoad = False Then
            If Dir(Base_General) <> "" Then
                AsignTec
            End If
        End If
    Else
        If Dir(App.Path & "\asign.tst") <> "" Then
            Kill App.Path & "\asign.tst"
        End If
        Call disableOrden
    End If
    'If Dir(App.Path & "\upper.txt") <> "" And Dir(App.Path & "\test.tst") <> "" Then
    '    lUpperNW = True
    'Else
    '    lUpperNW = False
    'End If


End Sub


Private Sub CmbProtocolo_DropDown()
    If OptSC10Comu(2).Value = True Then
        OptSC10Comu(2).Value = False
        TxtIPSC10.Enabled = False
        TxtTCPSC10.Enabled = False
        LblSC10IP.Enabled = False
        LblSC10TCP.Enabled = False
        SC10TCPIP = False
    End If

End Sub

'\\\\\\\
'///////
Public Sub Form_Load()
Dim Buffer As String
Dim Contador As Long
Dim PSerie As TRS232
Dim bucle As Integer
Dim MasterFile As Integer

    enLoad = True
    
    CheckAHO.Enabled = False
    LabelAHO.Enabled = False
    CheckAHO.Visible = False
    LabelAHO.Visible = False
    If AhorraMas Then
        CheckAHO.Enabled = True
        LabelAHO.Enabled = True
        CheckAHO.Visible = True
        LabelAHO.Visible = True
        If Dir(App.Path & "\killvnd.aho") <> "" Then
            CheckAHO.Value = vbChecked
        Else
            CheckAHO.Value = vbUnchecked
        End If
    End If
    If Dir(App.Path & "\test.tst") <> "" Then
        ChkPosFam.Value = vbChecked
    Else
        ChkPosFam.Value = vbUnchecked
    End If
    If Dir(App.Path & "\shname.txt") <> "" Then
        ChkSHname.Value = vbChecked
    Else
        ChkSHname.Value = vbUnchecked
    End If
    
    If Dir(App.Path & "\asign.tst") <> "" Then
        ChkAsignTec.Value = vbChecked
    Else
        ChkAsignTec.Value = vbUnchecked
    End If
    If Dir(App.Path & "\asign1.tst") <> "" Then
        Option5.Value = True
    Else
        Option5.Value = False
    End If
    If Dir(App.Path & "\asign2.tst") <> "" Then
        Option6.Value = True
    Else
        Option6.Value = False
    End If
    If Dir(App.Path & "\asign3.tst") <> "" Then
        Option7.Value = True
    Else
        Option7.Value = False
    End If
    If Dir(App.Path & "\asign4.tst") <> "" Then
        Option8.Value = True
    Else
        Option8.Value = False
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    If Dir(App.Path & "\nolabel.txt") <> "" Then
        ChkNoLabel.Value = vbChecked
    Else
        ChkNoLabel.Value = vbUnchecked
    End If
    If Dir(App.Path & "\nopack.txt") <> "" Then
        ChkNoPack.Value = vbChecked
    Else
        ChkNoPack.Value = vbUnchecked
    End If
    If Dir(App.Path & "\nopre.txt") <> "" Then
        ChkNoPre.Value = vbChecked
    Else
        ChkNoPre.Value = vbUnchecked
    End If
    If Dir(App.Path & "\acumulass") <> "" Then
        ChkAcuSS.Value = vbChecked
    Else
        ChkAcuSS.Value = vbUnchecked
    End If
    If Dir(App.Path & "\acumulaet") <> "" Then
        ChkAcuET.Value = vbChecked
    Else
        ChkAcuET.Value = vbUnchecked
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    If Dir(App.Path & "\nopark") <> "" Then
        ChkParked.Value = vbChecked
    Else
        ChkParked.Value = vbUnchecked
    End If
    
    '1.0.4
    Option2(0).Visible = False
    Option2(1).Visible = False
    Label38(1).Caption = ".ddm"
    '''''''
    TxtTCPSC10.TexT = SC10tcp
    TxtIPSC10.TexT = SC10IP
    TxtRetries.TexT = CommRetries
    ChkFamilias.Value = BoolToCheck(FamiliasFijas)
    ChkSubSec.Value = BoolToCheck(SubSeccionesFijas)
    Chk4Lineas.Value = BoolToCheck(SC104Lineas)
    If MultiEuroscale Then ChkMultiEuroscale.Value = vbChecked
    If EnvioContinuo Then ChkEnvioContinuo.Value = vbChecked
    ChkEstiloInterfaz.Value = vbUnchecked
    ChkEcoSoloUna.Value = vbUnchecked
    ChkPedidoIV4.Value = vbUnchecked
    If PedidoIV4 Then ChkPedidoIV4.Value = vbChecked
    If EcoSoloUna Then ChkEcoSoloUna.Value = vbChecked
    If EstiloInterfaz = 1 Then ChkEstiloInterfaz.Value = vbChecked
    For bucle = 0 To 11
        Tabla.TabCaption(bucle) = ""
    Next bucle
    If UsaEcoPrint Then
        ChkEcoPrint.Value = vbChecked
        ChkEcoPrint_Click
    Else
        ChkEcoPrint.Value = vbUnchecked
        ChkEcoPrint_Click
    End If
    OptIV4(TipoEcoPrint).Value = True
    CmbPortPrint.ListIndex = PuertoEcoPrint - 1
    If UsaECO1 Then
        ChkEcoNet.Value = vbChecked
        OptECO(0).Value = True
        If TipoV8 Then OptECO(1).Value = True
    End If
    
    
    If UsaECO2 Then
        ChkEcoLabel.Value = vbChecked
        If leco9600 Then option3.Value = vbChecked
        If leco19200 Then option4.Value = vbChecked
        If lMarteNew Then CheckMarteNew.Value = vbChecked
        If lMarteAlfa Then CheckMarteAlfa.Value = vbChecked
        If l49T Then Check49T.Value = vbChecked
        If l66T Then Check66T.Value = vbChecked
    Else
    ' inhabilitar ...
        option3.Enabled = False
        option4.Enabled = False
        CheckMarteNew.Enabled = False
        CheckMarteAlfa.Enabled = False
        Check49T.Enabled = False
        Check66T.Enabled = False
        If leco9600 Then option3.Value = vbUnchecked
        If leco19200 Then option4.Value = vbUnchecked
        If lMarteNew Then CheckMarteNew.Value = vbUnchecked
        If lMarteAlfa Then CheckMarteAlfa.Value = vbUnchecked
        If l49T Then Check49T.Value = vbUnchecked
        If l66T Then Check66T.Value = vbUnchecked
   
    '''''''''''''''''
    End If
    CmbPortV8.ListIndex = PuertoV8 - 1
    CmbPortLabel.ListIndex = PuertoLabel - 1
    TxtEquipoV8.TexT = MaquinaV8
    TxtEquipoLabel.TexT = MaquinaLabel
    ChkBorrarPedido.Value = BoolToCheck(BorrarPedido)
    ChkSolicitarPedido.Value = BoolToCheck(SolicitarPedido)
    ChkPedidoFinalDia.Value = BoolToCheck(BorrarPedidoFin)
    ChkBorrarInventario.Value = BoolToCheck(BorrarInventario)
    ChkSolicitarInventario.Value = BoolToCheck(SolicitarInventario)
    ChkInventarioFinal.Value = BoolToCheck(BorrarInventarioFin) 'BorrarInventario
    ChkSolicitarPedidoFin.Value = BoolToCheck(SolicitarPedidoFin)
    ChkSolicitarInventarioFin.Value = BoolToCheck(SolicitarInventarioFin)
    If RechazarPlu0 Then ChkPlu0.Value = vbChecked
    If HayMulti Then
        Label46.Visible = True
    Else
        Label46.Visible = False
    End If
    CambiarIdioma
    Antonio_Banderas
    'If Sabeco Then ChkPlu0.Enabled = False
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
'    ChkModem_Click
    If AhorraMas Then Frame15.Visible = True
    '************************
    ' modem
    '************************
    CmbModemPuerto.ListIndex = PuertoModem - 1
    CmbModemVelocidad.ListIndex = VelocidadModem - 1
    TxtModemTel.TexT = TelefonoModem
    TxtModemCadena(0).TexT = CadModem1
    TxtModemCadena(1).TexT = CadModem2
    TxtModemCadena(2).TexT = CadModem3
    '***************************
    ' Pasarela
    '***************************
    If Pasarela Then
        ChkPasarela.Value = vbChecked
    Else
        ChkPasarela.Value = vbUnchecked
    End If
    ChkPasarela_Click
    TxtPasarelaHost.TexT = PasarelaHost
    TxtPasarelaPort.TexT = PasarelaPuerto
    ChkClientes.Value = BoolToCheck(ClienteD)
    
    Me.chkLin2040.Value = BoolToCheck(versionG)
    
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
    Me.chkC9Cod.Value = var.C9Cod
    CmbTipoC9.TexT = CmbTipoC9.List(ModoC9)
    ChkGaReinit.Value = BoolToCheck(GA_Reinit)
    'If bunTicket Then
    '    Option2(1).Value = True
    'Else
    '    Option2(0).Value = True
    'End If
    TxtAllTicket.TexT = sunTicket
    If VacunoD = True Then
        Option1(1).Value = True
    Else
        Option1(0).Value = True
    End If
    '''1.1.1
    VacunoD = True
    ''''''''
    ChkTqDirecto.Value = BoolToCheck(TqDirecto)
    TxtNombreTicket.TexT = TqNombre
    If Exp_Diario = True Then
        Optdiario(0).Value = True
        Optdiario(1).Value = False
    Else
        Optdiario(1).Value = True
        Optdiario(0).Value = False
    End If
    If gedit = "1" Then
        ChkGedit.Value = vbChecked
    Else
        ChkGedit.Value = vbUnchecked
    End If
    If DiasBorrado > 0 Then
        Txtdias.TexT = DiasBorrado
    Else
        Txtdias.TexT = ""
    End If
    If BorrarBackup > 0 Then
        TxtBackupErase.TexT = BorrarBackup
    Else
        TxtBackupErase.TexT = ""
    End If
    'If Not UsaDual Then
    '    CmbMoneda.Enabled = False
    'End If
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
    TxtPuertoPantallas.TexT = PuertoPantallas
    '********************
    ' master/slave
    '********************
    If Dir(Miruta & "\master.cfg") <> "" Then
        'ChkMaestro.Value = vbChecked
        OptMaestro(0).Value = True
        MasterFile = FreeFile()
        Open Miruta & "\master.cfg" For Input As #MasterFile
        Line Input #MasterFile, Buffer
        TxtMasterLPort.TexT = Buffer
        Line Input #MasterFile, Buffer
        TxtMasterRHost.TexT = Buffer
        Line Input #MasterFile, Buffer
        TxtMasterPath.TexT = Buffer
        Close #MasterFile
    Else
        If Dir(Miruta & "\slave.cfg") <> "" Then
            'ChkMaestro.Value = vbChecked
            OptMaestro(1).Value = True
            MasterFile = FreeFile()
            Open Miruta & "\slave.cfg" For Input As #MasterFile
            Line Input #MasterFile, Buffer
            TxtMasterLPort.TexT = Buffer
            Line Input #MasterFile, Buffer
            TxtMasterRHost.TexT = Buffer
            Line Input #MasterFile, Buffer
            TxtMasterPath.TexT = Buffer
            Close #MasterFile
        Else
            ChkMaestro.Value = False
            Frame22.Enabled = False
        End If
    End If
    'If Dir(App.Path & "\eco9600.txt") <> "" Then
    If leco9600 Then
        option3.Value = 1
    Else
        option3.Value = 0
    End If
    'If Dir(App.Path & "\eco19200.txt") <> "" Then
    If leco19200 Then
        option4.Value = 1
    Else
        option4.Value = 0
    End If
    If lMarteNew Then '2.0.9 Dir(App.Path & "\marte.new") <> "" Then
        CheckMarteNew.Value = vbChecked
    Else
        CheckMarteNew.Value = vbUnchecked
    End If
    '*********************
    ' comandos libres
    '*********************
    For bucle = 0 To 4
        TxtComando(bucle).TexT = Trim(ComandoLibre(bucle))
        TxtEjecutar(bucle).TexT = Trim(EjecutarLibre(bucle))
    Next bucle
    '*********************
    ' /comandos libres
    '*********************
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
    Check_Store_Flow
    txtCI_HOST.TexT = CI_HOST
    txtCI_TCP.TexT = CI_TCP
    txtCI_UDP.TexT = CI_UDP
    '********
    ' automatización
    '*********
    If CI_HORA < 24 And CI_MINUTO < 60 Then
        TxtCI_HORA.TexT = CI_HORA
        txtCI_MINUTO.TexT = CI_MINUTO
    Else
        TxtCI_HORA.TexT = ""
        txtCI_MINUTO = ""
    End If
    If RE_HORA < 24 And RE_MINUTO < 60 Then
        TxtRE_HORA.TexT = RE_HORA
        TxtRE_MINUTO.TexT = RE_MINUTO
    Else
        TxtRE_HORA.TexT = ""
        TxtRE_MINUTO = ""
    End If
    ' ********
    ' General
    ' ********
    'TxtEmpresa.TexT = Empresa
    TxtEmpresa.TexT = Replace(Empresa, vbCrLf, "||")
    '******************
    ' Puerto Serie Euroscale
    '******************
    PSerie = ConvierteSerie(cngvelocidad)
    Select Case PSerie.baudios
        Case 19200
            CmbVelocidad.TexT = CmbVelocidad.List(4)
        Case 14400
            CmbVelocidad.TexT = CmbVelocidad.List(3)
        Case 9600
            CmbVelocidad.TexT = CmbVelocidad.List(2)
        Case 2400
            CmbVelocidad.TexT = CmbVelocidad.List(1)
        Case 1200
            CmbVelocidad.TexT = CmbVelocidad.List(0)
    End Select
    Select Case PSerie.bits
        Case 8
            CmbBits.TexT = CmbBits.List(0)
        Case 7
            CmbBits.TexT = CmbBits.List(1)
    End Select
    Select Case PSerie.paridad
        Case "E"
            CmbParidad.TexT = CmbParidad.List(0)
        Case "O"
            CmbParidad.TexT = CmbParidad.List(1)
        Case "N"
            CmbParidad.TexT = CmbParidad.List(2)
    End Select
    Select Case PSerie.puerto
        Case 1
            CmbCOM.TexT = CmbCOM.List(0)
        Case 2
            CmbCOM.TexT = CmbCOM.List(1)
        Case 3
            CmbCOM.TexT = CmbCOM.List(2)
        Case 4
            CmbCOM.TexT = CmbCOM.List(3)
        Case 5
            CmbCOM.TexT = CmbCOM.List(4)
    End Select
    '****************
    ' Puerto Serie SC10
    '****************
    PSerie = ConvierteSerie(SC10ParametrosSerie)
    Select Case PSerie.baudios
        Case 19200
            CMBSC10Baudios.TexT = CMBSC10Baudios.List(4)
        Case 14400
            CMBSC10Baudios.TexT = CMBSC10Baudios.List(3)
        Case 9600
            CMBSC10Baudios.TexT = CMBSC10Baudios.List(2)
        Case 2400
            CMBSC10Baudios.TexT = CMBSC10Baudios.List(1)
        Case 1200
            CMBSC10Baudios.TexT = CMBSC10Baudios.List(0)
    End Select
    Select Case PSerie.bits
        Case 8
            CmbSC10bits.TexT = CmbSC10bits.List(0)
        Case 7
            CmbSC10bits.TexT = CmbSC10bits.List(1)
    End Select
    Select Case PSerie.paridad
        Case "E"
            CmbSC10Paridad.TexT = CmbSC10Paridad.List(0)
        Case "O"
            CmbSC10Paridad.TexT = CmbSC10Paridad.List(1)
        Case "N"
            CmbSC10Paridad.TexT = CmbSC10Paridad.List(2)
    End Select
    Select Case PSerie.puerto
        Case 1
            CmbSC10Port.TexT = CmbSC10Port.List(0)
        Case 2
            CmbSC10Port.TexT = CmbSC10Port.List(1)
        Case 3
            CmbSC10Port.TexT = CmbSC10Port.List(2)
        Case 4
            CmbSC10Port.TexT = CmbSC10Port.List(3)
        Case 5
            CmbSC10Port.TexT = CmbSC10Port.List(4)
    End Select
    If TipoC9 Then
        If CmbSC10bits.TexT = "7" And CmbSC10Paridad.TexT = "E" And CMBSC10Baudios.TexT = "1200" Then
            CmbProtocolo.TexT = CmbProtocolo.List(0)
        Else
            If CmbSC10bits.TexT = "8" And CmbSC10Paridad.TexT = "N" And CMBSC10Baudios.TexT = "1200" Then
                CmbProtocolo.TexT = CmbProtocolo.List(1)
            Else
                If CmbSC10bits.TexT = "8" And CmbSC10Paridad.TexT = "N" And CMBSC10Baudios.TexT = "9600" Then
                    CmbProtocolo.TexT = CmbProtocolo.List(2)
                Else
                    If CmbSC10bits.TexT = "8" And CmbSC10Paridad.TexT = "E" And CMBSC10Baudios.TexT = "9600" Then
                        CmbProtocolo.TexT = CmbProtocolo.List(3)
                    Else
                        If CmbSC10bits.TexT = "8" And CmbSC10Paridad.TexT = "N" And CMBSC10Baudios.TexT = "19200" Then
                            CmbProtocolo.TexT = CmbProtocolo.List(4)
                        Else
                            CmbProtocolo.TexT = CmbProtocolo.List(0)
                            CmbSC10bits.TexT = "7"
                            CmbSC10Paridad.TexT = "E"
                            CMBSC10Baudios.TexT = "1200"
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
        ChkFichaVacuno.Enabled = True
        'Frame7.Enabled = True
        Frame14.Enabled = True
    Else
        ChkGamaAlta.Value = vbUnchecked
        ChkFichaVacuno.Enabled = False
        'Frame7.Enabled = False
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
        'Frame8.Enabled = True
    Else
        ChkSC10.Value = vbUnchecked
        'Frame8.Enabled = False
    End If
    ' ************
    ' Euroscale
    ' ************
    TxtPuerto.TexT = sPortBal
    '
    ' Desglosa configuración serie
    '
    If tcpip Then
        OptTCPIP.Value = True
        Label1.Enabled = True
        TxtPuerto.Enabled = True
        ChkMultiEuroscale.Enabled = True
        Frame4.Enabled = False
    Else
        'OPTRS232.Value = True
        'Frame4.Enabled = True
        'Label1.Enabled = False
        'TxtPuerto.Enabled = False
        'ChkMultiEuroscale.Enabled = False
    End If
    Chk20Lineas.Value = BoolToCheck(Usa20Lineas)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\Concentrador
    TxtPathSC10.TexT = pathSC10
    If SC10Rs232 = False Then
        OptSC10Comu(0).Value = True
        OptSC10Comu_Click 0
    Else
        If SC10TCPIP Then
            OptSC10Comu(2).Value = True
            OptSC10Comu_Click 0
        Else
            OptSC10Comu(1).Value = True
            OptSC10Comu_Click 0
        End If
    End If
    '/////
    '********
    ' Totales
    If descAuto Then
        ChkContinua.Value = vbChecked
        LblIntervalo.Enabled = True
        TxtIntervalo.Enabled = True
    Else
        ChkContinua.Value = vbUnchecked
        ChkEnvioContinuo.Enabled = False
        LblIntervalo.Enabled = False
        TxtIntervalo.Enabled = False
    End If
    If descautotime <> "" Then
        TxtIntervalo.TexT = descautotime
    Else
        TxtIntervalo.TexT = TxtIntervalo.List(5)
    End If
    '2.0.23
    If lVL000 = True Then
        FrmConfiguracion2.ChkL000.Value = vbChecked
    Else
        FrmConfiguracion2.ChkL000.Value = vbUnchecked
    End If
    '''''''
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
    ChkTotalesVacuno.Value = BoolToCheck(RecibirTotalVacuno)
    ChkBorrarTiquet.Value = BoolToCheck(SC10_Borrar_Ticket_dat)
    ChkBorrarVacuno.Value = BoolToCheck(BorrarTotalVacuno)
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
    TxtSimbolo.TexT = SimboloMonetario
    
    If IsNumeric(decimales) Then
        TxtDecimales.TexT = decimales
    Else
        decimales = 0
    End If
    QLMONEDA = True
    If UsaEuro Then
        CmbMoneda.TexT = CmbMoneda.List(1)
    Else
        CmbMoneda.TexT = CmbMoneda.List(0)
    End If
   QLMONEDA = False
    
    ' ***********************
    ' Especiales
    If UsaBackup Then
        ChkBackup.Value = vbChecked
    Else
        ChkBackup.Value = vbUnchecked
    End If
    TxtBackup.TexT = pathred
    TxtFamiliaVacuno.TexT = FamiliaEspecial
    TxtEtiquetaVacuno.TexT = EtiquetaEspecial
    TxtDbteclas = Trim(PathDBF)
    '***************+
    ' export
    CmbFormato(0).TexT = CmbFormato(0).List(Exp_Fecha)
    CmbFormato(1).TexT = CmbFormato(1).List(Exp_hora)
    TxtPositivo(0).TexT = Exp_Mas
    TxtPositivo(1).TexT = Exp_Menos
    ChkAbsoluto.Value = BoolToCheck(Exp_Absoluto)
    
    TxtPathExport.TexT = Exp_Path
    If StoreFlow Then OptSC10(1).Enabled = False
    '
    Opciones_Multi
    '
    Tabla.TabHeight = 1
    Tabla.TabVisible(0) = True
    Tabla.Tab = 0
    For bucle = 1 To 11
        Tabla.TabVisible(bucle) = False
    Next bucle
    
    If ChkBorrarTotalGA.Value = 0 Then
        ChkGaReinit.Enabled = False
        ChkBorrarTotalSC10.Enabled = False
        ChkBorrarEti.Enabled = False
        ChkParked.Enabled = False
    End If
    If lBelRos Then
        ChkPosFam.Enabled = False
        ChkAsignTec.Enabled = False
        Option5.Enabled = False
        Option6.Enabled = False
        Option7.Enabled = False
        Option8.Enabled = False
        
    End If
    enLoad = False
    
End Sub
Private Sub CambiarIdioma()
    Me.Caption = CargaCadena(258)
    ChkMultiEuroscale.Caption = CargaCadena(1209)
    Chk4Lineas.Caption = CargaCadena(1204)
    ChkEcoNet.Caption = CargaCadena(332) & " " & "V8"
    ChkEcoLabel.Caption = CargaCadena(332) & " " & "Marte IC-ILC"
    ChkEnvioContinuo.Caption = CargaCadena(1195)
    ChkEstiloInterfaz.Caption = CargaCadena(1181)
    ChkEcoSoloUna.Caption = CargaCadena(1184)
    'Frame29.Caption = CargaCadena(108) & "/" & CargaCadena(109)
    Label47.Caption = CargaCadena(616)
    Label50.Caption = CargaCadena(616)
    Label51.Caption = CargaCadena(616)
    Label48.Caption = CargaCadena(1119)
    Label49.Caption = CargaCadena(1119)
    ChkSolicitarPedido.Caption = CargaCadena(1098)
    ChkSolicitarInventario.Caption = CargaCadena(1100)
    ChkBorrarPedido.Caption = CargaCadena(1102)
    ChkBorrarInventario.Caption = CargaCadena(1104)
    ChkPedidoFinalDia.Caption = CargaCadena(1103)
    ChkInventarioFinal.Caption = CargaCadena(1105)
    ChkSolicitarInventarioFin.Caption = CargaCadena(1101)
    ChkSolicitarPedidoFin.Caption = CargaCadena(1099)
    ChkPlu0.Caption = CargaCadena(1078)
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
    'Frame1.Caption = CargaCadena(606)
    Frame2.Caption = CargaCadena(272)
    Frame5.Caption = CargaCadena(272)
    Frame6.Caption = CargaCadena(508)
    'Frame9.Caption = CargaCadena(607)
    Frame10.Caption = CargaCadena(609) & " (No Euro)"
    Frame12.Caption = CargaCadena(509)
    Frame13.Caption = CargaCadena(610)
    Frame14.Caption = CargaCadena(611)
    Frame15.Caption = CargaCadena(1259) & "  dbTeclas.dbf"
    'Frame16.Caption = CargaCadena(608)
    'Frame17.Caption = CargaCadena(615)
    Frame18.Caption = CargaCadena(613)
    Frame19.Caption = CargaCadena(614)
    Frame21.Caption = CargaCadena(334)
    Frame22.Caption = CargaCadena(612)
    CmdAceptar.Caption = CargaCadena(287)
    CmdCancelar.Caption = CargaCadena(288)
    Label1.Caption = CargaCadena(616)
    'Label2.Caption = CargaCadena(616)
    'Label3.Caption = CargaCadena(261)
    'Label4.Caption = CargaCadena(263)
    'Label5.Caption = CargaCadena(262)
    Label7.Caption = CargaCadena(292)
    Label45.Caption = CargaCadena(616)
    'Label8.Caption = CargaCadena(284)
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
    ChkBorrarEti.Caption = CargaCadena(38) & " " & " I.V.A (V.A.T)"
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
    'ChkMaestro.Caption = CargaCadena(297)
    ChkTaquionTiquets.Caption = CargaCadena(480)
    ChkIconificar.Caption = CargaCadena(481)
    ChkIconificarInicio.Caption = CargaCadena(482)
    ChkLogArticulos.Caption = CargaCadena(483)
    ChklogSistema.Caption = CargaCadena(484)
    ChkContinua.Caption = CargaCadena(508)
    ChkBorrarTotalSC10.Caption = CargaCadena(38) & " " & CargaCadena(643)
    ChkGaReinit.Caption = CargaCadena(943)
    'Frame28.Caption = CargaCadena(941)
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
    '**************
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
    CmbTipoC9.List(0) = "3 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
    CmbTipoC9.List(1) = "5 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
    CmbTipoC9.List(2) = "7 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
    ChkPosFam.Caption = CargaCadena(1457)
    ChkAsignTec.Caption = CargaCadena(1458)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v117
    Me.ChkFamilias.Caption = CargaCadena(1230)
    Me.ChkSubSec.Caption = CargaCadena(1231)
    Me.Label8.Caption = CargaCadena(1232)
    Me.ChkMultiEuroscale.Caption = CargaCadena(1233)
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v140
    Me.Label6.Caption = CargaCadena(1271)
    Me.Frame24.Caption = CargaCadena(1259)
    Me.OptSC10Comu(0).Caption = CargaCadena(266)
End Sub
'''''''
'''''''
Private Sub Chk20Lineas_Click()
    If (ChkClientes.Value = vbChecked) Or (ChkFichaVacuno.Value = vbChecked) Then
        Chk20Lineas.Value = vbChecked
    End If
End Sub
'''''''
'''''''
Private Sub ChkClientes_Click()
    If ChkClientes.Value = vbChecked Then
        Chk20Lineas.Value = vbChecked
    Else
        If Option1(1).Value Then
            If ChkFichaVacuno.Value = vbChecked Then
                ChkClientes.Value = vbChecked
            End If
        End If
    End If
End Sub
'''''''
'''''''
Private Sub chkLin2040_Click()
    If (Me.chkLin2040.Value = vbChecked) Then Me.chkLin2040.Value = vbChecked
End Sub
'''''''
'''''''
Private Sub ChkContinua_Click()
    If ChkContinua.Value = vbChecked Then
        LblIntervalo.Enabled = True
        TxtIntervalo.Enabled = True
        TxtIntervalo.TexT = TxtIntervalo.List(5)
        ChkEnvioContinuo.Enabled = True
    Else
       
        LblIntervalo.Enabled = False
        TxtIntervalo.Enabled = False
        ChkEnvioContinuo.Enabled = False
    End If
End Sub

Private Sub ChkEcoLabel_Click()
    DatosLista
    If ChkEcoLabel.Value = vbChecked Then
        Frame23.Enabled = True
        If ChkModem.Value = vbChecked Then
            ChkGamaAlta.Enabled = False
            ChkEcoNet.Enabled = False
            ChkSC10.Enabled = False
            ChkEcoPrint.Enabled = False
            ChkGamaAlta.Value = vbUnchecked
            ChkEcoNet.Value = vbUnchecked
            ChkSC10.Value = vbUnchecked
            ChkEcoPrint.Value = vbUnchecked
            
        End If
        ChkGruposConservacion.Enabled = True
        option3.Enabled = True
        option4.Enabled = True
        CheckMarteNew.Enabled = True
        CheckMarteAlfa.Enabled = True
        Check49T.Enabled = True
        Check66T.Enabled = True
        
    Else
        ChkGamaAlta.Enabled = True
        ChkSC10.Enabled = True
        ChkEcoNet.Enabled = True
        ChkEcoLabel.Enabled = True
        ChkEcoPrint.Enabled = True
        Frame23.Enabled = False
        If ChkGamaAlta.Value = vbUnchecked Then
            If ChkSC10.Value = vbUnchecked Or OptSC10(1).Value Then
                ChkGruposConservacion.Enabled = False
            End If
        End If
        lMarteNew = False
        option3.Value = 0
        option4.Value = 0
        CheckMarteNew.Value = 0
        CheckMarteAlfa.Value = 0
        Check49T.Value = 0
        Check66T.Value = 0
        
        option3.Enabled = False
        option4.Enabled = False
        CheckMarteNew.Enabled = False
        CheckMarteAlfa.Enabled = False
        Check49T.Enabled = False
        Check66T.Enabled = False
        
        option3.Refresh
        option4.Refresh
        CheckMarteNew.Refresh
        CheckMarteAlfa.Refresh
        Check49T.Refresh
        Check66T.Refresh
        
    End If
End Sub

Private Sub ChkEcoNet_Click()
    DatosLista
    If ChkEcoNet.Value = vbChecked Then
        OptECO(0).Enabled = True
        OptECO(1).Enabled = True
        Frame3.Enabled = True
        If ChkModem.Value = vbChecked Then
            ChkGamaAlta.Enabled = False
            ChkSC10.Enabled = False
            ChkEcoLabel.Enabled = False
            ChkEcoPrint.Enabled = False
            ChkGamaAlta.Value = vbUnchecked
            ChkSC10.Value = vbUnchecked
            ChkEcoLabel.Value = vbUnchecked
            ChkEcoPrint.Value = vbUnchecked
        End If
    Else
        ChkGamaAlta.Enabled = True
        ChkSC10.Enabled = True
        ChkEcoNet.Enabled = True
        ChkEcoLabel.Enabled = True
        ChkEcoPrint.Enabled = True
        OptECO(0).Enabled = False
        OptECO(1).Enabled = False
        Frame3.Enabled = False
    End If
    OptECO(1).Value = True
End Sub

Private Sub ChkEcoPrint_Click()
    Dim bucle As Integer
    DatosLista
    If ChkEcoPrint.Value = vbUnchecked Then
        For bucle = 0 To 3 '4 'c2f iv4 ic 3
            OptIV4(bucle).Enabled = False
        Next bucle
        ChkGamaAlta.Enabled = True
        ChkSC10.Enabled = True
        ChkEcoNet.Enabled = True
        ChkEcoLabel.Enabled = True
        ChkEcoPrint.Enabled = True
        Frame30.Enabled = False
    Else
        For bucle = 0 To 3 '4 'c2f iv4 ic 3
            OptIV4(bucle).Enabled = True
        Next bucle
        OptIV4(0).Value = True
        Frame30.Enabled = True
        If ChkModem.Value = vbChecked Then
            ChkGamaAlta.Enabled = False
            ChkEcoNet.Enabled = False
            ChkEcoLabel.Enabled = False
            ChkSC10.Enabled = False
            ChkGamaAlta.Value = vbUnchecked
            ChkEcoNet.Value = vbUnchecked
            ChkEcoLabel.Value = vbUnchecked
            ChkSC10.Value = vbUnchecked
        End If
    End If
End Sub

Private Sub ChkFichaVacuno_Click()
    If ChkFichaVacuno.Value = vbChecked Then
        Chk20Lineas.Value = vbChecked
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
Dim nMyFich As Integer

If ChkGamaAlta.Value = vbChecked Then
    'If Dir(App.Path & "\ahorramas.cfg") <> "" Then
    '    nMyFich = FreeFile()
    '    Open App.Path & "\tiqahip.txt" For Output As #nMyFich
    '    Close #nMyFich
    'End If
    DatosLista
    ChkFichaVacuno.Enabled = True
    'Frame7.Enabled = True
    Frame14.Enabled = True
    'ChkGruposConservacion.Enabled = True
    'If ChkModem.Value = vbChecked Then
    '    ChkSC10.Enabled = False
    '    ChkSC10.Value = vbUnchecked
    '    OPTRS232.Value = vbChecked
    '    ChkEcoNet.Enabled = False
    '    ChkEcoLabel.Enabled = False
    '    ChkEcoPrint.Enabled = False
    '    ChkEcoNet.Value = vbUnchecked
    '    ChkEcoLabel.Value = vbUnchecked
    '    ChkSC10.Value = vbUnchecked
    'End If
Else
    'If Dir(App.Path & "\tiqahip.txt") <> "" Then
    '    Kill App.Path & "\tiqahip.txt"
    'End If
    'ChkEcoNet.Enabled = True
    'ChkEcoLabel.Enabled = True
    'ChkEcoPrint.Enabled = True
    'ChkSC10.Enabled = True
    'DatosLista
    'ChkFichaVacuno.Enabled = False
    'Frame7.Enabled = False
    'Frame14.Enabled = False
    'If (OptSC10(1).Value) And (ChkEcoLabel.Value = vbUnchecked) Then ChkGruposConservacion.Enabled = False
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
        TxtMasterLPort.TexT = "32335"
        TxtMasterRHost.TexT = "127.000.000.001"
        TxtMasterPath.TexT = "y:\hydra"
        OptMaestro(0).Value = True
        Frame22.Enabled = True
    Else
        Frame22.Enabled = False
    End If
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
    Dim nMyFich As Integer
    With Me
        If .ChkSC10.Value = vbChecked Then
        
            'If Dir(App.Path & "\ahorramas.cfg") <> "" Then
            '    nMyFich = FreeFile()
            '    Open App.Path & "\tiqahsc.txt" For Output As #nMyFich
            '    Close #nMyFich
            'End If
            
            DatosLista
            '.Frame8.Enabled = True
            .OptSC10(0).Enabled = True
            .OptSC10(1).Enabled = True
            If .ChkModem.Value = vbChecked Then
                .ChkGamaAlta.Enabled = False
                .ChkEcoNet.Enabled = False
                .ChkEcoLabel.Enabled = False
                .ChkEcoPrint.Enabled = False
                .ChkGamaAlta.Value = vbUnchecked
                .ChkEcoNet.Value = vbUnchecked
                .ChkEcoLabel.Value = vbUnchecked
                .ChkEcoPrint.Value = vbUnchecked
            End If
        Else
            
            'If Dir(App.Path & "\ahorramas.cfg") <> "" And Dir(App.Path & "\tiqahsc.txt") <> "" Then
            '     Kill App.Path & "\tiqahsc.txt"
            'End If
            
            DatosLista
            '.Frame8.Enabled = False
            .OptSC10(0).Value = True
            .OptSC10(0).Enabled = False
            .OptSC10(1).Enabled = False
            .ChkGamaAlta.Enabled = True
            .ChkEcoNet.Enabled = True
            .ChkEcoLabel.Enabled = True
            .ChkEcoPrint.Enabled = True
        End If
    End With
End Sub





Private Sub ChkTaquion_Click()
Dim sruta As String
    If ChkTaquion.Value = vbChecked Then
        ChkStoreFlow.Enabled = False
        ChkStoreFlow.Value = vbUnchecked
        Frame21.Enabled = True
        sruta = App.Path & "\exporta"
        If (Dir(sruta, vbDirectory) = "") Then MkDir sruta
    Else
        'ChkStoreFlow.Enabled = True
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
        'ChkBorrarVacuno.Value = vbChecked
    End If
End Sub

Private Sub CmbMoneda_Click()
    If QLMONEDA Then Exit Sub
    FrmNewLogin.ElUsuario = "FABRICA"
    FrmNewLogin.Form_Load
    FrmNewLogin.Show vbModal
    If LoginSucceeded Then
        If UsaEuro = False And CmbMoneda.TexT = CmbMoneda.List(1) Then
            'msgbox "Recuerde Verificar el Factor Euro", vbInformation
            UsaEuro = True
        End If
        If UsaEuro = True And CmbMoneda.TexT = CmbMoneda.List(0) Then
            'MsgBox "Recuerde Verificar el número de decimales y el símbolo Monetario", vbInformation
            MsgBox CargaCadena(210), vbInformation
            UsaEuro = False
        End If
    Else
        QLMONEDA = True
        If UsaEuro Then
            CmbMoneda.ListIndex = 1
        Else
            CmbMoneda.ListIndex = 0
        End If
        QLMONEDA = False
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
    CmbSC10bits.TexT = MyParam(2)
    CMBSC10Baudios.TexT = MyParam(0)
    CmbSC10Paridad.TexT = MyParam(1)
    
    
End Sub
'\\\\\\
'//////
Private Sub Cmdaceptar_Click()
Dim LaRespuesta As Variant
Dim BufferSimbolo As String
Dim BufferDecimales As String
Dim Devuelve As Boolean
Dim bucle As Integer
Dim MasterFile As Integer
    If ChkGamaAlta.Value = vbUnchecked And ChkSC10.Value = vbUnchecked And _
    ChkEcoNet.Value = vbUnchecked And ChkEcoLabel.Value = vbUnchecked And _
    ChkEcoPrint.Value = vbUnchecked Then
        MsgBox CargaCadena(270), vbCritical
        Exit Sub
    End If
    If Trim(TxtNombreTicket.TexT) = "" Or Trim(TxtAllTicket.TexT) <> "" Then
        If Trim(TxtAllTicket.TexT) = "" Then
            TxtNombreTicket.TexT = "BBTTTT"
        Else
            If InStr(1, Trim(TxtAllTicket.TexT), ".") <> 0 Then
                TxtNombreTicket.TexT = Mid(Trim(TxtAllTicket.TexT), 1, InStr(1, Trim(TxtAllTicket.TexT), ".") - 1)
            Else
                TxtNombreTicket.TexT = Trim(TxtAllTicket.TexT)
            End If
        End If
    End If
    
    If Trim(Txtdias.TexT) = "" Then Txtdias.TexT = "0"
    If Trim(TxtBackupErase.TexT) = "" Then TxtBackupErase.TexT = "0"
    TxtPathExport.TexT = Trim(TxtPathExport.TexT)
    If TxtPathExport.TexT = "" Then TxtPathExport.TexT = Miruta
    If Right(TxtPathExport.TexT, 1) = "\" Then TxtPathExport.TexT = left(TxtPathExport.TexT, Len(TxtPathExport.TexT) - 1)
    '******************
    ' pantallas remotas
    '******************
    If Not IsNumeric(TxtPuertoPantallas.TexT) Then TxtPuertoPantallas.TexT = 32337
    If Val(TxtPuertoPantallas.TexT) < 1024 Or Val(TxtPuertoPantallas.TexT) > 65535 Then
        'MsgBox "Puerto para pantallas remotas no válido", vbCritical
        MsgBox CargaCadena(211), vbCritical
        Tabla.Tab = 1
        Exit Sub
    End If
    '***************
    ' Master / Slave
    '***************
    If ChkMaestro.Value = vbUnchecked Then
        If Dir(Miruta & "\master.cfg") <> "" Then Kill Miruta & "\master.cfg"
        If Dir(Miruta & "\slave.cfg") <> "" Then Kill Miruta & "\slave.cfg"
    Else
        MasterFile = FreeFile()
        If OptMaestro(0).Value = True Then
            If Dir(Miruta & "\slave.cfg") <> "" Then Kill Miruta & "\slave.cfg"
            Open Miruta & "\master.cfg" For Output As MasterFile
        Else
            If Dir(Miruta & "\master.cfg") <> "" Then Kill Miruta & "\master.cfg"
            Open Miruta & "\slave.cfg" For Output As MasterFile
        End If
        Print #MasterFile, TxtMasterLPort.TexT
        Print #MasterFile, TxtMasterRHost.TexT
        Print #MasterFile, TxtMasterPath.TexT
        Close #MasterFile
    End If
    '**************
    ' puertos serie
    '**************
    If (OPTRS232.Value = False) And (OptSC10Comu(1).Value = True) Then
        If CmbCOM.ListIndex = CmbSC10Port.ListIndex Then
            If CmbSC10Port.ListIndex < 3 Then
                CmbCOM.TexT = CmbCOM.List(CmbCOM.ListIndex + 1)
            Else
                CmbCOM.TexT = CmbCOM.List(0)
            End If
        End If
    Else
        If OPTRS232.Value = True And OptSC10Comu(1).Value = False Then
            If CmbCOM.ListIndex = CmbSC10Port.ListIndex Then
                If CmbCOM.ListIndex < 3 Then
                    CmbSC10Port.TexT = CmbSC10Port.List(CmbSC10Port.ListIndex + 1)
                Else
                    CmbSC10Port.TexT = CmbSC10Port.List(0)
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
    '****************
    ' Convesión Euros
    '****************
    BufferSimbolo = TxtSimbolo.TexT
    BufferDecimales = TxtDecimales.TexT
    '*************************
    '*************************
    If ChkStoreFlow.Value = vbChecked Then
        If comprueba_formato_ip(txtCI_HOST.TexT) = False Then
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
    End If
    '*************************
    ' fin storeflow
    '*************************
    If (Trim(TxtFamiliaVacuno.TexT) <> "" And Trim(TxtEtiquetaVacuno.TexT) = "") _
    Or (Trim(TxtFamiliaVacuno.TexT) = "" And Trim(TxtEtiquetaVacuno.TexT) <> "") Then
        'MsgBox "Para definir la Familia Especial debe definir tanto la familia como la etiqueta"
        MsgBox CargaCadena(271)
        Tabla.Tab = 5
        Exit Sub
    End If
    '****************
    ' comandos libres
    '****************
    For bucle = 0 To 4
        TxtComando(bucle).TexT = Trim(TxtComando(bucle).TexT)
        TxtEjecutar(bucle).TexT = Trim(TxtEjecutar(bucle).TexT)
    Next bucle
    EscribirParametrosInicio
    LeerParametrosInicio
    If frmEpelsa.Visible = True Then
        frmEpelsa.Enabled = True
    End If
    Unload Me
    'c2f iv4 ic+ ver si aquí es necesario... 1.7.19
    'If UsaEcoPrint And TipoEcoPrint = 3 Then
    '    If testICdb() Then
    '        MsgBox "IC+ dbasetouch.mdb Modified..."
    '        End
    '    End If
    'End If
    ''''''''''''
End Sub

Private Sub CmdCancelar_Click()
    If UsaEuro Then
        decimales = 2
    End If
    Unload Me
End Sub



Private Sub CmdCarpeta_Click(Index As Integer)
Dim Carpeta As New ClassPath
    Me.Enabled = False
    Carpeta.Show
    If Carpeta.Path <> "" Then
        Select Case Index
            Case 0
                TxtBackup.TexT = Carpeta.Path
            Case 1
                TxtPathExport.TexT = Carpeta.Path
        End Select
    End If
    Me.Enabled = True
End Sub

'Private Sub CmdElegirPAthSC10_Click()
'    frmSC10path.Show vbModal, Me
'End Sub
Public Sub Antonio_Banderas()
    '*****************************
    ' Para cubrir las peticiones
    ' del S.A.C. sobre idiomas
    ' solicitadas por Antonio,
    ' incluímos aquí la carga
    ' de banderas para los idiomas,
    ' de modo dinámico
    '******************************
    Dim bucle As Integer
    For bucle = 1 To 10
        If Dir(App.Path & "\langtouch\" & bucle & ".ico") <> "" Then
            Image1(bucle - 1).Visible = True
            Image1(bucle - 1).Picture = LoadPicture(App.Path & "\langtouch\" & bucle & ".ico")
        Else
            Image1(bucle - 1).Visible = False
        End If
    Next bucle
End Sub
Private Sub DatosLista()
    Lista.ListItems.Clear
    Lista.ListItems.Add , "L00", "General"
    'If ChkGamaAlta.Value = vbChecked Then
        Lista.ListItems.Add , "L01", "TouchScale"
    'End If
    'If ChkSC10.Value = vbChecked Then
    '    If OptSC10(0).Value = True Then
    '        Lista.ListItems.Add , "L02", "SC10"
    '    Else
    '        Lista.ListItems.Add , "L02", "C9/C10"
    '    End If
    'End If
    'If ChkSC10.Value = vbChecked Or ChkGamaAlta.Value = vbChecked Then
        Lista.ListItems.Add , "L03", CargaCadena(602)
    'End If
    Lista.ListItems.Add , "L04", CargaCadena(601)
    'If ChkEcoPrint.Value = vbChecked Then
    '    Lista.ListItems.Add , "L05", "IV4/EcoPrint"
    'End If
    If Not HayMulti Then Lista.ListItems.Add , "L06", CargaCadena(603)
    Lista.ListItems.Add , "L07", CargaCadena(604)
    Lista.ListItems.Add , "L08", CargaCadena(605)
    'If Not HayMulti Then Lista.ListItems.Add , "L09", "Master/Slave"
    Lista.ListItems.Add , "L10", CargaCadena(1052)
    'If ChkModem.Value = vbChecked Then
    '    Lista.ListItems.Add , "L11", "Modem"
    'End If
End Sub
Private Sub Check_Store_Flow()
    If StoreFlow Then
        ChkStoreFlow.Enabled = False
        ChkStoreFlow.Value = vbChecked
        ChkTaquion.Visible = False
        ChkTaquion.Value = vbUnchecked
        ChkModem.Enabled = False
        ChkModem.Value = vbUnchecked
        ChkPasarela.Enabled = False
        ChkPasarela.Value = vbUnchecked
        ChkPantallas.Enabled = False
        ChkPantallas.Value = vbUnchecked
        'Frame17.Enabled = True
        Frame19.Visible = False
        Frame21.Visible = False
    Else
        ChkStoreFlow.Enabled = False
        ChkStoreFlow.Value = vbUnchecked
        ChkTaquion.Enabled = True
        'Frame17.Enabled = False
    End If
End Sub

Private Sub CmdPath_Click(Index As Integer)
    Dialogo.CancelError = True
    On Error Resume Next
    Dialogo.ShowOpen
    If Err.Number = 0 Then
        TxtEjecutar(Index).TexT = Dialogo.FileName
    End If
    On Error GoTo 0
End Sub

Private Sub Opciones_Multi()
    If HayMulti Then
       
        ChkPantallas.Visible = False
        TxtPuertoPantallas.Visible = False
        Frame6.Visible = False
        Frame18.Visible = False
        Frame19.Visible = False
        Frame13.Visible = True 'marcos (barnapes)
       
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
    frmEpelsa.Enabled = True
End Sub



Private Sub Image1_Click(Index As Integer)
    id = (Index) * 1000
    CargaIdiomas
    CambiarIdioma
    DatosLista
End Sub







Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim MyVal As Integer
    Dim bucle As Integer
    MyVal = Val(Mid(Item.Key, 2))

    For bucle = 0 To 11
        
        If bucle <> MyVal Then
            Tabla.TabVisible(bucle) = False
        Else
            Tabla.TabVisible(bucle) = True
            Tabla.Tab = bucle
        End If
    Next bucle
End Sub



Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value Then
        Chk20Lineas.Value = vbChecked
        ChkClientes.Value = vbChecked
    End If
End Sub

'Private Sub Option2_Click(Index As Integer)
'    If Option2(0).Value = True Then
'        Label38(1).Caption = ".ddm"
'        TxtAllTicket.Visible = False
'    End If
'    If Option2(1).Value = True Then
'        Label38(1).Caption = ""
'        TxtAllTicket.Visible = True
'    End If
'End Sub

Private Sub Option3_Click()
    Dim n As Integer
    If (ChkEcoLabel.Value = vbChecked) Or (UsaEcoPrint And TipoEcoPrint = 3) Then
        If option3.Value = 1 Then
            'If Dir(App.Path & "\eco19200.txt") <> "" Then
            '    Kill App.Path & "\eco19200.txt"
            'End If
            option4.Value = 0
          
            'n = FreeFile()
            'Open App.Path & "\eco9600.txt" For Output As #n
            'Close #n
            leco9600 = True
        Else
            'If Dir(App.Path & "\eco9600.txt") <> "" Then
            '    Kill App.Path & "\eco9600.txt"
            'End If
            leco9600 = False
        End If
        option3.Refresh
        'option4.Refresh
    End If
End Sub
Private Sub Option4_Click()
    Dim n As Integer
    If (ChkEcoLabel.Value = vbChecked) Or (UsaEcoPrint And TipoEcoPrint = 3) Then
        If option4.Value = 1 Then
            'If Dir(App.Path & "\eco9600.txt") <> "" Then
            '    Kill App.Path & "\eco9600.txt"
            'End If
            leco9600 = False
            option3.Value = 0
            'n = FreeFile()
            'Open App.Path & "\eco19200.txt" For Output As #n
            'Close #n
            leco19200 = True
        Else
            'If Dir(App.Path & "\eco19200.txt") <> "" Then
            '    Kill App.Path & "\eco19200.txt"
            'End If
            leco19200 = False
        End If
        option3.Refresh
        option4.Refresh
    End If
End Sub



Private Sub OptIV4_Click(Index As Integer)
    Dim n As Integer
    
    Select Case Index
        Case 2
            ChkPedidoIV4.Visible = True
        Case Else
            ChkPedidoIV4.Visible = False
    End Select
    'c2f iv4 IC+
    For n = 0 To 3
        If Index = n Then
            OptIV4(Index).Value = True
        Else
            OptIV4(n).Value = False
        End If
    Next n
    ''''''''''''
End Sub

Private Sub OPTRS232_Click()
    If OPTRS232.Value = True Then
        Frame4.Enabled = True
        Label1.Enabled = False
        TxtPuerto.Enabled = False
        ChkMultiEuroscale.Enabled = False
    End If
End Sub

Private Sub OptSC10_Click(Index As Integer)
    DatosLista
    Select Case OptSC10(0).Value
        Case True
            OptSC10Comu(2).Visible = True
            LblSC10IP.Visible = True
            LblSC10TCP.Visible = True
            TxtIPSC10.Visible = True
            TxtTCPSC10.Visible = True
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
            'Frame8.Caption = "SC10"
            ChkBorrarTotalSC10.Caption = CargaCadena(38) & " " & CargaCadena(643)
            Chk4Lineas.Visible = True
            Me.chkC9Cod.Value = 0
            Me.chkC9Cod.Enabled = False
        Case False
            Chk4Lineas.Visible = False
            ChkTotalesSC10(1).Visible = False
            ChkTotalesSC10(2).Visible = False
            ChkTotalesSC10(3).Visible = False
            ChkBorrarTiquet.Visible = False
            ChkBorrarEti.Visible = False
            CmbProtocolo.TexT = CmbProtocolo.List(0)
            CmbSC10bits.TexT = "7"
            CmbSC10Paridad.TexT = "E"
            CMBSC10Baudios.TexT = "1200"
            OptSC10Comu(0).Visible = False
            OptSC10Comu(1).Visible = False
            
            '1.7.19 se pretende que C9/c10 tambien conecte TCP
            'OptSC10Comu(2).Visible = False
            'LblSC10IP.Visible = False
            'LblSC10TCP.Visible = False
            'TxtIPSC10.Visible = False
            'TxtTCPSC10.Visible = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            OptSC10Comu(2).Visible = True
            LblSC10IP.Visible = True
            LblSC10TCP.Visible = True
            TxtIPSC10.Visible = True
            TxtTCPSC10.Visible = True
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            
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
            If CmbTipoC9.TexT = "" Then CmbTipoC9.TexT = CmbTipoC9.List(0)
            If ChkGamaAlta.Value = vbUnchecked And ChkEcoLabel.Value = vbUnchecked Then
                ChkGruposConservacion.Enabled = False
                ChkGruposConservacion.Value = vbUnchecked
            End If
            'Frame8.Caption = "C9/C10"
            ChkBorrarTotalSC10.Caption = CargaCadena(38) & " " & CargaCadena(643)
            Me.chkC9Cod.Enabled = True
    End Select
End Sub
'//////
Private Sub chkC9Cod_Click()
Dim Res As Integer
    With Me
        If (.chkC9Cod.Value = 1) Then
            Res = verifica_C9Cod
            If Res <> 0 Then
                MsgBox db.CargaCadena(1235) & vbCrLf & _
                       db.CargaCadena(1236) & vbCrLf & _
                       db.CargaCadena(1237), vbCritical
                .chkC9Cod.Value = 0
                Exit Sub
            End If
            .CmbTipoC9.Clear
            .CmbTipoC9.AddItem db.CargaCadena(1238)
            .CmbTipoC9.ListIndex = 0
            .CmbTipoC9.Enabled = False
        Else
            .CmbTipoC9.Clear
            .CmbTipoC9.AddItem "3 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
            .CmbTipoC9.AddItem "5 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
            .CmbTipoC9.AddItem "7 " & CargaCadena(1053) & " + 1 " & CargaCadena(195)
            .CmbTipoC9.ListIndex = 0
            .CmbTipoC9.Enabled = True
        End If
        var.C9Cod = .chkC9Cod.Value
    End With
End Sub
Private Function verifica_C9Cod() As Integer
Dim db As dao.Database
Dim rst As dao.Recordset
Dim sSQL As String
Dim Res As Integer
    Set db = OpenDatabase(Base_General)
    sSQL = "SELECT DISTINCT secc_maqui " & _
           "FROM Seccion " & _
           "WHERE enviardatos=" & "'" & "SC10" & "'" & _
           " AND borrado=false " & _
           "ORDER BY secc_maqui"
    Set rst = db.OpenRecordset(sSQL)
    Res = 0
    If Not rst.EOF Then
        If rst.Recordcount > 1 Then
            Res = 1
        End If
    End If
    rst.Close
    Set rst = Nothing
    db.Close
    Set db = Nothing
    'Workspaces(0).close
    
    verifica_C9Cod = Res
End Function
'\\\\\\\\\\\
'///////////
Private Sub OptSC10Comu_Click(Index As Integer)
    CmbSC10Port.Enabled = False
    CmbSC10bits.Enabled = False
    CMBSC10Baudios.Enabled = False
    CmbSC10Paridad.Enabled = False
    Label14.Enabled = False
    Label15.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = False
    Label6.Enabled = False
    TxtPathSC10.Enabled = False
    CmdElegirPAthSC10.Enabled = False
    LblSC10IP.Enabled = False
    LblSC10TCP.Enabled = False
    TxtIPSC10.Enabled = False
    TxtTCPSC10.Enabled = False
    If OptSC10Comu(0).Value = True Then
        Label6.Enabled = True
        TxtPathSC10.Enabled = True
        CmdElegirPAthSC10.Enabled = True
    End If
    If OptSC10Comu(1).Value = True Then
        CmbSC10Port.Enabled = True
        CmbSC10bits.Enabled = True
        CMBSC10Baudios.Enabled = True
        CmbSC10Paridad.Enabled = True
        Label14.Enabled = True
        Label15.Enabled = True
        Label16.Enabled = True
        Label17.Enabled = True
    End If
    If OptSC10Comu(2).Value = True Then
        TxtIPSC10.Enabled = True
        TxtTCPSC10.Enabled = True
        LblSC10IP.Enabled = True
        LblSC10TCP.Enabled = True
    End If
End Sub

Private Sub OptTCPIP_Click()
    If OptTCPIP.Value = True Then
        Label1.Enabled = True
        TxtPuerto.Enabled = True
        ChkMultiEuroscale.Enabled = True
        Frame4.Enabled = False
    End If
End Sub








Private Sub TxtBackup_LostFocus()
    TxtBackup.TexT = Trim(TxtBackup.TexT)
    If Right(TxtBackup.TexT, 1) = "\" Then
        TxtBackup.TexT = left(TxtBackup.TexT, Len(TxtBackup.TexT) - 1)
    End If
End Sub












Private Sub TxtBackupErase_KeyPress(KeyAscii As Integer)
    TxtBackupErase.Locked = Checktexto(KeyAscii, 2, TxtBackupErase.TexT, True, False)
End Sub

Private Sub TxtCI_HORA_KeyPress(KeyAscii As Integer)
    TxtCI_HORA.Locked = Checktexto(KeyAscii, 2, TxtCI_HORA.TexT, True, False)
End Sub



Private Sub txtCI_MINUTO_KeyPress(KeyAscii As Integer)
    txtCI_MINUTO.Locked = Checktexto(KeyAscii, 2, txtCI_MINUTO.TexT, True, False)
End Sub

Private Sub txtCI_TCP_KeyPress(KeyAscii As Integer)
    txtCI_TCP.Locked = Checktexto(KeyAscii, 5, txtCI_TCP.TexT, True, False)
End Sub

Private Sub txtCI_UDP_keypress(KeyAscii As Integer)
    txtCI_UDP.Locked = Checktexto(KeyAscii, 5, txtCI_UDP.TexT, True, False)
End Sub

Private Sub TxtDbteclas_LostFocus()
    If Right(TxtDbteclas.TexT, 1) = "\" Then
        TxtDbteclas.TexT = left(TxtDbteclas.TexT, Len(TxtDbteclas.TexT) - 1)
    End If
End Sub



Private Sub TxtDecimales_KeyPress(KeyAscii As Integer)
    TxtDecimales.Locked = Checktexto(KeyAscii, 2, TxtDecimales.TexT, True, False)
End Sub





Private Sub Txtdias_KeyPress(KeyAscii As Integer)
    Txtdias.Locked = Checktexto(KeyAscii, 2, Txtdias.TexT, True, False)
End Sub

Private Sub TxtEjecutar_KeyPress(Index As Integer, KeyAscii As Integer)
    TxtEjecutar(Index).Locked = False
    If KeyAscii = 34 Then TxtEjecutar(Index).Locked = True
End Sub





Private Sub TxtEquipoLabel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtEquipoLabel_LostFocus()
    If TxtEquipoLabel.TexT = "" Then TxtEquipoLabel.TexT = "1"
End Sub

Private Sub TxtEquipoV8_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtEquipoV8_LostFocus()
    If TxtEquipoV8.TexT = "" Then TxtEquipoV8.TexT = "1"
End Sub

Private Sub TxtEtiquetaVacuno_keypress(KeyAscii As Integer)
    TxtEtiquetaVacuno.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtEtiquetaVacuno.Locked = True
    If Len(TxtEtiquetaVacuno.TexT) >= 2 Then TxtEtiquetaVacuno.Locked = True
    
End Sub

Private Sub TxtEtiquetaVacuno_LostFocus()
    If TxtEtiquetaVacuno.TexT <> "" Then
        If Val(TxtEtiquetaVacuno.TexT) > 16 Then
            MsgBox CargaCadena(485)  '"El valor de la etiqueta debe estar entre 0 y 16"
            Tabla.Tab = 5
            TxtEtiquetaVacuno.SetFocus
        End If
    End If
End Sub







Private Sub TxtFamiliaVacuno_KeyPress(KeyAscii As Integer)
    TxtFamiliaVacuno.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtFamiliaVacuno.Locked = True
    If Len(TxtFamiliaVacuno.TexT) >= 4 Then TxtFamiliaVacuno.Locked = True

End Sub

Private Sub TxtMasterLPort_keypress(KeyAscii As Integer)
    TxtMasterLPort.Locked = Checktexto(KeyAscii, 5, TxtMasterLPort.TexT, True, False)
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
    If (Not IsNumeric(left(TxtPasarelaHost.TexT, 3))) Or _
    (Not IsNumeric(Mid(TxtPasarelaHost.TexT, 5, 3))) Or _
    (Not IsNumeric(Mid(TxtPasarelaHost.TexT, 9, 3))) Or _
    (Not IsNumeric(Right(TxtPasarelaHost.TexT, 3))) Then
        MsgBox CargaCadena(957), vbExclamation  '"Dirección IP incorrecta", vbExclamation
        Tabla.Tab = 0
        TxtPasarelaHost.SetFocus
    Else
        If (Val(left(TxtPasarelaHost.TexT, 3)) > 255) Or _
           (Val(Mid(TxtPasarelaHost.TexT, 5, 3)) > 255) Or _
           (Val(Mid(TxtPasarelaHost.TexT, 9, 3)) > 255) Or _
           (Val(Right(TxtPasarelaHost.TexT, 3)) > 255) Then
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
    If Val(TxtPasarelaPort.TexT) < 1024 Or Val(TxtPasarelaPort.TexT) > 65535 Then
        MsgBox CargaCadena(958)  '"Valor de Puerto Incorrecto", vbExclamation
        Tabla.Tab = 0
        TxtPasarelaPort.SetFocus
    End If
End Sub


Private Sub TxtPathSC10_LostFocus()
    TxtPathSC10.TexT = Trim(TxtPathSC10.TexT)
    If Right(TxtPathSC10.TexT, 1) = "\" Then
        TxtPathSC10.TexT = left(TxtPathSC10.TexT, Len(TxtPathSC10.TexT) - 1)
    End If
End Sub

Private Sub TxtPuerto_KeyPress(KeyAscii As Integer)
    TxtPuerto.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then TxtPuerto.Enabled = True
    If Len(KeyAscii) >= 5 Then TxtPuerto.Enabled = True
End Sub

Private Sub TxtPuerto_LostFocus()
    If IsNumeric(TxtPuerto.TexT) Then
        If Val(TxtPuerto.TexT) < 1 Or Val(TxtPuerto.TexT) > 65535 Then
            MsgBox CargaCadena(486)  '"El puerto debe valer entre 1 y 65535"
            TxtPuerto.SetFocus
        End If
    End If
End Sub

Private Sub TxtPuertoPantallas_keypress(KeyAscii As Integer)
    TxtPuertoPantallas.Enabled = Checktexto(KeyAscii, 5, TxtPuertoPantallas.TexT, True, False)
End Sub



Private Sub TxtRetries_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub TxtRetries_LostFocus()
    If Not IsNumeric(TxtRetries.TexT) Then
        TxtRetries.TexT = "3"
    Else
        If Val(TxtRetries.TexT) < 3 Or Val(TxtRetries.TexT) > 9 Then TxtRetries.TexT = "3"
    End If
End Sub

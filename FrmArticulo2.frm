VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmArticulo2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Configuración de Artículos"
   ClientHeight    =   6510
   ClientLeft      =   3390
   ClientTop       =   2445
   ClientWidth     =   15405
   Icon            =   "FrmArticulo2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   15405
   Begin VB.CommandButton CmdOnKey 
      Caption         =   "Onkey ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   181
      Top             =   6240
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   28
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   420
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmArticulo2.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbOnKey"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TextUM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmbSub"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmbFam"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Chk100g"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtPreferente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmbTara"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtIVA"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtEuros"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtConservacion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmbBalenv"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmbPesado"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CmbCodigo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CmbSeccion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmbPlu"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtSubSeccion"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtFamilia"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtPrecio"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtCaducidad"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtTara"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "LblVisFam"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LabelUM"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Lbl3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label14"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label11"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "LblEuros"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "LblConservacion"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label8"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label2(0)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label2(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label2(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label2(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label2(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label2(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label2(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label2(7)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label2(9)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Adicionales."
      TabPicture(1)   =   "FrmArticulo2.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cmbOnKey 
         Height          =   315
         ItemData        =   "FrmArticulo2.frx":0044
         Left            =   -73320
         List            =   "FrmArticulo2.frx":004E
         TabIndex        =   180
         Top             =   4560
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4215
         Left            =   -71760
         TabIndex        =   29
         Top             =   600
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   7435
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   420
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Descriptivos"
         TabPicture(0)   =   "FrmArticulo2.frx":005A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label5(8)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5(6)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5(5)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label5(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label5(3)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label5(2)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label5(1)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label5(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label4"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label6(1)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label7(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "LabelICEAN"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label5(9)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label5(19)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label5(10)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label5(11)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label5(12)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label5(13)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label5(14)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label5(15)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label5(16)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label5(17)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label5(18)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label6(0)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Label7(0)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "CmbTipoLetra(9)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "CmbTipoLetra(8)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "CmbTipoLetra(7)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "CmbTipoLetra(6)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "CmbTipoLetra(5)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "CmbTipoLetra(4)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "CmbTipoLetra(3)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "CmbTipoLetra(2)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "CmbTipoLetra(1)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "CmbTipoLetra(0)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Text8(9)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Text8(8)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "Text8(7)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "Text8(6)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "Text8(5)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "Text8(4)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "Text8(3)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "Text8(2)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "Text8(1)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "Text8(0)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "TxtGrupo(1)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "TextICEAN"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "CmdPL"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "TxtGrupo(4)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "TxtGrupo(3)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "TxtGrupo(2)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "TxtGrupo(0)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "Text8(20)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "Text8(11)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "Text8(15)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "Text8(16)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "Text8(17)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "Text8(18)"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "Text8(19)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "Text8(10)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "Text8(12)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "Text8(13)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "Text8(14)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "CmbTipoLetra(10)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "CmbTipoLetra(20)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "CmbTipoLetra(11)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "CmbTipoLetra(12)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "CmbTipoLetra(13)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "CmbTipoLetra(14)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "CmbTipoLetra(15)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "CmbTipoLetra(16)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "CmbTipoLetra(17)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "CmbTipoLetra(18)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "CmbTipoLetra(19)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "cmdLin2040"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).ControlCount=   76
         TabCaption(1)   =   "Prcs."
         TabPicture(1)   =   "FrmArticulo2.frx":0076
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text3(9)"
         Tab(1).Control(1)=   "Text3(0)"
         Tab(1).Control(2)=   "Text3(1)"
         Tab(1).Control(3)=   "Text3(2)"
         Tab(1).Control(4)=   "Text3(3)"
         Tab(1).Control(5)=   "Text3(4)"
         Tab(1).Control(6)=   "Text3(5)"
         Tab(1).Control(7)=   "Text3(6)"
         Tab(1).Control(8)=   "Text3(7)"
         Tab(1).Control(9)=   "Text3(8)"
         Tab(1).Control(10)=   "Label18"
         Tab(1).Control(11)=   "Label19"
         Tab(1).Control(12)=   "Label20"
         Tab(1).Control(13)=   "Label21"
         Tab(1).Control(14)=   "Label22(0)"
         Tab(1).Control(15)=   "Label22(1)"
         Tab(1).Control(16)=   "Label22(2)"
         Tab(1).Control(17)=   "Label22(3)"
         Tab(1).Control(18)=   "Label22(4)"
         Tab(1).ControlCount=   19
         TabCaption(2)   =   "V. Text (Lin1-Lin20)"
         TabPicture(2)   =   "FrmArticulo2.frx":0092
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label22(5)"
         Tab(2).Control(1)=   "RTB1(4)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "V. text1"
         TabPicture(3)   =   "FrmArticulo2.frx":00AE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "RTB1(0)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "V. Text2"
         TabPicture(4)   =   "FrmArticulo2.frx":00CA
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "RTB1(1)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "V. Text3"
         TabPicture(5)   =   "FrmArticulo2.frx":00E6
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "RTB1(2)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "V. Text4"
         TabPicture(6)   =   "FrmArticulo2.frx":0102
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "RTB1(3)"
         Tab(6).ControlCount=   1
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   -71880
            TabIndex        =   177
            Top             =   3840
            Width           =   855
         End
         Begin VB.CommandButton cmdLin2040 
            Caption         =   "Textos adicionales líneas  21 - 40"
            Height          =   375
            Left            =   3960
            TabIndex        =   169
            Top             =   3720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   -71880
            TabIndex        =   159
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   -71880
            TabIndex        =   158
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   -71880
            TabIndex        =   157
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   -71880
            TabIndex        =   156
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   -71880
            TabIndex        =   155
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   -71880
            TabIndex        =   154
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   -71880
            TabIndex        =   153
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   -71880
            TabIndex        =   152
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   -71880
            TabIndex        =   151
            Top             =   3480
            Width           =   855
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   19
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   3120
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   18
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   2880
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   17
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   2640
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   16
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   2400
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   2160
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   14
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   1920
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   13
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   1680
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   1440
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   1200
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   20
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   3360
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   139
            Top             =   960
            Width           =   852
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   14
            Left            =   4800
            TabIndex        =   138
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   13
            Left            =   4800
            TabIndex        =   137
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   12
            Left            =   4800
            TabIndex        =   136
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   10
            Left            =   4800
            TabIndex        =   135
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   19
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   122
            Top             =   3120
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   18
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   121
            Top             =   2880
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   17
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   120
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   16
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   119
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   15
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   118
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   11
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   117
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   20
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   116
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox TxtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H80000017&
            Height          =   288
            Index           =   0
            Left            =   4800
            MaxLength       =   25
            TabIndex        =   115
            Text            =   "(Reservado)"
            Top             =   960
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox TxtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H80000017&
            Height          =   288
            Index           =   2
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   114
            Text            =   "(Reservado)"
            Top             =   1560
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox TxtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H80000017&
            Height          =   288
            Index           =   3
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   113
            Text            =   "(Reservado)"
            Top             =   1800
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox TxtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H80000017&
            Height          =   288
            Index           =   4
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   112
            Text            =   "(Reservado)"
            Top             =   1320
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton CmdPL 
            Caption         =   "Label's"
            Height          =   375
            Left            =   120
            TabIndex        =   110
            Top             =   3720
            Width           =   3015
         End
         Begin VB.TextBox TextICEAN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            MaxLength       =   13
            TabIndex        =   102
            Top             =   3720
            Width           =   615
         End
         Begin VB.TextBox TxtGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   288
            Index           =   1
            Left            =   960
            MaxLength       =   25
            TabIndex        =   80
            Text            =   "(Reservado)"
            Top             =   3240
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   960
            MaxLength       =   40
            TabIndex        =   14
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   960
            MaxLength       =   25
            TabIndex        =   18
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   2
            Left            =   960
            MaxLength       =   25
            TabIndex        =   19
            Top             =   1560
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   3
            Left            =   960
            MaxLength       =   25
            TabIndex        =   20
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   4
            Left            =   960
            MaxLength       =   25
            TabIndex        =   21
            Top             =   2040
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   5
            Left            =   960
            MaxLength       =   25
            TabIndex        =   22
            Top             =   2280
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   6
            Left            =   960
            MaxLength       =   25
            TabIndex        =   23
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   7
            Left            =   960
            MaxLength       =   25
            TabIndex        =   24
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   8
            Left            =   960
            MaxLength       =   25
            TabIndex        =   25
            Top             =   3000
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   9
            Left            =   960
            TabIndex        =   26
            Top             =   3240
            Width           =   2895
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   960
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1200
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1440
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1680
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1920
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2160
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2400
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2640
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2880
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.ComboBox CmbTipoLetra 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   3120
            Visible         =   0   'False
            Width           =   852
         End
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3495
            Index           =   3
            Left            =   -72600
            TabIndex        =   173
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmArticulo2.frx":011E
         End
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3495
            Index           =   2
            Left            =   -72600
            TabIndex        =   174
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmArticulo2.frx":01E7
         End
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3495
            Index           =   1
            Left            =   -72600
            TabIndex        =   175
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmArticulo2.frx":02B0
         End
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3495
            Index           =   0
            Left            =   -72600
            TabIndex        =   176
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmArticulo2.frx":0379
         End
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3495
            Index           =   4
            Left            =   -72600
            TabIndex        =   178
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmArticulo2.frx":0442
         End
         Begin VB.Label Label18 
            Caption         =   "Prc.1:"
            Height          =   255
            Left            =   -72480
            TabIndex        =   168
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Prc.2:"
            Height          =   255
            Left            =   -72480
            TabIndex        =   167
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "Prc.3:"
            Height          =   255
            Left            =   -72480
            TabIndex        =   166
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label21 
            Caption         =   "Prc.4:"
            Height          =   255
            Left            =   -72480
            TabIndex        =   165
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.5:"
            Height          =   255
            Index           =   0
            Left            =   -72480
            TabIndex        =   164
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.6:"
            Height          =   255
            Index           =   1
            Left            =   -72480
            TabIndex        =   163
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.7:"
            Height          =   255
            Index           =   2
            Left            =   -72480
            TabIndex        =   162
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.8:"
            Height          =   255
            Index           =   3
            Left            =   -72480
            TabIndex        =   161
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.9:"
            Height          =   255
            Index           =   4
            Left            =   -72480
            TabIndex        =   160
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Fuente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   150
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Texto"
            Height          =   255
            Index           =   0
            Left            =   4800
            TabIndex        =   134
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 19"
            Height          =   255
            Index           =   18
            Left            =   3960
            TabIndex        =   133
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 18"
            Height          =   255
            Index           =   17
            Left            =   3960
            TabIndex        =   132
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 17"
            Height          =   255
            Index           =   16
            Left            =   3960
            TabIndex        =   131
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 16"
            Height          =   255
            Index           =   15
            Left            =   3960
            TabIndex        =   130
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 15"
            Height          =   255
            Index           =   14
            Left            =   3960
            TabIndex        =   129
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 14"
            Height          =   255
            Index           =   13
            Left            =   3960
            TabIndex        =   128
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 13"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   127
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 12"
            Height          =   255
            Index           =   11
            Left            =   3960
            TabIndex        =   126
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 11"
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   125
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 20"
            Height          =   255
            Index           =   19
            Left            =   3960
            TabIndex        =   124
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 10"
            Height          =   255
            Index           =   9
            Left            =   3960
            TabIndex        =   123
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Prc.10:"
            Height          =   255
            Index           =   5
            Left            =   -66120
            TabIndex        =   109
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label LabelICEAN 
            Caption         =   "Posición:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Fuente"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Texto"
            Height          =   252
            Index           =   1
            Left            =   960
            TabIndex        =   66
            Top             =   840
            Width           =   1572
         End
         Begin VB.Label Label4 
            Caption         =   "Principal"
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
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 1"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 2"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   1560
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 3"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 4"
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 5"
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   2280
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 6"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   39
            Top             =   2520
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 7"
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   2760
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 8"
            Height          =   252
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   3000
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Linea 9"
            Height          =   252
            Index           =   8
            Left            =   120
            TabIndex        =   33
            Top             =   3240
            Width           =   732
         End
      End
      Begin VB.TextBox TextUM 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73320
         TabIndex        =   100
         Top             =   4080
         Width           =   1455
      End
      Begin VB.ComboBox CmbSub 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1200
         Width           =   1692
      End
      Begin VB.ComboBox CmbFam 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   1692
      End
      Begin VB.CheckBox Chk100g 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   -74520
         TabIndex        =   92
         Top             =   1920
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.TextBox TxtPreferente 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72600
         TabIndex        =   90
         Top             =   3720
         Width           =   732
      End
      Begin VB.ComboBox CmbTara 
         Height          =   315
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   2760
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.TextBox TxtIVA 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   10
         Top             =   3000
         Width           =   1092
      End
      Begin VB.TextBox TxtEuros 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   7
         Top             =   2280
         Width           =   1092
      End
      Begin VB.TextBox TxtConservacion 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -67440
         TabIndex        =   13
         Top             =   3960
         Width           =   732
      End
      Begin VB.Frame Frame2 
         Caption         =   "Precios por Tramos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   68
         Top             =   2400
         Width           =   7452
         Begin VB.TextBox TxtPrc 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   1320
            Width           =   1932
         End
         Begin VB.TextBox TxtPrc 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   1080
            Width           =   1932
         End
         Begin VB.TextBox TxtTramo 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   1320
            Width           =   1452
         End
         Begin VB.TextBox TxtTramo 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   1080
            Width           =   1452
         End
         Begin VB.CheckBox ChkTramo 
            Caption         =   "Activar Precios por tramos para este Artículo"
            Height          =   372
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   7212
         End
         Begin VB.Label Label10 
            Caption         =   "Precio del Tramo :"
            Height          =   252
            Index           =   1
            Left            =   3360
            TabIndex        =   78
            Top             =   1320
            Width           =   1692
         End
         Begin VB.Label Label10 
            Caption         =   "Precio del Tramo :"
            Height          =   252
            Index           =   0
            Left            =   3360
            TabIndex        =   77
            Top             =   1080
            Width           =   1692
         End
         Begin VB.Label LblTramo2 
            Caption         =   "Label10"
            Height          =   252
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   1572
         End
         Begin VB.Label LblTramo1 
            Caption         =   "Label10"
            Height          =   252
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   1332
         End
         Begin VB.Label lblPrecioBase 
            Caption         =   "Precio Base"
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
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   3612
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Adicionales del Artículo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   61
         Top             =   480
         Width           =   7452
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            MaxLength       =   14
            TabIndex        =   106
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   14
            TabIndex        =   104
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox TextPoid 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   98
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox CheckMix 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   96
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox TxtPresel 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   6360
            MaxLength       =   2
            TabIndex        =   87
            Top             =   1200
            Width           =   732
         End
         Begin VB.TextBox TxtMerma 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   85
            Top             =   1560
            Width           =   972
         End
         Begin VB.TextBox TxtEAN13 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   4920
            TabIndex        =   65
            Top             =   480
            Width           =   2172
         End
         Begin VB.TextBox TxtEtiqueta 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1800
            MaxLength       =   14
            TabIndex        =   63
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "EAN 14:"
            Height          =   255
            Left            =   4200
            TabIndex        =   107
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "2. Lbl.:"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Múltiplo de 10 (Multiple of 10)."
            Height          =   255
            Left            =   2880
            TabIndex        =   103
            Top             =   1560
            Width           =   3375
         End
         Begin VB.Label LabelPoid 
            Caption         =   "Poid:"
            Height          =   255
            Left            =   1080
            TabIndex        =   97
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label LabelMix 
            Caption         =   "Mix.:"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Discount (%)"
            Height          =   255
            Left            =   5040
            TabIndex        =   86
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Merma (%) :"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "EAN 13:"
            Height          =   255
            Left            =   4200
            TabIndex        =   64
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo de Etiqueta :"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.ComboBox CmbBalenv 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmArticulo2.frx":050B
         Left            =   -67080
         List            =   "FrmArticulo2.frx":0515
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.ComboBox CmbPesado 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmArticulo2.frx":0521
         Left            =   -73200
         List            =   "FrmArticulo2.frx":052B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3360
         Width           =   1332
      End
      Begin VB.ComboBox CmbCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72960
         TabIndex        =   0
         Text            =   "CmbCodigo"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox CmbSeccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1092
      End
      Begin VB.ComboBox CmbPlu 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -72960
         TabIndex        =   2
         Top             =   960
         Width           =   1092
      End
      Begin VB.TextBox TxtSubSeccion 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   3
         Top             =   1200
         Width           =   1092
      End
      Begin VB.TextBox TxtFamilia 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   4
         Top             =   1560
         Width           =   1092
      End
      Begin VB.TextBox TxtPrecio 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   6
         Top             =   2040
         Width           =   1092
      End
      Begin VB.TextBox TxtCaducidad 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   8
         Top             =   2520
         Width           =   1092
      End
      Begin VB.TextBox TxtTara 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -72960
         TabIndex        =   9
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label LblVisFam 
         Caption         =   "On Key:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   179
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label LabelUM 
         Caption         =   "Última Modificación:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Lbl3 
         Caption         =   "Consumo Preferente :"
         Height          =   255
         Left            =   -74520
         TabIndex        =   91
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   -71760
         TabIndex        =   88
         Top             =   4800
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo de IVA"
         Height          =   255
         Left            =   -74880
         TabIndex        =   82
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label LblEuros 
         Caption         =   "Precio (euros)"
         Height          =   252
         Left            =   -74520
         TabIndex        =   81
         Top             =   2280
         Width           =   1212
      End
      Begin VB.Label LblConservacion 
         Caption         =   "Grupo de Conservación"
         Height          =   255
         Left            =   -69960
         TabIndex        =   79
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Envío a ambos Sistemas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69720
         TabIndex        =   59
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
         Height          =   252
         Index           =   0
         Left            =   -74520
         TabIndex        =   58
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Mostrador"
         Height          =   252
         Index           =   1
         Left            =   -74520
         TabIndex        =   57
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "PLU"
         Height          =   252
         Index           =   2
         Left            =   -74520
         TabIndex        =   56
         Top             =   960
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Subsección"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   55
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Familia"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   54
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Precio (nacional)"
         Height          =   252
         Index           =   5
         Left            =   -74520
         TabIndex        =   53
         Top             =   2040
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Caducidad"
         Height          =   252
         Index           =   6
         Left            =   -74520
         TabIndex        =   52
         Top             =   2520
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Tara"
         Height          =   255
         Index           =   7
         Left            =   -74520
         TabIndex        =   51
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Forma de Venta"
         Height          =   252
         Index           =   9
         Left            =   -74520
         TabIndex        =   50
         Top             =   3360
         Width           =   1332
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Text1-Text5 (Lin1-Lin5)"
      Height          =   255
      Left            =   11520
      TabIndex        =   172
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Trans. LongText= (lin2-lin40)"
      Height          =   255
      Left            =   11520
      TabIndex        =   171
      Top             =   960
      Width           =   3735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Trans. LongText= (lin1-lin40)"
      Height          =   255
      Left            =   11520
      TabIndex        =   170
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton CmdOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenar Teclas"
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
      Left            =   9720
      TabIndex        =   111
      Top             =   5760
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmArticulo2.frx":0542
            Key             =   "disabled"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmArticulo2.frx":069C
            Key             =   "enabled"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmArticulo2.frx":07F6
            Key             =   "no"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   375
      Left            =   8160
      TabIndex        =   93
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   83
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
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
      Height          =   372
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   1212
   End
   Begin VB.CommandButton CmdBorrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Borrar"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1212
   End
   Begin VB.CommandButton CmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Añadir"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label17 
      Caption         =   "Prc.1:"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   108
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image picture1 
      Height          =   3015
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5160
      TabIndex        =   60
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
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
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu mnuListados 
      Caption         =   "Listados"
      Begin VB.Menu mnuTodos 
         Caption         =   "Todos"
      End
      Begin VB.Menu mnuseccion 
         Caption         =   "Por Mostrador"
      End
      Begin VB.Menu mnuCambios 
         Caption         =   "Cambios de Precio"
      End
   End
End
Attribute VB_Name = "FrmArticulo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CodigoAnterior As Long
Private SeccionAnterior As Long
Private AnulaSeccion As Boolean
Private EnPlu As Boolean
Private PluGB As Long
Private Nocarga As Boolean
Private PorCodigo As Boolean
Private lCharging As Boolean

#Const BALSAM = False

'Private sTxtCod As String
'


Private Sub CambiarIdioma()
    Dim bucle As Integer
    'Label8.Caption = "Euroscale + V-12" 'CargaCadena(566)
    'If id <> 0 Then CmbBalenv.List(1) = "Yes"
    mnuTodos.Caption = CargaCadena(311)
    mnuseccion.Caption = CargaCadena(312)
    mnuListados.Caption = CargaCadena(490)
    CmdAceptar.Caption = CargaCadena(550)
    Label6(0).Caption = CargaCadena(539)
    Label6(1).Caption = CargaCadena(539)
    LblInfo.Caption = CargaCadena(524)
    Chk100g.Caption = CargaCadena(1194)
    FrmArticulo2.Caption = CargaCadena(526)
    Label2(0).Caption = CargaCadena(69)
    Label2(1).Caption = CargaCadena(302)
    'Label2(3).Caption = CargaCadena(10)
    Label2(3).Visible = False
    Label2(4).Caption = CargaCadena(67)
    Label2(5).Caption = CargaCadena(534)
    LblEuros.Caption = CargaCadena(535)
    Label2(6).Caption = CargaCadena(72)
    Label2(7).Caption = CargaCadena(74) & "(kg)"
    'Label2(8).Caption = CargaCadena( 535)
    Label2(9).Caption = CargaCadena(76)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
    CmbPesado.List(0) = CargaCadena(79)
    CmbPesado.List(1) = CargaCadena(80)
    For bucle = 0 To 19
        Label5(bucle).Caption = CargaCadena(537) & " " & bucle + 1
    Next bucle
    If usaSC10 And Not UsaGamaAlta And TipoC9 Then
        SSTab2.TabCaption(0) = CargaCadena(536)
    Else
        SSTab2.TabCaption(0) = CargaCadena(536) & " 1-20"
    End If
    SSTab2.TabCaption(1) = "Prcs."    'CargaCadena(536) & " 11-20"
    Label4.Caption = CargaCadena(538)
    Label11.Caption = CargaCadena(553)
    LblConservacion.Caption = CargaCadena(638)
    Lbl3.Caption = CargaCadena(73)
    Frame1.Caption = CargaCadena(646)
    Frame2.Caption = CargaCadena(647)
    ChkTramo.Caption = CargaCadena(648)
    Label3.Caption = CargaCadena(77)
    lblPrecioBase.Caption = CargaCadena(649)
    Label10(0).Caption = CargaCadena(650)
    Label10(1).Caption = CargaCadena(650)
    Label7(0).Caption = CargaCadena(893)
    Label7(1).Caption = CargaCadena(893)
    mnuCambios.Caption = CargaCadena(24)
    CmdLimpiar.Caption = CargaCadena(1045)
    Label12.Caption = CargaCadena(1046)
    Label13.Caption = CargaCadena(1047)
End Sub
Private Sub Refresca_Codigos()

    Dim Registro As New RecordNet
    CmbCodigo.Clear

    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        Registro.OpenRecordset _
                ("select codigo,borrado from articulo where borrado=false and codi_ident=" & _
                 UsuarioActual.Asociado & " order by codigo")
    Else
        Registro.OpenRecordset _
                ("select codigo,borrado from articulo where borrado=false order by codigo")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                CmbCodigo.AddItem .Fields("codigo")
                .Movenext
            Loop
        End If
    End With
End Sub
Private Sub Refresca_Plus()
    Dim Registro As New RecordNet
    Dim sSQL As String
    If EnPlu Then Exit Sub
    CmbPlu.Clear

    If Val(CmbSeccion.TexT) > 0 Then
        sSQL = "SELECT plu " & _
               "FROM articulo " & _
               "WHERE borrado=false" & _
             " AND codi_ident=" & CmbSeccion.TexT & " " & _
               "ORDER BY plu"
        Registro.OpenRecordset (sSQL)
        With Registro
            If Not .EOF Then
                .MoveFirst
                Do Until .EOF
                    CmbPlu.AddItem .Fields("plu")
                    .Movenext
                Loop
            End If
        End With
    Else
        CmbPlu.AddItem 0
    End If
End Sub
Private Sub Refresca_Datos()    'aquí Aho
    Dim bucle As Integer
    Dim Art As DB_Articulo
    Dim MiNombre As String
    Dim ok As Boolean
    Dim sSQL As String
    Dim rst As dao.Recordset
    Dim db As dao.Database
    Dim bPlu As Long
    Dim Ret As Boolean
    Dim sD As String
    Dim Token As Long
    Dim sTextLong As String
    lCharging = True
    nQueImagen = 1
    sQuePathImagen = App.Path & "\images\"

    If Not ((PorCodigo And CmbCodigo.TexT = "") Or (Not PorCodigo And CmbPlu.TexT = "")) Then
        If PorCodigo Then
            Art.codigo = Val(CmbCodigo.TexT)
        Else
            Art.Plu = Val(CmbPlu.TexT)
            Art.Mostrador = Val(CmbSeccion.TexT)
        End If
    End If
    ok = DB_Consulta_Art(Art)

    sQueNombreImagen = Art.Imagen

    If ok Then
        '2.0.23
        'If Art.WGH Or (CmbPesado.TexT = CmbPesado.List(1)) Then
        '    TextPoid.Enabled = True
        'Else
        '    TextPoid.Enabled = False
        '    TextPoid.TexT = "0"
        'End If
        ''''''''
        LblInfo.Caption = CargaCadena(525)  '"Modificación"
        LblInfo.BackColor = vbYellow
        CmdAceptar.Caption = CargaCadena(36)  '"Modificar"
        AnulaSeccion = True
        On Error Resume Next
        CmbSeccion.TexT = CStr(Art.Mostrador)
        If Err.Number <> 0 Then
            On Error GoTo 0
            CmbSeccion.AddItem CStr(Art.Mostrador)
            CmbSeccion.TexT = CStr(Art.Mostrador)
        End If
        On Error GoTo 0
        AnulaSeccion = False
        CmbSeccion.Enabled = False
        TxtPreferente.TexT = CStr(Art.Preferente)
        Chk100g.Value = BoolToCheck(Art.PRC100G)
        TxtIVA.TexT = CStr(Art.IVA)
        TxtEtiqueta.TexT = CStr(Art.Etiqueta)
        If Art.Etiqueta = 255 Then TxtEtiqueta.TexT = ""
        If Art.onkey = 1 Then
            cmbOnKey.ListIndex = 0
        Else
            cmbOnKey.ListIndex = 1
        End If
        If PorCodigo Then
            CmbPlu.TexT = CStr(Art.Plu)
            CmbPlu.Enabled = False
            CmbCodigo.Enabled = True
        Else
            CmbCodigo.TexT = CStr(Art.codigo)
            CmbPlu.Enabled = True
            CmbCodigo.Enabled = False
        End If
        #If BALSAM Then
            For bucle = 0 To 8
                Text3(bucle).TexT = Art.tarifa(bucle)
            Next bucle
        #Else
            For bucle = 0 To 9
                Text3(bucle).TexT = Art.tarifa(bucle)
            Next bucle
        #End If
        If FamiliasFijas Then
            If CmbFam.ListCount > 0 Then
                For bucle = 0 To CmbFam.ListCount - 1
                    If Val(left(CmbFam.List(bucle), 4)) = Art.familia Then
                        CmbFam.ListIndex = bucle
                        Exit For
                    End If
                Next bucle
            End If
        Else
            TxtFamilia.TexT = CStr(Art.familia)
        End If
        'If SubSeccionesFijas Then
        '    If CmbSub.ListCount > 0 Then
        '        For Bucle = 0 To CmbSub.ListCount - 1 '///cas.v118
        '            If Val(left(CmbSub.List(Bucle), 4)) = Art.subsec Then
        '                CmbSub.ListIndex = Bucle
        '                Exit For
        '            End If
        '        Next Bucle
        '    End If
        'Else
        TxtSubSeccion.TexT = CStr(Art.subsec)
        'End If
        If UsaEuro Then
            TxtEuros.TexT = CStr(Art.precio)
        Else
            TxtPrecio.TexT = CStr(Art.precio)
        End If
        TxtCaducidad.TexT = CStr(Art.caducidad)
        TxtConservacion.TexT = CStr(Art.GrpConserv)
        TxtPresel.TexT = CStr(Art.Presel)
        #If BALSAM Then
            TxtMerma.TexT = CStr(Art.Merma)
        #End If
        TxtTara.TexT = CStr(Art.tara)
        'If Dir(App.Path & "\mixmode.txt") <> "" Then
        If Art.lMix Then
            CheckMix.Value = vbChecked
        Else
            CheckMix.Value = vbUnchecked
        End If
        TextPoid.TexT = CStr(Art.nPoid)
        'End If

        '2.0.22
        'If lVL000 And Dir(App.Path & "\mixmode.txt") = "" Then
        '    Art.lMix = False
        '    TextPoid.TexT = CStr(Art.nPoid)
        'End If
        '''''''

        ' Combo de tipos de tara EcoLabel
        If CmbTara.Visible = True Then
            If Art.tara <> 0 Then
                For bucle = 0 To CmbTara.ListCount - 1
                    If Art.tara = Val(CmbTara.List(bucle)) Then
                        CmbTara.ListIndex = bucle
                        Exit For
                    End If
                Next bucle
            Else
                CmbTara.TexT = CmbTara.List(0)
            End If
        End If
        TxtEAN13.TexT = Trim(Art.ean13)

        'If (UsaEcoPrint And TipoEcoPrint = 3) Or lMarteAlfa Then 'IC+
        '    TextICEAN.TexT = Trim(Art.EAN13)
        'End If

        'TextICEAN.TexT = CStr(Art.Posicion)
        '''''''''''''''''''''''''''''''''''
        CmbPesado.TexT = CmbPesado.List(0)
        If Art.WGH Then CmbPesado.TexT = CmbPesado.List(1)
        'CmbBalenv.TexT = CmbBalenv.List(0)
        'If Art.Balenv Then CmbBalenv.TexT = CmbBalenv.List(1)
        '*******

        Text1.TexT = Art.Label2

        #If BALSAM = True Then
            TxtEtiqueta.TexT = Art.Label1
            Text2.TexT = Art.ean14
        #End If
        'RTB1(4).TextRTF = ""
        sTextLong = ""
        For bucle = 0 To 20
            Text8(bucle).TexT = Art.Descriptivos(bucle)
            If bucle < 10 Then
                CmbTipoLetra(bucle).TexT = Art.TipoLetra(bucle)
            End If
            If bucle > 0 Then sTextLong = sTextLong & Text8(bucle).TexT    'Art.Descriptivos(bucle)
        Next bucle
        RTB1(4).TexT = sTextLong
        RTB1(4).SelStart = 0
        RTB1(4).SelStart = 1

        SSTab2.Tab = 0
        '*************
        'Tramos
        '*************
        ChkTramo.Value = vbUnchecked
        TxtPrc(0).TexT = ""
        TxtTramo(0).TexT = ""
        TxtPrc(1).TexT = ""
        TxtTramo(1).TexT = ""
        TxtPrc(0).Enabled = False
        TxtTramo(0).Enabled = False
        TxtPrc(1).Enabled = False
        TxtTramo(1).Enabled = False

        'If Dir(App.Path & "\digitonuba.txt") <> "" Then
        '    ChkTramo.Value = vbChecked
        '    TxtPrc(0).Enabled = True
        '    'TxtTramo(0).Enabled = True
        '    TxtPrc(1).Enabled = True
        '    'TxtTramo(1).Enabled = True
        '    TxtPrc(0).TexT = ToString(Art.PRC1)
        '    'TxtTramo(0).TexT = ToString(Art.TRM1)
        '    'If Art.PRC2 >= 0 And Art.TRM2 >= 0 Then
        '        TxtPrc(1).TexT = ToString(Art.PRC2)
        '        'TxtTramo(1).TexT = ToString(Art.TRM2)
        '    'Else
        '    '    TxtPrc(1).TexT = ""
        '    '    TxtTramo(1).TexT = ""
        '    'End If
        '
        'Else
        If Art.TRM1 >= 0 Or Art.TRM2 >= 0 Then
            ChkTramo.Value = vbChecked
            TxtPrc(0).Enabled = True
            TxtTramo(0).Enabled = True
            TxtPrc(1).Enabled = True
            TxtTramo(1).Enabled = True
            TxtPrc(0).TexT = ToString(Art.PRC1)
            TxtTramo(0).TexT = ToString(Art.TRM1)
            If Art.PRC2 >= 0 And Art.TRM2 >= 0 Then
                TxtPrc(1).TexT = ToString(Art.PRC2)
                TxtTramo(1).TexT = ToString(Art.TRM2)
            Else
                TxtPrc(1).TexT = ""
                TxtTramo(1).TexT = ""
            End If
        End If
        'End If
        Set db = AbrirBase
        Set rst = db.OpenRecordset("select * from text15 where codigo=" & CStr(Val(CmbCodigo.TexT)))
        If Not rst.EOF Then
            If IsNull(rst.Fields("text1")) Then
                RTB1(0).TextRTF = ""
            Else
                RTB1(0).TextRTF = rst.Fields("text1")
            End If
            If IsNull(rst.Fields("text2")) Then
                RTB1(1).TextRTF = ""
            Else
                RTB1(1).TextRTF = rst.Fields("text2")
            End If
            If IsNull(rst.Fields("text3")) Then
                RTB1(2).TextRTF = ""
            Else
                RTB1(2).TextRTF = rst.Fields("text3")
            End If
            If IsNull(rst.Fields("text4")) Then
                RTB1(3).TextRTF = ""
            Else
                RTB1(3).TextRTF = rst.Fields("text4")
            End If
        End If
        rst.Close
        db.Close
        Set db = Nothing
        'If AhorraMas Or Dir(App.Path & "\upper.txt") <> "" Then
        sD = Format(Art.PRC3, "0000000000")
        sD = Mid(sD, 5, 2) & "/" & Mid(sD, 3, 2) & "/" & Mid(sD, 1, 2) & "-" & Mid(sD, 7, 2) & ":" & Mid(sD, 9)
        TextUM.TexT = sD

        'End If

        'If Mid(Trim(sQueNombreImagen), 1, 1) <> "#" Then
        picture1.Height = 3135
        picture1.Width = 3615
        'picture1.top = 0
        'picture1.left = CmdColor(1).left

        On Error Resume Next
        Token = InitGDIPlus
        If Dir(Art.Imagen) = "" Then
            If Dir(App.Path & "\images\" & Art.Imagen) <> "" Then
                Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\images\" & Art.Imagen)
            Else
                Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
            End If
        Else
            Set picture1.Picture = LoadPictureGDIPlus(Art.Imagen)
        End If
        FreeGDIPlus Token
        On Error GoTo 0
        'Else
        '    Call ColorToImage(Trim(sQueNombreImagen))
        'End If

    Else

        RTB1(0).TextRTF = ""
        RTB1(1).TextRTF = ""
        RTB1(2).TextRTF = ""
        RTB1(3).TextRTF = ""
        RTB1(4).TextRTF = ""
        cmbOnKey.ListIndex = 0

        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
        If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then
            CmbSeccion.Enabled = True
        End If
        CmbPlu.Enabled = True
        CmbCodigo.Enabled = True
        'CmbSeccion.Text = ""
        If PorCodigo = True Then
            CmbPlu.TexT = ""
        Else
            'CmbCodigo.Text = ""
        End If
        TxtSubSeccion.TexT = ""
        TxtFamilia.TexT = ""
        TxtPrecio.TexT = ""
        TxtEuros.TexT = ""
        TxtCaducidad.TexT = ""
        TxtTara.TexT = ""
        TxtEAN13.TexT = ""
        'If (UsaEcoPrint And TipoEcoPrint = 3) Or lMarteAlfa Then 'IC+
        '    TextICEAN.TexT = ""
        'End If
        'TextICEAN.TexT = "0"
        For bucle = 0 To 20
            Text8(bucle).TexT = ""
            If bucle < 10 Then CmbTipoLetra(bucle).TexT = "0"
        Next bucle
        #If BALSAM Then
            For bucle = 0 To 8
                Text3(bucle).TexT = ""
            Next bucle
        #Else
            For bucle = 0 To 9
                Text3(bucle).TexT = ""
            Next bucle
        #End If
        ChkTramo.Value = vbUnchecked
        TxtPrc(0) = ""
        TxtTramo(0) = ""
        TxtPrc(1) = ""
        TxtTramo(1) = ""
        TxtPrc(0).Enabled = False
        TxtTramo(0).Enabled = False
        TxtPrc(1).Enabled = False
        TxtTramo(1).Enabled = False
        'If Dir(App.Path & "\hiper") <> "" Then '**********************************cas.v160
        '    'CmbPlu.Text = CStr((db.carga_plu_max(Me.CmbCodigo.Text))) + 2
        '    If (CStr(Me.CmbSeccion.TexT) <> "") And (CStr(Me.CmbCodigo.TexT) <> "") Then
        '            Set db = OpenDatabase(Base_General)
        '            sSQL = "SELECT plu " & _
                     '                           "FROM articulo " & _
                     '                           "WHERE borrado=false" & _
                     '                           " AND secc_maqui=" & CStr(Me.CmbSeccion.TexT)
        '            Set rst = db.OpenRecordset(sSQL)
        '            With rst
        '                bPlu = 1
        '                If Not .EOF Then
        '                    For bPlu = 1 To 9999
        '                        .MoveFirst
        '                        ret = True
        '                        Do While Not .EOF
        '                            If bPlu = rst.Fields("plu") Then
        '                                ret = False
        '                                Exit Do
        '                            End If
        '                            .Movenext
        '                        Loop
        '                        If ret Then Exit For
        '                    Next bPlu
        '                End If
        '                .Close
        '            End With
        '            Set rst = Nothing
        '            db.Close
        '            Set db = Nothing
        '            '1.7.2
        '            'Workspaces(0).close
        '            Me.CmbPlu.TexT = bPlu
        '    End If
        'End If
    End If

    'MsgBox "des1:" & Text8(1)
    lCharging = False
End Sub




Private Sub Check1_Click()
    Dim nf As Integer
    If Check1.Value Then
        nf = FreeFile()
        Open App.Path & "\5l.txt" For Output As #nf
        Close #nf
        Check2.Value = vbUnchecked
        Check3.Value = vbUnchecked
    Else
        If Dir(App.Path & "\5l.txt") <> "" Then
            Kill App.Path & "\5l.txt"
        End If
    End If

End Sub
Private Sub Check2_Click()    ''
    Dim nf As Integer
    If Check2.Value Then
        nf = FreeFile()
        Open App.Path & "\40l.txt" For Output As #nf
        Close #nf
        Check1.Value = vbUnchecked
        Check3.Value = vbUnchecked
    Else
        If Check3.Value = vbUnchecked Then
            Check2.Value = vbChecked
        Else
            If Dir(App.Path & "\40l.txt") <> "" Then
                Kill App.Path & "\40l.txt"
            End If
        End If
    End If
End Sub

Private Sub Check3_Click()
    Dim nf As Integer
    If Check3.Value Then
        nf = FreeFile()
        Open App.Path & "\39l.txt" For Output As #nf
        Close #nf
        Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
    Else
        If Dir(App.Path & "\39l.txt") <> "" Then
            Kill App.Path & "\39l.txt"
        End If
        Check2.Value = vbChecked
    End If
End Sub

Private Sub Chk100g_Click()
    #If BALSAM = False Then
    #Else
        If Chk100g.Value = vbChecked Then
            If Chk100g.Visible = True Then
                CmbPesado.Enabled = False
                CmbPesado.ListIndex = 0
            End If
        Else
            CmbPesado.Enabled = True
        End If
    #End If
End Sub

Private Sub ChkTramo_Click()
    If ChkTramo.Value = vbChecked Then
        'If Dir(App.Path & "\digitonuba.txt") <> "" Then
        '    'TxtTramo(0).Enabled = True
        '    'TxtTramo(1).Enabled = True
        '    TxtPrc(0).Enabled = True
        '    TxtPrc(1).Enabled = True
        '    'TxtTramo(0).Locked = False
        '    'TxtTramo(1).Locked = False
        '    TxtPrc(0).Locked = False
        '    TxtPrc(1).Locked = False
        '    If UsaEuro Then
        '        'lblPrecioBase.Caption = "Precio Base " & TxtEuros.Text
        '        lblPrecioBase.Caption = CargaCadena(649) & " " & TxtEuros.TexT
        '    Else
        '        'lblPrecioBase.Caption = "Precio Base " & TxtPrecio.Text
        '        lblPrecioBase.Caption = CargaCadena(649) & " " & TxtPrecio.TexT
        '    End If
        '
        'Else
        TxtTramo(0).Enabled = True
        TxtTramo(1).Enabled = True
        TxtPrc(0).Enabled = True
        TxtPrc(1).Enabled = True
        TxtTramo(0).Locked = False
        TxtTramo(1).Locked = False
        TxtPrc(0).Locked = False
        TxtPrc(1).Locked = False
        If UsaEuro Then
            'lblPrecioBase.Caption = "Precio Base " & TxtEuros.Text
            lblPrecioBase.Caption = CargaCadena(649) & " " & TxtEuros.TexT
        Else
            'lblPrecioBase.Caption = "Precio Base " & TxtPrecio.Text
            lblPrecioBase.Caption = CargaCadena(649) & " " & TxtPrecio.TexT
        End If
        'End If
    Else
        'If Dir(App.Path & "\digitonuba.txt") <> "" Then

        'Else

        TxtTramo(0).Enabled = False
        TxtTramo(1).Enabled = False
        TxtPrc(0).Enabled = False
        TxtPrc(1).Enabled = False
        TxtTramo(0).TexT = ""
        TxtTramo(1).TexT = ""
        TxtPrc(0).TexT = ""
        TxtPrc(1).TexT = ""
        lblPrecioBase.Caption = CargaCadena(649)

        'End If
    End If
End Sub



Private Sub CmbCodigo_Change()
    LblInfo2.Caption = ""
    PorCodigo = True
    Refresca_Datos
    PorCodigo = False
    Ver_Multi

End Sub

'Private Sub CmbBalenv_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If TxtConservacion.Visible = True Then
'            TxtConservacion.SetFocus
'       Else
'           SSTab2.Tab = 0
'            Text8(0).SetFocus
'        End If
'    End If
'End Sub

'Private Sub CmbCodigo_Change()
'    sTxtCod = CmbCodigo.TexT
'End Sub

Private Sub CmbCodigo_Click()
    Dim Token As Long
    Token = InitGDIPlus
    Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
    FreeGDIPlus Token    '
    PorCodigo = True
    LblInfo2.Caption = ""
    Refresca_Datos
    PorCodigo = False
    Ver_Multi
End Sub

'Private Sub CmbCodigo_GotFocus()
'    If CmbCodigo.TexT <> sTxtCod Then
'        sTxtCod = sTxtCod
'        CmbCodigo.TexT = sTxtCod
'        Refresca_Datos (True)
'    End If
'End Sub

Private Sub CmbCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorCodigo = True
        Refresca_Datos
        PorCodigo = False
        'CmbCodigo_LostFocus
        Ver_Multi
        If Val(CmbSeccion.TexT) = 0 Then
            CmbPlu.Enabled = False
        End If
        If CmbPlu.Enabled = False Then
            CmbFam.SetFocus
            Exit Sub
        Else
            CmbSeccion.SetFocus
            Exit Sub
        End If
    End If
    CmbCodigo.Locked = Checktexto(KeyAscii, 6, CmbCodigo.TexT, True, False)
End Sub
'\\\\\\\\\\\\\\\\
'////////////////
'PROBLEMA REFRESCO...
'Private Sub CmbCodigo_LostFocus()
'    If Nocarga = False Then
'        LblInfo2.Caption = ""
'        Refresca_Datos (True)
'        Ver_Multi
'        If Dir(App.Path & "\hiper") <> "" Then
'            CmbPlu.Enabled = False
'        End If
'    End If
'End Sub




'Private Sub CmbCodigo_LostFocus()
'    'If Nocarga = False Then
'        LblInfo2.Caption = ""
'        PorCodigo = True
'        Refresca_Datos
'        PorCodigo = False
'        Ver_Multi
'        If Dir(App.Path & "\hiper") <> "" Then
'            CmbPlu.Enabled = False
'        End If
'    'End If
'
'End Sub

'\\\\\\\\\\\\\\\\
'////////////////
Private Sub CmbPesado_Click()
    If CmbPesado.TexT = CmbPesado.List(0) Then
        LblTramo1.Caption = CargaCadena(651) & " 1 (KG)"
        LblTramo2.Caption = CargaCadena(651) & " 2 (KG)"
    Else
        LblTramo1.Caption = CargaCadena(651) & " 1 (Ud.)"
        LblTramo2.Caption = CargaCadena(651) & " 2 (Ud.)"
    End If


End Sub

Private Sub CmbPesado_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And CmbBalenv.Visible = True Then
'    CmbBalenv.SetFocus
'Else
    SSTab2.Tab = 0
    Text8(0).SetFocus
    'End If
End Sub

Private Sub CmbPesado_LostFocus()
'2.0.23
'If (CmbPesado.TexT = CmbPesado.List(1)) Then
'    TextPoid.Enabled = True
'Else
'    TextPoid.Enabled = False
'    TextPoid.TexT = "0"
'End If
''''''''

End Sub

Private Sub CmbPlu_Click()
    Dim Token As Long
    Token = InitGDIPlus
    Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
    FreeGDIPlus Token    '
    PorCodigo = False
    EnPlu = True
    LblInfo2.Caption = ""
    Refresca_Datos
    EnPlu = False
    Ver_Multi
End Sub

Private Sub CmbPlu_KeyPress(KeyAscii As Integer)
    If Val(CmbSeccion.TexT) > 0 Then
        If KeyAscii = 13 Then
            Refresca_Datos
            If TxtSubSeccion.Visible = True Then
                TxtSubSeccion.SetFocus
                'Else
                '    CmbSub.SetFocus
            End If
        End If
        If PluGB = 0 Or PluGB = 2000 Then    'Or id = 6000 Then
            CmbPlu.Locked = Checktexto(KeyAscii, 4, CmbPlu.TexT, True, False)
        Else

            CmbPlu.Locked = Checktexto(KeyAscii, 3, CmbPlu.TexT, True, False)
        End If
    Else
        CmbPlu.TexT = ""
        CmbPlu.TexT = "0"
    End If
End Sub

'c2f iv4 ic+ es tipoecoprint=3
Private Sub CmbPlu_LostFocus()
    Dim MyOK As Boolean
    If Val(CmbSeccion.TexT) > 0 Then
        If Not (usaSC10 And TipoC9) Then
            MyOK = True
        Else
            MyOK = Check_C9Plu()
            If Not MyOK Then
                MsgBox CargaCadena(963), vbExclamation    '"Este número de PLU no es válido para C9, y no será enviado a dicho sistema", vbExclamation
            End If
        End If
        If PluGB <> 0 Then
            If IsNumeric(CmbPlu.TexT) Then
                If (Val(CmbPlu.TexT) > PluGB And TipoEcoPrint <> 3 And (Not (lMarteAlfa))) Or (Val(CmbPlu.TexT) > 1000 And TipoEcoPrint = 3 And (Not (lMarteAlfa))) Then
                    MyOK = False
                    MsgBox "PLU : 1 - " & Format(PluGB, "000"), vbCritical
                    If CmbPlu.Enabled = True Then CmbPlu.SetFocus
                End If
            End If
        End If
        If MyOK Then
            LblInfo2.Caption = ""
            Refresca_Datos
        End If
        Ver_Multi
        CmbFam.SetFocus
    Else
        CmbPlu.TexT = "0"
        CmbFam.SetFocus
    End If
End Sub
Private Function Check_C9Plu() As Boolean
    Dim SecEnv As Integer
    Dim MyMaq As Integer
    Dim MyOK As Boolean
    Dim Registro As New RecordNet
    MyOK = True
    If (Me.CmbPlu.TexT <> "") And (var.C9Cod = 1) Then   '////////cas.v118
        If Val(Me.CmbPlu.TexT) > var.IDMAXPLU_C9COD Then
            'MsgBox "El sistema solo puede trabajar con un máximo de 6999 artículos.", vbInformation
            'Me.CmbPlu.Text = ""
            'Me.CmbPlu.SetFocus
            'Exit Function
            MyOK = False
        End If
    Else
        Select Case ModoC9
        Case 0 To 1
            SecEnv = 6
        Case 2
            SecEnv = 8
        End Select
        If IsNumeric(CmbSeccion.TexT) And IsNumeric(CmbPlu.TexT) Then
            Registro.OpenRecordset ("select secc_maqui from seccion where borrado=false and codi_ident=" & Val(CmbSeccion.TexT))
            With Registro
                If Not .EOF Then
                    MyMaq = .Fields("secc_maqui")
                End If
            End With
            Select Case ModoC9
            Case 0
                If MyMaq = 1 Or MyMaq = 2 Or MyMaq = 3 Then
                    If CmbPlu.TexT >= 999 Then MyOK = False
                Else
                    If MyMaq = SecEnv Then
                        If CmbPlu.TexT > 499 Then MyOK = False
                    End If
                End If
            Case 1
                If MyMaq = 1 Or MyMaq = 2 Or MyMaq = 3 Or MyMaq = 4 Or MyMaq = 5 Then
                    If CmbPlu.TexT >= 599 Then MyOK = False
                Else
                    If MyMaq = SecEnv Then
                        If CmbPlu.TexT > 499 Then MyOK = False
                    End If
                End If
            Case 2
                If MyMaq = 1 Or MyMaq = 2 Or MyMaq = 3 Or MyMaq = 4 Or MyMaq = 5 Or MyMaq = 6 Or MyMaq = 7 Then
                    If CmbPlu.TexT >= 399 Then MyOK = False
                Else
                    If MyMaq = SecEnv Then
                        If CmbPlu.TexT > 499 Then MyOK = False
                    End If
                End If
            End Select
        End If

    End If
    Check_C9Plu = MyOK
End Function
Private Sub Refresca_Taras()

    Dim Registro As New RecordNet
    CmbTara.Clear
    CmbTara.AddItem 0

    Registro.OpenRecordset ("select distinct(valor) from tipotara order by valor")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If .Fields("Valor") <> 0 Then CmbTara.AddItem .Fields("Valor")
                .Movenext
            Loop
        End If
    End With

    CmbTara.ListIndex = CmbTara.List(0)
End Sub
Private Sub CmbSeccion_Click()
    Dim bucle As Integer
    Dim SecEnv As Integer
    Dim TxtFam As String
    Dim Registro As New RecordNet
    Dim RegCont As New RecordNet
    Dim NoOcultar As Boolean
    PluGB = 0
    '*********
    ' familias
    '*********
    'If FamiliasFijas Then
    CmbFam.Clear
    Registro.OpenRecordset ("select codi_fam,txt_fam from fam_code order by codi_fam")
    If Registro.EOF Then
        CmbFam.Enabled = False
    Else
        Registro.MoveFirst
        Do Until Registro.EOF
            TxtFam = Registro.Fields("codi_fam")
            If Len(TxtFam) < 4 Then TxtFam = TxtFam & Space(4 - Len(TxtFam))
            TxtFam = TxtFam & "-" & Registro.Fields("txt_fam")
            CmbFam.AddItem TxtFam
            Registro.Movenext
        Loop
        CmbFam.ListIndex = 0
    End If
    'End If
    'If SubSeccionesFijas Then
    '    CmbSub.Clear
    '    Registro.OpenRecordset ("select codi_sub,txt_sub from sub_code where codi_ident=" & Val(CmbSeccion.TexT) & " order by codi_sub")
    '    If Registro.EOF Then
    '        CmbSub.Enabled = False
    '    Else
    '        Registro.MoveFirst
    '        Do Until Registro.EOF
    '            TxtFam = Registro.Fields("codi_sub")
    '            If Len(TxtFam) < 4 Then TxtFam = TxtFam & Space(4 - Len(TxtFam))
    '            TxtFam = TxtFam & "-" & Registro.Fields("txt_sub")
    '            CmbSub.AddItem TxtFam
    '            Registro.Movenext
    '        Loop
    '        CmbSub.ListIndex = 0
    '    End If
    'End If
    If AnulaSeccion Then Exit Sub
    Label14.Caption = ""
    #If BALSAM = True Then
        Chk100g.Visible = True
    #Else
        Chk100g.Visible = False
    #End If
    For bucle = 1 To 20
        Label5(bucle - 1).Enabled = True
        Text8(bucle).Enabled = True
    Next bucle

    Refresca_NumPlus
    Refresca_Plus

End Sub
Public Sub Refresca_NumPlus()

    Dim Registro As New RecordNet
    Dim RegCont As New RecordNet

    SeccionAnterior = CmbSeccion.TexT
    If IsNumeric(CmbSeccion.TexT) Then

        Registro.OpenRecordset ("select secc_maqui from seccion where codi_ident=" & Val(CmbSeccion.TexT))
        If Not Registro.EOF Then
            RegCont.OpenRecordset ("select count(codigo) from articulo where borrado=false and secc_maqui=" & Registro.Fields("secc_maqui"))
            With RegCont
                If Not .EOF Then
                    Label14.Caption = CargaCadena(964) & " " & .Fields(1)
                End If
            End With
        End If

    End If


End Sub


Private Sub CmbSeccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbPlu.SetFocus
        Exit Sub
    End If
End Sub

Private Sub CmbSeccion_LostFocus()
'    FrmLinTxt2040.resetear
    If Val(CmbSeccion.TexT) = 0 Then
        CmbPlu.TexT = "0"
        CmbFam.SetFocus
    End If
End Sub

Private Sub CmbTara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtIVA.SetFocus
        Exit Sub
    End If
    TxtTara.Locked = Checktexto(KeyAscii, 5, TxtTara.TexT, True, True)
End Sub

Private Sub Cmdaceptar_Click()
    Dim Miarticulo As DB_Articulo
    Dim bucle As Integer
    Dim Buclelite As Integer
    Dim Resultado As Integer
    Dim BufBase As String

    Me.CmbCodigo.Enabled = True
    Me.CmbSeccion.Enabled = True
    Me.CmbPlu.Enabled = True
    If UsaEuro Then
        If ToDouble(TxtEuros.TexT) > 99999.99 Then
            MsgBox CargaCadena(1375), vbCritical
            TxtEuros.SetFocus
            Exit Sub
        End If
    Else
        If ToDouble(TxtPrecio.TexT) > (9999999 / (10 ^ decimales)) Then
            MsgBox CargaCadena(1375), vbCritical
            TxtPrecio.SetFocus
            Exit Sub
        End If
    End If
    'If FamiliasFijas Then
    If Not IsNumeric(Trim(left(CmbFam.TexT, 4))) Then
        MsgBox CargaCadena(1205), vbCritical
        Exit Sub
    End If
    'End If
    'If SubSeccionesFijas Then
    '    If Not IsNumeric(Trim(left(CmbSub.TexT, 4))) Then
    '        MsgBox CargaCadena(1212), vbCritical
    '        Exit Sub
    '    End If
    'End If
    If CmbTara.Visible = True Then TxtTara.TexT = CmbTara.TexT
    If TxtEuros.TexT = "" Then TxtEuros.TexT = "0"
    If TxtPrecio.TexT = "" Then TxtPrecio.TexT = "0"
    'If RechazarPlu0 Then
    '    If (UsaEuro = True And CDbl(TxtEuros.TexT) = 0) Or (UsaEuro = False And CDbl(TxtPrecio.TexT) = 0) Then
    '        MsgBox CargaCadena(368), vbCritical
    '        Exit Sub
    '    End If
    'End If
    If TxtIVA.TexT = "" Then
        'If UsaECO2 And Val(TxtIVA.TexT) = 0 Then
        '    TxtIVA.TexT = "1"
        'Else
        TxtIVA.TexT = "0"
        'End If
    Else
        'If UsaECO2 And Val(TxtIVA.TexT) = 0 Then
        '    TxtIVA.TexT = "1"
        'End If
    End If
    If Trim(TxtPresel.TexT) = "" Then TxtPresel.TexT = "0"
    #If BALSAM Then
        If Trim(TxtMerma.TexT) = "" Then TxtMerma.TexT = "0"
    #End If
    'c2f 26/07/05...
    'If Trim(TxtEAN13.Text) <> "" Then
    '    If Len(TxtEAN13.Text) <> 12 Then
    '        MsgBox CargaCadena(548)  '"El código EAN13 debe tener 12 caracteres"
    '        SSTab1.Tab = 1
    '        TxtEAN13.SetFocus
    '        Exit Sub
    '    End If
    'End If
    If CmbCodigo.TexT = "" Then
        LblInfo2.Caption = CargaCadena(547)  '"Introduzca el código del artículo"
        CmbCodigo.SetFocus
        Exit Sub
    End If
    If CmbSeccion.TexT = "" Then
        LblInfo2.Caption = CargaCadena(546)  '"Introduzca la Seccion"
        CmbSeccion.SetFocus
        Exit Sub
    End If


    If ChkTramo.Value = vbChecked Then
        'If Dir(App.Path & "\digitonuba.txt") = "" Then

        '2.0.21
        If TxtTramo(0).TexT = "" Then TxtTramo(0).TexT = "0"
        'If TxtTramo(1).TexT = "" Then TxtTramo(1).TexT = "0"
        If TxtPrc(0).TexT = "" Then TxtPrc(0).TexT = "0"
        'If TxtPrc(1).TexT = "" Then TxtPrc(1).TexT = "0"
        '''''''''''''' Verificación tramos...
        If Val(TxtTramo(0).TexT) <> 0 And Val(TxtTramo(1).TexT) <> 0 Then
            If (TxtTramo(1).TexT <> "" Or TxtPrc(1).TexT <> "") Then
                If (ToDouble(TxtTramo(1).TexT) <= ToDouble(TxtTramo(0).TexT)) Or (ToDouble(TxtPrc(1).TexT) >= ToDouble(TxtPrc(0).TexT)) Then
                    MsgBox CargaCadena(653)  '"Complete los datos del segundo tramo"
                    SSTab1.TabVisible(1) = True
                    Exit Sub

                End If
            End If
        End If
        '''''''''''''''''''''''''''''''''''''

        If TxtTramo(0).TexT = "" And TxtPrc(0).TexT = "" Then
            MsgBox CargaCadena(652)  '"Introduzca al menos un tramo"
            SSTab1.TabVisible(1) = True
            TxtTramo(0).SetFocus
            Exit Sub
        End If
        '2.0.21
        'If (TxtTramo(1).TexT <> "" And TxtPrc(1).TexT = "") Or _
         '(TxtPrc(1).TexT <> "" And TxtTramo(1).TexT = "") Then
        '    MsgBox CargaCadena(653)  '"Complete los datos del segundo tramo"
        '    SSTab1.TabVisible(1) = True
        '    Exit Sub
        'End If
        ''''''''
        If (TxtTramo(0).TexT <> "" And TxtPrc(0).TexT = "") Or _
           (TxtPrc(0).TexT <> "" And TxtTramo(0).TexT = "") Then
            MsgBox CargaCadena(654)  '"Complete los datos del primer tramo"
            SSTab1.TabVisible(1) = True
            Exit Sub
        End If

        'End If

    End If
    'MsgBox "A"
    If CmbPlu.TexT = "" Then CmbPlu.TexT = "0"
    'msgbox "B"
    If TxtConservacion.TexT = "" Then TxtConservacion.TexT = "0"
    'msgbox "C"
    If TxtPrecio.TexT = "" Then TxtPrecio.TexT = "0"
    'msgbox "D"
    If TxtEuros.TexT = "" Then TxtEuros.TexT = "0"
    'msgbox "E"
    If TxtCaducidad.TexT = "" Then TxtCaducidad.TexT = "0"
    'msgbox "F"
    If TxtTara.TexT = "" Then TxtTara.TexT = "0"
    'msgbox "G"
    If TxtSubSeccion.TexT = "" Then TxtSubSeccion.TexT = "0"
    'msgbox "H"
    If TxtFamilia.TexT = "" Then TxtFamilia.TexT = "0"
    'msgbox "I"
    If TxtPreferente.TexT = "" Then TxtPreferente.TexT = "0"
    'msgbox "J"
    If TxtEtiqueta.TexT = "" Then TxtEtiqueta.TexT = "0"
    'msgbox "K"
    If Val(Trim(left(CmbFam.TexT, 4))) = Val(FamiliaEspecial) Then
        TxtEtiqueta.TexT = EtiquetaEspecial
    End If

    'msgbox "Bloque..."

    'If TextICEAN.TexT = "" Then TextICEAN.TexT = "0"
    'If Val(TxtPrecio.TexT) = 0 Then
    '   ' MsgBox "El precio del Artículo es CERO", vbExclamation + vbOKOnly
    '   'MsgBox CargaCadena( 368), vbExclamation + vbOKOnly
    'End If
    If Chk100g.Value = vbChecked Then
        Miarticulo.PRC100G = True
    Else
        Miarticulo.PRC100G = False
    End If
    Miarticulo.codigo = CmbCodigo.TexT
    Miarticulo.Plu = CmbPlu.TexT
    Miarticulo.Mostrador = CmbSeccion.TexT
    Miarticulo.caducidad = TxtCaducidad.TexT
    Miarticulo.ean13 = TxtEAN13.TexT
    Miarticulo.Label2 = Trim(Text1.TexT)
    #If BALSAM = True Then
        Miarticulo.Label1 = Trim(TxtEtiqueta.TexT)

        Miarticulo.ean14 = Trim(Text2.TexT)
    #Else
        Miarticulo.Etiqueta = TxtEtiqueta.TexT
    #End If
    'If FamiliasFijas Then
    Miarticulo.familia = Val(Trim(left(CmbFam.TexT, 4)))
    'Else
    '    Miarticulo.familia = TxtFamilia.TexT
    'End If
    'If SubSeccionesFijas Then
    '    Miarticulo.subsec = Val(Trim(left(CmbSub.TexT, 4)))
    'Else
    Miarticulo.subsec = 0    'TxtSubSeccion.TexT
    'End If
    Miarticulo.GrpConserv = TxtConservacion.TexT
    Miarticulo.Preferente = TxtPreferente.TexT
    Miarticulo.tara = TxtTara.TexT
    Miarticulo.IVA = TxtIVA.TexT
    #If BALSAM Then
        Miarticulo.Merma = TxtMerma.TexT
    #End If
    Miarticulo.Presel = TxtPresel.TexT
    If CmbPesado.TexT = CmbPesado.List(0) Then
        Miarticulo.WGH = False
    Else
        Miarticulo.WGH = True
    End If
    'If CmbBalenv.TexT = CmbBalenv.List(1) Then
    '    Miarticulo.Balenv = True
    'Else
    '    Miarticulo.Balenv = False
    'End If
    If UsaEuro Then
        Miarticulo.precio = ToDouble(TxtEuros.TexT)
    Else
        Miarticulo.precio = ToDouble(TxtPrecio.TexT)
    End If
    If cmbOnKey.TexT = "SI" Then
        Miarticulo.onkey = 1
    Else
        Miarticulo.onkey = 0
    End If
    'msgbox "Tramos..."

    If ChkTramo.Value = vbChecked Then
        'If Dir(App.Path & "\digitonuba.txt") = "" Then
        Miarticulo.TRM1 = ToDouble(TxtTramo(0).TexT)
        Miarticulo.PRC1 = ToDouble(TxtPrc(0).TexT)
        If TxtTramo(1).TexT <> "" Then
            Miarticulo.TRM2 = ToDouble(TxtTramo(1).TexT)
            Miarticulo.PRC2 = ToDouble(TxtPrc(1).TexT)
        Else
            Miarticulo.TRM2 = -1
            Miarticulo.PRC2 = -1
        End If
        '2.0.21
        If Miarticulo.WGH = False And (Miarticulo.TRM1 > 30 Or Miarticulo.TRM2 > 30) Then
            MsgBox "Error en Tramos."
            SSTab1.TabVisible(1) = True
            Exit Sub
        End If
        '''''''

        'Else
        '    Miarticulo.TRM1 = 0
        '    Miarticulo.PRC1 = ToDouble(TxtPrc(0).TexT)
        '    Miarticulo.TRM2 = 0
        '    Miarticulo.PRC2 = ToDouble(TxtPrc(1).TexT)
        'End If
    Else
        Miarticulo.TRM1 = 0    '-1
        Miarticulo.PRC1 = 0    '-1
        Miarticulo.TRM2 = 0    '-1
        Miarticulo.PRC2 = 0    '-1
    End If

    'msgbox "descriptivos..."
    #If BALSAM Then
        For bucle = 0 To 8
            If (IsNull(Text3(bucle).TexT) Or Text3(bucle).TexT = "") Then Text3(bucle).TexT = "0"
            Miarticulo.tarifa(bucle) = ToDouble(Text3(bucle).TexT)
        Next bucle
    #Else
        For bucle = 0 To 9
            If (IsNull(Text3(bucle).TexT) Or Text3(bucle).TexT = "") Then Text3(bucle).TexT = "0"
            Miarticulo.tarifa(bucle) = ToDouble(Text3(bucle).TexT)
        Next bucle
    #End If
    For bucle = 0 To 20
        Miarticulo.Descriptivos(bucle) = Text8(bucle).TexT
    Next bucle
    For bucle = 0 To 10
        Miarticulo.TipoLetra(bucle) = CmbTipoLetra(bucle).TexT
    Next bucle

    'If Dir(App.Path & "\mixmode.txt") <> "" Then
    If CheckMix.Value = 1 Then
        Miarticulo.lMix = True
    Else
        Miarticulo.lMix = False
    End If
    Miarticulo.nPoid = Val(TextPoid.TexT)
    'End If
    'If lVL000 And Dir(App.Path & "\mixmode.txt") = "" Then '2.0.22
    '    Miarticulo.lMix = False
    '    '2.0.24
    '    'If Val(Mid(TextPoid.TexT, Len(Trim(TextPoid.TexT)), 1)) <> 0 Then
    '    '    If Val(Mid(TextPoid.TexT, Len(Trim(TextPoid.TexT)), 1)) >= 5 Then
    '    '
    '    '    Else
    '    '
    '    '    End If
    '    'End If
    '    '''''''
    '    'Miarticulo.nPoid = Val(TextPoid.TexT)
    'End If
    If Dir(App.Path & "\famplu.pos") <> "" Then
        Miarticulo.posicion = Miarticulo.Plu
    Else
        Miarticulo.posicion = 0
    End If

    'MsgBox "sQueNombreImagen"

    sQueNombreImagen = Trim(sQueNombreImagen)
    'If Mid(sQueNombreImagen, 1, 1) <> "#" Then
    'If InStr(1, sQueNombreImagen, "\") <> 0 Then
    Miarticulo.Imagen = sQueNombreImagen
    'Else
    '    Miarticulo.Imagen = sQuePathImagen & sQueNombreImagen
    'End If
    'Else
    '    Miarticulo.Imagen = sQueNombreImagen
    'End If

    Resultado = Alta_Articulo(Miarticulo, , , True)

    'MsgBox "RESULTADO:" & CStr(Resultado)

    If (Resultado = 0) Then
        FrmLinTxt2040.alta
    ElseIf (Resultado = 1) Then
        FrmLinTxt2040.alta
    ElseIf (Resultado = 4) Then
        FrmLinTxt2040.alta
    End If
    Select Case Resultado
    Case 0
        LblInfo2.Caption = CargaCadena(527)  '"Datos Añadidos.Envíe Modificaciones"
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
        CmbCodigo.Enabled = True
        If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then
            CmbSeccion.Enabled = True
        End If
        CmbPlu.Enabled = True
        CmbCodigo.TexT = ""
        CmbPlu.TexT = ""
        SSTab1.Tab = 0
        SSTab2.Tab = 0
        Refresca_Codigos
        Refresca_Plus
        Refresca_Datos
        CmbPlu.TexT = ""
    Case 1
        LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados.Envíe Modificaciones"
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
        CmbCodigo.Enabled = True
        If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then
            CmbSeccion.Enabled = True
        End If
        CmbPlu.Enabled = True
        CmbCodigo.TexT = ""
        CmbPlu.TexT = ""
        SSTab1.Tab = 0
        SSTab2.Tab = 0
        Refresca_Codigos
        Refresca_Plus
        Refresca_Datos

    Case 4
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
        CmbCodigo.Enabled = True
        If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then
            CmbSeccion.Enabled = True
        End If
        CmbPlu.Enabled = True
        CmbCodigo.TexT = ""
        CmbPlu.TexT = ""
        Me.TxtIVA.TexT = ""
        ChkTramo.Value = vbUnchecked
        TxtTramo(0).Enabled = False
        TxtTramo(1).Enabled = False
        TxtPrc(0).Enabled = False
        TxtPrc(1).Enabled = False
        TxtTramo(0).TexT = ""
        TxtTramo(1).TexT = ""
        TxtPrc(0).TexT = ""
        TxtPrc(1).TexT = ""
        SSTab1.Tab = 0
        SSTab2.Tab = 0
        Refresca_Codigos
        Refresca_Plus
        Refresca_Datos
        Me.CmbPlu.TexT = ""

    Case 17
        ' MsgBox "No es posible incluir más artículos con precios por tramos. El máximo es de 200", vbInformation
        MsgBox CargaCadena(655), vbInformation
    Case 60
        MsgBox CargaCadena(1202), vbCritical
        Refresca_Codigos
        Refresca_Plus
        Refresca_Datos
        Refresca_NumPlus

    End Select

    If CmbCodigo.Enabled = True Then
        CmbCodigo.SetFocus
    Else
        If CmbPlu.Enabled = True Then
            CmbPlu.SetFocus
        End If
    End If

    If Resultado = 0 Or Resultado = 1 Then
        Refresca_NumPlus
        ModificacionesPendientes(0) = ModificacionesPendientes(0) Or 128
    End If
    'If AhorraMas Or Dir(App.Path & "\upper.txt") <> "" Then
    TextUM.TexT = ""
    'End If
    If Nocarga = False Then
        Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
    End If
    Chk100g.Value = vbUnchecked
    Text1.TexT = ""
    Text2.TexT = ""
End Sub

Private Sub CmdBorrar_Click()
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim Reg3 As New RecordNet
    Dim Miarticulo As DB_Articulo
    Dim Token As Long

    If CmbCodigo.TexT = "" Then Exit Sub
    Me.CmbCodigo.Enabled = True
    Me.CmbSeccion.Enabled = True
    Me.CmbPlu.Enabled = True
    Miarticulo.codigo = CmbCodigo.TexT
    If Baja_Articulo(Miarticulo) = 0 Then
        Load FrmLinTxt2040
        FrmLinTxt2040.codigo (Me.CmbCodigo.TexT)
        FrmLinTxt2040.Plu (Me.CmbPlu.TexT)
        FrmLinTxt2040.Mostrador (Me.CmbSeccion.TexT)
        FrmLinTxt2040.baja
        LblInfo2.Caption = CargaCadena(528)  '"Datos Eliminados."
        CmbCodigo.TexT = ""
        CmbPlu.TexT = ""
        Refresca_Codigos
        Refresca_Plus
        Refresca_Datos
        CmbPlu.TexT = ""
    End If
    Refresca_NumPlus
    If Dir(App.Path & "\hiper") <> "" Then    '***cas.v160
        CmbPlu.Enabled = False
    End If
    'If AhorraMas Or Dir(App.Path & "\upper.txt") <> "" Then
    TextUM.TexT = ""
    'End If
    Token = InitGDIPlus
    Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
    FreeGDIPlus Token    '

End Sub

'Private Sub CmdColor_Click(Index As Integer)
'    Select Case Index
'    Case 0
'    sQueNombreImagen = "#FFE0E0E0"
'    Case 1
'    sQueNombreImagen = "#FFC1C1C1"
'    Case 2
'    sQueNombreImagen = "#FFA2A2A2"
'    Case 3
'    sQueNombreImagen = "#FF6D6F81"
'    Case 4
'    sQueNombreImagen = "#FF838383"'
'    Case 5
'    sQueNombreImagen = "#FF646464"
'    Case 6
'    sQueNombreImagen = "#FF454545"
'    Case 7
'    sQueNombreImagen = "#FF000000"
'    Case 8
'    sQueNombreImagen = "#FFF0FF00"
'    Case 9
'    sQueNombreImagen = "#FFC7DE28"
'    Case 10
'    sQueNombreImagen = "#FFFFDA28"
'    Case 11
'    sQueNombreImagen = "#FFBBA125"
'    Case 12
'    sQueNombreImagen = "#FFFFA500"
'    Case 13
'    sQueNombreImagen = "#FFFF7E34"
'    Case 14
'    sQueNombreImagen = "#FFFF7400"
'    Case 15
'    sQueNombreImagen = "#FF974500"
'    Case 16
'    sQueNombreImagen = "#FFC2FF49"
'    Case 17
'    sQueNombreImagen = "#FF4EFB22"
'    Case 18
'    sQueNombreImagen = "#FF16C700"
'    Case 19
'    sQueNombreImagen = "#FF00B761"
'    Case 20
'    sQueNombreImagen = "#FF00C7B4"
'    Case 21
'    sQueNombreImagen = "#FF00BAFF"
'    Case 22
'    sQueNombreImagen = "#FF00C7B4"
'    Case 23
'    sQueNombreImagen = "#FF242DA3"
'    Case 24
'    sQueNombreImagen = "#FFC088FF"
'    Case 25
'    sQueNombreImagen = "#FFE082FF"
'    Case 26
'    sQueNombreImagen = "#FFD555FF"
'    Case 27
'    sQueNombreImagen = "#FFC518FF"
'    Case 28
'    sQueNombreImagen = "#FFFF00BE"
'    Case 29
'    sQueNombreImagen = "#FFFF005F"
'    Case 30
'    sQueNombreImagen = "#FFE20000"
'    Case 31
'    sQueNombreImagen = "#FFB70000"
'    End Select
'    picture1.Picture = LoadPicture(App.Path & "\images\" & CStr(Index + 1) & ".jpg")
'    picture1.Width = CmdColor(0).Width
'    picture1.Height = CmdColor(0).Height
'    picture1.left = CmdColor(0).left
'    picture1.top = CmdColor(0).top - CmdColor(0).Height - 100
'End Sub
'Private Sub ColorToImage(ByVal sColor As String)
'Dim nI As Integer
'    Select Case sColor
'    Case "#FFE0E0E0"
'    nI = 1
'    Case "#FFC1C1C1"
'    nI = 2
'    Case "#FFA2A2A2"
'    nI = 3
'    Case "#FF6D6F81"
'    nI = 4
'    Case "#FF838383"
'    nI = 5
'    Case "#FF646464"
'    nI = 6
'    Case "#FF454545"
'    nI = 7
'    Case "#FF000000"
'    nI = 8
'    Case "#FFF0FF00"
'    nI = 9
'    Case "#FFC7DE28"
'    nI = 10
'    Case "#FFFFDA28"
'    nI = 11
'    Case "#FFBBA125"
'    nI = 12
'    Case "#FFFFA500"
'    nI = 13
'    Case "#FFFF7E34"
'    nI = 14
'    Case "#FFFF7400"
'    nI = 15
'    Case "#FF974500"
'    nI = 16
'    Case "#FFC2FF49"
'    nI = 17
'    Case "#FF4EFB22"
'    nI = 18
'    Case "#FF16C700"
'    nI = 19
'    Case "#FF00B761"
'    nI = 20
'    Case "#FF00C7B4"
'    nI = 21
'    Case "#FF00BAFF"
'    nI = 22
'    Case "#FF00C7B4"
'    nI = 23
'    Case "#FF242DA3"
'    nI = 24
'    Case "#FFC088FF"
'    nI = 25
'    Case "#FFE082FF"
'    nI = 26
'    Case "#FFD555FF"
'    nI = 27
'    Case "#FFC518FF"
'    nI = 28
'    Case "#FFFF00BE"
'    nI = 29
'    Case "#FFFF005F"
'    nI = 30
'    Case "#FFE20000"
'    nI = 31
'    Case "#FFB70000"
'    nI = 32
'    Case Else
'    nI = 1
'    End Select
'    picture1.Picture = LoadPicture(App.Path & "\images\" & CStr(nI) & ".jpg")
'    picture1.Width = CmdColor(0).Width
'    picture1.Height = CmdColor(0).Height
'    picture1.left = CmdColor(0).left
'    picture1.top = CmdColor(0).top - CmdColor(0).Height - 100
'End Sub

Private Sub CmdLimpiar_Click()
    LblInfo2.Caption = ""
    CmbCodigo.TexT = ""
    CmbPlu.TexT = ""
    Refresca_Datos
    CmbSeccion_Click
    If Dir(App.Path & "\hiper") <> "" Then    '***cas.v160
        CmbPlu.Enabled = False
    End If
    FrmLinTxt2040.resetear
    Me.CmbCodigo.Enabled = True
    Me.CmbSeccion.Enabled = True
    Me.CmbPlu.Enabled = True
End Sub
'''''''
'''''''
Private Sub cmdLin2040_Click()
    Dim cnt As Long
    Dim sCampo As String
    On Error GoTo Ret
    With Me
        If (.CmbCodigo.TexT = "") Then Err.Raise 1000
        If (.CmbSeccion.TexT = "") Then Err.Raise 1001
        'If (.CmbPlu.TexT = "") Then Err.Raise 1002
        Load FrmLinTxt2040
        FrmLinTxt2040.resetear
        FrmLinTxt2040.codigo (.CmbCodigo.TexT)
        FrmLinTxt2040.Plu (.CmbPlu.TexT)
        FrmLinTxt2040.Mostrador (.CmbSeccion.TexT)
        FrmLinTxt2040.Show vbModal
        If ((FrmLinTxt2040.abm = 1) Or (FrmLinTxt2040.abm = 3)) Then
            .CmbCodigo.Enabled = False
            .CmbSeccion.Enabled = False
            .CmbPlu.Enabled = False
        End If
    End With
Ret:
    If Err.Number <> 0 Then
        If (Err.Number = 1000) Then
            sCampo = Me.Label2.Item(0).Caption
        ElseIf (Err.Number = 1001) Then
            sCampo = Me.Label2.Item(1).Caption
        ElseIf (Err.Number = 1002) Then
            sCampo = Me.Label2.Item(2).Caption
        End If
        MsgBox CargaCadena(1376) & sCampo, vbInformation
    End If
End Sub

Private Sub CmdOnKey_Click()
    Dim db As Database
    On Error Resume Next
    Set db = OpenDatabase(Base_General)
    db.Execute "update articulo set onkey=true"
    db.Close
    On Error GoTo 0
    If Dir(App.Path & "\gigante.txt") <> "" Then
        Call generaGIAN
    End If
    If Dir(App.Path & "\40l.txt") = "" And Dir(App.Path & "\39l.txt") = "" And Dir(App.Path & "\5l.txt") = "" Then
        Check2.Value = vbChecked
    End If
    Unload Me
End Sub

Private Sub CmdOrder_Click()
    If Dir(App.Path & "\asign.tst") <> "" Then
        Me.Enabled = False
        Call AsignTec
        Me.Enabled = True
    End If
End Sub

'''''''
'''''''
Private Sub CmdSalir_Click()
    If Dir(App.Path & "\gigante.txt") <> "" Then
        Call generaGIAN
    End If
    If Dir(App.Path & "\40l.txt") = "" And Dir(App.Path & "\39l.txt") = "" And Dir(App.Path & "\5l.txt") = "" Then
        Check2.Value = vbChecked
    End If
    Unload Me
End Sub
Private Sub CmdPL_Click()
    If Dir(App.Path & "\teclasp.exe") <> "" Then

        FrmArticulo2.Hide
        Do_Events

        MostrarFrmLabel FrmPrintLabel

        Do While FrmPrintLabel.Visible = True
            Do_Events
        Loop

        MostrarModal FrmArticulo2
        FrmArticulo2.Enabled = True
    Else
        'MostrarFrmLabel FrmPrintLabelGeneric
        FrmArticulo2.Hide
        Do_Events

        'Original
        MostrarFrmLabel FrmPrintLabelGeneric
        Do While FrmPrintLabelGeneric.Visible = True
            Do_Events
        Loop

        MostrarModal FrmArticulo2
        FrmArticulo2.Enabled = True
    End If

End Sub

'Private Sub Form_Activate()
'    Timer1.Enabled = True
'End Sub

Private Sub Form_Load()
    Dim Registro As New RecordNet
    Dim RegTramo As New RecordNet
    Dim bucle As Integer
    Dim Bucle2 As Integer

    For bucle = 0 To 20
        Text8(bucle).MaxLength = 40
    Next bucle

    sQueNombreImagen = ""
    If Dir(App.Path & "\veronkeyall.ver") <> "" Then
        CmdOnKey.Visible = True
    Else
        CmdOnKey.Visible = False
    End If
    'LabelUM.Visible = False
    'TextUM.Visible = False
    'If Dir(App.Path & "\ahorramas.cfg") <> "" Or Dir(App.Path & "\upper.txt") <> "" Then
    LabelUM.Visible = True
    TextUM.Visible = True
    'End If
    FrmArticulo2.SSTab1.TabCaption(0) = CargaCadena(753)
    FrmArticulo2.SSTab1.TabCaption(1) = CargaCadena(1371)
    FrmArticulo2.LabelPoid.Caption = CargaCadena(130)
    'LabelICEAN.Enabled = False
    'TextICEAN.Enabled = False
    LabelICEAN.Visible = False
    TextICEAN.Visible = False
    'If (UsaEcoPrint And TipoEcoPrint = 3) Or lMarteAlfa Then 'IC+
    '    LabelICEAN.Enabled = True
    '    TextICEAN.Enabled = True
    '    LabelICEAN.Visible = True
    '    TextICEAN.Visible = True
    'End If

    Label1.Caption = ""

    'If Dir(App.Path & "\mixmode.txt") <> "" Then
    '
    LabelMix.Visible = True
    CheckMix.Visible = True
    LabelPoid.Visible = True
    TextPoid.Visible = True
    LabelMix.Enabled = True
    CheckMix.Enabled = True
    LabelPoid.Enabled = True
    TextPoid.Enabled = True
    'LabelPoid.Caption = "Poid :"
    '
    'Else

    '    If lVL000 Then '2.0.22 Peso aproximado
    '        LabelPoid.Visible = True
    '        TextPoid.Visible = True
    '        LabelPoid.Enabled = True
    '        TextPoid.Enabled = True
    '        LabelMix.Visible = False
    '        CheckMix.Visible = False
    '        LabelMix.Enabled = False
    '        CheckMix.Enabled = False
    '        LabelPoid.Caption = "Peso Apr.(Apr. Weight):"
    '        Label1.Caption = "Múltiplo de 10(Multiple of 10)."
    '    Else

    '        LabelMix.Visible = False
    '        CheckMix.Visible = False
    '        LabelPoid.Visible = False
    '        TextPoid.Visible = False
    '        LabelMix.Enabled = False
    '        CheckMix.Enabled = False
    '        LabelPoid.Enabled = False
    '        TextPoid.Enabled = False

    '    End If

    'End If

    'If FamiliasFijas Then
    CmbFam.Visible = True
    TxtFamilia.Visible = False
    'Else
    '    TxtFamilia.Visible = True
    '    CmbFam.Visible = False
    'End If
    'If SubSeccionesFijas Then
    '    CmbSub.Visible = True
    '    TxtSubSeccion.Visible = False
    'Else
    TxtSubSeccion.Visible = False
    CmbSub.Visible = False
    'End If
    'If Dir(App.Path & "\linux.txt") <> "" Then
    Me.Height = 7200 * (1)
    'DrawGradient Me, 0, 16, 118, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    CambiarIdioma

    Text3(9).Visible = False
    Label22(5).Visible = False

    #If BALSAM = True Then
        'Label18.Visible = False
        'Text3(0).Visible = False
        TxtMerma.Visible = True
        Label12.Visible = True
        Label1.Visible = True
    #Else
        TxtMerma.Visible = False
        Label12.Visible = False
        Label1.Visible = False
    #End If
    If UsaEuro Then
        TxtPrecio.Visible = False
        Label2(5).Visible = False
    Else
        TxtEuros.Visible = False
        LblEuros.Visible = False
    End If
    Image1.Picture = Form2.Imagen.Picture
    Me.Icon = Form2.Icon
    'If Not UsaGamaAlta Then
    '    SSTab2.TabVisible(1) = False
    'End If
    'If Not Usa20Lineas Then
    '    For Bucle = 11 To 20
    '        Text8(Bucle).Visible = False
    '        Label5(Bucle - 1).Visible = False
    '    Next Bucle
    'End If
    Label2(5).Caption = CargaCadena(71) & "(" & SimboloMonetario & ")"
    '********************
    ' Ahora no se emplean
    ' Los tipos de letra
    ' de descriptivos
    ' de Euroscale (10-20)
    ' están aquí
    ' Para uso futuro, deshabilitados
    For bucle = 10 To 20
        CmbTipoLetra(bucle).Visible = False
    Next bucle
    Label7(0).Visible = False
    '********************
    'If AhorraMas Then
    '    TxtGrupo(2).Visible = True
    '    TxtGrupo(3).Visible = True
    '    TxtGrupo(4).Visible = True
    'End If

    'If UsaEcoPrint Or UsaECO2 Or UsaECO1 Then
    '    If (Not UsaGamaAlta) And (Not usaSC10) Then
    '        CmbBalenv.Visible = False
    '        Label8.Visible = False
    '        LblConservacion.top = Label8.top
    '        TxtConservacion.top = CmbBalenv.top
    '    End If
    'Else
    '    If Not (UsaGamaAlta And usaSC10) Then
    '        CmbBalenv.Visible = False
    '        Label8.Visible = False
    '        LblConservacion.top = Label8.top
    '        TxtConservacion.top = CmbBalenv.top
    '    End If
    'End If
    For bucle = 0 To 20
        CmbTipoLetra(bucle).Clear
        For Bucle2 = 0 To 3
            CmbTipoLetra(bucle).AddItem Bucle2
        Next Bucle2
        CmbTipoLetra(bucle).TexT = 0
        If bucle > 9 Then CmbTipoLetra(bucle).Enabled = False
    Next bucle
    LblTramo1.Caption = CargaCadena(651) & " 1 (KG)"
    LblTramo2.Caption = CargaCadena(651) & " 2 (KG)"
    CodigoAnterior = 0
    SeccionAnterior = 0
    'If Not UsaGamaAlta Then SSTab1.TabVisible(1) = False
    CmbSeccion.Clear
    If Not (HaySeguridad And UsuarioActual.Asociado <> 0) Then
        Registro.OpenRecordset _
                ("select codi_ident,secc_maqui from seccion where borrado=false")
    Else
        Registro.OpenRecordset _
                ("select codi_ident,secc_maqui from seccion where borrado=false and ((codi_ident=" & UsuarioActual.Asociado & "and codi_ident<>0) or (secc_maqui=" & UsuarioActual.Asociado & "and secc_maqui<>0))")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If .Fields("codi_ident") <> 0 Then
                    CmbSeccion.AddItem .Fields("codi_ident")
                Else
                    CmbSeccion.AddItem .Fields("secc_maqui")
                End If
                .Movenext
            Loop
            CmbSeccion.Enabled = True
        Else
            CmbSeccion.Enabled = False
            CmbPlu.Enabled = False
            CmbCodigo.Enabled = False
            CmdAceptar.Enabled = False
            CmdBorrar.Enabled = False
            LblInfo2.Caption = CargaCadena(219)  '"La sección no existe"
        End If
    End With
    '*****************************************cas.v160
    'If Dir(App.Path & "\hiper") <> "" Then
    '    CmbPlu.Enabled = False
    'End If
    '*****************************************cas.v160
    RegTramo.OpenRecordset ("select count(codigo) from tramos")
    Refresca_Codigos
    If CmbSeccion.ListCount > 0 Then
        CmbSeccion.TexT = CmbSeccion.List(0)
    Else
        LblInfo2.Caption = CargaCadena(399)  '"Debe definir primero Mostradores para poder continuar"
        SSTab1.Enabled = False
        SSTab1.Tab = 0
        CmdBorrar.Enabled = False
        CmdAceptar.Enabled = False
        Exit Sub
    End If
    'If UsaGruposdeConservacion Then
    '    TxtGrupo(0).Visible = True
    '    TxtGrupo(1).Visible = True
    '    Text8(9).Visible = False
    '    Text8(10).Visible = False
    'Else
    TxtConservacion.Visible = False
    LblConservacion.Visible = False
    'End If
    CmbPesado.TexT = CmbPesado.List(0)
    'CmbBalenv.TexT = CmbBalenv.List(0)
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    If HaySeguridad Then
        If ConsultarPermiso(24) = False Then
            CmdAceptar.Enabled = False
            CmdBorrar.Enabled = False
        End If
    End If
    LabelUM.Caption = CargaCadena(1370)
    Label9.Caption = "EAN13"    'CargaCadena(1372)
    'cmdLin2040.Caption = CargaCadena(1373)
    Me.cmdLin2040.Visible = True
    cmdLin2040.Caption = CargaCadena(1373)
    'touchScale Barna --> sin precio por tramos
    'Frame2.Visible = False
    '''''''''''''''''''''''''''''''''''''''''''

    Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")

    Label15.Visible = True
    Text1.Visible = True

    #If BALSAM = True Then
        Check1.Enabled = False
        Check1.Visible = False
        Check2.Visible = False
        Check2.Enabled = False
        Check3.Visible = False
        Check3.Enabled = False
        'Label15.Visible = True
        Label16.Visible = True
        'Text1.Visible = True

        Text2.Visible = True
    #Else
        If Dir(App.Path & "\5l.txt") <> "" Then
            Check1.Value = vbChecked
        End If
        If Dir(App.Path & "\40l.txt") <> "" Then
            Check2.Value = vbChecked
        End If
        If Dir(App.Path & "\39l.txt") <> "" Then
            Check3.Value = vbChecked
        End If

        'Label15.Visible = False
        Label16.Visible = False
        'Text1.Visible = False
        Text2.Visible = False
    #End If
    If Dir(App.Path & "\teclasp.exe") <> "" Then
        CmdPL.Visible = True
    Else
        CmdPL.Visible = False
    End If
    If Dir(App.Path & "\asign.tst") <> "" Then
        CmdOrder.Enabled = True
    Else
        CmdOrder.Enabled = False
    End If
    Me.Width = frmEpelsa.Width
    Me.left = frmEpelsa.left
    Me.top = frmEpelsa.top
End Sub


'Private Sub Form_Terminate()
'Timer1.Enabled = False
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    LockBase = False
    'Timer1.Enabled = False
    'If Dir(App.Path & "\asign.tst") <> "" Then
    '    Call AsignTec
    'End If

    Unload FrmLinTxt2040
    Unload FrmArticulo2
    FrmArticulo2.Hide
    frmEpelsa.Enabled = True
End Sub


Private Sub mnuCambios_Click()
    Aux_Ejecuta "rpcprecios"
End Sub

Private Sub mnuseccion_Click()
    Me.Enabled = False
    Aux_Ejecuta "rpartsec", CmbSeccion.TexT
    Me.Enabled = True
End Sub

Private Sub mnuTodos_Click()
    Me.Enabled = False
    Aux_Ejecuta "rpart"
    Me.Enabled = True
End Sub



Private Sub Picture1_Click()
    Dim Token As Long
    If Val(CmbCodigo.TexT) <> 0 Then
        'sQueNombreImagen = Trim(CmbCodigo.TexT)
        Load FrmImagen
        FrmImagen.Show
        Do While FrmImagen.Visible = True
            Do_Events
        Loop
        On Error Resume Next
        Token = InitGDIPlus
        'Set picture1.Picture = LoadPictureGDIPlus(sQuePathImagen & sQueNombreImagen)
        Set picture1.Picture = LoadPictureGDIPlus(sQueNombreImagen)
        FreeGDIPlus Token    '
        If Err.Number <> 0 Then
            On Error GoTo 0
        Else
            picture1.Height = 3135
            picture1.Width = 3615
            'picture1.top = 0
            'picture1.left = CmdColor(1).left
        End If
        On Error GoTo 0
        picture1.Refresh
    End If
End Sub


Private Sub RTB1_Change(Index As Integer)
    Dim db As dao.Database
    Dim rst As dao.Recordset
    Dim bucle As Integer
    Dim sLongText As String
    If Not lCharging Then
        lCharging = True
        Set db = AbrirBase
        Set rst = db.OpenRecordset("select * from text15 where codigo=" & CStr(Val(CmbCodigo.TexT)))
        If Not rst.EOF Then
            RTB1(0).TextRTF = rst.Fields("text1")
            RTB1(1).TextRTF = rst.Fields("text2")
            RTB1(2).TextRTF = rst.Fields("text3")
            RTB1(3).TextRTF = rst.Fields("text4")
        Else
            RTB1(0).TextRTF = ""
            RTB1(1).TextRTF = ""
            RTB1(2).TextRTF = ""
            RTB1(3).TextRTF = ""
        End If
        rst.Close
        db.Close
        Set db = Nothing
        'RTB1(4).TextRTF = ""

        For bucle = 0 To 20
            sLongText = sLongText & Text8(bucle).TexT
        Next bucle
        RTB1(4).TextRTF = sLongText
        RTB1(4).SelStart = 0
        RTB1(4).SelStart = 1

        Me.Refresh
        lCharging = False
    End If
End Sub

Private Sub RTB1_Click(Index As Integer)
    If Trim(CmbCodigo.TexT) <> "" Then
        Select Case Index
        Case 0
            FrmArtText.ChkT(1).Value = vbChecked
        Case 1
            FrmArtText.ChkT(2).Value = vbChecked

        Case 2
            FrmArtText.ChkT(3).Value = vbChecked

        Case 3
            FrmArtText.ChkT(4).Value = vbChecked
        Case 4
            FrmArtText.ChkT(0).Value = vbChecked
        End Select
        FrmArticulo2.Enabled = False
        FrmArtText.tXTcODIGO = CmbCodigo.TexT
        'Load FrmArtText
        FrmArtText.Show
        'Do While FrmArtText.Visible = True
        '    Do_Events
        'Loop
        FrmArticulo2.Enabled = True
    Else
        CmbCodigo.SetFocus
    End If
End Sub

Private Sub RTB1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text3_Change(Index As Integer)
'If TxtEuros.Text = "," Or TxtEuros.Text = "." Then TxtEuros.Text = ""
    If Text3(Index).Enabled = True Then
        If Text3(Index).TexT = "" Then
            Text3(Index).TexT = ""
        Else
            'On Error Resume Next
            If Trim(Text3(Index).TexT) = "," Or Trim(Text3(Index).TexT) = "." Then
                Text3(Index).TexT = "0"
            Else
                If left(LTrim(Text3(Index).TexT), 1) = "," Then Text3(Index).TexT = "0" & Text3(Index).TexT
                If left(LTrim(Text3(Index).TexT), 1) = "." Then Text3(Index).TexT = "0" & Text3(Index).TexT
            End If
            'If UsaEuro Then Text3(Index).TexT = 1

            'On Error GoTo 0
        End If
    End If

End Sub

Private Sub text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim bucle As Integer
    Dim Comas As Boolean
    'If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    '***********************************
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    '***********************************
    Text3(Index).Locked = Checktexto(KeyAscii, 7, Text3(Index).TexT, True, True)

    If KeyAscii = 8 Then Exit Sub
    Comas = False
    For bucle = 1 To Len(Text3(Index).TexT)
        If Mid(Text3(Index).TexT, bucle, 1) = "," Or Mid(Text3(Index).TexT, bucle, 1) = "." Then
            Comas = True
            Exit For
        End If
    Next bucle
    If Comas Then
        Text3(Index).Locked = Checktexto(KeyAscii, 8, Text3(Index).TexT, True, True)
        If Len(Text3(Index).TexT) - bucle >= 2 Then

            Text3(Index).Locked = True
        End If
    End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
    If lMarteNew Then    '2.0.9 If Dir(App.Path & "\marte.new") <> "" Then
        If CmbPlu.TexT <> "" Then
            If Val(CmbPlu.TexT) > 800 Then
                If CmbCodigo.Enabled Then
                    If Index > 0 Then CmbCodigo.SetFocus
                Else
                    If CmbPlu.Enabled Then
                        If Index > 0 Then CmbPlu.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Dir(App.Path & "\polish.cht") <> "" Or Dir(App.Path & "\greek.cht") <> "" Or Dir(App.Path & "\tk.cht") <> "") And UsaECO2 Then
        If Asc(UCase(Chr(KeyAscii))) <> KeyAscii Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    Else
        If (UsaEcoPrint And TipoEcoPrint = 3) And KeyAscii = Asc("ñ") Then
            KeyAscii = Asc("Ñ")
        End If
        If (Dir(App.Path & "\polish.cht") = "" And Dir(App.Path & "\greek.cht") = "" And Dir(App.Path & "\tk.cht") = "") And UsaECO2 Then
            If Asc(UCase(Chr(KeyAscii))) <> KeyAscii Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If

        End If

    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then
            If CmdAceptar.Enabled = True Then CmdAceptar.SetFocus
            Exit Sub
        End If
        If Not UsaGamaAlta And Index = 9 Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        If Index = 10 And UsaGamaAlta And Not Usa20Lineas Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        If Index = 20 Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        If Index = 8 And (Not UsaGamaAlta Or Not Usa20Lineas) And UsaGruposdeConservacion Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        If Index = 8 And UsaGamaAlta And UsaGruposdeConservacion Then
            SSTab2.Tab = 1
            Text8(11).SetFocus
            Exit Sub
        End If
        If AhorraMas And Index = 11 Then
            Text8(15).SetFocus
            Exit Sub
        End If
        If Index = 9 And UsaGamaAlta Then
            SSTab2.Tab = 1
            Text8(10).SetFocus
            Exit Sub
        End If
        If Index = 9 And Not UsaGamaAlta Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        On Error Resume Next
        Text8(Index + 1).SetFocus
        Exit Sub
    End If
    If Index = 0 Then
        Text8(Index).Locked = Checktexto(KeyAscii, 40, Text8(Index).TexT, False, False)
    Else
        'mrosas --> era 25
        Text8(Index).Locked = Checktexto(KeyAscii, 40, Text8(Index).TexT, False, False)
        ''''''''''''''''''
    End If
End Sub


Private Sub TextICEAN_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    If Val(TextICEAN.TexT) > 999 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextPoid_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    'If lVL000 And Dir(App.Path & "\mixmode.txt") = "" Then
    If Val(TextPoid.TexT) > 15000 Then
        KeyAscii = 0
    End If
    'Else
    '    If Val(TextPoid.TexT) > 9999 Then
    '        KeyAscii = 0
    '    End If
    'End If
End Sub

'Private Sub Timer1_Timer()
'    If frmEpelsa.Visible Then
'        Me.Show
'        Me.Visible = True
'    End If
'End Sub

'Private Sub TextPoid_LostFocus()
'
'    If lVL000 And Dir(App.Path & "\mixmode.txt") = "" Then
'        If Mid(TextPoid.TexT, Len(TextPoid.TexT), 1) <> "0" Then
'            TextPoid.TexT = Mid(TextPoid.TexT, 1, Len(TextPoid.TexT) - 1) & "0"
'       End If
'    End If
'End Sub

Private Sub TxtCaducidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If TxtTara.Visible = True Then
            TxtTara.SetFocus
        Else
            CmbTara.SetFocus
        End If
        Exit Sub
    End If
    TxtCaducidad.Locked = Checktexto(KeyAscii, 3, TxtCaducidad.TexT, True, False)
End Sub



Private Sub TxtEAN13_KeyPress(KeyAscii As Integer)
    Dim Matriz As String
    Dim bucle As Integer
    Dim Conforme As Boolean
    Matriz = "SWZITPCVALFMXDNYQJUKswzitpcvalfmxdnyqjuk"
    TxtEAN13.Locked = False
    If KeyAscii = 8 Then Exit Sub

    If Len(TxtEAN13.TexT) >= 13 Then
        TxtEAN13.Locked = True
        Exit Sub
    End If
    Conforme = False
    If IsNumeric(Chr(KeyAscii)) Then
        Conforme = True
    Else
        'If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32

        For bucle = 1 To Len(Matriz)
            If Mid(Matriz, bucle, 1) = Chr(KeyAscii) Then
                Conforme = True
            End If
        Next bucle
    End If
    If Not Conforme Then
        TxtEAN13.Locked = True
    End If

End Sub
'Private Sub texticean_KeyPress(KeyAscii As Integer)
'    Dim Matriz As String
'    Dim Bucle As Integer
'    Dim Conforme As Boolean
'    Matriz = "0123456789"
'    TextICEAN.Locked = False
'    If KeyAscii = 8 Then Exit Sub
'
'    If Len(TextICEAN.TexT) >= 12 Then
'        TextICEAN.Locked = True
'        Exit Sub
'    End If
'    Conforme = False
'    If IsNumeric(Chr(KeyAscii)) Then
'        Conforme = True
'    Else
'        If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'
'        For Bucle = 1 To Len(Matriz)
'            If Mid(Matriz, Bucle, 1) = Chr(KeyAscii) Then
'                Conforme = True
'            End If
'        Next Bucle
'    End If
'    If Not Conforme Then
'        TextICEAN.Locked = True
'    End If
'
'End Sub


#If BALSAM = False Then

#Else
Private Sub TxtEtiqueta_KeyPress(KeyAscii As Integer)
    TxtEtiqueta.Locked = Checktexto(KeyAscii, 2, TxtEtiqueta.TexT, True, False)
End Sub

Private Sub TxtEtiqueta_LostFocus()
    If Trim(TxtEtiqueta.TexT) <> "" Then
        If Val(TxtEtiqueta.TexT) > 16 Then
            MsgBox CargaCadena(485)  '"El valor de la etiqueta debe estar entre 0 y 16"
            SSTab1.Tab = 1
            TxtEtiqueta.SetFocus
        End If
    End If
End Sub
#End If



Private Sub TxtEuros_Change()
'If TxtEuros.Text = "," Or TxtEuros.Text = "." Then TxtEuros.Text = ""
    If TxtEuros.Enabled = True Then
        If TxtEuros.TexT = "" Then
            TxtPrecio.TexT = ""
        Else
            'On Error Resume Next
            If Trim(TxtEuros.TexT) = "," Or Trim(TxtEuros.TexT) = "." Then
                TxtEuros.TexT = "0"
            Else
                If left(LTrim(TxtEuros.TexT), 1) = "," Then TxtEuros.TexT = "0" & TxtEuros.TexT
                If left(LTrim(TxtEuros.TexT), 1) = "." Then TxtEuros.TexT = "0" & TxtEuros.TexT
            End If
            If UsaEuro Then TxtPrecio.TexT = 1

            'On Error GoTo 0
        End If
    End If
End Sub

Private Sub TxtEuros_KeyPress(KeyAscii As Integer)
    Dim bucle As Integer
    Dim Comas As Boolean
    'If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    '***********************************
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    '***********************************
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtCaducidad.SetFocus
        Exit Sub
    End If
    TxtEuros.Locked = Checktexto(KeyAscii, 7, TxtEuros.TexT, True, True)

    If KeyAscii = 8 Then Exit Sub
    Comas = False
    For bucle = 1 To Len(TxtEuros.TexT)
        If Mid(TxtEuros.TexT, bucle, 1) = "," Or Mid(TxtEuros.TexT, bucle, 1) = "." Then
            Comas = True
            Exit For
        End If
    Next bucle
    If Comas Then
        TxtEuros.Locked = Checktexto(KeyAscii, 8, TxtEuros.TexT, True, True)
        If Len(TxtEuros.TexT) - bucle >= 2 Then

            TxtEuros.Locked = True
        End If
    End If
End Sub

Private Sub TxtFamilia_Change()
    TxtEtiqueta.Enabled = True
    If FamiliaEspecial <> "" Then
        If Val(TxtFamilia.TexT) = Val(FamiliaEspecial) Then
            TxtEtiqueta.Enabled = False
        End If
    End If
End Sub

Private Sub TxtFamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If UsaEuro Then TxtEuros.SetFocus
        If Not UsaEuro Then TxtPrecio.SetFocus
        Exit Sub
    End If
    If KeyAscii = 8 Then
        TxtFamilia.Locked = False
        Exit Sub
    End If
    If Not Supeco Then
        If Len(TxtFamilia.TexT) = 2 Then
            If Val(left(TxtFamilia.TexT, 1)) > 2 Then
                TxtFamilia.Locked = True
                Exit Sub
            Else
                If Val(left(TxtFamilia.TexT, 1)) = 2 And Val(Right(TxtFamilia, 1)) > 5 Then
                    TxtFamilia.Locked = True
                    Exit Sub
                Else
                    If Val(left(TxtFamilia.TexT, 1)) = 2 And Val(Right(TxtFamilia, 1)) = 5 And KeyAscii > 53 Then
                        TxtFamilia.Locked = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    TxtFamilia.Locked = Checktexto(KeyAscii, 3, TxtFamilia.TexT, True, False)
End Sub





Private Sub TxtFamilia_LostFocus()
    TxtEtiqueta.Enabled = True
    If FamiliaEspecial <> "" Then
        If TxtFamilia.TexT <> "" Then
            If Val(TxtFamilia.TexT) = Val(FamiliaEspecial) Then
                TxtEtiqueta.TexT = EtiquetaEspecial
                TxtEtiqueta.Enabled = False
            End If
        End If
    End If
    If TxtFamilia.TexT <> "" Then
        If Val(TxtFamilia.TexT) > 255 And Not Supeco Then
            MsgBox CargaCadena(426)  '"El valor máximo de Familia admitido es 255"
            TxtFamilia.SetFocus
        End If
        If (UsaEcoPrint Or UsaECO2) And Val(TxtFamilia.TexT) > 15 Then
            MsgBox Mid(Trim(CargaCadena(426)), 1, Len(Trim(CargaCadena(426))) - 4) & " : 15"    '"El valor máximo de Familia admitido es 255"
            TxtFamilia.SetFocus

        End If
    End If
End Sub





Private Sub TxtIVA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        CmbPesado.SetFocus
        Exit Sub
    End If
    TxtIVA.Locked = Checktexto(KeyAscii, 1, TxtIVA.TexT, True, False)
    If KeyAscii <> 8 Then
        If UsaEcoPrint And TipoEcoPrint = 3 Then    'IC+
            If KeyAscii < 48 Or KeyAscii > 53 Then TxtIVA.Locked = True
        Else
            If Not (UsaECO2) Then
                If KeyAscii < 48 Or KeyAscii > 52 Then TxtIVA.Locked = True
            Else
                If KeyAscii < 49 Or KeyAscii > 53 Then TxtIVA.Locked = True
            End If
        End If
    End If
End Sub

Private Sub TxtMerma_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyDelete
    Case vbKeyBack
    Case 48 To 57
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub TxtMerma_LostFocus()
    If Trim(TxtMerma.TexT) = "" Then TxtMerma.TexT = "0"
End Sub

Private Sub TxtPrc_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim bucle As Integer
    Dim Comas As Boolean
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 Then TxtTramo(1).SetFocus
        If Index = 1 Then CmdAceptar.SetFocus
        Exit Sub
    End If
    If UsaEuro Then

        TxtPrc(Index).Locked = Checktexto(KeyAscii, 7, TxtPrc(Index).TexT, True, True)

        If KeyAscii = 8 Then Exit Sub
        Comas = False
        For bucle = 1 To Len(TxtPrc(Index).TexT)
            If Mid(TxtPrc(Index).TexT, bucle, 1) = "," Or Mid(TxtPrc(Index).TexT, bucle, 1) = "." Then
                Comas = True
                Exit For
            End If
        Next bucle
        If Comas Then
            TxtPrc(Index).Locked = Checktexto(KeyAscii, 8, TxtPrc(Index).TexT, True, True)
            If Len(TxtPrc(Index).TexT) - bucle >= 2 Then

                TxtPrc(Index).Locked = True
            End If
        End If
    Else

        If decimales > 0 Then
            TxtPrc(Index).Locked = Checktexto(KeyAscii, 7, TxtPrc(Index).TexT, True, True)
        Else
            TxtPrc(Index).Locked = Checktexto(KeyAscii, 7, TxtPrc(Index).TexT, True, False)
        End If
        If KeyAscii = 8 Then Exit Sub
        Comas = False
        For bucle = 1 To Len(TxtPrc(Index).TexT)
            If Mid(TxtPrc(Index).TexT, bucle, 1) = "," Or Mid(TxtPrc(Index).TexT, bucle, 1) = "." Then
                Comas = True
                Exit For
            End If
        Next bucle
        If Comas Then
            TxtPrc(Index).Locked = Checktexto(KeyAscii, 8, TxtPrc(Index).TexT, True, True)
            If Len(TxtPrc(Index).TexT) - bucle >= decimales Then
                TxtPrc(Index).Locked = True
            End If
        End If
    End If
End Sub

Private Sub TxtPrecio_Change()
    If TxtPrecio.Enabled = True Then
        If TxtPrecio.TexT = "" Then
            TxtEuros.TexT = ""
        Else
            'On Error Resume Next
            If Trim(TxtPrecio.TexT) = "," Or Trim(TxtPrecio.TexT) = "." Then
                TxtPrecio.TexT = "0"
            Else
                If left(LTrim(TxtPrecio.TexT), 1) = "," Then TxtPrecio.TexT = "0" & TxtPrecio.TexT
                If left(LTrim(TxtPrecio.TexT), 1) = "." Then TxtPrecio.TexT = "0" & TxtPrecio.TexT
            End If
            If Not UsaEuro Then TxtEuros.TexT = 1    'EUROtoPTA(TxtPrecio.Text)

            'On Error GoTo 0
        End If
    End If
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
    Dim bucle As Integer
    Dim Comas As Boolean
    'If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtCaducidad.SetFocus
        Exit Sub
    End If
    If decimales > 0 Then
        TxtPrecio.Locked = Checktexto(KeyAscii, 7, TxtPrecio.TexT, True, True)
    Else
        TxtPrecio.Locked = Checktexto(KeyAscii, 7, TxtPrecio.TexT, True, False)
    End If
    If KeyAscii = 8 Then Exit Sub
    Comas = False
    For bucle = 1 To Len(TxtPrecio.TexT)
        If Mid(TxtPrecio.TexT, bucle, 1) = "," Or Mid(TxtPrecio.TexT, bucle, 1) = "." Then
            Comas = True
            Exit For
        End If
    Next bucle
    If Comas Then
        TxtPrecio.Locked = Checktexto(KeyAscii, 8, TxtPrecio.TexT, True, True)
        If Len(TxtPrecio.TexT) - bucle >= decimales Then
            TxtPrecio.Locked = True
        End If
    End If
End Sub



Private Sub TxtPreferente_KeyPress(KeyAscii As Integer)
    TxtPreferente.Locked = Checktexto(KeyAscii, 3, TxtPreferente.TexT, True, False)
End Sub

Private Sub TxtPresel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyDelete
    Case vbKeyBack
    Case 48 To 57
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub TxtSubSeccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If TxtFamilia.Visible = True Then
            TxtFamilia.SetFocus
        Else
            If CmbFam.Enabled = True Then CmbFam.SetFocus
        End If
        Exit Sub
    End If
    TxtSubSeccion.Locked = Checktexto(KeyAscii, 3, TxtSubSeccion.TexT, True, False)
End Sub



Private Sub TxtTara_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtIVA.SetFocus
        Exit Sub
    End If
    TxtTara.Locked = Checktexto(KeyAscii, 5, TxtTara.TexT, True, True)
End Sub



Private Sub TxtTramo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Max As Integer
    Dim lDec As Boolean
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    If CmbPesado.TexT = CmbPesado.List(0) Then    'pesado
        Max = 6
        lDec = True
    Else
        Max = 3
        lDec = False
    End If
    TxtTramo(Index).Locked = Checktexto(KeyAscii, Max, TxtTramo(Index).TexT, True, lDec)

    If (Val(TxtTramo(Index).TexT & Chr(KeyAscii)) > 30) And CmbPesado.TexT = CmbPesado.List(0) Then
        TxtTramo(Index).TexT = ""
        TxtPrc(Index).TexT = ""
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtPrc(Index).SetFocus
    End If
End Sub
Private Sub Ver_Multi()

    Dim bucle As Long
    Dim Tiendas() As TTienda_Data
    Dim NTiendas As Long
    Dim nResp As Long
    '!!!
    Lista.Visible = False
    Exit Sub
    HayMulti = True
    '!!!
    If Not HayMulti Then Exit Sub
    Lista.ListItems.Clear
    If Not IsNumeric(CmbCodigo.TexT) Then Exit Sub
    If Not IsNumeric(CmbPlu.TexT) Then Exit Sub
    'If Not IsNumeric(CmbPlu.TexT) Then Exit Sub
    Dame_Tiendas Tiendas, NTiendas
    If NTiendas > 0 Then
        For bucle = 0 To NTiendas - 1
            nResp = ArticuloCompatible(Tiendas(bucle).numero, Val(CmbCodigo.TexT), Val(CmbPlu.TexT), Val(CmbSeccion.TexT))
            Select Case nResp
            Case 0
                Lista.ListItems.Add , "T" & Format(Tiendas(bucle).numero, "000"), Tiendas(bucle).numero, , "no"
            Case 1
                Lista.ListItems.Add , "T" & Format(Tiendas(bucle).numero, "000"), Tiendas(bucle).numero, , "enabled"
            Case Else
                Lista.ListItems.Add , "T" & Format(Tiendas(bucle).numero, "000"), Tiendas(bucle).numero, , "disabled"
            End Select
        Next bucle
    End If

End Sub
Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.SmallIcon <> "disabled" Then
        If Item.SmallIcon = "no" Then
            Lista.ListItems(Item.Key).SmallIcon = "enabled"
        Else
            Lista.ListItems(Item.Key).SmallIcon = "no"
        End If
    End If
End Sub

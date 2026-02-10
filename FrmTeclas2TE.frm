VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTeclas2TE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Teclas TouchScale"
   ClientHeight    =   7305
   ClientLeft      =   2655
   ClientTop       =   2430
   ClientWidth     =   12165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdRetro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Picture         =   "FrmTeclas2TE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   6480
      Width           =   615
   End
   Begin VB.CommandButton CmdAvan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Picture         =   "FrmTeclas2TE.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 12"
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
      Index           =   11
      Left            =   3240
      TabIndex        =   236
      Top             =   6840
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 11"
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
      Index           =   10
      Left            =   1680
      TabIndex        =   235
      Top             =   6840
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 10"
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
      Index           =   9
      Left            =   120
      TabIndex        =   234
      Top             =   6840
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 9"
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
      Index           =   8
      Left            =   3240
      TabIndex        =   233
      Top             =   6480
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 8"
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
      Index           =   7
      Left            =   1680
      TabIndex        =   232
      Top             =   6480
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 7"
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
      Index           =   6
      Left            =   120
      TabIndex        =   231
      Top             =   6480
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 6"
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
      Index           =   5
      Left            =   3240
      TabIndex        =   228
      Top             =   6120
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 5"
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
      Index           =   4
      Left            =   1680
      TabIndex        =   227
      Top             =   6120
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 4"
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
      Index           =   3
      Left            =   120
      TabIndex        =   226
      Top             =   6120
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 3"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   225
      Top             =   5760
      Width           =   1452
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Imprimir"
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
      Left            =   8400
      TabIndex        =   224
      Top             =   6000
      Width           =   1572
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5295
      Left            =   7800
      TabIndex        =   223
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text."
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 2"
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
      Index           =   1
      Left            =   1680
      TabIndex        =   219
      Top             =   5760
      Width           =   1452
   End
   Begin VB.CommandButton CmdParte 
      Caption         =   "Parte 1"
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
      Left            =   120
      TabIndex        =   218
      Top             =   5760
      Width           =   1452
   End
   Begin VB.CommandButton Command3 
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
      Height          =   252
      Left            =   10440
      TabIndex        =   217
      Top             =   6000
      Width           =   1095
   End
   Begin VB.VScrollBar Vertical 
      Height          =   3735
      Left            =   7440
      Max             =   6132
      TabIndex        =   215
      Top             =   1680
      Width           =   375
   End
   Begin VB.HScrollBar Horizontal 
      Height          =   252
      Left            =   120
      Max             =   7572
      TabIndex        =   214
      Top             =   5400
      Width           =   7332
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3852
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   7332
      Begin VB.PictureBox PTeclas 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6132
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   7575
         TabIndex        =   13
         Top             =   0
         Width           =   7572
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   99
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   213
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   98
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   211
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   97
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   209
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   96
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   207
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   95
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   205
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   94
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   203
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   93
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   201
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   92
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   199
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   91
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   197
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   90
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   195
            Text            =   "Text1"
            Top             =   5760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   89
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   193
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   88
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   191
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   87
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   189
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   86
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   187
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   85
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   185
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   84
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   183
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   83
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   181
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   82
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   179
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   81
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   177
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   80
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   175
            Text            =   "Text1"
            Top             =   5160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   79
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   173
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   78
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   171
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   77
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   169
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   76
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   167
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   75
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   165
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   74
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   163
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   73
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   161
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   72
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   159
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   71
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   157
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   70
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   155
            Text            =   "Text1"
            Top             =   4560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   69
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   153
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   68
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   151
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   67
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   66
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   147
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   65
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   145
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   64
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   143
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   63
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   141
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   62
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   139
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   61
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   137
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   60
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   135
            Text            =   "Text1"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   59
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   133
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   58
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   131
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   57
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   129
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   56
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   55
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   54
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   53
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   121
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   52
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   119
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   51
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   117
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   50
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   115
            Text            =   "Text1"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   49
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   113
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   48
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   111
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   47
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   109
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   46
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   107
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   45
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   105
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   44
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   103
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   43
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   101
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   42
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   41
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   97
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   40
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   95
            Text            =   "Text1"
            Top             =   2760
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   39
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   93
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   38
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   37
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   36
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   35
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   34
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   33
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   32
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   31
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   30
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   75
            Text            =   "Text1"
            Top             =   2160
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   29
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   28
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   27
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   26
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   25
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   24
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   23
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   22
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   21
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   20
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   1560
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   19
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   18
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   17
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   16
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   15
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   14
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   13
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   12
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   11
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   10
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   960
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   9
            Left            =   6600
            OLEDropMode     =   1  'Manual
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   8
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   7
            Left            =   5160
            OLEDropMode     =   1  'Manual
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   6
            Left            =   4440
            OLEDropMode     =   1  'Manual
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   5
            Left            =   3720
            OLEDropMode     =   1  'Manual
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   4
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   3
            Left            =   2280
            OLEDropMode     =   1  'Manual
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   2
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   1
            Left            =   840
            OLEDropMode     =   1  'Manual
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   99
            Left            =   6600
            TabIndex        =   212
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   98
            Left            =   5880
            TabIndex        =   210
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   97
            Left            =   5160
            TabIndex        =   208
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   96
            Left            =   4440
            TabIndex        =   206
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   95
            Left            =   3720
            TabIndex        =   204
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   94
            Left            =   3000
            TabIndex        =   202
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   93
            Left            =   2280
            TabIndex        =   200
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   92
            Left            =   1560
            TabIndex        =   198
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   91
            Left            =   840
            TabIndex        =   196
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   90
            Left            =   120
            TabIndex        =   194
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   89
            Left            =   6600
            TabIndex        =   192
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   88
            Left            =   5880
            TabIndex        =   190
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   87
            Left            =   5160
            TabIndex        =   188
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   86
            Left            =   4440
            TabIndex        =   186
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   85
            Left            =   3720
            TabIndex        =   184
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   84
            Left            =   3000
            TabIndex        =   182
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   83
            Left            =   2280
            TabIndex        =   180
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   82
            Left            =   1560
            TabIndex        =   178
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   81
            Left            =   840
            TabIndex        =   176
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   80
            Left            =   120
            TabIndex        =   174
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   79
            Left            =   6600
            TabIndex        =   172
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   78
            Left            =   5880
            TabIndex        =   170
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   77
            Left            =   5160
            TabIndex        =   168
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   76
            Left            =   4440
            TabIndex        =   166
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   75
            Left            =   3720
            TabIndex        =   164
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   74
            Left            =   3000
            TabIndex        =   162
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   73
            Left            =   2280
            TabIndex        =   160
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   72
            Left            =   1560
            TabIndex        =   158
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   71
            Left            =   840
            TabIndex        =   156
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   70
            Left            =   120
            TabIndex        =   154
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   69
            Left            =   6600
            TabIndex        =   152
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   68
            Left            =   5880
            TabIndex        =   150
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   67
            Left            =   5160
            TabIndex        =   148
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   66
            Left            =   4440
            TabIndex        =   146
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   65
            Left            =   3720
            TabIndex        =   144
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   64
            Left            =   3000
            TabIndex        =   142
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   63
            Left            =   2280
            TabIndex        =   140
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   62
            Left            =   1560
            TabIndex        =   138
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   61
            Left            =   840
            TabIndex        =   136
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   60
            Left            =   120
            TabIndex        =   134
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   59
            Left            =   6600
            TabIndex        =   132
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   58
            Left            =   5880
            TabIndex        =   130
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   57
            Left            =   5160
            TabIndex        =   128
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   56
            Left            =   4440
            TabIndex        =   126
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   55
            Left            =   3720
            TabIndex        =   124
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   54
            Left            =   3000
            TabIndex        =   122
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   53
            Left            =   2280
            TabIndex        =   120
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   52
            Left            =   1560
            TabIndex        =   118
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   51
            Left            =   840
            TabIndex        =   116
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   50
            Left            =   120
            TabIndex        =   114
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   49
            Left            =   6600
            TabIndex        =   112
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   48
            Left            =   5880
            TabIndex        =   110
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   47
            Left            =   5160
            TabIndex        =   108
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   46
            Left            =   4440
            TabIndex        =   106
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   45
            Left            =   3720
            TabIndex        =   104
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   44
            Left            =   3000
            TabIndex        =   102
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   43
            Left            =   2280
            TabIndex        =   100
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   42
            Left            =   1560
            TabIndex        =   98
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   41
            Left            =   840
            TabIndex        =   96
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   40
            Left            =   120
            TabIndex        =   94
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   39
            Left            =   6600
            TabIndex        =   92
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   38
            Left            =   5880
            TabIndex        =   90
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   37
            Left            =   5160
            TabIndex        =   88
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   36
            Left            =   4440
            TabIndex        =   86
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   35
            Left            =   3720
            TabIndex        =   84
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   34
            Left            =   3000
            TabIndex        =   82
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   33
            Left            =   2280
            TabIndex        =   80
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   32
            Left            =   1560
            TabIndex        =   78
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   31
            Left            =   840
            TabIndex        =   76
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   30
            Left            =   120
            TabIndex        =   74
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   29
            Left            =   6600
            TabIndex        =   72
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   28
            Left            =   5880
            TabIndex        =   70
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   27
            Left            =   5160
            TabIndex        =   68
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   26
            Left            =   4440
            TabIndex        =   66
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   25
            Left            =   3720
            TabIndex        =   64
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   24
            Left            =   3000
            TabIndex        =   62
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   23
            Left            =   2280
            TabIndex        =   60
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   22
            Left            =   1560
            TabIndex        =   58
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   21
            Left            =   840
            TabIndex        =   56
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   20
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   19
            Left            =   6600
            TabIndex        =   52
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   18
            Left            =   5880
            TabIndex        =   50
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   17
            Left            =   5160
            TabIndex        =   48
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   16
            Left            =   4440
            TabIndex        =   46
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   15
            Left            =   3720
            TabIndex        =   44
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   14
            Left            =   3000
            TabIndex        =   42
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   13
            Left            =   2280
            TabIndex        =   40
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   12
            Left            =   1560
            TabIndex        =   38
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   11
            Left            =   840
            TabIndex        =   36
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   9
            Left            =   6600
            TabIndex        =   32
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   8
            Left            =   5880
            TabIndex        =   30
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   7
            Left            =   5160
            TabIndex        =   28
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   6
            Left            =   4440
            TabIndex        =   26
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   5
            Left            =   3720
            TabIndex        =   24
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   4
            Left            =   3000
            TabIndex        =   22
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   3
            Left            =   2280
            TabIndex        =   20
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   2
            Left            =   1560
            TabIndex        =   18
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   16
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   612
         End
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Grupos de Teclas/Equipos"
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
      Height          =   1572
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.TextBox cmbgrupo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   239
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox CmbGrupoA 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   229
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox CmbNumero 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1692
      End
      Begin VB.OptionButton OptTipo 
         Appearance      =   0  'Flat
         Caption         =   "Teclas de un  Equipo"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2292
      End
      Begin VB.OptionButton OptTipo 
         Appearance      =   0  'Flat
         Caption         =   "Teclas de un  Mostrador"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2532
      End
      Begin VB.Label LblGrupo 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   120
         TabIndex        =   230
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblMostrador 
         Caption         =   "Equipo:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Teclado"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4485
      Begin VB.TextBox Txtconsulta 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   600
         TabIndex        =   221
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptSeccion 
         Appearance      =   0  'Flat
         Caption         =   "Por PLU"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1332
      End
      Begin VB.OptionButton OptSeccion 
         Appearance      =   0  'Flat
         Caption         =   "Por Código "
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1332
      End
      Begin VB.ComboBox CmbJuego 
         Height          =   315
         ItemData        =   "FrmTeclas2TE.frx":0884
         Left            =   2640
         List            =   "FrmTeclas2TE.frx":08A6
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox CmbTeclas 
         Height          =   315
         ItemData        =   "FrmTeclas2TE.frx":08C8
         Left            =   3240
         List            =   "FrmTeclas2TE.frx":08D5
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblresul 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   222
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblconsulta 
         Caption         =   "Cod:"
         Height          =   252
         Left            =   120
         TabIndex        =   220
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Tabla :"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Teclas :"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   6015
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   216
      Top             =   6480
      Width           =   6015
   End
End
Attribute VB_Name = "FrmTeclas2TE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ClipLinux As Boolean
Private sLinux As String
Private QueParte As Integer
Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
Private QuantumLeap As Boolean
Private XX As Long
Private YY As Long
Private Type TipoElArticulo
    codigo() As Long
    Plu() As Integer
    Mostrador() As Integer
    cantidad As Long
End Type
Private ElArticulo As TipoElArticulo
Private lss165 As Boolean
'Private lEsSalir As Boolean


Private Sub CambiarIdioma()
    OptSeccion(0).Caption = CargaCadena(709) & " " & CargaCadena(69)
    OptSeccion(1).Caption = CargaCadena(709) & " " & CargaCadena(68)
    OptTipo(0).Caption = CargaCadena(710)
    OptTipo(1).Caption = CargaCadena(711)
    Label2.Caption = CargaCadena(321)
    Label3.Caption = CargaCadena(223)
    LblMostrador.Caption = CargaCadena(147)    '"Equipos:" ''CargaCadena(302)
    CmdParte(0).Caption = "1 a 60"    'CargaCadena(712) & " 1"
    CmdParte(1).Caption = "61 a 120"    'CargaCadena(712) & " 2"
    CmdParte(2).Caption = "121 a 180"    'CargaCadena(712) & " 3"
    CmdParte(3).Caption = "181 a 240"    'CargaCadena(712) & " 4"
    CmdParte(4).Caption = "241 a 300"    'CargaCadena(712) & " 5"
    CmdParte(5).Caption = "301 a 360"    'CargaCadena(712) & " 6"
    CmdParte(6).Caption = "361 a 420"    'CargaCadena(712) & " 7"
    CmdParte(7).Caption = "421 a 480"    'CargaCadena(712) & " 8"
    CmdParte(8).Caption = "481 a 540"    'CargaCadena(712) & " 9"
    CmdParte(9).Caption = "541 a 600"    'CargaCadena(712) & " 10"
    CmdParte(10).Caption = "601 a 660"    'CargaCadena(712) & " 11"
    CmdParte(11).Caption = "661 a 720"    'CargaCadena(712) & " 12"

    Command3.Caption = CargaCadena(7)
    Frame1.Caption = CargaCadena(713)
    Frame2.Caption = CargaCadena(1406) & "/" & CargaCadena(147)  'CargaCadena(714)
    FrmTeclas2T.Caption = CargaCadena(1397)
    FrmTeclas2TE.Caption = CargaCadena(1398)
    CmdPrint.Caption = CargaCadena(466)
    LblGrupo.Caption = CargaCadena(442)
    'CmdConsulta.Caption = CargaCadena(742)
End Sub
Private Sub Refresca_Numero()

    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim Agregar As Boolean
    Dim Buffer As String
    'Select Case OptTipo(0).Value
    '    Case True
    '        Aux_Refresca_Mostradores CmbNumero
    '    Case False
    Aux_Refresca_equipos CmbNumero, True
    'End Select
    If CmbNumero.ListCount > 0 Then
        CmbNumero.TexT = CmbNumero.List(0)
        HabilitarTeclas True
    Else
        HabilitarTeclas False
    End If

End Sub
Private Function Desglosa_Codigos(LaCadena As String) As TipoCodigos
    Dim LosCodigos As TipoCodigos
    Dim Buffer As String
    Dim bucle As Long
    Dim cantidad As Integer
    cantidad = 0
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            cantidad = cantidad + 1
            ReDim Preserve LosCodigos.codigo(cantidad + 1)
            LosCodigos.codigo(cantidad) = Buffer
            LosCodigos.cantidad = cantidad
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Desglosa_Codigos = LosCodigos
End Function

Private Sub Refresca_Datos()
    Dim Modificacion As Boolean

    Dim Registro As New RecordNet
    Dim bucle As Integer
    Dim LaSeccion As TipoCodigos
    Dim Mimaximo As Integer
    Dim MiFactor As Integer
    Dim MiMinimo As Integer

    '1.9.0
    lss165 = False
    ''''''''
    QuantumLeap = True

    If CmbNumero.ListCount = 0 Then Exit Sub

    For bucle = 0 To 99
        Text1(bucle).TexT = ""
        Text1(bucle).ToolTipText = ""
        Text1(bucle).BackColor = vbWhite
    Next bucle

    'If OptTipo(1).Value = True Then
    Registro.OpenRecordset ("select * from equipos where numero_eqp=" & CmbNumero.TexT & " and borrado=false")
    'If Not Registro.EOF Then
    '    lss165 = Registro.Fields("ss165")
    'End If
    'End If

    'If lss165 Then
    '    CmbTeclas.ListIndex = 4
    '    CmbTeclas.Enabled = False
    '    CmbJuego.ListIndex = 0
    '    CmbJuego.Enabled = False
    'Else
    CmbTeclas.Enabled = True
    CmbJuego.Enabled = True
    'End If

    If QueParte = 0 Then
        MiMinimo = 1
        Select Case Val(CmbTeclas.TexT)
        Case 30
            Mimaximo = 30
        Case 56
            Mimaximo = 56
        Case 80
            Mimaximo = 40
        Case 96
            Mimaximo = 50
        Case 165
            If lss165 Then
                Mimaximo = 88
            Else
                Mimaximo = 90
            End If
        Case 240
            Mimaximo = 60
        Case 360
            Mimaximo = 60
        Case 720
            Mimaximo = 60

        End Select
    Else
        Select Case Val(CmbTeclas.TexT)
        Case 30
            Mimaximo = 30
            MiMinimo = 1
        Case 56
            Mimaximo = 56
            MiMinimo = 1
        Case 80
            Mimaximo = 80
            MiMinimo = 41
        Case 96
            Mimaximo = 96
            MiMinimo = 50
        Case 165
            If lss165 Then
                Mimaximo = 165
                MiMinimo = 89
            Else
                Mimaximo = 165
                MiMinimo = 91
            End If
        Case 240
            Select Case QueParte
            Case 1
                MiMinimo = 61
                Mimaximo = 120

            Case 2
                MiMinimo = 121
                Mimaximo = 180

            Case 3
                MiMinimo = 181
                Mimaximo = 240

            End Select

        Case 360
            Select Case QueParte
            Case 1
                MiMinimo = 61
                Mimaximo = 120

            Case 2
                MiMinimo = 121
                Mimaximo = 180

            Case 3
                MiMinimo = 181
                Mimaximo = 240

            Case 4
                MiMinimo = 241
                Mimaximo = 300

            Case 5
                MiMinimo = 301
                Mimaximo = 360
            End Select

        Case 720
            Select Case QueParte
            Case 1
                MiMinimo = 61
                Mimaximo = 120

            Case 2
                MiMinimo = 121
                Mimaximo = 180

            Case 3
                MiMinimo = 181
                Mimaximo = 240

            Case 4
                MiMinimo = 241
                Mimaximo = 300

            Case 5
                MiMinimo = 301
                Mimaximo = 360

            Case 6
                MiMinimo = 361
                Mimaximo = 420
            Case 7
                MiMinimo = 421
                Mimaximo = 480
            Case 8
                MiMinimo = 481
                Mimaximo = 540
            Case 9
                MiMinimo = 541
                Mimaximo = 600
            Case 10
                MiMinimo = 601
                Mimaximo = 660
            Case 11
                MiMinimo = 661
                Mimaximo = 720

            End Select

        End Select
    End If
    If CmbNumero.TexT = "" Then
        For bucle = 0 To 99
            Text1(bucle).TexT = ""

        Next bucle
        QuantumLeap = False
        Exit Sub
    End If
    Modificacion = False


    'If OptTipo(1).Value = True Then
    '    If lss165 Then
    '        If QueParte = 0 Then
    '        Registro.OpenRecordset _
             '                 ("select * from teclas where borrado=false and numero_eqp=" & CmbNumero.TexT & " and prog_tec='TOUCH' and tabla=0 and tran_tecsc10='" & cmbgrupo.TexT & "'")
    '
    '        Else
    '        Registro.OpenRecordset _
             '                 ("select * from teclas where borrado=false and numero_eqp=" & CmbNumero.TexT & " and prog_tec='TOUCH' and tabla=1 and tran_tecsc10='" & cmbgrupo.TexT & "'")
    '        End If
    '    Else
    Registro.OpenRecordset _
            ("select * from teclas where borrado=false and tabla=" & CmbJuego.TexT & " and prog_tec='TOUCH' and tran_tecsc10='" & cmbgrupo.TexT & "'")
    '("select * from teclas where borrado=false and tabla=" & CmbJuego.TexT & "  and numero_eqp=" & CmbNumero.TexT & " and prog_tec='TOUCH' and tran_tecsc10='" & cmbgrupo.TexT & "'")
    '    End If
    With Registro
        If Not .EOF Then

            .MoveFirst
            Do Until .EOF

                'If lss165 Then
                '    If (.Fields("tabla") = 0) Then
                '        If (.Fields("codi_tec")) >= MiMinimo And (.Fields("codi_tec")) <= Mimaximo Then
                '            Text1(.Fields("codi_tec") - MiMinimo).BackColor = vbGreen
                '            Text1(.Fields("codi_tec") - MiMinimo).ToolTipText = .Fields("desc_plu")
                '            If OptSeccion(0).Value = True Then
                '
                '                 Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("codigo")
                '             Else
                '                 Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("plu")
                '             End If
                '         End If
                '     Else
                '         If (.Fields("codi_tec") + 88) >= MiMinimo And (.Fields("codi_tec") + 88) <= Mimaximo Then
                '             Text1(.Fields("codi_tec") - (MiMinimo - 88)).BackColor = vbGreen
                '             Text1(.Fields("codi_tec") - (MiMinimo - 88)).ToolTipText = .Fields("desc_plu")
                '             If OptSeccion(0).Value = True Then
                '
                '                 Text1(.Fields("codi_tec") - (MiMinimo - 88)).TexT = .Fields("codigo")
                '             Else
                '                 Text1(.Fields("codi_tec") - (MiMinimo - 88)).TexT = .Fields("plu")
                '             End If
                '         End If
                '
                '
                '     End If

                ' Else
                If (.Fields("codi_tec")) >= MiMinimo And (.Fields("codi_tec")) <= Mimaximo Then
                    Text1(.Fields("codi_tec") - MiMinimo).BackColor = vbGreen
                    Text1(.Fields("codi_tec") - MiMinimo).ToolTipText = .Fields("desc_plu")
                    If OptSeccion(0).Value = True Then

                        Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("codigo")
                    Else
                        Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("plu")
                    End If
                End If
                'End If
                .Movenext
            Loop
            Modificacion = True

        Else
            For bucle = 0 To 99
                Text1(bucle).TexT = ""
                Text1(bucle).BackColor = vbWhite
            Next bucle
        End If
    End With
    'Else
    '    LaSeccion = Desglosa_Codigos(CmbNumero.TexT)
    '    Registro.OpenRecordset _
         '    ("select * from teclas where borrado=false  and tabla=" & CmbJuego.TexT _
         '    & " and codi_ident=" & LaSeccion.codigo(1) & " and prog_tec='TOUCH' and tran_tecsc10='" & cmbgrupo.TexT & "'")
    '    With Registro
    '        If Not .EOF Then
    '            .MoveFirst
    '            Do Until .EOF
    '                If (.Fields("codi_tec")) >= MiMinimo And (.Fields("codi_tec")) <= Mimaximo Then
    '                    Text1(.Fields("codi_tec") - MiMinimo).BackColor = vbGreen
    '                    Text1(.Fields("codi_tec") - MiMinimo).ToolTipText = .Fields("desc_plu")
    '                    If OptSeccion(0).Value = True Then
    '
    '                        Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("codigo")
    '                    Else
    '
    '                        Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("plu")
    '                    End If
    '                End If
    '            .Movenext
    '        Loop
    '        Modificacion = True
    '        Else
    '            For Bucle = 0 To 99
    '                Text1(Bucle).TexT = ""
    '                Text1(Bucle).BackColor = vbWhite
    '            Next Bucle
    '        End If
    '    End With
    'End If

    QuantumLeap = False
End Sub


'Private Sub CmbGrupo_Click()
'    QueParte = 0
'    Refresca_Teclas
'    Refresca_Datos'
'
'End Sub

Private Sub cmbgrupo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    If KeyAscii = 48 And cmbgrupo.TexT = "" Then
        KeyAscii = 0
    End If
End Sub

'Private Sub cmbgrupo_change()
'    Refresca_Teclas
'    Refresca_Datos
'End Sub

Private Sub CmbJuego_Click()
    QueParte = 0
    Refresca_Teclas
    Refresca_Datos
    'If Val(cmbgrupo.Text) = 8 Or Val(cmbgrupo.Text) = 9 Then
    '    Select Case Val(CmbJuego.Text)
    '    Case 0
    '        lblresul.Caption = "  FRUTA  "
    '    Case 1
    '        lblresul.Caption = "  GRANEL  "
    '    Case Else
    '        lblresul.Caption = ""
    '    End Select
    'Else
    '    lblresul.Caption = ""
    'End If

End Sub

Private Sub CmbNumero_Click()
'Dim Registro As New RecordNet
'CmdConsulta.Visible = False
'If OptTipo(1).Value And IsNumeric(CmbNumero.Text) Then
'Registro.OpenRecordset "select secc_maqui from seccion where codi_ident=" & CStr(Val(CmbNumero.TexT))
'If Not Registro.EOF Then
'CmdConsulta.Visible = True
    cmbgrupo.TexT = CStr(Val(CmbNumero.TexT) + 10)    'CStr(Registro.Fields("secc_maqui"))
'End If
'End If
'If Dir(App.Path & "\ikea.exe") <> "" Then
'    CmdConsulta.Visible = True
'End If
    QueParte = 0

    'Refresca_Numero

    Refresca_Teclas
    Refresca_Datos
    Refresca_Teclas    'añadido
    'If Val(cmbgrupo.Text) = 8 Or Val(cmbgrupo.Text) = 9 Then
    '    Select Case Val(CmbJuego.Text)
    '    Case 0
    '        lblresul.Caption = "  FRUTA  "
    '    Case 1
    '        lblresul.Caption = "  GRANEL  "
    '    Case Else
    '        lblresul.Caption = ""
    '    End Select
    'Else
    '    lblresul.Caption = ""
    'End If

    '''''''''''''''
    refresca_lista
End Sub

Private Sub CmbTeclas_Click()
    QueParte = 0
    Refresca_Numero
    Refresca_Teclas
    Refresca_Datos
End Sub

Private Sub CmdAvan_Click()
    If (QueParte < 11 And Val(CmbTeclas.TexT) = 720) Or (QueParte < 5 And Val(CmbTeclas.TexT) = 360) Or (QueParte < 3 And Val(CmbTeclas.TexT) = 240) Then _
       QueParte = QueParte + 1

    'Refresca_Numero

    Refresca_Teclas
    Refresca_Datos

End Sub


'ADAPTAR TOUCHSCALE BARNA
'Private Sub CmdConsulta_Click()
'    Dim Registro As New RecordNet
'    Dim Bucle As Long
'    Dim BSec As Long
'    Dim Encontrado As Long
'    Dim sCad As String
'    Dim MyTecla As DB_Teclas
'    Dim NTeclas As Long
'    Dim CodTecla As Long
'    Dim NumTecla As Long
'    Encontrado = -1
'    Me.Enabled = False
'    If UsaGamaAlta Then
'        MUE_Sistemas
'        If SisEur.NSistemas > 0 Then
'            Registro.OpenRecordset ("select codi_ident from equipos where numero_eqp=" & Val(CmbNumero.TexT) & " and borrado=false and (modelo=0 or modelo=1)")
'            If Registro.EOF Then
'                MsgBox "Esta Balanza no es Euroscale", vbCritical
'            Else
'                For Bucle = 0 To SisEur.NSistemas - 1
'                    For BSec = 0 To SisEur.Sistemas(Bucle).NSecciones - 1
'                        If SisEur.Sistemas(Bucle).Secciones(BSec) = Registro.Fields("codi_ident") Then
'                            Encontrado = Bucle
'                            Exit For
'                        End If
'                    Next BSec
'                    If Encontrado <> -1 Then Exit For
'                Next Bucle
'                If Encontrado = -1 Then
'                    MsgBox "Esta balanza no es Euroscale", vbCritical
'                Else
'                    LblInfo.Caption = "Conectando..."
'                    frmControl.Switch_mode 700 + Encontrado
'                    If frmControl.EstadoBal = 0 Then
'                        If Dir(App.Path & "\ikea.exe") <> "" Then
'                            sCad = frmControl.Balanza.Query_Keys(0, Val(CmbNumero.TexT), Val(CmbJuego.TexT), 1)
'                        Else
'                            sCad = frmControl.Balanza.Query_Keys(1, Val(CmbNumero.TexT), Val(CmbJuego.TexT), 1)
'                        End If
'                        If sCad = "" Then
'                            MsgBox "Error de comunicaciones", vbCritical
'                        Else
'                            LblInfo.Caption = "recibiendo datos..."
'                            Do_Events
'                            NTeclas = left(sCad, 4)
'                            If NTeclas > 0 Then
'                                sCad = Mid(sCad, 5)
'                                For Bucle = 0 To NTeclas - 1
'                                    NumTecla = left(sCad, 3)
'                                    CodTecla = Mid(sCad, 4, 6)
'                                    sCad = Mid(sCad, 10)
'                                    MyTecla.CodArticulo = CodTecla
'                                    MyTecla.PluArticulo = 0
'                                    MyTecla.SecArticulo = 0
'                                    MyTecla.Destino_CodiTec = NumTecla
'                                    If Dir(App.Path & "\ikea.exe") <> "" Then
'                                        MyTecla.Destino_NMostrador = 0
'                                        MyTecla.Destino_NBalanza = Val(CmbNumero.TexT)
'
'                                    Else
'                                        MyTecla.Destino_NMostrador = 0
'                                        MyTecla.Destino_NBalanza = Val(CmbNumero.TexT)
'                                    End If
'                                    MyTecla.Destino_Tabla = Val(CmbJuego.TexT)
'                                    Alta_Teclas MyTecla
'                                Next Bucle
'                            End If
'                            Refresca_Teclas
'                            Refresca_Datos
'                            refresca_lista
'                            LblInfo.Caption = "OK"
'                            MsgBox "OK", vbInformation
'                        End If
'                    Else
'                        MsgBox "Error de comunicaciones", vbCritical
'                    End If
'                End If
'            End If
'        Else
'            MsgBox "Esta balanza no es Euroscale", vbCritical
'        End If
'    End If
'    LblInfo.Caption = ""
'    Me.Enabled = True
'End Sub

Private Sub CmdParte_Click(Index As Integer)
    QueParte = Index
    'Refresca_Numero
    Refresca_Teclas
    Refresca_Datos
End Sub

Private Sub CmdParte_GotFocus(Index As Integer)
    QuantumLeap = True
End Sub

Private Sub CmdPrint_Click()
    Dim MyCad As String
    Dim PP As TipoCodigos
    'If OptTipo(0).Value And (CmbNumero.TexT <> "") Then
    '    PP = Desglosa_Codigos(CmbNumero.TexT)
    '    MyCad = "1" & PP.codigo(1)
    'End If
    'If OptTipo(1).Value And (CmbNumero.TexT <> "") Then
    MyCad = "0" & CmbNumero.TexT
    'End If
    If MyCad <> "" Then
        Aux_Ejecuta "rteclas", MyCad
    End If
End Sub

Private Sub CmdRetro_Click()
    If QueParte > 0 Then QueParte = QueParte - 1

    'Refresca_Numero
    Refresca_Teclas
    Refresca_Datos

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub HabilitarTeclas(Si_o_no As Boolean)
    Dim bucle As Integer
    For bucle = 0 To 99
        Text1(bucle).Enabled = Si_o_no
    Next bucle
End Sub


Private Sub Consulta_Art()

    Dim Registro As New RecordNet
    Dim RegAux As New RecordNet
    Dim MiCodigo As TipoCodigos

    LblInfo.Caption = ""
    'If OptSeccion(0).Value = True Then
    '    Registro.OpenRecordset ("select des_plu1 from articulo where borrado=false and codigo=" & Txtconsulta.TexT)
    '    If Not Registro.EOF Then
    '        lblresul.Caption = Registro.Fields("des_plu1")
    '    Else
    '        lblinfo.Caption = CargaCadena(218)  '"El código indicado no existe"
    '    End If
    'Else
    If CmbNumero.TexT <> "" Then
        MiCodigo = Desglosa_Codigos(CmbNumero.TexT)
        'If OptTipo(0).Value = True Then
        '    RegAux.OpenRecordset ("select secc_maqui from seccion where borrado=false and codi_ident=" & MiCodigo.codigo(1))
        'Else
        RegAux.OpenRecordset ("select secc_maqui from equipos where borrado=false and numero_eqp=" & MiCodigo.codigo(1))
        'End If
        If Not RegAux.EOF Then
            Registro.OpenRecordset _
                    ("select des_plu1 from articulo where borrado=false and secc_maqui=" & RegAux.Fields("secc_Maqui") & " and borrado=false and plu=" & Txtconsulta.TexT)
            If Not Registro.EOF Then
                lblresul.Caption = Registro.Fields("des_plu1")
            Else
                LblInfo.Caption = CargaCadena(221)  '"El PLU indicado no existe"
            End If
        End If
    End If
    'End If

End Sub
Private Sub refresca_lista()

    Dim Registro As New RecordNet
    Dim Cad1 As String
    Dim MiSec As Integer
    Dim MyError As Boolean
    Dim lAccess As Boolean

    Lista.ListItems.Clear

    Select Case OptSeccion(0).Value
    Case True
        Cad1 = "codigo="
        lAccess = True
    Case False
        Cad1 = "plu="
        lAccess = False
    End Select
    MyError = False
    If Trim(CmbNumero.TexT) <> "" Then
        'Select Case OptTipo(0).Value
        '    Case True
        '        If InStr(1, CmbNumero.TexT, ",") Then
        '            MiSec = Val(left(CmbNumero.TexT, InStr(1, CmbNumero.TexT, ",") - 1))
        '        Else
        '            MiSec = Val(CmbNumero.TexT)
        '        End If
        '        Registro.OpenRecordset ("select secc_maqui from seccion where codi_ident = " & MiSec & " and borrado=false")
        '        With Registro
        '            If Not .EOF Then
        '                MiSec = .Fields("secc_Maqui")
        '            Else
        '                MyError = True
        '            End If
        '        End With
        '    Case False
        Registro.OpenRecordset ("select secc_maqui from equipos where numero_eqp=" & Val(CmbNumero.TexT) & " and borrado=false")
        With Registro
            If Not .EOF Then
                MiSec = .Fields("secc_Maqui")
            Else
                MyError = True
            End If
        End With
        'End Select
    Else
        MyError = True
    End If
    If Not MyError Then
        'c2f 1.7.19... todos los códigos visibles.
        'If OptTipo(0).Value = False Then
        '    If lAccess Then
        Registro.OpenRecordset ("select codigo,plu,des_plu1 from articulo where borrado=false order by codigo")
        '    Else
        '        Registro.OpenRecordset ("select codigo,plu,des_plu1 from articulo where borrado=false order by plu")
        '    End If
        'Else
        '    If lAccess Then
        '        Registro.OpenRecordset ("select codigo,plu,des_plu1 from articulo where borrado=false and secc_maqui=" & MiSec & " order by codigo")
        '    Else
        '        Registro.OpenRecordset ("select codigo,plu,des_plu1 from articulo where borrado=false and secc_maqui=" & MiSec & " order by plu")
        '    End If
        'End If
        With Registro
            If Not .EOF Then
                Do Until .EOF
                    If OptSeccion(0).Value Then
                        Lista.ListItems.Add , , .Fields("codigo")
                    Else
                        Lista.ListItems.Add , , .Fields("plu")
                    End If
                    If Not IsNull(.Fields("des_plu1")) Then Lista.ListItems(Lista.ListItems.Count).SubItems(1) = .Fields("des_plu1")


                    .Movenext
                Loop
            End If
        End With
    End If

End Sub


'Private Sub Command3_GotFocus()
'lEsSalir = True
'End Sub

'Private Sub Command3_LostFocus()
'lEsSalir = False
'End Sub

Private Sub Form_Load()
    Dim nJ As Integer
    Dim Registro As New RecordNet
    Dim Contador As Long
    Dim bucle As Integer
    Dim MyPref As String
    If Dir(App.Path & "\etc\l_keys.txt") <> "" Then
        ClipLinux = True
    End If
    CambiarIdioma
    Me.Icon = Form2.Icon
    Me.Caption = Me.Caption & " TouchScale"
    Image1.Picture = Form2.Imagen.Picture
    Contador = 0
    For bucle = 1 To 99
        Text1(bucle).Appearance = 0
    Next bucle

    'CmbGrupo.Text = CmbGrupo.List(0)
    cmbgrupo.TexT = "11"
    CmbJuego.ListIndex = 0
    ''''''''''''''''''''''''''''''''
    Registro.OpenRecordset ("select codigo,codi_ident,plu from articulo where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Contador = Contador + 1
                ReDim Preserve ElArticulo.codigo(Contador + 1)
                ReDim Preserve ElArticulo.Mostrador(Contador + 1)
                ReDim Preserve ElArticulo.Plu(Contador + 1)
                ElArticulo.codigo(Contador) = .Fields("codigo")
                ElArticulo.Mostrador(Contador) = .Fields("codi_ident")
                ElArticulo.Plu(Contador) = .Fields("plu")
                .Movenext
            Loop
            ElArticulo.cantidad = Contador
        End If
    End With

    If ElArticulo.cantidad = 0 Then
        LblInfo.Caption = CargaCadena(715)
        HabilitarTeclas False

    End If
    XX = PTeclas.Width
    YY = PTeclas.Height
    Horizontal.Max = XX / 100
    Vertical.Max = YY / 100
    Horizontal.Value = 0
    Vertical.Value = 0

    For bucle = 0 To 99
        Label1(bucle).Caption = bucle + 1
        Text1(bucle).TexT = ""
    Next bucle
    CmbTeclas.ListIndex = 2

    'CmbJuego.TexT = CmbJuego.List(0)
    OptTipo(0).Value = True
    OptTipo(1).Value = False
    OptSeccion(0).Value = True
    MyPref = RPreferencias("TECLAS    ", 1)
    On Error Resume Next
    If MyPref <> "" Then
        If left(MyPref, 1) = "1" Then OptSeccion(1).Value = True
        If Mid(MyPref, 2, 1) = "1" Then OptTipo(1).Value = True
        CmbTeclas.TexT = Val(Mid(MyPref, 3, 3))
    End If
    On Error GoTo 0
    'If Not UsaGamaAlta Then
    'OptTipo(0).Value = True
    OptTipo(0).Value = True
    OptTipo(1).Enabled = False
    OptTipo(0).Visible = False
    OptTipo(1).Visible = False

    'End If

    'If Val(cmbgrupo.Text) = 8 Or Val(cmbgrupo.Text) = 9 Then
    '    Select Case Val(CmbJuego.Text)
    '    Case 0
    '        lblresul.Caption = "  FRUTA  "
    '    Case 1
    '        lblresul.Caption = "  GRANEL  "
    '    Case Else
    '        lblresul.Caption = ""
    '    End Select
    'Else
    '    lblresul.Caption = ""
    'End If
    'If Dir(App.Path & "\10tablas") <> "" Then
    CmbJuego.Clear
    For nJ = 0 To 99
        CmbJuego.AddItem CStr(nJ)
    Next nJ
    'Else
    '    CmbJuego.Clear
    '    For nJ = 0 To 1
    '        CmbJuego.AddItem CStr(nJ)
    '    Next nJ
    '
    'End If

    'CmbJuego.ListIndex = 1
    CmbJuego.TexT = "0"

    'If Val(cmbgrupo.Text) = 8 Or Val(cmbgrupo.Text) = 9 Then
    '    Select Case Val(CmbJuego.Text)
    '    Case 0
    '        lblresul.Caption = "  FRUTA  "
    '    Case 1
    '        lblresul.Caption = "  GRANEL  "
    '    Case Else
    '        lblresul.Caption = ""
    '    End Select
    'Else
    '    lblresul.Caption = ""
    'End If

    Refresca_Numero
    If CmbNumero.ListCount = 0 Then
        MsgBox "NO EXISTEN EQUIPOS...", vbCritical
        Unload Me
        Exit Sub
    End If
    Refresca_Teclas


End Sub
Private Sub Refresca_Teclas()
    Dim bucle As Integer
    Dim Y1 As Integer
    QuantumLeap = True
    If Val(CmbTeclas.TexT) < 80 Then
        CmdParte(0).Visible = False
        CmdParte(1).Visible = False
        CmdParte(2).Visible = False
        CmdParte(3).Visible = False
        CmdParte(4).Visible = False
        CmdParte(5).Visible = False
        CmdParte(6).Visible = False
        CmdParte(7).Visible = False
        CmdParte(8).Visible = False
        CmdParte(9).Visible = False
        CmdParte(10).Visible = False
        CmdParte(11).Visible = False

    Else
        CmdParte(0).Visible = True
        CmdParte(1).Visible = True
        If Val(CmbTeclas.TexT) = 240 Then
            CmdParte(2).Visible = True
            CmdParte(3).Visible = True
            CmdParte(4).Visible = False
            CmdParte(5).Visible = False
            CmdParte(6).Visible = False
            CmdParte(7).Visible = False
            CmdParte(8).Visible = False
            CmdParte(9).Visible = False
            CmdParte(10).Visible = False
            CmdParte(11).Visible = False
        End If
        If Val(CmbTeclas.TexT) = 360 Then
            CmdParte(2).Visible = True
            CmdParte(3).Visible = True
            CmdParte(4).Visible = True
            CmdParte(5).Visible = True
            CmdParte(6).Visible = False
            CmdParte(7).Visible = False
            CmdParte(8).Visible = False
            CmdParte(9).Visible = False
            CmdParte(10).Visible = False
            CmdParte(11).Visible = False
        End If
        If Val(CmbTeclas.TexT) = 720 Then
            CmdParte(2).Visible = True
            CmdParte(3).Visible = True
            CmdParte(4).Visible = True
            CmdParte(5).Visible = True
            CmdParte(6).Visible = True
            CmdParte(7).Visible = True
            CmdParte(8).Visible = True
            CmdParte(9).Visible = True
            CmdParte(10).Visible = True
            CmdParte(11).Visible = True
        End If
    End If
    For bucle = 0 To 99
        Label1(bucle).Visible = False
        Text1(bucle).Visible = False
    Next bucle
    Horizontal.Value = 0
    Vertical.Value = 0
    If CmbTeclas.TexT = "30" Then
        For Y1 = 0 To 5
            For bucle = 0 To 4
                Label1(Y1 * 5 + bucle).Caption = Y1 * 5 + bucle + 1
                Label1(Y1 * 5 + bucle).BackColor = vbYellow
                Label1(Y1 * 5 + bucle).top = 120 + (600 * Y1)
                Label1(Y1 * 5 + bucle).left = 120 + (720 * bucle)
                Text1(Y1 * 5 + bucle).top = 360 + (600 * Y1)
                Text1(Y1 * 5 + bucle).left = 120 + (720 * bucle)
                Label1(Y1 * 5 + bucle).Visible = True
                Text1(Y1 * 5 + bucle).Visible = True
            Next bucle

        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = False
        PTeclas.Width = 5 * 720 + 120
        PTeclas.Height = 6 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    If CmbTeclas.TexT = "80" Then
        For Y1 = 0 To 4
            For bucle = 0 To 7

                Label1(Y1 * 8 + bucle).BackColor = vbYellow
                Label1(Y1 * 8 + bucle).top = 120 + (600 * Y1)
                Label1(Y1 * 8 + bucle).left = 120 + (720 * bucle)
                Text1(Y1 * 8 + bucle).top = 360 + (600 * Y1)
                Text1(Y1 * 8 + bucle).left = 120 + (720 * bucle)
                Label1(Y1 * 8 + bucle).Visible = True
                Text1(Y1 * 8 + bucle).Visible = True
                Label1(Y1 * 8 + bucle).Caption = ((Y1 * 8) + bucle + 1 + (40 * QueParte))
            Next bucle

        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = True
        PTeclas.Width = 8 * 720 + 120
        PTeclas.Height = 10 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    If CmbTeclas.TexT = "240" Then
        For Y1 = 0 To 5
            For bucle = 0 To 9
                If ((Y1 * 10) + bucle + (60 * QueParte)) < 240 Then
                    Label1(Y1 * 10 + bucle).BackColor = vbYellow
                    Label1(Y1 * 10 + bucle).top = 120 + (600 * Y1)
                    Label1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Text1(Y1 * 10 + bucle).top = 360 + (600 * Y1)
                    Text1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Label1(Y1 * 10 + bucle).Visible = True
                    Text1(Y1 * 10 + bucle).Visible = True
                    Label1(Y1 * 10 + bucle).Caption = ((Y1 * 10) + bucle + 1 + (60 * QueParte))
                End If
            Next bucle
        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = False
        PTeclas.Width = 10 * 720 + 120
        PTeclas.Height = 10 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100

    End If

    If CmbTeclas.TexT = "720" Then
        For Y1 = 0 To 5
            For bucle = 0 To 9
                If ((Y1 * 10) + bucle + (60 * QueParte)) < 720 Then
                    Label1(Y1 * 10 + bucle).BackColor = vbYellow
                    Label1(Y1 * 10 + bucle).top = 120 + (600 * Y1)
                    Label1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Text1(Y1 * 10 + bucle).top = 360 + (600 * Y1)
                    Text1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Label1(Y1 * 10 + bucle).Visible = True
                    Text1(Y1 * 10 + bucle).Visible = True
                    Label1(Y1 * 10 + bucle).Caption = ((Y1 * 10) + bucle + 1 + (60 * QueParte))
                End If
            Next bucle
        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = False
        PTeclas.Width = 10 * 720 + 120
        PTeclas.Height = 10 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100


    End If

    If CmbTeclas.TexT = "360" Then
        For Y1 = 0 To 5
            For bucle = 0 To 9
                If ((Y1 * 10) + bucle + (60 * QueParte)) < 360 Then
                    Label1(Y1 * 10 + bucle).BackColor = vbYellow
                    Label1(Y1 * 10 + bucle).top = 120 + (600 * Y1)
                    Label1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Text1(Y1 * 10 + bucle).top = 360 + (600 * Y1)
                    Text1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                    Label1(Y1 * 10 + bucle).Visible = True
                    Text1(Y1 * 10 + bucle).Visible = True
                    Label1(Y1 * 10 + bucle).Caption = ((Y1 * 10) + bucle + 1 + (60 * QueParte))
                End If
            Next bucle
        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = False
        PTeclas.Width = 10 * 720 + 120
        PTeclas.Height = 10 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100


    End If

    If CmbTeclas.TexT = "165" Then

        If lss165 Then
            For Y1 = 0 To 8
                For bucle = 0 To 9
                    If ((Y1 * 10) + bucle + (88 * QueParte)) < 165 Then
                        Label1(Y1 * 10 + bucle).BackColor = vbYellow
                        Label1(Y1 * 10 + bucle).top = 120 + (600 * Y1)
                        Label1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                        Text1(Y1 * 10 + bucle).top = 360 + (600 * Y1)
                        Text1(Y1 * 10 + bucle).left = 120 + (720 * bucle)

                        If QueParte = 0 And (Y1 * 10 + bucle) > 87 Then
                        Else
                            Label1(Y1 * 10 + bucle).Visible = True
                            Text1(Y1 * 10 + bucle).Visible = True
                        End If

                        Label1(Y1 * 10 + bucle).Caption = ((Y1 * 10) + bucle + 1 + (88 * QueParte))
                    End If
                Next bucle

            Next Y1

        Else
            For Y1 = 0 To 8
                For bucle = 0 To 9
                    If ((Y1 * 10) + bucle + (90 * QueParte)) < 165 Then
                        Label1(Y1 * 10 + bucle).BackColor = vbYellow
                        Label1(Y1 * 10 + bucle).top = 120 + (600 * Y1)
                        Label1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                        Text1(Y1 * 10 + bucle).top = 360 + (600 * Y1)
                        Text1(Y1 * 10 + bucle).left = 120 + (720 * bucle)
                        Label1(Y1 * 10 + bucle).Visible = True
                        Text1(Y1 * 10 + bucle).Visible = True
                        Label1(Y1 * 10 + bucle).Caption = ((Y1 * 10) + bucle + 1 + (90 * QueParte))
                    End If
                Next bucle

            Next Y1
        End If
        Horizontal.Enabled = True
        Vertical.Enabled = True
        PTeclas.Width = 10 * 720 + 120
        PTeclas.Height = 10 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    If CmbTeclas.TexT = "96" Then

        For Y1 = 0 To 6
            For bucle = 0 To 6
                If (Y1 * 7 + bucle + QueParte * 50) < 97 Then
                    Label1(Y1 * 7 + bucle).BackColor = vbYellow
                    Label1(Y1 * 7 + bucle).top = 120 + (600 * bucle)
                    Label1(Y1 * 7 + bucle).left = 120 + (720 * Y1)
                    Text1(Y1 * 7 + bucle).top = 360 + (600 * bucle)
                    Text1(Y1 * 7 + bucle).left = 120 + (720 * Y1)
                    Label1(Y1 * 7 + bucle).Visible = True
                    If QueParte = 0 Then
                        Label1(Y1 * 7 + bucle).Caption = 1 + (Y1 * 7 + bucle + QueParte * 50)
                    Else
                        Label1(Y1 * 7 + bucle).Caption = (Y1 * 7 + bucle + QueParte * 50)
                    End If
                    Text1(Y1 * 7 + bucle).Visible = True
                End If
            Next bucle
        Next Y1

        PTeclas.Width = Y1 * 720 + 120
        Horizontal.Enabled = True
        Vertical.Enabled = True
        PTeclas.Width = 14 * 720 + 120
        PTeclas.Height = 7 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    If CmbTeclas.TexT = "56" Then
        For bucle = 13 To 0 Step -1
            Label1(bucle).BackColor = vbYellow
            Label1(bucle).top = 120
            Label1(bucle).Caption = bucle + 1
            Label1(bucle).left = 120 + (720 * (13 - bucle))
            Text1(bucle).top = 360
            Text1(bucle).left = 120 + (720 * (13 - bucle))
            Label1(bucle).Visible = True
            Text1(bucle).Visible = True
        Next bucle

        For Y1 = 1 To 6
            For bucle = 6 To 0 Step -1
                Label1(7 + (Y1 * 7 + bucle)).Caption = 8 + (Y1 * 7 + bucle)
                Label1(7 + (Y1 * 7 + bucle)).BackColor = vbYellow
                Label1(7 + (Y1 * 7 + bucle)).top = 120 + 600 * Y1
                Label1(7 + (Y1 * 7 + bucle)).left = 120 + (720 * (bucle))
                Text1(7 + (Y1 * 7 + bucle)).top = 360 + 600 * Y1
                Text1(7 + (Y1 * 7 + bucle)).left = 120 + (720 * (bucle))
                Label1(7 + (Y1 * 7 + bucle)).Visible = True
                Text1(7 + (Y1 * 7 + bucle)).Visible = True
            Next bucle
        Next Y1
        Horizontal.Enabled = True
        Vertical.Enabled = True
        PTeclas.Width = 14 * 720 + 120
        PTeclas.Height = 7 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    If CmbTeclas.TexT = "65" Then
        For Y1 = 0 To 7
            For bucle = 6 To 0 Step -1
                Label1(Y1 * 7 + bucle).BackColor = vbYellow
                Label1(Y1 * 7 + bucle).top = 120 + 600 * Y1
                Label1(Y1 * 7 + bucle).left = 120 + (720 * (6 - bucle))
                Text1(Y1 * 7 + bucle).top = 360 + 600 * Y1
                Text1(Y1 * 7 + bucle).left = 120 + (720 * (6 - bucle))
                Label1(Y1 * 7 + bucle).Visible = True
                Text1(Y1 * 7 + bucle).Visible = True
            Next bucle
        Next Y1
        Horizontal.Enabled = False
        Vertical.Enabled = True
        PTeclas.Width = 8 * 720 + 120
        PTeclas.Height = 7 * 720
        Horizontal.Max = (PTeclas.Width - Frame3.Width) / 100
        Vertical.Max = (PTeclas.Height - Frame3.Height) / 100
    End If
    QuantumLeap = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim MyPref As String
    frmEpelsa.Enabled = True
    LockBase = False
    If OptSeccion(1).Value = True Then
        MyPref = "1"
    Else
        MyPref = "0"
    End If
    If OptTipo(1).Value = True Then
        MyPref = MyPref & "1"
    Else
        MyPref = MyPref & "0"
    End If
    MyPref = MyPref & Format(Val(CmbTeclas.TexT), "000")
    WPreferencias "TECLAS    ", 1, MyPref
    frmEpelsa.Enabled = True
End Sub


Private Sub Horizontal_Change()
    PTeclas.left = 0 - Horizontal.Value * 100
End Sub


Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not ClipLinux Then
        Clipboard.SetText Item.TexT
    Else
        sLinux = Item.TexT
    End If
End Sub

'Private Sub OptSeccion_Click(Index As Integer)
'    QueParte = 0
'    Txtconsulta.TexT = ""
'    'lblresul.Caption = ""
'    If OptSeccion(0).Value = True Then
'        lblconsulta.Caption = "Cod:"
'        Lista.ColumnHeaders(1).TexT = CargaCadena(89)
'    Else
'        lblconsulta.Caption = "PLU:"
'        Lista.ColumnHeaders(1).TexT = "PLU"
'    End If
'
'    Refresca_Teclas
'    Refresca_Datos
'    refresca_lista
'End Sub

'Private Sub OptTipo_Click(Index As Integer)
'    If Dir(App.Path & "\ikea.exe") <> "" Then
'        OptTipo(0).Value = True
'    End If
'    'If OptTipo(0).Value = True Then
'    '    lblMostrador.Caption = "Equipos:" 'CargaCadena(302)  '"Mostrador"
'    '    If Dir(App.Path & "\ikea.exe") <> "" Then
'    '        CmdConsulta.Visible = True
'    '    Else
'    '        CmdConsulta.Visible = False
'    '    End If
'    'Else
'        LblMostrador.Caption = "Equipos:" ''CargaCadena(303)  '"Equipo "
'        If UsaGamaAlta Then CmdConsulta.Visible = True
'    'End If
'    Refresca_Numero
'
'    'QueParte = 1
'    '1.9.0
'    Refresca_Teclas
'    QueParte = 0
'    ''''''
'    Refresca_Datos
'    refresca_lista
'End Sub






Private Sub Text1_Click(Index As Integer)
    Dim Buf As String
    If Not ClipLinux Then
        Buf = Clipboard.GetText
        If IsNumeric(Buf) Then Text1(Index).TexT = Buf
        '1.9.0
        'next
        'If lss165 And Index > 87 Then
        '    CmbJuego.ListIndex = 1
        'Else
        '    CmbJuego.ListIndex = 0
        'End If
        ''''''
        Clipboard.SetText ""
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If ClipLinux Then
        If IsNumeric(sLinux) Then Text1(Index).TexT = sLinux
        sLinux = ""
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Text1(Index).Locked = False
    Select Case KeyAscii
    Case vbKeyDelete
    Case vbKeyBack
    Case 48 To 57
        If OptSeccion(0).Value = False Then
            If Len(Text1(Index).TexT) >= 4 Then Text1(Index).Locked = True
        Else
            If Len(Text1(Index).TexT) >= 6 Then Text1(Index).Locked = True
        End If
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub Text1_LostFocus(Index As Integer)

    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim bucle As Long
    Dim BucleGrande As Long
    Dim eNCONTRADO As Boolean
    Dim ElMostrador As TipoCodigos
    Dim MisTeclas As DB_Teclas
    Dim MiRespuesta As Integer
    Dim artBase As dao.Database
    Dim RegArt As dao.Recordset
    Dim ncont As Long

    'If lEsSalir Then Exit Sub

    ElMostrador = Desglosa_Codigos(CmbNumero.TexT)

    eNCONTRADO = False
    If Text1(Index).TexT = "" Then
        Text1(Index).BackColor = vbWhite
        MisTeclas.CodArticulo = 0
        MisTeclas.PluArticulo = 0
        MisTeclas.SecArticulo = 0
        MisTeclas.Destino_CodiTec = Label1(Index).Caption
        If OptTipo(0).Value = True Then
            MisTeclas.Destino_NBalanza = 0
            MisTeclas.Destino_NMostrador = ElMostrador.codigo(1)
        Else
            MisTeclas.Destino_NBalanza = CmbNumero.TexT
            MisTeclas.Destino_NMostrador = 0
        End If
        If lss165 Then
            If MisTeclas.Destino_CodiTec > 88 Then
                MisTeclas.Destino_Tabla = 1
                MisTeclas.Destino_CodiTec = MisTeclas.Destino_CodiTec - 88
            Else
                MisTeclas.Destino_Tabla = 0
            End If
        Else
            MisTeclas.Destino_Tabla = Val(CmbJuego.TexT)
        End If

        MisTeclas.Tipo = "TOUCH"

        MisTeclas.Section = Trim(cmbgrupo.TexT)

        MiRespuesta = Baja_Teclas(MisTeclas)
        Text1(Index).BackColor = vbWhite
        Text1(Index).TexT = ""
        Exit Sub
    End If
    '***************
    ' mira si existe el
    ' artículo indicado
    '*****************

    If OptSeccion(0).Value = True Then
        ' por codigo
        Registro.OpenRecordset _
                ("select codigo from articulo where borrado=false and codigo=" & Val(Text1(Index).TexT))
        If Registro.EOF Then
            eNCONTRADO = False
        Else
            eNCONTRADO = True
            '2.0.9 If Dir(App.Path & "\marte.new") <> "" Then
            'If lMarteNew Then
            '    Set artBase = OpenDatabase(Base_General)
            '    Set RegArt = artBase.OpenRecordset("select codigo,codi_tec from teclas where codigo=" & CStr(Val(Text1(Index).TexT)))
            '    If Not RegArt.EOF Then
            '
            '        If Not RegArt.EOF And Index <> (RegArt.Fields("codi_tec") - 1) Then
            '            Text1(Index).TexT = ""
            '
            '            Text1(RegArt.Fields("codi_tec") - 1).SetFocus
            '            Beep
            '            'nCont = Text1(regArt.Fields("codi_tec") - 1).BackColor
            '            Beep
            '            Text1(RegArt.Fields("codi_tec") - 1).BackColor = vbRed
            '            Beep
            '            'Text1(regArt.Fields("codi_tec") - 1).BackColor = nCont
            '
            '            RegArt.Close
            '            Set RegArt = Nothing
            '            artBase.Close
            '            Set artBase = Nothing
            '
            '            Exit Sub
            '
            '        End If
            '
            '    End If
            '
            '    Set RegArt = artBase.OpenRecordset("select tran_plu,tran_el,tran_textoel from articulo where codigo=" & CStr(Val(Text1(Index).TexT)))
            '    RegArt.Edit
            '    RegArt.Fields("tran_plu") = ""
            '    RegArt.Fields("tran_EL") = ""
            '    RegArt.Fields("tran_textoel") = ""
            '
            '    RegArt.Update
            '    RegArt.Close
            '    Set RegArt = Nothing
            '    artBase.Close
            '    Set artBase = Nothing
            '
            'End If
        End If
    Else
        ' por plu
        If OptTipo(0).Value = True Then
            ' teclas de un mostrador
            Reg2.OpenRecordset _
                    ("select secc_maqui from seccion where borrado=false and codi_ident=" & ElMostrador.codigo(1))
        Else
            ' teclas de un equipo
            Reg2.OpenRecordset _
                    ("select secc_maqui from equipos where borrado=false and numero_eqp=" & ElMostrador.codigo(1))
        End If

        If Reg2.EOF Then
            eNCONTRADO = False
        Else
            Registro.OpenRecordset _
                    ("select codi_ident from articulo where borrado=false and plu=" & Text1(Index).TexT & " and secc_maqui=" & Reg2.Fields("secc_Maqui"))
            If Registro.EOF Then
                eNCONTRADO = False
            Else
                eNCONTRADO = True
                MisTeclas.SecArticulo = Registro.Fields("codi_ident")

                '2.0.9 If Dir(App.Path & "\marte.new") <> "" Then
                'If lMarteNew Then
                '
                '     Set artBase = OpenDatabase(Base_General)
                '     Set RegArt = artBase.OpenRecordset("select codigo,codi_tec from teclas where plu=" & Text1(Index).TexT & " and secc_maqui=" & Reg2.Fields("secc_Maqui"))
                '     If Not RegArt.EOF Then
                '
                '         If Not RegArt.EOF And Index <> (RegArt.Fields("codi_tec") - 1) Then
                '             Text1(Index).TexT = ""
                '
                '             Text1(RegArt.Fields("codi_tec") - 1).SetFocus
                '             Beep
                '             'nCont = Text1(regArt.Fields("codi_tec") - 1).BackColor
                '             Beep
                '             Text1(RegArt.Fields("codi_tec") - 1).BackColor = vbRed
                '             Beep
                '             'Text1(regArt.Fields("codi_tec") - 1).BackColor = nCont
                '
                '             RegArt.Close
                '             Set RegArt = Nothing
                '             artBase.Close
                '             Set artBase = Nothing
                '
                '             Exit Sub
                '
                '         End If
                '
                '     End If
                '
                '    Set artBase = OpenDatabase(Base_General)
                '    Set RegArt = artBase.OpenRecordset("select tran_plu,tran_el,tran_textoel from articulo where plu=" & Text1(Index).TexT & " and secc_maqui=" & Reg2.Fields("secc_Maqui"))
                '    RegArt.Edit
                '    RegArt.Fields("tran_plu") = ""
                '    RegArt.Fields("tran_EL") = ""
                '    RegArt.Fields("tran_textoel") = ""
                '
                '    RegArt.Update
                '    RegArt.Close
                '    Set RegArt = Nothing
                '    artBase.Close
                '    Set artBase = Nothing
                'End If

            End If
        End If
    End If

    If eNCONTRADO = False Then
        If OptSeccion(0).Value Then
            LblInfo.Caption = CargaCadena(218)  '("El código indicado no existe")
            Text1(Index).BackColor = vbRed
        Else
            LblInfo.Caption = CargaCadena(221)  '("El PLU indicado no existe")
            Text1(Index).BackColor = vbRed
        End If
    Else
        LblInfo.Caption = ""
        Text1(Index).BackColor = vbWhite
        If OptSeccion(0).Value = True Then
            MisTeclas.CodArticulo = Val(Text1(Index).TexT)
            MisTeclas.PluArticulo = 0
            MisTeclas.SecArticulo = 0
        Else
            MisTeclas.CodArticulo = 0
            MisTeclas.PluArticulo = Val(Text1(Index).TexT)
        End If
        MisTeclas.Destino_CodiTec = Label1(Index).Caption

        'If OptTipo(0).Value = True Then
        '    MisTeclas.Destino_NBalanza = 0
        '    MisTeclas.Destino_NMostrador = ElMostrador.codigo(1)
        'Else
        MisTeclas.Destino_NBalanza = CmbNumero.TexT
        MisTeclas.Destino_NMostrador = 0
        'End If

        'MisTeclas.Destino_Tabla = Val(CmbJuego.TexT)
        If lss165 Then
            If MisTeclas.Destino_CodiTec > 88 Then
                MisTeclas.Destino_Tabla = 1
                MisTeclas.Destino_CodiTec = MisTeclas.Destino_CodiTec - 88
            Else
                MisTeclas.Destino_Tabla = 0
            End If
        Else
            MisTeclas.Destino_Tabla = Val(CmbJuego.TexT)
        End If
        MisTeclas.Tipo = "TOUCH"
        MisTeclas.Section = Trim(cmbgrupo.TexT)

        MiRespuesta = Alta_Teclas(MisTeclas)
        If MiRespuesta = 0 Or MiRespuesta = 1 Then Text1(Index).BackColor = vbGreen
    End If
    Refresca_Datos
End Sub



Private Sub Txtconsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Txtconsulta.TexT <> "" Then
        KeyAscii = 0
        Consulta_Art
    End If

    If OptSeccion(0).Value = True Then
        Txtconsulta.Locked = Checktexto(KeyAscii, 6, Txtconsulta.TexT, True, False)
    Else
        Txtconsulta.Locked = Checktexto(KeyAscii, 4, Txtconsulta.TexT, True, False)
    End If
End Sub

Private Sub Vertical_Change()
    PTeclas.top = 0 - Vertical.Value * 100
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmOrdTec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Artículos en Familias"
   ClientHeight    =   7305
   ClientLeft      =   1185
   ClientTop       =   1770
   ClientWidth     =   12165
   LinkTopic       =   "FrmOrdTec"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12165
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
      Picture         =   "FrmOrdTec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   226
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
      Picture         =   "FrmOrdTec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   225
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
      TabIndex        =   224
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
      TabIndex        =   223
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
      TabIndex        =   222
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
      TabIndex        =   221
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
      TabIndex        =   220
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
      TabIndex        =   219
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
      TabIndex        =   218
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
      TabIndex        =   217
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
      TabIndex        =   216
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
      TabIndex        =   215
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
      Left            =   6120
      TabIndex        =   214
      Top             =   5400
      Width           =   1572
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5295
      Left            =   7800
      TabIndex        =   213
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
      TabIndex        =   212
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
      TabIndex        =   211
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
      TabIndex        =   210
      Top             =   6000
      Width           =   1095
   End
   Begin VB.VScrollBar Vertical 
      Height          =   3735
      Left            =   7440
      Max             =   6132
      TabIndex        =   208
      Top             =   1680
      Width           =   375
   End
   Begin VB.HScrollBar Horizontal 
      Height          =   252
      Left            =   120
      Max             =   7572
      TabIndex        =   207
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
      TabIndex        =   5
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
         TabIndex        =   6
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
            TabIndex        =   206
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
            TabIndex        =   204
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
            TabIndex        =   202
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
            TabIndex        =   200
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
            TabIndex        =   198
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
            TabIndex        =   196
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
            TabIndex        =   194
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
            TabIndex        =   192
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
            TabIndex        =   190
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
            TabIndex        =   188
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
            TabIndex        =   186
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
            TabIndex        =   184
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
            TabIndex        =   182
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
            TabIndex        =   180
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
            TabIndex        =   178
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
            TabIndex        =   176
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
            TabIndex        =   174
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
            TabIndex        =   172
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
            TabIndex        =   170
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
            TabIndex        =   168
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
            TabIndex        =   166
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
            TabIndex        =   164
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
            TabIndex        =   162
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
            TabIndex        =   160
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
            TabIndex        =   158
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
            TabIndex        =   156
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
            TabIndex        =   154
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
            TabIndex        =   152
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
            TabIndex        =   150
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
            TabIndex        =   148
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
            TabIndex        =   146
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
            TabIndex        =   144
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
            TabIndex        =   142
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
            TabIndex        =   140
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
            TabIndex        =   138
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
            TabIndex        =   136
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
            TabIndex        =   134
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
            TabIndex        =   132
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
            TabIndex        =   130
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
            TabIndex        =   128
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
            TabIndex        =   126
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
            TabIndex        =   124
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
            TabIndex        =   122
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
            TabIndex        =   120
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
            TabIndex        =   118
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
            TabIndex        =   116
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
            TabIndex        =   114
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
            TabIndex        =   112
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
            TabIndex        =   110
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
            TabIndex        =   108
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
            TabIndex        =   106
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
            TabIndex        =   104
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
            TabIndex        =   102
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
            TabIndex        =   100
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
            TabIndex        =   98
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
            TabIndex        =   96
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
            TabIndex        =   94
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
            TabIndex        =   92
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
            TabIndex        =   90
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
            TabIndex        =   88
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
            TabIndex        =   86
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
            TabIndex        =   84
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
            TabIndex        =   82
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
            TabIndex        =   80
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
            TabIndex        =   78
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
            TabIndex        =   76
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
            TabIndex        =   74
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
            TabIndex        =   72
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
            TabIndex        =   70
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
            TabIndex        =   68
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
            TabIndex        =   66
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
            TabIndex        =   64
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
            TabIndex        =   62
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
            TabIndex        =   60
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
            TabIndex        =   58
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
            TabIndex        =   56
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
            TabIndex        =   54
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
            TabIndex        =   52
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
            TabIndex        =   50
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
            TabIndex        =   48
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
            TabIndex        =   46
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
            TabIndex        =   44
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
            TabIndex        =   42
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
            TabIndex        =   40
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
            TabIndex        =   38
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
            TabIndex        =   36
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
            TabIndex        =   34
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
            TabIndex        =   32
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
            TabIndex        =   30
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
            TabIndex        =   28
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
            TabIndex        =   26
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
            TabIndex        =   24
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
            TabIndex        =   22
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
            TabIndex        =   20
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
            TabIndex        =   18
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
            TabIndex        =   16
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
            TabIndex        =   14
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
            TabIndex        =   12
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
            TabIndex        =   10
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
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   99
            Left            =   6600
            TabIndex        =   205
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   98
            Left            =   5880
            TabIndex        =   203
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   97
            Left            =   5160
            TabIndex        =   201
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   96
            Left            =   4440
            TabIndex        =   199
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   95
            Left            =   3720
            TabIndex        =   197
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   94
            Left            =   3000
            TabIndex        =   195
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   93
            Left            =   2280
            TabIndex        =   193
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   92
            Left            =   1560
            TabIndex        =   191
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   91
            Left            =   840
            TabIndex        =   189
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   90
            Left            =   120
            TabIndex        =   187
            Top             =   5520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   89
            Left            =   6600
            TabIndex        =   185
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   88
            Left            =   5880
            TabIndex        =   183
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   87
            Left            =   5160
            TabIndex        =   181
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   86
            Left            =   4440
            TabIndex        =   179
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   85
            Left            =   3720
            TabIndex        =   177
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   84
            Left            =   3000
            TabIndex        =   175
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   83
            Left            =   2280
            TabIndex        =   173
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   82
            Left            =   1560
            TabIndex        =   171
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   81
            Left            =   840
            TabIndex        =   169
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   80
            Left            =   120
            TabIndex        =   167
            Top             =   4920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   79
            Left            =   6600
            TabIndex        =   165
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   78
            Left            =   5880
            TabIndex        =   163
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   77
            Left            =   5160
            TabIndex        =   161
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   76
            Left            =   4440
            TabIndex        =   159
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   75
            Left            =   3720
            TabIndex        =   157
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   74
            Left            =   3000
            TabIndex        =   155
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   73
            Left            =   2280
            TabIndex        =   153
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   72
            Left            =   1560
            TabIndex        =   151
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   71
            Left            =   840
            TabIndex        =   149
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   70
            Left            =   120
            TabIndex        =   147
            Top             =   4320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   69
            Left            =   6600
            TabIndex        =   145
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   68
            Left            =   5880
            TabIndex        =   143
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   67
            Left            =   5160
            TabIndex        =   141
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   66
            Left            =   4440
            TabIndex        =   139
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   65
            Left            =   3720
            TabIndex        =   137
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   64
            Left            =   3000
            TabIndex        =   135
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   63
            Left            =   2280
            TabIndex        =   133
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   62
            Left            =   1560
            TabIndex        =   131
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   61
            Left            =   840
            TabIndex        =   129
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   60
            Left            =   120
            TabIndex        =   127
            Top             =   3720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   59
            Left            =   6600
            TabIndex        =   125
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   58
            Left            =   5880
            TabIndex        =   123
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   57
            Left            =   5160
            TabIndex        =   121
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   56
            Left            =   4440
            TabIndex        =   119
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   55
            Left            =   3720
            TabIndex        =   117
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   54
            Left            =   3000
            TabIndex        =   115
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   53
            Left            =   2280
            TabIndex        =   113
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   52
            Left            =   1560
            TabIndex        =   111
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   51
            Left            =   840
            TabIndex        =   109
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   50
            Left            =   120
            TabIndex        =   107
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   49
            Left            =   6600
            TabIndex        =   105
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   48
            Left            =   5880
            TabIndex        =   103
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   47
            Left            =   5160
            TabIndex        =   101
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   46
            Left            =   4440
            TabIndex        =   99
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   45
            Left            =   3720
            TabIndex        =   97
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   44
            Left            =   3000
            TabIndex        =   95
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   43
            Left            =   2280
            TabIndex        =   93
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   42
            Left            =   1560
            TabIndex        =   91
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   41
            Left            =   840
            TabIndex        =   89
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   40
            Left            =   120
            TabIndex        =   87
            Top             =   2520
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   39
            Left            =   6600
            TabIndex        =   85
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   38
            Left            =   5880
            TabIndex        =   83
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   37
            Left            =   5160
            TabIndex        =   81
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   36
            Left            =   4440
            TabIndex        =   79
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   35
            Left            =   3720
            TabIndex        =   77
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   34
            Left            =   3000
            TabIndex        =   75
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   33
            Left            =   2280
            TabIndex        =   73
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   32
            Left            =   1560
            TabIndex        =   71
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   31
            Left            =   840
            TabIndex        =   69
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   30
            Left            =   120
            TabIndex        =   67
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   29
            Left            =   6600
            TabIndex        =   65
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   28
            Left            =   5880
            TabIndex        =   63
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   27
            Left            =   5160
            TabIndex        =   61
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   26
            Left            =   4440
            TabIndex        =   59
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   25
            Left            =   3720
            TabIndex        =   57
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   24
            Left            =   3000
            TabIndex        =   55
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   23
            Left            =   2280
            TabIndex        =   53
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   22
            Left            =   1560
            TabIndex        =   51
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   21
            Left            =   840
            TabIndex        =   49
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   20
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   19
            Left            =   6600
            TabIndex        =   45
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   18
            Left            =   5880
            TabIndex        =   43
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   17
            Left            =   5160
            TabIndex        =   41
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   16
            Left            =   4440
            TabIndex        =   39
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   15
            Left            =   3720
            TabIndex        =   37
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   14
            Left            =   3000
            TabIndex        =   35
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   13
            Left            =   2280
            TabIndex        =   33
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   12
            Left            =   1560
            TabIndex        =   31
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   11
            Left            =   840
            TabIndex        =   29
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   9
            Left            =   6600
            TabIndex        =   25
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   8
            Left            =   5880
            TabIndex        =   23
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   7
            Left            =   5160
            TabIndex        =   21
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   6
            Left            =   4440
            TabIndex        =   19
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   5
            Left            =   3720
            TabIndex        =   17
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   4
            Left            =   3000
            TabIndex        =   15
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   3
            Left            =   2280
            TabIndex        =   13
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   2
            Left            =   1560
            TabIndex        =   11
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   612
         End
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
      Begin VB.ComboBox CmbNumero 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   227
         Top             =   720
         Width           =   1692
      End
      Begin VB.ComboBox CmbJuego 
         Height          =   315
         ItemData        =   "FrmOrdTec.frx":0884
         Left            =   2640
         List            =   "FrmOrdTec.frx":0886
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox CmbTeclas 
         Height          =   315
         ItemData        =   "FrmOrdTec.frx":0888
         Left            =   2640
         List            =   "FrmOrdTec.frx":0895
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label LblMostrador 
         Caption         =   "Mostrador"
         Height          =   255
         Left            =   1560
         TabIndex        =   228
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Familia :"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Posiciones:"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   855
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
      TabIndex        =   209
      Top             =   6480
      Width           =   6015
   End
End
Attribute VB_Name = "FrmOrdTec"
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
Private noCLICK As Boolean
'Private lEsSalir As Boolean

Private Sub CambiarIdioma()
    Label2.Caption = "Posiciones:" 'CargaCadena(321)
    Label3.Caption = "Familia:" 'CargaCadena(223)
    
    LblMostrador.Caption = CargaCadena(302) & ":"
    
    CmdParte(0).Caption = "1 a 60" 'CargaCadena(712) & " 1"
    CmdParte(1).Caption = "61 a 120" 'CargaCadena(712) & " 2"
    CmdParte(2).Caption = "121 a 180" 'CargaCadena(712) & " 3"
    CmdParte(3).Caption = "181 a 240" 'CargaCadena(712) & " 4"
    CmdParte(4).Caption = "241 a 300" 'CargaCadena(712) & " 5"
    CmdParte(5).Caption = "301 a 360" 'CargaCadena(712) & " 6"
    CmdParte(6).Caption = "361 a 420" 'CargaCadena(712) & " 7"
    CmdParte(7).Caption = "421 a 480" 'CargaCadena(712) & " 8"
    CmdParte(8).Caption = "481 a 540" 'CargaCadena(712) & " 9"
    CmdParte(9).Caption = "541 a 600" 'CargaCadena(712) & " 10"
    CmdParte(10).Caption = "601 a 660" 'CargaCadena(712) & " 11"
    CmdParte(11).Caption = "661 a 720" 'CargaCadena(712) & " 12"
    
    Command3.Caption = CargaCadena(7)
    Frame1.Caption = CargaCadena(713)
    FrmOrdTec.Caption = "Orden de Artículos en Familias" 'CargaCadena(1398)
    
    CmdPrint.Caption = CargaCadena(466)
    CmdPrint.Visible = False
End Sub
Private Sub Refresca_Numero()
    
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim Agregar As Boolean
    Dim Buffer As String
    
    Aux_Refresca_Mostradores_Con_Familias CmbNumero
    If CmbNumero.ListCount > 0 Then
        
        Aux_Refresca_Familias_Mostrador CmbJuego, CmbNumero.List(0)
        If CmbJuego.ListCount > 0 Then
            HabilitarTeclas True
        Else
            HabilitarTeclas False
        End If
        
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
    Dim nMos As Integer
    
    lss165 = False
    
    QuantumLeap = True
    
    For bucle = 0 To 99
        Text1(bucle).TexT = ""
        Text1(bucle).ToolTipText = ""
        Text1(bucle).BackColor = vbWhite
    Next bucle
    
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
            Case Else
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
    
    'LaSeccion = Desglosa_Codigos(CmbNumero.TexT)
    nMos = Val(CmbNumero.TexT)
    noCLICK = True
    If CmbJuego.TexT = "" Then CmbJuego.ListIndex = 0
    noCLICK = False
    Registro.OpenRecordset _
    ("select * from teclas where borrado=false and secc_maqui=" & Trim(CmbNumero.TexT) & " and tabla=" & CmbJuego.TexT _
    & " and prog_tec='TOUCH' and tran_tecsc10='" & Trim(CmbNumero.TexT) & "'")
    '''''''''
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If (.Fields("codi_tec")) >= MiMinimo And (.Fields("codi_tec")) <= Mimaximo Then
                    Text1(.Fields("codi_tec") - MiMinimo).BackColor = vbGreen
                    Text1(.Fields("codi_tec") - MiMinimo).ToolTipText = .Fields("desc_plu")
                    Text1(.Fields("codi_tec") - MiMinimo).TexT = .Fields("codigo")
                End If
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
    QuantumLeap = False
End Sub


Private Sub CmbJuego_Click()
    If noCLICK Then Exit Sub
    QueParte = 0
    Refresca_Teclas
    Refresca_Datos
    refresca_lista
End Sub

Private Sub CmbNumero_Click()
    Dim Registro As New RecordNet
    If noCLICK Then Exit Sub
    
    Aux_Refresca_Familias_Mostrador CmbJuego, CmbNumero.TexT
    
    QueParte = 0
        
    Refresca_Datos
    
    CmbJuego.ListIndex = 0
    
    Refresca_Teclas 'añadido
    Refresca_Datos
    '''''''''''''''
    refresca_lista
End Sub

Private Sub CmbTeclas_Click()
    If noCLICK Then Exit Sub
    
    QueParte = 0
    Refresca_Teclas
    Refresca_Datos
End Sub

Private Sub CmdAvan_Click()
    If (QueParte < 11 And Val(CmbTeclas.TexT) = 720) Or (QueParte < 5 And Val(CmbTeclas.TexT) = 360) Or (QueParte < 3 And Val(CmbTeclas.TexT) = 240) Then _
        QueParte = QueParte + 1
    
    Refresca_Teclas
    Refresca_Datos

End Sub


Private Sub CmdParte_Click(Index As Integer)
    QueParte = Index
    Refresca_Teclas
    Refresca_Datos
End Sub

Private Sub CmdParte_GotFocus(Index As Integer)
    QuantumLeap = True
End Sub

'Private Sub CmdPrint_Click()
'    Dim MyCad As String
'    Dim PP As TipoCodigos
'    MyCad = "0" & CmbNumero.TexT
'    If MyCad <> "" Then
'        Aux_Ejecuta "rteclas", MyCad
'    End If
'End Sub

Private Sub CmdRetro_Click()
If QueParte > 0 Then QueParte = QueParte - 1
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


'Private Sub Consulta_Art()
'
'    Dim Registro As New RecordNet
'    Dim RegAux As New RecordNet
'    Dim MiCodigo As TipoCodigos
'
'    LblInfo.Caption = ""
'    Registro.OpenRecordset ("select des_plu1 from articulo where borrado=false and codigo=" & Txtconsulta.TexT)
'    If Not Registro.EOF Then
'        lblresul.Caption = Registro.Fields("des_plu1")
'    Else
'        LblInfo.Caption = CargaCadena(218)  '"El código indicado no existe"
'    End If
'End Sub

Private Sub refresca_lista()
    
    Dim Registro As New RecordNet
    Dim Cad1 As String
    Dim MiSec As Integer
    Dim MyError As Boolean
    Dim lAccess As Boolean
    
    Lista.ListItems.Clear
    
    Cad1 = "codigo="
    lAccess = True
    MyError = False
    If Trim(CmbNumero.TexT) <> "" Then
        If InStr(1, CmbNumero.TexT, ",") Then
            MiSec = Val(left(CmbNumero.TexT, InStr(1, CmbNumero.TexT, ",") - 1))
        Else
            MiSec = Val(CmbNumero.TexT)
        End If
    Else
        MyError = True
    End If
    If Not MyError Then
        Registro.OpenRecordset ("select secc_maqui,codigo,plu,des_plu1 from articulo where borrado=false and codi_fam=" & Trim(CmbJuego.TexT) & " and secc_maqui=" & Trim(CmbNumero.TexT) & " order by codigo")
        With Registro
            If Not .EOF Then
                Do Until .EOF
                    Lista.ListItems.Add , , .Fields("codigo")
                    If Not IsNull(.Fields("des_plu1")) Then Lista.ListItems(Lista.ListItems.Count).SubItems(1) = .Fields("des_plu1")
                    .Movenext
                Loop
            End If
        End With
    End If

End Sub


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
    
    Refresca_Numero
    
    MyPref = RPreferencias("TECLAS    ", 1)
    On Error Resume Next
    If MyPref <> "" Then
        CmbTeclas.TexT = Val(Mid(MyPref, 3, 3))
    End If
    On Error GoTo 0
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
    MyPref = "1"
    MyPref = MyPref & "0"
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

Private Sub Text1_Click(Index As Integer)
    Dim Buf As String
    If Not ClipLinux Then
        Buf = Clipboard.GetText
        If IsNumeric(Buf) Then Text1(Index).TexT = Buf
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
            'If OptSeccion(0).Value = False Then
            '    If Len(Text1(Index).TexT) >= 4 Then Text1(Index).Locked = True
            'Else
                If Len(Text1(Index).TexT) >= 6 Then Text1(Index).Locked = True
            'End If
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
    Dim nMos As Integer
    
    '3.2.5 ElMostrador = Desglosa_Codigos(CmbNumero.TexT)
    nMos = Val(CmbNumero.TexT)
    
    eNCONTRADO = False
    If Text1(Index).TexT = "" Then
        On Error GoTo NoCodLen
        'If Len(ElMostrador.codigo(0)) > 0 Then
            Text1(Index).BackColor = vbWhite
            MisTeclas.CodArticulo = 0
            MisTeclas.PluArticulo = 0
            MisTeclas.SecArticulo = 0
            MisTeclas.Destino_CodiTec = Label1(Index).Caption
            MisTeclas.Destino_NBalanza = 0
            '3.2.5 MisTeclas.Destino_NMostrador = ElMostrador.codigo(1)
            MisTeclas.Destino_NMostrador = nMos
            'If lss165 Then
            '    If MisTeclas.Destino_CodiTec > 88 Then
            '        MisTeclas.Destino_Tabla = 1
            '        MisTeclas.Destino_CodiTec = MisTeclas.Destino_CodiTec - 88
            '    Else
            '        MisTeclas.Destino_Tabla = 0
            '    End If
            'Else
                MisTeclas.Destino_Tabla = Val(CmbJuego.TexT)
            'End If
            
            MisTeclas.Tipo = "TOUCH"
            
            MisTeclas.Section = nMos '"1" 09/09/17
            
            MiRespuesta = Baja_Teclas(MisTeclas)
            Text1(Index).BackColor = vbWhite
            Text1(Index).TexT = ""
        'End If
NoCodLen:
        On Error GoTo 0
        GoTo salRefresh
    End If
    '***************
    ' mira si existe el
    ' artículo indicado
    '*****************
    
        Registro.OpenRecordset _
        ("select codigo from articulo where borrado=false and codigo=" & Val(Text1(Index).TexT))
        If Registro.EOF Then
            eNCONTRADO = False
        Else
            eNCONTRADO = True
        End If

    If eNCONTRADO = False Then
        LblInfo.Caption = CargaCadena(218)  '("El código indicado no existe")
        Text1(Index).BackColor = vbRed
    Else
        LblInfo.Caption = ""
        Text1(Index).BackColor = vbWhite
        MisTeclas.CodArticulo = Val(Text1(Index).TexT)
        MisTeclas.PluArticulo = 0
        MisTeclas.SecArticulo = 0
        MisTeclas.Destino_CodiTec = Label1(Index).Caption
        MisTeclas.Destino_NBalanza = 0
        '3.2.5 MisTeclas.Destino_NMostrador = ElMostrador.codigo(1)
        MisTeclas.Destino_NMostrador = nMos
        ''''
        'If lss165 Then
        '    If MisTeclas.Destino_CodiTec > 88 Then
        '        MisTeclas.Destino_Tabla = 1
        '        MisTeclas.Destino_CodiTec = MisTeclas.Destino_CodiTec - 88
        '    Else
        '        MisTeclas.Destino_Tabla = 0
        '    End If
        'Else
            MisTeclas.Destino_Tabla = Val(CmbJuego.TexT)
        'End If
        MisTeclas.Tipo = "TOUCH"
        MisTeclas.Section = nMos '"1" 'Trim(CmbNumero.TexT) 09/09/17
        
        MiRespuesta = Alta_Teclas(MisTeclas)
        If MiRespuesta = 0 Or MiRespuesta = 1 Then Text1(Index).BackColor = vbGreen
    End If
    
salRefresh:
    Refresca_Datos
End Sub



'Private Sub Txtconsulta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 And Txtconsulta.TexT <> "" Then
'        KeyAscii = 0
'        Consulta_Art
'    End If
'
'    If OptSeccion(0).Value = True Then
'        Txtconsulta.Locked = Checktexto(KeyAscii, 6, Txtconsulta.TexT, True, False)
'    Else
'        Txtconsulta.Locked = Checktexto(KeyAscii, 4, Txtconsulta.TexT, True, False)
'    End If
'End Sub

Private Sub Vertical_Change()
    PTeclas.top = 0 - Vertical.Value * 100
End Sub

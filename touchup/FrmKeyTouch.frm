VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmKeyTouch 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOUCH-SCALE TECLAS DIRECTAS."
   ClientHeight    =   10950
   ClientLeft      =   1050
   ClientTop       =   600
   ClientWidth     =   15420
   Icon            =   "FrmKeyTouch.frx":0000
   LinkTopic       =   "touch-scale"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15420
   Begin MSComctlLib.ListView Lista 
      Height          =   5175
      Left            =   0
      TabIndex        =   36
      Top             =   5760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descriptivo"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Timer TmrEvents 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   0
   End
   Begin VB.TextBox TxtVer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   41
      Top             =   9600
      Width           =   7095
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Salir."
      Height          =   735
      Left            =   11280
      Picture         =   "FrmKeyTouch.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton CmdEnvCambios 
      Caption         =   "Envio Cambios."
      Height          =   735
      Left            =   13920
      Picture         =   "FrmKeyTouch.frx":1FBC
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   10200
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListaImg 
      Height          =   1095
      Left            =   12360
      TabIndex        =   38
      Top             =   9720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   6240
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.FileListBox imgDir 
      Height          =   480
      Left            =   6960
      TabIndex        =   37
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   18
      Left            =   6840
      TabIndex        =   94
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   93
      Top             =   10440
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   92
      Top             =   10440
      Width           =   1575
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   91
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   90
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   89
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   39
      Left            =   14040
      Picture         =   "FrmKeyTouch.frx":23FE
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   39
      Left            =   14040
      TabIndex        =   88
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   39
      Left            =   14040
      TabIndex        =   87
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   38
      Left            =   12600
      Picture         =   "FrmKeyTouch.frx":4208
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   38
      Left            =   12600
      TabIndex        =   86
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   38
      Left            =   12600
      TabIndex        =   85
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   37
      Left            =   11160
      Picture         =   "FrmKeyTouch.frx":6012
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   37
      Left            =   11160
      TabIndex        =   84
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   37
      Left            =   11160
      TabIndex        =   83
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   36
      Left            =   9720
      Picture         =   "FrmKeyTouch.frx":7E1C
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   36
      Left            =   9720
      TabIndex        =   82
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   36
      Left            =   9720
      TabIndex        =   81
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   35
      Left            =   8280
      Picture         =   "FrmKeyTouch.frx":9C26
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   35
      Left            =   8280
      TabIndex        =   80
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   35
      Left            =   8280
      TabIndex        =   79
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   34
      Left            =   6840
      Picture         =   "FrmKeyTouch.frx":BA30
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   34
      Left            =   6840
      TabIndex        =   78
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   34
      Left            =   6840
      TabIndex        =   77
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   33
      Left            =   5400
      Picture         =   "FrmKeyTouch.frx":D83A
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   33
      Left            =   5400
      TabIndex        =   76
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   33
      Left            =   5400
      TabIndex        =   75
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   5
      X1              =   3840
      X2              =   15360
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   32
      Left            =   3960
      Picture         =   "FrmKeyTouch.frx":F644
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   32
      Left            =   3960
      TabIndex        =   74
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   32
      Left            =   3960
      TabIndex        =   73
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   31
      Left            =   14040
      Picture         =   "FrmKeyTouch.frx":1144E
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   31
      Left            =   14040
      TabIndex        =   72
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   31
      Left            =   14040
      TabIndex        =   71
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   30
      Left            =   12600
      Picture         =   "FrmKeyTouch.frx":13258
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   30
      Left            =   12600
      TabIndex        =   70
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   30
      Left            =   12600
      TabIndex        =   69
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   29
      Left            =   11160
      Picture         =   "FrmKeyTouch.frx":15062
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   29
      Left            =   11160
      TabIndex        =   68
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   29
      Left            =   11160
      TabIndex        =   67
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   28
      Left            =   9720
      Picture         =   "FrmKeyTouch.frx":16E6C
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   28
      Left            =   9720
      TabIndex        =   66
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   28
      Left            =   9720
      TabIndex        =   65
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   27
      Left            =   8280
      Picture         =   "FrmKeyTouch.frx":18C76
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   27
      Left            =   8280
      TabIndex        =   64
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   27
      Left            =   8280
      TabIndex        =   63
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   26
      Left            =   6840
      Picture         =   "FrmKeyTouch.frx":1AA80
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   26
      Left            =   6840
      TabIndex        =   62
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   26
      Left            =   6840
      TabIndex        =   61
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   25
      Left            =   5400
      Picture         =   "FrmKeyTouch.frx":1C88A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   25
      Left            =   5400
      TabIndex        =   60
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   25
      Left            =   5400
      TabIndex        =   59
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   24
      Left            =   3960
      Picture         =   "FrmKeyTouch.frx":1E694
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   24
      Left            =   3960
      TabIndex        =   58
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   24
      Left            =   3960
      TabIndex        =   57
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   23
      Left            =   14040
      Picture         =   "FrmKeyTouch.frx":2049E
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   23
      Left            =   14040
      TabIndex        =   56
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   23
      Left            =   14040
      TabIndex        =   55
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   22
      Left            =   12600
      Picture         =   "FrmKeyTouch.frx":222A8
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   22
      Left            =   12600
      TabIndex        =   54
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   22
      Left            =   12600
      TabIndex        =   53
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   21
      Left            =   11160
      Picture         =   "FrmKeyTouch.frx":240B2
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   21
      Left            =   11160
      TabIndex        =   52
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   21
      Left            =   11160
      TabIndex        =   51
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   20
      Left            =   9720
      Picture         =   "FrmKeyTouch.frx":25EBC
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   20
      Left            =   9720
      TabIndex        =   50
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   20
      Left            =   9720
      TabIndex        =   49
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   19
      Left            =   8280
      Picture         =   "FrmKeyTouch.frx":27CC6
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   19
      Left            =   8280
      TabIndex        =   48
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   19
      Left            =   8280
      TabIndex        =   47
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   18
      Left            =   6840
      Picture         =   "FrmKeyTouch.frx":29AD0
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   18
      Left            =   6840
      TabIndex        =   46
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   17
      Left            =   5400
      Picture         =   "FrmKeyTouch.frx":2B8DA
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   17
      Left            =   5400
      TabIndex        =   45
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   17
      Left            =   5400
      TabIndex        =   44
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   16
      Left            =   3960
      Picture         =   "FrmKeyTouch.frx":2D6E4
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   16
      Left            =   3960
      TabIndex        =   43
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   16
      Left            =   3960
      TabIndex        =   42
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   4
      X1              =   3840
      X2              =   15360
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   13920
      X2              =   13920
      Y1              =   120
      Y2              =   9600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   12480
      X2              =   12480
      Y1              =   120
      Y2              =   9600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   11040
      X2              =   11040
      Y1              =   120
      Y2              =   9600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   15350
      X2              =   15350
      Y1              =   0
      Y2              =   10920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9600
      X2              =   9600
      Y1              =   120
      Y2              =   9720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   8160
      X2              =   8160
      Y1              =   120
      Y2              =   9720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   6720
      X2              =   6720
      Y1              =   120
      Y2              =   9720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   9600
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   3840
      X2              =   15360
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   3840
      X2              =   15360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   3840
      X2              =   15360
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LineV 
      BorderColor     =   &H80000005&
      X1              =   3840
      X2              =   3840
      Y1              =   10920
      Y2              =   0
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   3840
      X2              =   15360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label LblGrp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   3960
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblGrp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgGrp 
      Height          =   1455
      Index           =   2
      Left            =   2280
      Picture         =   "FrmKeyTouch.frx":2F4EE
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label LblGrp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   360
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgGrp 
      Height          =   1455
      Index           =   1
      Left            =   2280
      Picture         =   "FrmKeyTouch.frx":312F8
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   15
      Left            =   14040
      TabIndex        =   32
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   14
      Left            =   12600
      TabIndex        =   31
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   13
      Left            =   11160
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   12
      Left            =   9720
      TabIndex        =   29
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   8280
      TabIndex        =   28
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   6840
      TabIndex        =   27
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   5400
      TabIndex        =   26
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   14040
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   12600
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   11160
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   21
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   15
      Left            =   14040
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   14
      Left            =   12600
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   13
      Left            =   11160
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   12
      Left            =   9720
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA HOJA"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   8280
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA DEL ARBOL"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   6840
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   5400
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "CLEMENTINA GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   14160
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "POMELO AMARILLO"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   12600
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "POMELO ROJO"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   11160
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "LIMON GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA DE MESA"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA ZUMO"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   15
      Left            =   14040
      Picture         =   "FrmKeyTouch.frx":33102
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   14
      Left            =   12600
      Picture         =   "FrmKeyTouch.frx":34F0C
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   13
      Left            =   11160
      Picture         =   "FrmKeyTouch.frx":36D16
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   12
      Left            =   9720
      Picture         =   "FrmKeyTouch.frx":38B20
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   11
      Left            =   8280
      Picture         =   "FrmKeyTouch.frx":3A92A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   10
      Left            =   6840
      Picture         =   "FrmKeyTouch.frx":3C734
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   9
      Left            =   5400
      Picture         =   "FrmKeyTouch.frx":3E53E
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   8
      Left            =   3960
      Picture         =   "FrmKeyTouch.frx":40348
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   7
      Left            =   14040
      Picture         =   "FrmKeyTouch.frx":42152
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   6
      Left            =   12600
      Picture         =   "FrmKeyTouch.frx":43F5C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   5
      Left            =   11160
      Picture         =   "FrmKeyTouch.frx":45D66
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   4
      Left            =   9720
      Picture         =   "FrmKeyTouch.frx":47B70
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   3
      Left            =   8280
      Picture         =   "FrmKeyTouch.frx":4997A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   2
      Left            =   6840
      Picture         =   "FrmKeyTouch.frx":4B784
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   1
      Left            =   5400
      Picture         =   "FrmKeyTouch.frx":4D58E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   5
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":4F398
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   4
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":511A2
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   3
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":52FAC
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   2
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":54DB6
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   1
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":56BC0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image ImgTec 
      Height          =   855
      Index           =   0
      Left            =   3960
      Picture         =   "FrmKeyTouch.frx":589CA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image ImgFam 
      Height          =   1095
      Index           =   0
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":5A7D4
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Image ImgGrp 
      Height          =   1455
      Index           =   0
      Left            =   2280
      Picture         =   "FrmKeyTouch.frx":5C5DE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu CmdServidora 
      Caption         =   "Servidora"
   End
End
Attribute VB_Name = "FrmKeyTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CmdEnvCambios_Click()
    Call mainEnvCambios
End Sub


Private Sub CmdExit_Click()
    Call Form_Unload(0)
End Sub




Public Sub CmdServidora_Click()
    FrmKeyTouch.Enabled = False
    Load FrmConfiguracion
    FrmConfiguracion.Show
    Do While FrmConfiguracion.Visible
        Do_Events
    Loop
    FrmKeyTouch.Enabled = True
End Sub

Private Sub Form_Activate()
    Call ImgGrp_Click(0)
End Sub


Private Sub Form_Load()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyBaseDES As DAO.Database
Dim MyRsDes As DAO.Recordset
Dim MyRsITM As DAO.Recordset
Dim nF As Integer
Dim nnF As Integer
Dim nnnF As Integer
Dim sPath() As String
Dim nCont As Integer
Dim sCfgSS(6) As String
Dim sGrpSS() As String
Dim sSubGrpSS() As String
Dim nGrp As Integer
Dim sS As String
Dim Token As Long
Dim n As Integer
Dim nfich As Integer

'nfich = FreeFile()
'Open App.Path & "\tiempo.txt" For Append As #nfich
'Print #nfich, Now
    
    'Solo se harán visibles los grupos-->Familias que estén dados de alta en tabla grp de touchscale.mdb
    For nnF = 0 To 2
        FrmKeyTouch.ImgGrp(nnF).Enabled = False
        FrmKeyTouch.LblGrp(nnF).Enabled = False
        FrmKeyTouch.ImgGrp(nnF).Visible = False
        FrmKeyTouch.LblGrp(nnF).Visible = False
    Next nnF
    'Carga de Grupos --> Familias, Subgrupos --> Particiones en pantallas de teclas de 40 en 40
    ' y artículos según su key (posición), familia
    Set MyBaseDES = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsDes = MyBaseDES.OpenRecordset("select * from grp order by grp")
    MyRsDes.MoveLast
    ReDim TTecGrupos.Indice(MyRsDes.RecordCount)
    nGrupos = MyRsDes.RecordCount
    MyRsDes.MoveFirst
    Do While Not MyRsDes.EOF
        nnF = MyRsDes.Fields("grp") - 1
        TTecGrupos.Indice(nnF).Indice = nnF
        FrmKeyTouch.ImgGrp(nnF).Enabled = True
        FrmKeyTouch.LblGrp(nnF).Enabled = True
        FrmKeyTouch.ImgGrp(nnF).Visible = True
        FrmKeyTouch.LblGrp(nnF).Visible = True
        TTecGrupos.Indice(nnF).lblTxt = MyRsDes.Fields("text")
        TTecGrupos.Indice(nnF).SubGroupnr = MyRsDes.Fields("subgroupnr")
        TTecGrupos.Indice(nnF).IniKey = MyRsDes.Fields("inikey")
        TTecGrupos.Indice(nnF).EndKey = MyRsDes.Fields("endkey")
        TTecGrupos.Indice(nnF).ImgNoSelect = MyRsDes.Fields("imgnoselect")
        TTecGrupos.Indice(nnF).ImgSelect = MyRsDes.Fields("imgselect")
        ReDim TTecGrupos.Indice(nnF).SubGrupo(MyRsDes.Fields("subgroupnr"))
        Token = InitGDIPlus()
        'If nnF = 0 Then
            FrmKeyTouch.ImgGrp(nnF).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsDes.Fields("imgselect")))
        'Else
        '    FrmKeyTouch.ImgGrp(nnF).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsDes.Fields("imgnoselect")))
        'End If
        FreeGDIPlus Token
        FrmKeyTouch.LblGrp(nnF).Caption = MyRsDes.Fields("text")
        Set MyRsORG = MyBaseDES.OpenRecordset("select * from subgroup where grp=" & CStr(nnF + 1) & " order by subgroup")
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            nnnF = MyRsORG.Fields("subgroup") - 1
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).Grupo = MyRsORG.Fields("grp")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).Indice = MyRsORG.Fields("subgroup") - 1
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).lblTxt = MyRsORG.Fields("text")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey = MyRsORG.Fields("inikey")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey = MyRsORG.Fields("endkey")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgNoSelect = MyRsORG.Fields("imgnoselect")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgSelect = MyRsORG.Fields("imgselect")
            If nnF = 0 Then
                'If nnnF = 0 Then
                    Token = InitGDIPlus()
                    FrmKeyTouch.ImgFam(nnnF).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsORG.Fields("imgselect")))
                    FreeGDIPlus Token
                    'carga items primera pantalla
                    For nCont = 0 To 39
                        FrmKeyTouch.ImgTec(nCont).Visible = False
                        FrmKeyTouch.LblTec1(nCont).Visible = False
                        FrmKeyTouch.LblTec2(nCont).Visible = False
                    Next nCont
                    nCont = 0
                    For nF = TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey To TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey
                        FrmKeyTouch.ImgTec(nCont).Visible = True
                        FrmKeyTouch.LblTec1(nCont).Visible = True
                        FrmKeyTouch.LblTec2(nCont).Visible = True
                        FrmKeyTouch.LblTec1(nCont).Caption = ""
                        FrmKeyTouch.LblTec2(nCont).Caption = CStr(nF)
                        Set MyRsITM = MyBaseDES.OpenRecordset("select * from items where key=" & CStr(nF) & " and grp=" & CStr(MyRsDes.Fields("grp")))
                        If Not MyRsITM.EOF Then
                            Token = InitGDIPlus()
                            FrmKeyTouch.ImgTec(nCont).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsITM.Fields("icon_name")))
                            FreeGDIPlus Token
                            FrmKeyTouch.LblTec1(nCont).Caption = Trim(MyRsITM.Fields("shortname"))
                        Else
                            'Token = InitGDIPlus()
                            'FrmKeyTouch.ImgTec(nCont).Picture = LoadPictureGDIPlus(App.Path & "\images\" & "LCarrefour.png")
                            'FreeGDIPlus Token
                            FrmKeyTouch.ImgTec(nCont).Picture = Nothing
                        End If
                        nCont = nCont + 1
                    Next nF
                'Else
                '    Token = InitGDIPlus()
                '    FrmKeyTouch.ImgFam(nnnF).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsORG.Fields("imgnoselect")))
                '    FreeGDIPlus Token
                'End If
                FrmKeyTouch.LblFam(nnnF).Caption = MyRsORG.Fields("text")
            End If
            
            
            MyRsORG.MoveNext
        Loop
        MyRsDes.MoveNext
    Loop
    
    
    nGrupoActual = 0
    nSubGrupoActual = 0
    
    On Error Resume Next
    MyRsDes.Close
    MyRsORG.Close
    MyBaseDES.Close
    Set MyBaseDES = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    Set MyRsDes = Nothing
    Set MyRsORG = Nothing
    On Error GoTo 0
    
    LblGrp(0).ForeColor = vbBlack
    LblFam(0).ForeColor = vbBlack
    
    If Lista.ListItems.Count = 0 And (LCase(Trim(Command)) = "hydra" Or Trim(Command) = "") Then
    
        'FrmKeyTouch.ListaImg.ListItems.Clear
        'Set FrmKeyTouch.ListaImg.Icons = Nothing
        'FrmKeyTouch.Imagenes.ListImages.Clear
        FrmKeyTouch.Lista.ListItems.Clear
        If lAllItem Then
            Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where borrado=false order by codigo")
            With MyRsORG
                
                If Not .EOF Then
                    Do Until .EOF
                        
                        FrmKeyTouch.Lista.ListItems.Add , , Format(!codigo, "000000")
                        'FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = Format(!PLU, "0000")
                        
                        If Not IsNull(!des_plu1) Then
                            FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = !des_plu1
                        Else
                            FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = Space(25)
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            
        Else
            Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where (borrado=false and code<>0) order by code")
        
            With MyRsORG
                
                If Not .EOF Then
                    Do Until .EOF
                        
                        FrmKeyTouch.Lista.ListItems.Add , , Format(!Code, "000000")
                        'FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = Format(!PLU, "0000")
                        
                        If Not IsNull(!ShortName) Then
                            FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = !ShortName
                        Else
                            FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = Space(25)
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        
        End If
    
        MyRsORG.Close
        Set MyRsORG = Nothing
        MyBaseORG.Close
        Set MyBaseORG = Nothing
    
        'versión 3
        'FrmKeyTouch.imgDir.Path = App.Path & "\images"
        'FrmKeyTouch.imgDir.Pattern = "*.*"
        'FrmKeyTouch.imgDir.Refresh
        'For n = 1 To FrmKeyTouch.imgDir.ListCount
        '    If InStr(1, Trim(imgDir.List(n - 1)), ".bmp") <> 0 Or InStr(1, Trim(imgDir.List(n - 1)), ".jpg") <> 0 Or InStr(1, Trim(imgDir.List(n - 1)), ".png") <> 0 Then
        '        Token = InitGDIPlus()
        '        On Error Resume Next
        '        FrmKeyTouch.Imagenes.ListImages.Add n, , LoadPictureGDIPlus(App.Path & "\images\" & Trim(imgDir.List(n - 1)))
        '        On Error GoTo 0
        '        FreeGDIPlus Token
        '    End If
        'Next n
        'FrmKeyTouch.ListaImg.Icons = Imagenes
        'For n = 1 To FrmKeyTouch.imgDir.ListCount
        '    On Error Resume Next
        '    FrmKeyTouch.ListaImg.ListItems.Add n, , FrmKeyTouch.imgDir.List(n - 1), n
        '    On Error GoTo 0
        'Next n
        ''''''''''
    
    End If

'Print #nfich, Now
'Close #nfich
    
End Sub



''Secuencia MouseDown
''          CLICK
''          MouseUp
'Private Sub ImgTec_Click(Index As Integer)
'Dim a As Variant
'a = a'
'
'End Sub

'Private Sub ImgTec_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim a As Variant
'a = a''
'
'End Sub

Public Sub cargaSubGrupoTeclas(ByVal Indice As Integer, ByVal n As Integer)
Dim nn As Integer
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim Token As Long
Dim cm As String
    
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")

    For nn = 0 To 39
        ImgTec(nn).Visible = False
        LblTec1(nn).Visible = False
        LblTec2(nn).Visible = False
        'ImgQ(nn).Visible = False
    Next nn
    
    For nn = TTecGrupos.Indice(Indice).SubGrupo(n).IniKey To TTecGrupos.Indice(Indice).SubGrupo(n).EndKey
        ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec2(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec2(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = CStr(nn)
        Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where (code <>0 and key=" & CStr(nn) & " and family=" & CStr(MyFam) & ")")
        If Not MyRsORG.EOF Then
            If Trim(MyRsORG.Fields("icon_name")) <> "" Then
                Token = InitGDIPlus()
                ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(MyRsORG.Fields("icon_name")))
                FreeGDIPlus Token
            Else
                'Token = InitGDIPlus()
                'ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPictureGDIPlus(App.Path & "\images\" & "LCarrefour.png")
                'FreeGDIPlus Token
                ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = Nothing
            End If
            LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = Trim(MyRsORG.Fields("shortname"))
        Else
            'Token = InitGDIPlus()
            'ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPictureGDIPlus(App.Path & "\images\LCarrefour.png")
            'FreeGDIPlus Token
            ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = Nothing
            LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = ""
        End If
    Next nn
    
    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    
End Sub

Public Sub cargaSubGrupos(ByVal Index As Integer)
Dim n As Integer
Dim Token As Long
    For n = 0 To 5
        ImgFam(n).Visible = False
        LblFam(n).Visible = False
    Next n
    For n = 0 To TTecGrupos.Indice(Index).SubGroupnr - 1
        ImgFam(n).Visible = True
        LblFam(n).Visible = True
        If n = 0 Then
            Token = InitGDIPlus()
            ImgFam(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(Index).SubGrupo(n).ImgSelect))
            FreeGDIPlus Token
            LblFam(n).ForeColor = vbBlack
            Call cargaSubGrupoTeclas(Index, n)
        Else
            Token = InitGDIPlus()
            ImgFam(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(Index).SubGrupo(n).ImgNoSelect))
            FreeGDIPlus Token
            LblFam(n).ForeColor = vbWhite
        End If
        LblFam(n).Caption = TTecGrupos.Indice(Index).SubGrupo(n).lblTxt
    Next n
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmSelect
    End
End Sub


Private Sub ImgTec_Carga(Index As Integer)
Dim n As Variant
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
    TeclaTmp.Key = Val(LblTec2(Index).Caption)
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(TeclaTmp.Key))
    If Not MyRsORG.EOF Then 'edición
        TeclaTmp.Code = MyRsORG.Fields("code")
        'TeclaTmp.Counter = MyRsORG.Fields("counter")
        If IsNull(Trim(MyRsORG.Fields("icon_name"))) Then
            TeclaTmp.icon_name = "" '"LCarrefour.png"
        Else
            TeclaTmp.icon_name = Trim(MyRsORG.Fields("icon_name"))
        End If
        If IsNull(MyRsORG.Fields("shortname")) Then
            TeclaTmp.Name = ""
        Else
            TeclaTmp.Name = MyRsORG.Fields("shortname")
        End If
        TeclaTmp.PLU = MyRsORG.Fields("plu")
        TeclaTmp.qLevel = MyRsORG.Fields("qualitynr")
        'If Not IsNull(MyRsORG.Fields("qualityname")) Then
        '    TeclaTmp.qLevelName = Trim(MyRsORG.Fields("qualityname"))
        'Else
        '    TeclaTmp.qLevelName = ""
        'End If
        TeclaTmp.Section = MyRsORG.Fields("section")
        TeclaTmp.ShortName = MyRsORG.Fields("shortname")

        ImgTec(Index).ToolTipText = "Código:" & CStr(TeclaTmp.Code) & " PLU:" & CStr(TeclaTmp.PLU) & " Imagen:" & TeclaTmp.icon_name
    Else 'ne existe pa ¿Se escribe así?
        ImgTec(Index).ToolTipText = ""
    End If

    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing '

End Sub

'Private Sub ImgTec_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim n
'n = n
'End Sub

Private Sub LblGrp_Click(Index As Integer)
    Call ImgGrp_Click(Index)
End Sub
Private Sub ImgGrp_Click(Index As Integer)
Dim n As Integer
Dim Token As Long

    If FrmSelect.Visible = False Then
    
    nGrupoActual = Index
    MyFam = Index + 1
    
    For n = 0 To nGrupos - 1
        If Index = n Then
            Token = InitGDIPlus()
            ImgGrp(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(n).ImgSelect))
            FreeGDIPlus Token
            LblGrp(n).ForeColor = vbGreen
            LblGrp(n).Font.Underline = True
            LblGrp(n).Font.Bold = True
            LblGrp(n).Font.size = 16
            Call cargaSubGrupos(Index)
        Else
            Token = InitGDIPlus()
            ImgGrp(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(n).ImgNoSelect))
            FreeGDIPlus Token
            LblGrp(n).ForeColor = vbWhite
            LblGrp(n).Font.Underline = False
            LblGrp(n).Font.Bold = False
            LblGrp(n).Font.size = 9
        End If
    Next n
    'ojo
    'If FrmSelect.Visible = True Then
        Call ImgFam_Click(0)
    'End If
    ''''
    FrmKeyTouch.Refresh
    
    End If
    
End Sub
Private Sub LblFam_Click(Index As Integer)
    Call ImgFam_Click(Index)
End Sub
Private Sub ImgFam_Click(Index As Integer)
Dim n As Integer
Dim Token As Long

If FrmSelect.Visible = False Then
    nSubGrupoActual = Index
    For n = 0 To TTecGrupos.Indice(nGrupoActual).SubGroupnr - 1
        If Index = n Then
            Token = InitGDIPlus()
            ImgFam(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(nGrupoActual).SubGrupo(n).ImgSelect))
            FreeGDIPlus Token
            
            LblFam(n).ForeColor = vbGreen
            LblFam(n).Font.Underline = True
            
            Call cargaSubGrupoTeclas(nGrupoActual, nSubGrupoActual)
            'ojo
            If FrmSelect.Visible = True Then
                Lista.Sorted = False
                Call imgTecClick(1000) 'prueba
                Lista.Sorted = False
                FrmSelect.Caption = "TECLAS:" & Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey, "000") & " Hasta " & _
                         Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).EndKey, "000")
            End If
            ''''
        Else
            Token = InitGDIPlus()
            ImgFam(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(nGrupoActual).SubGrupo(n).ImgNoSelect))
            FreeGDIPlus Token
            
            LblFam(n).ForeColor = vbWhite
            LblFam(n).Font.Underline = False

        End If
    Next n
    FrmKeyTouch.Refresh
End If
End Sub
Private Sub LblTec1_Click(Index As Integer)
    Call ImgTec_Click(Index)
    Call ImgTec_Carga(Index)
End Sub
Private Sub LblTec2_Click(Index As Integer)
    Call ImgTec_Click(Index)
    Call ImgTec_Carga(Index)
End Sub

Private Sub ImgTec_Click(Index As Integer)

    Call imgTecClick(Index)
    Call ImgTec_Carga(Index)
    
End Sub

Public Sub imgTecClick(ByVal Index As Integer, Optional ByVal nTec As Integer)
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim n As Integer
Dim nn As Integer
Dim nF As Integer
Dim Token As Long


    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    If Index < 1000 Then
        TeclaTmp.Key = Val(LblTec2(Index).Caption)
    Else
        If nTec <> 0 Then
            TeclaTmp.Key = nTec
        Else
            TeclaTmp.Key = TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey
        End If
    End If
    TeclaTmp.Group = nGrupoActual
    TeclaTmp.SubGroup = nSubGrupoActual
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(TeclaTmp.Key) & " and family=" & CStr(MyFam))
    If Not MyRsORG.EOF Then 'edición
        TeclaTmp.Code = MyRsORG.Fields("code")
        TeclaTmp.Counter = MyRsORG.Fields("counter")
        If IsNull(MyRsORG.Fields("icon_name")) Then
            TeclaTmp.icon_name = "" '"LCarrefour.png"
        Else
            TeclaTmp.icon_name = Trim(MyRsORG.Fields("icon_name"))
        End If
        If IsNull(MyRsORG.Fields("shortname")) Then
            TeclaTmp.Name = ""
        Else
            TeclaTmp.Name = MyRsORG.Fields("shortname")
        End If
        TeclaTmp.PLU = MyRsORG.Fields("plu")
        TeclaTmp.qLevel = MyRsORG.Fields("qualitynr")
        If Not IsNull(MyRsORG.Fields("qualityname")) Then
            TeclaTmp.qLevelName = Trim(MyRsORG.Fields("qualityname"))
        Else
            TeclaTmp.qLevelName = ""
        End If
        TeclaTmp.Section = MyRsORG.Fields("section")
        TeclaTmp.ShortName = MyRsORG.Fields("shortname")
    Else 'ne existe pa ¿Se escribe así?
        TeclaTmp.Code = 0
        TeclaTmp.icon_name = "" '"LCarrefour.png"
        TeclaTmp.Name = ""
        TeclaTmp.PLU = 0
        TeclaTmp.qLevel = 0
        TeclaTmp.qLevelName = ""
        TeclaTmp.ShortName = ""
        TeclaTmp.Counter = nCodiIdent
        TeclaTmp.Section = nSeccMaqui
    End If
    
    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    
    If Index < 1000 Then

        For n = 0 To 39 'veraquí
            LblTec1(n).Enabled = False
            LblTec2(n).Enabled = False
            ImgTec(n).Enabled = False
            'ImgQ(n).Enabled = False
        Next n
    
    End If
    
    FrmSelect.LblTec1(0).Caption = TeclaTmp.ShortName
    FrmSelect.LblTec2(0).Caption = CStr(TeclaTmp.Key)
    If TeclaTmp.icon_name = "" Then
        TeclaTmp.icon_name = "" '"LCarrefour.png"
        FrmSelect.ImgTec(0).Picture = Nothing
    Else
        Token = InitGDIPlus()
        FrmSelect.ImgTec(0).Picture = LoadPictureGDIPlus(Trim(App.Path & "\images\" & Trim(TeclaTmp.icon_name)))
        FreeGDIPlus Token
    End If
    
    Load FrmSelect
    FrmSelect.Show
    
    If Index < 1000 Then
        Lista.Visible = True
        'ListaImg.Visible = True
    End If

End Sub


Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Do_Events
    If Lista.SortOrder = lvwAscending Then
        Lista.SortOrder = lvwDescending
    Else
        Lista.SortOrder = lvwAscending
    End If
    Lista.SortKey = ColumnHeader.Index - 1
    Lista.Sorted = True
End Sub

Public Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim nCod As Long
    Dim sDes As String
    Dim nPlu As Integer
    
    nCod = Item
    sDes = Item.SubItems(1)
    TeclaTmp.Code = nCod
    TeclaTmp.Name = sDes
    TeclaTmp.ShortName = sDes
    FrmSelect.LblTec1(0).Caption = Trim(sDes)
    FrmSelect.LblTec2(0).Caption = TeclaTmp.Key
        
End Sub

'Private Sub ListaImg_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    Dim sIcon As String
'    Dim Token As Long
'    TeclaTmp.icon_name = Item
'    Token = InitGDIPlus()
'    FrmSelect.ImgTec(0).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TeclaTmp.icon_name))
'    FreeGDIPlus Token
'End Sub

Private Sub TmrEvents_Timer()
    Do_Events
End Sub

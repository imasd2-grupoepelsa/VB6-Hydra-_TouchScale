VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmKeyTouch 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOUCH-SCALE TECLAS DIRECTAS."
   ClientHeight    =   10965
   ClientLeft      =   -30
   ClientTop       =   255
   ClientWidth     =   15420
   Icon            =   "FrmKeyTouch.frx":0000
   LinkTopic       =   "touch-scale"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   15420
   Begin VB.Timer TmrEvents 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   7200
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
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   56
      Top             =   9600
      Width           =   7095
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Salir."
      Height          =   735
      Left            =   11280
      Picture         =   "FrmKeyTouch.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton CmdEnvTodo 
      Caption         =   "Envio Datos."
      Height          =   735
      Left            =   12480
      Picture         =   "FrmKeyTouch.frx":1FBC
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9960
      Width           =   1335
   End
   Begin VB.CommandButton CmdEnvCambios 
      Caption         =   "Envio Cambios."
      Height          =   735
      Left            =   13920
      Picture         =   "FrmKeyTouch.frx":2106
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   9960
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImagenesCalidad 
      Left            =   1680
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListaCalidad 
      Height          =   4815
      Left            =   8880
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8493
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
   Begin MSComctlLib.ListView ListaImg 
      Height          =   6375
      Left            =   3840
      TabIndex        =   51
      Top             =   4800
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11245
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
      TabIndex        =   50
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4815
      Left            =   3840
      TabIndex        =   49
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8493
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PLU"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descriptivo"
         Object.Width           =   5292
      EndProperty
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
      X1              =   13080
      X2              =   13080
      Y1              =   0
      Y2              =   9600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   10800
      X2              =   10800
      Y1              =   0
      Y2              =   9600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   8520
      X2              =   8520
      Y1              =   0
      Y2              =   9600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   6240
      X2              =   6240
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
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   3840
      X2              =   15360
      Y1              =   4800
      Y2              =   4800
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
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   19
      Left            =   13200
      Top             =   7800
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   18
      Left            =   10920
      Top             =   7800
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   17
      Left            =   8640
      Top             =   7800
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   16
      Left            =   6360
      Top             =   7800
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   15
      Left            =   4080
      Top             =   7800
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   14
      Left            =   13200
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   13
      Left            =   10920
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   12
      Left            =   8640
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   11
      Left            =   6360
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   10
      Left            =   4080
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   9
      Left            =   13200
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   8
      Left            =   10920
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   7
      Left            =   8640
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   6
      Left            =   6360
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   5
      Left            =   4080
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   4
      Left            =   13200
      Top             =   600
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   3
      Left            =   10920
      Top             =   600
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   2
      Left            =   8640
      Top             =   600
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   1
      Left            =   6360
      Top             =   600
      Width           =   975
   End
   Begin VB.Image ImgQ 
      Height          =   615
      Index           =   0
      Left            =   4080
      Top             =   600
      Width           =   975
   End
   Begin VB.Label LblGrp 
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
      Height          =   1335
      Index           =   2
      Left            =   240
      TabIndex        =   48
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label LblGrp 
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
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   47
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image ImgGrp 
      Height          =   1575
      Index           =   2
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label LblGrp 
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
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   46
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image ImgGrp 
      Height          =   1575
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "20"
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
      Left            =   13200
      TabIndex        =   45
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "19"
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
      Left            =   10920
      TabIndex        =   44
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "18"
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
      Left            =   8640
      TabIndex        =   43
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "17"
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
      Left            =   6360
      TabIndex        =   42
      Top             =   9120
      Width           =   2055
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
      Height          =   375
      Index           =   15
      Left            =   4080
      TabIndex        =   41
      Top             =   9120
      Width           =   2055
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
      Left            =   13200
      TabIndex        =   40
      Top             =   6720
      Width           =   2055
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
      Left            =   10920
      TabIndex        =   39
      Top             =   6720
      Width           =   2055
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
      Left            =   8640
      TabIndex        =   38
      Top             =   6720
      Width           =   2055
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
      Left            =   6360
      TabIndex        =   37
      Top             =   6720
      Width           =   2055
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
      Left            =   4080
      TabIndex        =   36
      Top             =   6720
      Width           =   2055
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
      Left            =   13200
      TabIndex        =   35
      Top             =   4320
      Width           =   2055
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
      Left            =   10920
      TabIndex        =   34
      Top             =   4320
      Width           =   2055
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
      Left            =   8640
      TabIndex        =   33
      Top             =   4320
      Width           =   2055
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
      Left            =   6360
      TabIndex        =   32
      Top             =   4320
      Width           =   2055
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
      Left            =   4080
      TabIndex        =   31
      Top             =   4320
      Width           =   2055
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
      Left            =   13200
      TabIndex        =   30
      Top             =   1920
      Width           =   2055
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
      Left            =   10920
      TabIndex        =   29
      Top             =   1920
      Width           =   2055
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
      Left            =   8640
      TabIndex        =   28
      Top             =   1920
      Width           =   2055
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
      Left            =   6360
      TabIndex        =   27
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   19
      Left            =   13200
      TabIndex        =   26
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   18
      Left            =   10920
      TabIndex        =   25
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   17
      Left            =   8640
      TabIndex        =   24
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   16
      Left            =   6360
      TabIndex        =   23
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   15
      Left            =   4080
      TabIndex        =   22
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   14
      Left            =   13200
      TabIndex        =   21
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   13
      Left            =   10920
      TabIndex        =   20
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   12
      Left            =   8640
      TabIndex        =   19
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA HOJA"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   11
      Left            =   6360
      TabIndex        =   18
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA DEL ARBOL"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   10
      Left            =   4080
      TabIndex        =   17
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   9
      Left            =   13200
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "CLEMENTINA GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   8
      Left            =   10920
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MANDARINA GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   7
      Left            =   8640
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "POMELO AMARILLO"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   6
      Left            =   6360
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "POMELO ROJO"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   5
      Left            =   4080
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "LIMON GRANEL"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   13200
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA DE MESA"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   3
      Left            =   10920
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA ZUMO"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   19
      Left            =   13200
      Picture         =   "FrmKeyTouch.frx":2548
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   18
      Left            =   10920
      Picture         =   "FrmKeyTouch.frx":1344A
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   17
      Left            =   8640
      Picture         =   "FrmKeyTouch.frx":2434C
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   16
      Left            =   6360
      Picture         =   "FrmKeyTouch.frx":3524E
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   15
      Left            =   4080
      Picture         =   "FrmKeyTouch.frx":46150
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   14
      Left            =   13200
      Picture         =   "FrmKeyTouch.frx":57052
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   13
      Left            =   10920
      Picture         =   "FrmKeyTouch.frx":67F54
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   12
      Left            =   8640
      Picture         =   "FrmKeyTouch.frx":78E56
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   11
      Left            =   6360
      Picture         =   "FrmKeyTouch.frx":89D58
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   10
      Left            =   4080
      Picture         =   "FrmKeyTouch.frx":9AC5A
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   9
      Left            =   13200
      Picture         =   "FrmKeyTouch.frx":ABB5C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   8
      Left            =   10920
      Picture         =   "FrmKeyTouch.frx":BCA5E
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   7
      Left            =   8640
      Picture         =   "FrmKeyTouch.frx":CD960
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   6
      Left            =   6360
      Picture         =   "FrmKeyTouch.frx":DE862
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   5
      Left            =   4080
      Picture         =   "FrmKeyTouch.frx":EF764
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   4
      Left            =   13200
      Picture         =   "FrmKeyTouch.frx":100666
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   3
      Left            =   10920
      Picture         =   "FrmKeyTouch.frx":111568
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   2
      Left            =   8640
      Picture         =   "FrmKeyTouch.frx":12246A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   1
      Left            =   6360
      Picture         =   "FrmKeyTouch.frx":13336C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   " Otras Frutas           teclas 101-120"
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
      Height          =   615
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   " Melón/Sandia y    Plátanos 81-100"
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
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   " Frutas de Hueso   teclas 61 - 80"
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
      Height          =   615
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   " Peras teclas                41 - 60"
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
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   " Manzanas                teclas 21 - 40"
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
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   5
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":14426E
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   4
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":194170
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   3
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":1E4072
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   2
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":233F74
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   1
      Left            =   2040
      Picture         =   "FrmKeyTouch.frx":283E76
      Stretch         =   -1  'True
      Top             =   5640
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image ImgTec 
      Height          =   1335
      Index           =   0
      Left            =   4080
      Picture         =   "FrmKeyTouch.frx":2D3D78
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label LblFam 
      BackStyle       =   0  'Transparent
      Caption         =   "   Cítricos teclas             1 - 20"
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
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Image ImgFam 
      Height          =   1575
      Index           =   0
      Left            =   120
      Picture         =   "FrmKeyTouch.frx":2E4C7A
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Image ImgGrp 
      Height          =   1575
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "FrmKeyTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTodo As Boolean

Public Sub CmdEnvTodo_Click()
    Me.BorderStyle = 3
    CmdEnvTodo.Enabled = False
    CmdEnvCambios.Enabled = False
    CmdExit.Enabled = False
    lTodo = True
    Call CmdEnvCambios_Click
    CmdEnvTodo.Enabled = True
    CmdEnvCambios.Enabled = True
    CmdExit.Enabled = True
    Me.BorderStyle = 1
End Sub


Public Sub CmdEnvCambios_Click()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyBaseORG1 As DAO.Database
Dim MyRsORG1 As DAO.Recordset
Dim MyBaseBAL As DAO.Database
Dim MyRsBAL As DAO.Recordset
Dim sIp As String
Dim MyRsDes As New ADODB.Recordset
Dim myRsKEYS As New ADODB.Recordset
Dim myRsTMP As New ADODB.Recordset
Dim Cm As String
Dim nF As Integer

nF = FreeFile()
Open App.Path & "\LOGTOUCH.TXT" For Output As #nF
Print #nF, "Sincronización con TOUCH-SCALE.MDB -- DBASE.MDB"

    lComunica = True
    Me.BorderStyle = 3
    CmdEnvTodo.Enabled = False
    CmdEnvCambios.Enabled = False
    CmdExit.Enabled = False
    TmrEvents.Enabled = True

    'recogida datos articulos de dbase.mdb.
    'primero ver las bajas.
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyBaseORG1 = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where (borrado=true and secc_maqui=8)")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            Set MyRsORG1 = MyBaseORG1.OpenRecordset("select * from items where code=" & CStr(MyRsORG.Fields("codigo")))
            If Not MyRsORG1.EOF Then
                Do While Not MyRsORG1.EOF
                    MyRsORG1.Delete
                    MyRsORG1.MoveNext
                Loop
    
            End If
    
            MyRsORG.MoveNext
        Loop
    Else
        'NO HAY ARTÍCULOS EN DBASE.MDB
        'Ver acciones a realizar.
    End If
    'traspaso de datos... Dbase.mdb a TouchScale.mdb
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where (borrado=false and secc_maqui=8)")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            Set MyRsORG1 = MyBaseORG1.OpenRecordset("select * from items where code=" & CStr(MyRsORG.Fields("codigo")))
            If Not MyRsORG1.EOF Then
                Do While Not MyRsORG1.EOF
                    MyRsORG1.Edit
    
                    MyRsORG1.Fields("name") = Trim(MyRsORG.Fields("des_plu1"))
                    MyRsORG1.Fields("shortname") = Trim(MyRsORG.Fields("des_plu1"))
                    MyRsORG1.Fields("text") = "" '

                    MyRsORG1.Fields("family-codi_sub") = MyRsORG.Fields("codi_sub")
                    MyRsORG1.Fields("price") = MyRsORG.Fields("euros")
                    If MyRsORG.Fields("codi_pes") = "W" Then
                        MyRsORG1.Fields("saleform") = 1
                    Else
                        MyRsORG1.Fields("saleform") = 0
                    End If
                    MyRsORG1.Fields("plu") = MyRsORG.Fields("plu")
                    MyRsORG1.Fields("counter") = MyRsORG.Fields("secc_maqui")
                    'MyRsDes.Fields("vat") = MyRsORG.Fields("iva") 'localizar valor
                    MyRsORG1.Fields("ean13") = Trim(MyRsORG.Fields("ean13"))
                    MyRsORG1.Fields("expiring") = MyRsORG.Fields("caducidad")
                    MyRsORG1.Fields("bestbefore") = MyRsORG.Fields("pref")
                    MyRsORG1.Fields("tare") = MyRsORG.Fields("tara")
                    MyRsORG1.Fields("label_text") = "" 'stoutf8(Trim(MyRsORG.Fields("label_text")))
                    If IsNull(Trim(MyRsORG1.Fields("icon_name"))) Then
                        MyRsORG1.Fields("icon") = "pralcampo.bmp"
                    End If
                    'MyRsDes.Fields("qlevel") = MyRsORG.Fields("qualitynr")
    
                    MyRsORG1.Update
                    MyRsORG1.MoveNext
                Loop
    
            End If
    
            MyRsORG.MoveNext
        Loop
    End If
    MyRsORG.Close
    MyRsORG1.Close
    MyBaseORG.Close
    MyBaseORG1.Close
    
    
'CREATE TABLE IF NOT EXISTS QLevels
'       ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Code VARCHAR(8) UNIQUE NOT NULL, Icon TEXT NOT NULL,
'         Description TEXT NOT NULL );
'    Cm = "select * from mensajes where estado=0 order by orden"
'    FDb1.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
'    If Not FDb1.EOF Then
TxtVer.Text = ""
Set MyBaseBAL = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
Set MyRsBAL = MyBaseBAL.OpenRecordset("select * from balanzas where ip<>'127-0-0-1' order by scalenr")
If Not MyRsBAL.EOF Then
MyRsBAL.MoveFirst
Do While Not MyRsBAL.EOF

sIp = toIpAdd(Trim(MyRsBAL.Fields("ip")))

TxtVer.Text = TxtVer.Text & Now() & " --> Iniciando conexión nodo:" & sIp & vbCrLf
TxtVer.Refresh
Print #nF, Now() & " --> Iniciando conexión nodo:" & sIp

If ElPing(sIp) = 1 Then

Call AbrirMySQL(sIp)

If lMySqlAbierta Then
    TxtVer.Text = TxtVer.Text & Now() & " --> Conexión Correcta." & vbCrLf
    TxtVer.Refresh
    
Print #nF, Now() & " --> Conexión Correcta." & vbCrLf

    If lTodo Then
Print #nF, "BORRADO TABLAS TOUCH-SCALE..."
        lTodo = False
        Cm = "delete from pcscale.items"
        DbMysqli.Execute Cm
        Cm = "delete from pcscale.plukeys"
        DbMysqli.Execute Cm
        Cm = "delete from pcscale.qlevels"
        DbMysqli.Execute Cm
    End If
    
Print #nF, "Bajas en Touch-Scale en función contenido TOUCH-SCALE.MDB..."
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    ' Realizar bajas en Base de datos de Balanza
    'QLEVELS
    Cm = "select * from pcscale.qlevels"
    MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    If Not MyRsDes.EOF Then
        MyRsDes.MoveFirst
        Do While Not MyRsDes.EOF
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from qlevels where trim(code)=" & Chr(34) & Trim(MyRsDes.Fields("code")) & Chr(34))
            If MyRsORG.EOF Then
                Cm = "delete from pcscale.qlevels where code=" & Chr(34) & Trim(MyRsDes.Fields("code")) & Chr(34)
                DbMysqli.Execute Cm
            End If
            MyRsDes.MoveNext
        Loop
    End If
    MyRsDes.Close
    Set MyRsDes = Nothing
    'PLUKEYS
    'Set MyRsORG = MyBaseORG.OpenRecordset("select * from items order by key")
    Cm = "select * from pcscale.plukeys order by plukey"
    myRsKEYS.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    If Not myRsKEYS.EOF Then
        myRsKEYS.MoveFirst
        Do While Not myRsKEYS.EOF
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(myRsKEYS.Fields("plukey")))
            If MyRsORG.EOF Then
                Cm = "delete from pcscale.plukeys where plukey=" & CStr(myRsKEYS.Fields("plukey"))
                DbMysqli.Execute Cm
            End If
            myRsKEYS.MoveNext
        Loop
    End If
    myRsKEYS.Close
    Set myRsKEYS = Nothing
    'ITEMS y PLUKEYS por código
    Cm = "select * from pcscale.items order by code"
    MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    If Not MyRsDes.EOF Then
        MyRsDes.MoveFirst
        Do While Not MyRsDes.EOF
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where code=" & CStr(MyRsDes.Fields("code")))
            If MyRsORG.EOF Then
                Cm = "delete from pcscale.items where code=" & CStr(MyRsDes.Fields("code"))
                DbMysqli.Execute Cm
                Cm = "delete from pcscale.plukeys where code=" & CStr(MyRsDes.Fields("code"))
                DbMysqli.Execute Cm
            End If
            MyRsORG.Close
            Set MyRsORG = Nothing
            
            MyRsDes.MoveNext
        Loop
    
    End If
    MyRsDes.Close
    Set MyRsDes = Nothing
    
    
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from qlevels")
    'Cm = "select * from pcscale.qlevels"
    'MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    'If Not MyRsDes.EOF And Not MyRsORG.EOF Then
    '    'no funciona...
    '    'MyRsDes.MoveFirst
    '    'Do While Not MyRsDes.EOF
    '    '    MyRsDes.Delete
    '    '    MyRsDes.MoveNext
    '    'Loop
    '    '''''''''''''''
    '    'debe realizarse como sigue
    '    Cm = "delete from pcscale.qlevels"
    '    DbMysqli.Execute Cm
    '    '''''''''''''''''''''''''''
    'End If
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        ''en el caso de haber borrado previamente los datos de la tabla...
        'Do While Not MyRsORG.EOF
        '    MyRsDes.AddNew
        '    MyRsDes.Fields("code") = Trim(MyRsORG.Fields("code"))
        '
        '    MyRsDes.Fields("icon") = bmpTopng(Trim(MyRsORG.Fields("icon")))
        '
        '    MyRsDes.Fields("description") = Trim(MyRsORG.Fields("description"))
        '
        '    MyRsDes.Update
        '
        '    MyRsORG.MoveNext
        'Loop
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Do While Not MyRsORG.EOF
            Cm = "select * from pcscale.qlevels where code=" & Chr(34) & Trim(MyRsORG.Fields("code")) & Chr(34)
            MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
            If MyRsDes.EOF Then
                MyRsDes.AddNew
                MyRsDes!Code = Trim(MyRsORG.Fields("code"))
                MyRsDes!Icon = sToUTF8(bmpTopng(Trim(MyRsORG.Fields("icon"))))
                MyRsDes!Description = sToUTF8(Trim(MyRsORG.Fields("description")))
                MyRsDes.Update
            Else
                Cm = "update pcscale.qlevels set icon=" & "'" & sToUTF8(bmpTopng(Trim(MyRsORG.Fields("icon")))) & "'" & ", description=" & "'" & sToUTF8(Trim(MyRsORG.Fields("description"))) & "'" & " where code=" & Chr(34) & Trim(MyRsORG.Fields("code")) & Chr(34)
                DbMysqli.Execute Cm
                MyRsDes.Update
            End If
            MyRsORG.MoveNext
            MyRsDes.Close
        Loop
        'MyRsDes.Close
        Set MyRsDes = Nothing
    End If
    
Print #nF, "TRASPASO DE DATOS TOUCHSCALE.MDB a TOUCH-SCALE..."

    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items order by key")
    Cm = "select * from pcscale.items"
    MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    Cm = "select * from pcscale.plukeys"
    myRsKEYS.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
    
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            
            Cm = "select * from pcscale.items where code=" & CStr(Trim(MyRsORG.Fields("code")))
            myRsTMP.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
            
            If myRsTMP.EOF Then
            '''''''''''''''''''
                MyRsDes.AddNew
                'MyRsDes.Fields("code") = MyRsORG.Fields("code")
                'If IsNull(Trim(MyRsORG.Fields("name"))) Then
                '    MyRsDes.Fields("name") = ""
                'Else
                '    MyRsDes.Fields("name") = sToUTF8(UCase(Trim(MyRsORG.Fields("name"))))
                'End If
                'If IsNull(Trim(MyRsORG.Fields("shortname"))) Then
                '    MyRsDes.Fields("shortname") = ""
                'Else
                '    MyRsDes.Fields("shortname") = sToUTF8(UCase(Trim(MyRsORG.Fields("shortname"))))
                'End If
                'If IsNull(Trim(MyRsORG.Fields("text"))) Then
                '    MyRsDes.Fields("text") = ""
                'Else
                '    MyRsDes.Fields("text") = sToUTF8(Trim(MyRsORG.Fields("text")))
                'End If
                'MyRsDes.Fields("family") = MyRsORG.Fields("family-codi_sub")
                'MyRsDes.Fields("price") = MyRsORG.Fields("price")
                'MyRsDes.Fields("saleform") = MyRsORG.Fields("saleform")
                'MyRsDes.Fields("plu") = MyRsORG.Fields("plu")
                'MyRsDes.Fields("counter") = MyRsORG.Fields("section")
                'MyRsDes.Fields("vat") = MyRsORG.Fields("vat_value")
                'If IsNull(Trim(MyRsORG.Fields("ean13"))) Then
                '    MyRsDes.Fields("ean13") = ""
                'Else
                '    MyRsDes.Fields("ean13") = sToUTF8(Trim(MyRsORG.Fields("ean13")))
                'End If
                'MyRsDes.Fields("expiring") = MyRsORG.Fields("expiring")
                'MyRsDes.Fields("bestbefore") = MyRsORG.Fields("bestbefore")
                'MyRsDes.Fields("tare") = MyRsORG.Fields("tare")
                'If IsNull(Trim(MyRsORG.Fields("label_text"))) Then
                '    MyRsDes.Fields("label") = ""
                'Else
                '    MyRsDes.Fields("label") = "" 'stoutf8(Trim(MyRsORG.Fields("label_text")))
                'End If
                'If IsNull(Trim(MyRsORG.Fields("icon_name"))) Then
                '    MyRsDes.Fields("icon") = "pralcampo.bmp"
                'Else
                '    MyRsDes.Fields("icon") = sToUTF8(Trim(MyRsORG.Fields("icon_name")))
                'End If
                'MyRsDes.Fields("qlevel") = MyRsORG.Fields("qualitynr")
                'MyRsDes.Update
                MyRsDes!Code = MyRsORG.Fields("code")
                If IsNull(Trim(MyRsORG.Fields("name"))) Then
                    MyRsDes!Name = ""
                Else
                    MyRsDes!Name = sToUTF8(UCase(Trim(MyRsORG.Fields("name"))))
                End If
                If IsNull(Trim(MyRsORG.Fields("shortname"))) Then
                    MyRsDes!ShortName = ""
                Else
                    MyRsDes!ShortName = sToUTF8(UCase(Trim(MyRsORG.Fields("shortname"))))
                End If
                If IsNull(Trim(MyRsORG.Fields("text"))) Then
                    MyRsDes!Text = ""
                Else
                    MyRsDes!Text = sToUTF8(Trim(MyRsORG.Fields("text")))
                End If
                MyRsDes!family = MyRsORG.Fields("family-codi_sub")
                MyRsDes!price = MyRsORG.Fields("price")
                MyRsDes!saleform = MyRsORG.Fields("saleform")
                MyRsDes!PLU = MyRsORG.Fields("plu")
                MyRsDes!Counter = MyRsORG.Fields("section")
                MyRsDes!vat = MyRsORG.Fields("vat_value")
                If IsNull(Trim(MyRsORG.Fields("ean13"))) Then
                    MyRsDes!ean13 = ""
                Else
                    MyRsDes!ean13 = sToUTF8(Trim(MyRsORG.Fields("ean13")))
                End If
                MyRsDes!expiring = MyRsORG.Fields("expiring")
                MyRsDes!bestbefore = MyRsORG.Fields("bestbefore")
                MyRsDes!tare = MyRsORG.Fields("tare")
                If IsNull(Trim(MyRsORG.Fields("label_text"))) Then
                    MyRsDes!Label = ""
                Else
                    MyRsDes!Label = "" 'stoutf8(Trim(MyRsORG.Fields("label_text")))
                End If
                If IsNull(Trim(MyRsORG.Fields("icon_name"))) Then
                    MyRsDes!Icon = "pralcampo.bmp"
                Else
                    MyRsDes!Icon = sToUTF8(Trim(MyRsORG.Fields("icon_name")))
                End If
                MyRsDes!qLevel = MyRsORG.Fields("qualitynr")
                MyRsDes.Update
            Else
                'Cm = "update pcscale.items set name=" & "'" & sToUTF8(UCase(Trim(MyRsORG.Fields("name")))) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set shortname=" & "'" & sToUTF8(UCase(Trim(MyRsORG.Fields("shortname")))) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'If Not IsNull(Trim(MyRsORG.Fields("text"))) Then
                '    Cm = "update pcscale.items set text=" & "'" & sToUTF8(Trim(MyRsORG.Fields("text"))) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'End If
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set family=" & CStr(Trim(MyRsORG.Fields("family-codi_sub"))) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set price=" & comaTopunto(CStr(MyRsORG.Fields("price"))) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set saleform=" & CStr(MyRsORG.Fields("saleform")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set plu=" & CStr(MyRsORG.Fields("plu")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set counter=" & CStr(MyRsORG.Fields("section")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set vat=" & CStr(MyRsORG.Fields("vat_value")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'If Not IsNull(Trim(MyRsORG.Fields("ean13"))) Then
                '    Cm = "update pcscale.items set ean13=" & "'" & sToUTF8(Trim(MyRsORG.Fields("ean13"))) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'End If
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set expiring=" & CStr(MyRsORG.Fields("expiring")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set bestbefore=" & CStr(MyRsORG.Fields("bestbefore")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set tare=" & CStr(MyRsORG.Fields("tare")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set label=" & "'" & Trim(MyRsORG.Fields("label_text")) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                'If IsNull(Trim(MyRsORG.Fields("icon_name"))) Then
                '    Cm = "update pcscale.items set icon='pralcampo.bmp'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'Else
                '    Cm = "update pcscale.items set icon=" & "'" & sToUTF8(Trim(MyRsORG.Fields("icon_name"))) & "'" & " where code=" & CStr(MyRsORG.Fields("code"))
                'End If
                'DbMysqli.Execute Cm
                'Cm = "update pcscale.items set qlevel=" & CStr(MyRsORG.Fields("qualitynr")) & " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                
                
                Cm = "update pcscale.items set name=" & "'" & sToUTF8(UCase(Trim(MyRsORG.Fields("name")))) & "'" & " ," '& " where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "shortname=" & "'" & sToUTF8(UCase(Trim(MyRsORG.Fields("shortname")))) & "'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                If Not IsNull(Trim(MyRsORG.Fields("text"))) Then
                    Cm = Cm & "text=" & "'" & sToUTF8(Trim(MyRsORG.Fields("text"))) & "'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                End If
                'DbMysqli.Execute Cm
                Cm = Cm & "family=" & CStr(Trim(MyRsORG.Fields("family-codi_sub"))) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "price=" & comaTopunto(CStr(MyRsORG.Fields("price"))) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "saleform=" & CStr(MyRsORG.Fields("saleform")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "plu=" & CStr(MyRsORG.Fields("plu")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "counter=" & CStr(MyRsORG.Fields("section")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "vat=" & CStr(MyRsORG.Fields("vat_value")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                If Not IsNull(Trim(MyRsORG.Fields("ean13"))) Then
                    Cm = Cm & "ean13=" & "'" & sToUTF8(Trim(MyRsORG.Fields("ean13"))) & "'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                End If
                'DbMysqli.Execute Cm
                Cm = Cm & "expiring=" & CStr(MyRsORG.Fields("expiring")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "bestbefore=" & CStr(MyRsORG.Fields("bestbefore")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "tare=" & CStr(MyRsORG.Fields("tare")) & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                Cm = Cm & "label=" & "'" & Trim(MyRsORG.Fields("label_text")) & "'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                'DbMysqli.Execute Cm
                If IsNull(Trim(MyRsORG.Fields("icon_name"))) Then
                    Cm = Cm & "icon='pralcampo.bmp'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                Else
                    Cm = Cm & "icon=" & "'" & sToUTF8(Trim(MyRsORG.Fields("icon_name"))) & "'" & " ," '" where code=" & CStr(MyRsORG.Fields("code"))
                End If
                'DbMysqli.Execute Cm
                Cm = Cm & "qlevel=" & CStr(MyRsORG.Fields("qualitynr")) & " where code=" & CStr(MyRsORG.Fields("code"))
                DbMysqli.Execute Cm
            
                MyRsDes.Update
            
            End If
            myRsTMP.Close
            
            Cm = "select * from pcscale.plukeys where plukey=" & CStr(Trim(MyRsORG.Fields("key")))
            myRsTMP.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
            If myRsTMP.EOF Then
                myRsKEYS.AddNew
                myRsKEYS.Fields("section") = 1 'MyRsORG.Fields("section")
                myRsKEYS.Fields("board") = 0 'MyRsORG.Fields
                myRsKEYS.Fields("plukey") = MyRsORG.Fields("key")
                myRsKEYS.Fields("code") = MyRsORG.Fields("code")
                myRsKEYS.Update
            Else
                Cm = "update pcscale.plukeys set code=" & CStr(MyRsORG.Fields("code")) & " where plukey=" & CStr(Trim(MyRsORG.Fields("key")))
                DbMysqli.Execute Cm
                myRsKEYS.Update
            End If
            
            myRsTMP.Close
            
            MyRsORG.MoveNext
        Loop
        myRsKEYS.Close
        Cm = "select * from pcscale.barcodeean13"
        myRsKEYS.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
        If Not myRsKEYS.EOF Then
            myRsKEYS.MoveFirst
            Cm = "update pcscale.barcodeean13 set networkss=" & "'20CCCCCiiiii'" & ", localss=" & "'20CCCCCiiiii'"
            DbMysqli.Execute Cm
'NetworkPacket1 CHAR(12) NOT NULL, NetworkPacket2
            Cm = "update pcscale.barcodeean13 set NetworkPacket1=" & "'20CCCCCIIIII'" & ", NetworkPacket2=" & "'20CCCCCIIIII'"
            DbMysqli.Execute Cm
'LocalPacket1 CHAR(12) NOT NULL,  LocalPacket2 CHAR(12) NOT NULL
            Cm = "update pcscale.barcodeean13 set LocalPacket1=" & "'20CCCCCIIIII'" & ", LocalPacket2=" & "'20CCCCCIIIII'"
            DbMysqli.Execute Cm

        End If
    End If
    
    
    Set myRsTMP = Nothing
    MyRsDes.Close
    Set MyRsDes = Nothing
    myRsKEYS.Close
    Set myRsKEYS = Nothing
    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    
    If lMySqlAbierta Then
        DbMysqli.Close
        Set DbMysqli = Nothing
    End If
    'MsgBox "Finalizado..."
    TxtVer.Text = TxtVer.Text & Now() & " --> Datos enviados." & vbCrLf
    TxtVer.Refresh

Print #nF, Now() & " --> Datos enviados."

Else
    TxtVer.Text = TxtVer.Text & Now() & " --> No ha sido posible la Conexión." & vbCrLf
    TxtVer.Refresh
    'MsgBox "NO CONECTADA", vbCritical
Print #nF, Now() & " --> No ha sido posible la Conexión."

End If

Else
    TxtVer.Text = TxtVer.Text & Now() & " --> No ha sido posible la Conexión.(sin ping)" & vbCrLf
    TxtVer.Refresh
Print #nF, Now() & " --> No ha sido posible la Conexión.(sin ping)"

End If

MyRsBAL.MoveNext
Loop

End If

MyRsBAL.Close
Set MyRsBAL = Nothing

MyBaseBAL.Close
Set MyBaseBAL = Nothing

    TmrEvents.Enabled = False
    CmdEnvTodo.Enabled = True
    CmdEnvCambios.Enabled = True
    CmdExit.Enabled = True
    Me.BorderStyle = 1
    lComunica = False
    Close #nF
    
End Sub


Private Sub CmdExit_Click()
    Call Form_Unload(0)
End Sub



'DEJAR ESTOS COMENTARIOS TAL CUÁL ESTAN...
'Inicio Trabajos TouchScale --> 30/03/2011, fin 31/03/2011 se reinicia 09/04/2011
'FRUTAS 1-120. ("                                FRUTAS             1 - 120"
'fruta dispone de 6 familias
'VERDURAS 121-200 ("                             VERDURAS         121 - 200"
'Verduras dispone de 4 familias
'EXÓTICOS Y FRUTOS SECOS 201-240 (" EXÓTICOS   Y  FRUTOS           SECOS                  201 - 240"
'Exóticos dispone de 2 familias
''''''''''''''''''''''''''
'Familias de FRUTAS
''''''''''''''''''''''''''
'Familia CÍTRICOS 1-20 ("   Cítricos teclas             1 - 20"
'Familia MANZANAS 21-40 (" Manzanas                teclas 21 - 40"
'Familia PERAS 41-60 (" Peras teclas                41 - 60"
'Familia Frutas de Hueso 61-80 (" Frutas de Hueso   teclas 61 - 80"
'Familia Melón/Sandia y Plátanos 81-100 (" Melón/Sandia y    Plátanos 81-100"
'Familia Otras Frutas 101-120 (" Otras Frutas           teclas 101-120"
''''''''''''''''''''''''''
'Familias de Verduras
''''''''''''''''''''''''''
'Familia de Lechugas 121-140 ("  Lechugas               teclas 121-140"
'Familia de Tomates 141-160 (" Tomates                 teclas 141-160"
'Familia de Hortalizas 161-180 ("      Hortalizas          teclas 141-160"
'Familia de Verduras de Cocer 181-200 (" Verduras Cocer      teclas 181-200"
'''''''''''''''''''''''''''
'Familias de EXÓTICOS
'''''''''''''''''''''''''''
'Familia de Exóticos 201-220 ("  Exóticos teclas     teclas 201-220"
'Familia de Frutos Secos. 221-240 ("  Frutos Secos         teclas 221-240"
'''''''''''''''''''''''''''
'Carga de las condiciones iniciales tal cúal se han recibido de Alcampo
'Base de datos para TouchScale:(versión a Viernes 01/04/2011
'-- CREATE DATABASE IF NOT EXISTS pcscale;
'-- GRANT ALL PRIVILEGES ON pcscale.* TO 'pcscale'@'%' IDENTIFIED BY 'epelsa';
'-- USE pcscale;
'CREATE TABLE IF NOT EXISTS Items
'( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Code NUMERIC UNIQUE, Name TEXT NOT NULL, ShortName VARCHAR(32) NOT NULL,
'  Text TEXT, Family NUMERIC, Price DOUBLE DEFAULT 0, SaleForm NUMERIC DEFAULT 1, PLU NUMERIC DEFAULT 0,
'  Counter NUMERIC DEFAULT 0, VAT NUMERIC DEFAULT 0, EAN13 TEXT, Expiring NUMERIC DEFAULT 0, BestBefore NUMERIC DEFAULT 0
'  , Tare DOUBLE DEFAULT 0, Label TEXT, Icon TEXT, UNIQUE INDEX(Code), INDEX(Name(19)));
'CREATE TABLE IF NOT EXISTS Vendors       ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Code NUMERIC UNIQUE, Name TEXT NOT NULL);
'CREATE TABLE IF NOT EXISTS Counters
'( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Code NUMERIC UNIQUE, Name TEXT NOT NULL, Description TEXT NOT NULL);
'CREATE TABLE IF NOT EXISTS Scales        ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, MAC VARCHAR(17) UNIQUE NOT NULL, IP VARCHAR(16) NOT NULL, Number NUMERIC, Name TEXT NOT NULL, Type CHAR(1));
'CREATE TABLE IF NOT EXISTS HTickets      ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, tckType NUMERIC, Status CHAR(1), Number NUMERIC DEFAULT 0, Counter NUMERIC DEFAULT 0, Scale NUMERIC DEFAULT 0, Vendor NUMERIC, DatTim DATETIME, ReloadFrom NUMERIC DEFAULT 0, Customer NUMERIC DEFAULT 0, tckDiscount DOUBLE DEFAULT 0, UNIQUE INDEX (Vendor,Id) );
'CREATE TABLE IF NOT EXISTS LTickets      ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, IdHTicket NUMERIC, lineType NUMERIC, SaleForm NUMERIC, Item NUMERIC, Weight DOUBLE, lineDiscount DOUBLE, Price DOUBLE);
'CREATE TABLE IF NOT EXISTS PayTickets    ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, IdHTicket NUMERIC, PayMethod NUMERIC, Delivered DOUBLE DEFAULT 0 );
'CREATE TABLE IF NOT EXISTS PayMethods    ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(30) UNIQUE NOT NULL );
'CREATE TABLE IF NOT EXISTS VendorKeys    ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Counter NUMERIC, VendorKey NUMERIC, Vendor NUMERIC, UNIQUE INDEX(Counter, VendorKey));
'CREATE TABLE IF NOT EXISTS PLUKeys
'( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Section NUMERIC, Board NUMERIC, PLUKey NUMERIC, Code NUMERIC,
'  UNIQUE INDEX(Section, Board, PLUKey));
'CREATE TABLE IF NOT EXISTS Bovine        ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, EarTag TEXT NOT NULL, BirthLand NUMERIC, BirthDate DATE, BreedingLand NUMERIC, SlaughterLand NUMERIC, SlaughterReg TEXT, SlaughterDate DATE, ButcherLand NUMERIC, ButcherReg TEXT, Lot TEXT, UNIQUE INDEX(Eartag(10)));
'CREATE TABLE IF NOT EXISTS BarcodeEAN13
'( Id INTEGER PRIMARY KEY AUTO_INCREMENT, NumberType CHAR(1), Number NUMERIC DEFAULT 0, NetworkSales CHAR(12) NOT NULL,
'  NetworkMixed CHAR(12) NOT NULL, NetworkSuper CHAR(12) NOT NULL, NetworkSS CHAR(12) NOT NULL,
'  NetworkPacket1 CHAR(12) NOT NULL, NetworkPacket2 CHAR(12) NOT NULL, LocalSales CHAR(12) NOT NULL,
'  LocalMixed CHAR(12) NOT NULL, LocalSuper CHAR(12) NOT NULL, LocalSS CHAR(12) NOT NULL, LocalPacket1 CHAR(12) NOT NULL,
'  LocalPacket2 CHAR(12) NOT NULL, LocalTotal CHAR(12) NOT NULL, UNIQUE INDEX(NumberType, Number)  );
'CREATE TABLE IF NOT EXISTS Status        ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Version NUMERIC DEFAULT 0, TicketNumberType NUMERIC DEFAULT 0, ResetTicketNumber BOOLEAN DEFAULT FALSE);
'CREATE TABLE IF NOT EXISTS NumTickets    ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Identifier NUMERIC UNIQUE, Number NUMERIC);
'CREATE TABLE IF NOT EXISTS EraseTickets  ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, DatTim DATETIME);
'CREATE TABLE IF NOT EXISTS Trademark     ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, CO_Name TEXT NOT NULL, Description TEXT NOT NULL, Business TEXT NOT NULL, Slogan TEXT NOT NULL, Headquarter TEXT NOT NULL, CO_WEB TEXT NOT NULL, Shop_Name TEXT NOT NULL, Address TEXT NOT NULL, City TEXT NOT NULL, Phone TEXT NOT NULL, TIN TEXT NOT NULL, Shop_WEB TEXT NOT NULL, Text1 TEXT NOT NULL, Text2 TEXT NOT NULL, Text3 TEXT NOT NULL, Text4 TEXT NOT NULL, Text5 TEXT NOT NULL, Text6 TEXT NOT NULL);
'
'-- Local tables
'CREATE TABLE IF NOT EXISTS LocalStatus   ( Id INTEGER PRIMARY KEY AUTO_INCREMENT, Version NUMERIC DEFAULT 0, ItemsUpdated BOOLEAN DEFAULT FALSE, EraseTickets BOOLEAN DEFAULT FALSE, RemoveScale BOOLEAN DEFAULT FALSE );
'CREATE TABLE IF NOT EXISTS LocalHTickets    LIKE HTickets;
'CREATE TABLE IF NOT EXISTS LocalLTickets    LIKE LTickets;
'CREATE TABLE IF NOT EXISTS LocalPayTickets  LIKE PayTickets;
'CREATE TABLE IF NOT EXISTS LocalNumTickets  LIKE NumTickets;'
'
'-- Initialization: Barcode EAN13 table
'REPLACE INTO BarcodeEAN13   VALUES(1,'G',0, '2YTTTTTIIIII', '20VVViiiiiii', '10VVVIIIIIII', '2YTTTTTIIIII', '2CCCCCCIIIII', '2CCCCCCIIIII', '2YTTTTTIIIII', '20VVViiiiiii', '10VVVIIIIIII', '2YTTTTTIIIII', '2CCCCCCIIIII', '2CCCCCCIIIII',  '2CCCCCCIIIII');
'
'-- Initialization: Status tables
'REPLACE INTO Status         VALUES(1,0,0,FALSE);
'REPLACE INTO LocalStatus    VALUES(1,0,FALSE,FALSE,FALSE);
'REPLACE INTO EraseTickets   VALUES(1,NOW());
'REPLACE INTO Trademark      VALUES(1,'','','','','','','','','','','','','','','','','','');

' Debe existir un fichero TOUCHPATH.CFG conteniendo PATH's TOUCHSCALE
'                         contenido: (Los paths con un máximo de 50 caracteres)
'                         PATH de las imágenes de los Items
'                         PATH de las imágenes de Publicidad
'                         PATH de las imágenes de Calidad
'.........................En principio no hay más path's en la base de datos...
' Debe existir un fichero TOUCHSS.CFG conteniendo las características del autoservicio
'                         contenido:
'                         Mostrador gestión autoservicio
'                         Sección máquina AutoServicio
'                         Número máximo teclas autoServicio (en principio máximo 240 teclas)(se recalculan en función de los datos siguientes)
'                         Número máximo Grupos              (en principio máximo 3 grupos)
'                         Número máximo de SubGrupos por Grupo (en principio 6 máximo 6 subgrupos por grupo)
'                         Número máximo de teclas por subgrupo (en principio 20 teclas máximo por subgrupo)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debe existir un fichero para configuración Grupos...
'FRUTAS 1-120. ("                                FRUTAS             1 - 120"
'fruta dispone de 6 familias
'VERDURAS 121-200 ("                             VERDURAS         121 - 200"
'Verduras dispone de 4 familias
'EXÓTICOS Y FRUTOS SECOS 201-240 (" EXÓTICOS   Y  FRUTOS           SECOS                  201 - 240"
'Exóticos dispone de 2 familias
'           Fichero TOUCHGRP.CFG
'           Contenido: 1 registro con cada grupo... No deben existir más de los indicados en TOUCHSS.CFG
'           Grupo
'           Descripción (Máximo 70 caracteres)
'           Número de Subgrupos
'           Tecla Inicial
'           Tecla Final
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debe existir un fichero para configuración Sub-Grupos...
''''''''''''''''''''''''''
'Subgrupos del Grupo 1 FRUTAS
''''''''''''''''''''''''''
'Familia CÍTRICOS 1-20 ("   Cítricos teclas             1 - 20"
'Familia MANZANAS 21-40 (" Manzanas                teclas 21 - 40"
'Familia PERAS 41-60 (" Peras teclas                41 - 60"
'Familia Frutas de Hueso 61-80 (" Frutas de Hueso   teclas 61 - 80"
'Familia Melón/Sandia y Plátanos 81-100 (" Melón/Sandia y    Plátanos 81-100"
'Familia Otras Frutas 101-120 (" Otras Frutas           teclas 101-120"
''''''''''''''''''''''''''
'Familias de Verduras
''''''''''''''''''''''''''
'Familia de Lechugas 121-140 ("  Lechugas               teclas 121-140"
'Familia de Tomates 141-160 (" Tomates                 teclas 141-160"
'Familia de Hortalizas 161-180 ("      Hortalizas          teclas 141-160"
'Familia de Verduras de Cocer 181-200 (" Verduras Cocer      teclas 181-200"
'''''''''''''''''''''''''''
'Familias de EXÓTICOS
'''''''''''''''''''''''''''
'Familia de Exóticos 201-220 ("  Exóticos teclas     teclas 201-220"
'Familia de Frutos Secos. 221-240 ("  Frutos Secos         teclas 221-240"
'''''''''''''''''''''''''''

            
Private Sub Form_Load()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim n As Integer

    'Call ImgGrp_Click(0)
    'Call ImgGrp_Click(0)
    LblGrp(0).ForeColor = vbBlack
    LblFam(0).ForeColor = vbBlack
    
   '''''''''''''''''''''''''''''''''''
    'versión 2
    If Lista.ListItems.Count = 0 And LCase(Trim(Command)) = "hydra" Then
    
        FrmKeyTouch.ListaImg.ListItems.Clear
        Set FrmKeyTouch.ListaImg.Icons = Nothing
        FrmKeyTouch.Imagenes.ListImages.Clear
        FrmKeyTouch.Lista.ListItems.Clear
        
        Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
        Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where borrado=false and secc_maqui=" & CStr(8) & " order by plu")
        With MyRsORG
            
            If Not .EOF Then
                Do Until .EOF
                    
                    FrmKeyTouch.Lista.ListItems.Add , , Format(!codigo, "000000")
                    FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(1) = Format(!PLU, "0000")
                    
                    If Not IsNull(!des_plu1) Then
                        FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(2) = !des_plu1
                    Else
                        FrmKeyTouch.Lista.ListItems(FrmKeyTouch.Lista.ListItems.Count).SubItems(2) = Space(25)
                    End If
                    .MoveNext
                Loop
            End If
        End With
    
        MyRsORG.Close
        Set MyRsORG = Nothing
        MyBaseORG.Close
        Set MyBaseORG = Nothing
    
        'versión 3
        FrmKeyTouch.imgDir.Path = App.Path & "\items"
        FrmKeyTouch.imgDir.Pattern = "*.bmp"
        FrmKeyTouch.imgDir.Refresh
        
        For n = 1 To FrmKeyTouch.imgDir.ListCount
            FrmKeyTouch.Imagenes.ListImages.Add n, , LoadPicture(App.Path & "\items\" & imgDir.List(n - 1))
        Next n
        FrmKeyTouch.ListaImg.Icons = Imagenes
        For n = 1 To FrmKeyTouch.imgDir.ListCount
            'ListaImg.ListItems.Add n, , App.Path & "\items\" & imgDir.List(n - 1), n
            FrmKeyTouch.ListaImg.ListItems.Add n, , FrmKeyTouch.imgDir.List(n - 1), n
        Next n
        ''''''''''
    
    End If

    
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

    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    'Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key>=" & CStr(TTecGrupos.Indice(Index).SubGrupo(n).IniKey) & _
    '    " key<=" & CStr(TTecGrupos.Indice(Indice).SubGrupo(n).EndKey) & " order by key")
    For nn = 0 To 19
        ImgTec(nn).Visible = False
        LblTec1(nn).Visible = False
        LblTec2(nn).Visible = False
        ImgQ(nn).Visible = False
    Next nn
    
    For nn = TTecGrupos.Indice(Indice).SubGrupo(n).IniKey To TTecGrupos.Indice(Indice).SubGrupo(n).EndKey
        ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec2(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
        LblTec2(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = CStr(nn)
        Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(nn))
        If Not MyRsORG.EOF Then
            If Trim(MyRsORG.Fields("icon_name")) <> "" Then
                ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPicture(App.Path & "\items\" & Trim(MyRsORG.Fields("icon_name")))
            Else
                ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPicture(App.Path & "\items\" & "pralcampo.bmp")
            End If
            LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = Trim(MyRsORG.Fields("shortname"))
            If MyRsORG.Fields("qualitynr") <> 0 Then
                ImgQ(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPicture(App.Path & "\quality\" & UCase(Trim(MyRsORG.Fields("qualityname"))))
                ImgQ(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Visible = True
            End If
            
        Else
            ImgTec(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Picture = LoadPicture(App.Path & "\items\pralcampo.bmp")
            LblTec1(nn - TTecGrupos.Indice(Indice).SubGrupo(n).IniKey).Caption = ""
        End If
    Next nn
    
    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    
End Sub

'TTecGrupos.Indice(nnF).SubGrupo(nnnF).lblTxt = MyRsORG.Fields("text")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey = MyRsORG.Fields("inikey")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey = MyRsORG.Fields("endkey")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgNoSelect = MyRsORG.Fields("imgnoselect")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgSelect = MyRsORG.Fields("imgselect")
Public Sub cargaSubGrupos(ByVal Index As Integer)
Dim n As Integer
    For n = 0 To 5
        ImgFam(n).Visible = False
        LblFam(n).Visible = False
    Next n
    For n = 0 To TTecGrupos.Indice(Index).SubGroupnr - 1
        ImgFam(n).Visible = True
        LblFam(n).Visible = True
        If n = 0 Then
            ImgFam(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(Index).SubGrupo(n).ImgSelect))
            LblFam(n).ForeColor = vbBlack
            Call cargaSubGrupoTeclas(Index, n)
        Else
            ImgFam(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(Index).SubGrupo(n).ImgNoSelect))
            LblFam(n).ForeColor = vbWhite
        End If
        LblFam(n).Caption = TTecGrupos.Indice(Index).SubGrupo(n).lblTxt
    Next n
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmSelect
    End
End Sub

Private Sub ImgTec_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Variant
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
    TeclaTmp.Key = Val(LblTec2(Index).Caption)
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(TeclaTmp.Key))
    If Not MyRsORG.EOF Then 'edición
        TeclaTmp.Code = MyRsORG.Fields("code")
        'TeclaTmp.Counter = MyRsORG.Fields("counter")
        If IsNull(MyRsORG.Fields("icon_name")) Then
            TeclaTmp.icon_name = "pralcampo.bmp"
        Else
            TeclaTmp.icon_name = MyRsORG.Fields("icon_name")
        End If
        TeclaTmp.Name = MyRsORG.Fields("name")
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
    Set MyBaseORG = Nothing

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
    nGrupoActual = Index
    For n = 0 To nGrupos - 1
        If Index = n Then
            ImgGrp(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(n).ImgSelect))
            LblGrp(n).ForeColor = vbBlack
            'FrmSelect.ImgGrp(0).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(n).ImgSelect))
            Call cargaSubGrupos(Index)
        Else
            ImgGrp(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(n).ImgNoSelect))
            LblGrp(n).ForeColor = vbWhite
        End If
    Next n
    'ojo
    If FrmSelect.Visible = True Then
        Call ImgFam_Click(0)
    End If
    ''''
    FrmKeyTouch.Refresh
End Sub
Private Sub LblFam_Click(Index As Integer)
    Call ImgFam_Click(Index)
End Sub
Private Sub ImgFam_Click(Index As Integer)
Dim n As Integer
    nSubGrupoActual = Index
    For n = 0 To TTecGrupos.Indice(nGrupoActual).SubGroupnr - 1
        If Index = n Then
            ImgFam(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(nGrupoActual).SubGrupo(n).ImgSelect))
            LblFam(n).ForeColor = vbBlack

            'FrmSelect.ImgFam(0).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(nGrupoActual).SubGrupo(n).ImgSelect))
            Call cargaSubGrupoTeclas(nGrupoActual, nSubGrupoActual)
            'ojo
            If FrmSelect.Visible = True Then
                'Unload FrmSelect
                'FrmKeyTouch.Lista.Visible = True
                'FrmKeyTouch.ListaImg.Visible = True
                'FrmKeyTouch.ListaCalidad.Visible = True
                Lista.Sorted = False
                Call imgTecClick(1000) 'prueba
                'If Lista.SortOrder = lvwAscending Then
                '    Lista.SortOrder = lvwDescending
                'Else
                '    Lista.SortOrder = lvwAscending
                'End If
                'Lista.SortKey = 1
                Lista.Sorted = False
                FrmSelect.Caption = "TECLAS:" & Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey, "000") & " Hasta " & _
                         Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).EndKey, "000")
                'FrmKeyTouch.Lista.Refresh
                'Load FrmSelect
                'FrmSelect.Show
            End If
            ''''
        Else
            ImgFam(n).Picture = LoadPicture(App.Path & "\imggrpfam\" & Trim(TTecGrupos.Indice(nGrupoActual).SubGrupo(n).ImgNoSelect))
            LblFam(n).ForeColor = vbWhite
        End If
    Next n
    FrmKeyTouch.Refresh
End Sub
Private Sub LblTec1_Click(Index As Integer)
    Call ImgTec_Click(Index)
End Sub
Private Sub LblTec2_Click(Index As Integer)
    Call ImgTec_Click(Index)
End Sub
Private Sub imgq_Click(Index As Integer)
    Call ImgTec_Click(Index)
End Sub

Private Sub ImgTec_Click(Index As Integer)

    Call imgTecClick(Index)
    
End Sub

Public Sub imgTecClick(ByVal Index As Integer, Optional ByVal nTec As Integer)
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim n As Integer
Dim nn As Integer
Dim nF As Integer

nF = FreeFile()
Open App.Path & "\logcargapantalla.txt" For Output As #nF
Print #nF, Now() & "-->Entrada..."
    If Index < 1000 Then
    '    Lista.Visible = True
    '    ListaImg.Visible = True
    '    ListaCalidad.Visible = True
        
        'versión 2
        'Lista.ListItems.Clear
        ''''''''''
        
        'ListaImg.ListItems.Clear
        ListaCalidad.ListItems.Clear
        'Set ListaImg.Icons = Nothing
        Set ListaCalidad.Icons = Nothing
        'Imagenes.ListImages.Clear
        ImagenesCalidad.ListImages.Clear
Print #nF, Now() & "-->Borradas listas..."
    End If

Print #nF, Now() & "-->Carga Datos..."

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
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where key=" & CStr(TeclaTmp.Key))
    If Not MyRsORG.EOF Then 'edición
        TeclaTmp.Code = MyRsORG.Fields("code")
        TeclaTmp.Counter = MyRsORG.Fields("counter")
        If IsNull(MyRsORG.Fields("icon_name")) Then
            TeclaTmp.icon_name = "pralcampo.bmp"
        Else
            TeclaTmp.icon_name = MyRsORG.Fields("icon_name")
        End If
        TeclaTmp.Name = MyRsORG.Fields("name")
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
        TeclaTmp.icon_name = "pralcampo.bmp"
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
        'versión 2
        'Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
        'Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where borrado=false and secc_maqui=" & CStr(nSeccMaqui) & " order by plu")
        'With MyRsORG
        '    If Not .EOF Then
        '        Do Until .EOF
        '
        '            Lista.ListItems.Add , , Format(!codigo, "000000")
        '            Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Format(!PLU, "0000")
        '
        '            If Not IsNull(!des_plu1) Then
        '                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = !des_plu1
        '            Else
        '                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = Space(25)
        '            End If
        '            .MoveNext
        '        Loop
        '    End If
        'End With
    
        'MyRsORG.Close
        'Set MyRsORG = Nothing
        'MyBaseORG.Close
        'Set MyBaseORG = Nothing
        '''''''''''''
        
'Print #nF, Now() & "-->Cargada lista artículos..."
        'versión 3
        'imgDir.Path = App.Path & "\items"
        'imgDir.Pattern = "*.bmp"
        'imgDir.Refresh
       '
       ' For n = 1 To imgDir.ListCount
       '     Imagenes.ListImages.Add n, , LoadPicture(App.Path & "\items\" & imgDir.List(n - 1))
       ' Next n
       ' ListaImg.Icons = Imagenes
       ' For n = 1 To imgDir.ListCount
       '     'ListaImg.ListItems.Add n, , App.Path & "\items\" & imgDir.List(n - 1), n
       '     ListaImg.ListItems.Add n, , imgDir.List(n - 1), n
       ' Next n

'Print #nF, Now() & "-->Cargada lista imágenes artículos..."
        '''''''''''''''''''''''''''
        
        imgDir.Path = App.Path & "\quality"
        imgDir.Pattern = "*.bmp"
        imgDir.Refresh
        'nn = 1
        For n = 1 To imgDir.ListCount
            ImagenesCalidad.ListImages.Add n, , LoadPicture(App.Path & "\quality\" & imgDir.List(n - 1))
            'nn = nn + 1
        Next n
        ListaCalidad.Icons = ImagenesCalidad
        For n = 1 To imgDir.ListCount
            ListaCalidad.ListItems.Add n, , App.Path & "\quality\" & imgDir.List(n - 1), n
        Next n
        
Print #nF, Now() & "-->Cargada lista imágenes calidad..."

        For n = 0 To 19
            LblTec1(n).Enabled = False
            LblTec2(n).Enabled = False
            ImgTec(n).Enabled = False
            ImgQ(n).Enabled = False
        Next n
    
    End If
    
    FrmSelect.LblTec1(0).Caption = TeclaTmp.ShortName
    FrmSelect.LblTec2(0).Caption = CStr(TeclaTmp.Key)
    If TeclaTmp.icon_name = "" Then
        TeclaTmp.icon_name = "pralcampo.bmp"
    End If
    FrmSelect.ImgTec(0).Picture = LoadPicture(Trim(App.Path & "\items\" & TeclaTmp.icon_name))
    If TeclaTmp.qLevel <> 0 Then
        FrmSelect.ImgQ(0).Picture = LoadPicture(App.Path & "\quality\" & Trim(TeclaTmp.qLevelName))
        FrmSelect.ImgQ(0).Visible = True
    Else
        FrmSelect.ImgQ(0).Visible = False
    End If
    Load FrmSelect
    FrmSelect.Show
    
    If Index < 1000 Then
        Lista.Visible = True
        ListaImg.Visible = True
        ListaCalidad.Visible = True
    End If
    
Print #nF, Now() & "-->Salida."
Close #nF

End Sub


Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
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
    nPlu = Item.SubItems(1)
    sDes = Item.SubItems(2)
    TeclaTmp.Code = nCod
    TeclaTmp.Name = sDes
    TeclaTmp.PLU = nPlu
    TeclaTmp.ShortName = sDes
    FrmSelect.LblTec1(0).Caption = Trim(sDes)
End Sub

Private Sub ListaImg_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sIcon As String
    TeclaTmp.icon_name = Item
    FrmSelect.ImgTec(0).Picture = LoadPicture(App.Path & "\items\" & TeclaTmp.icon_name)
    If FrmSelect.ImgQ(0).Picture <> 0 Then
        FrmSelect.ImgQ(0).Visible = True
    End If
End Sub
Private Sub Listacalidad_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sIcon As String
    TeclaTmp.qLevelName = Item
    'FrmSelect.ImgTec(0).Picture = LoadPicture(TeclaTmp.icon_name)
    FrmSelect.ImgQ(0).Picture = LoadPicture(TeclaTmp.qLevelName)
    If FrmSelect.ImgQ(0).Picture <> 0 Then
        FrmSelect.ImgQ(0).Visible = True
    End If
    If FrmSelect.Visible = True Then
        FrmSelect.CmdEliQ.Enabled = True
    End If
End Sub

Private Sub TmrEvents_Timer()
    Do_Events
End Sub

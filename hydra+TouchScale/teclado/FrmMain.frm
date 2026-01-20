VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   1065
   ClientTop       =   2985
   ClientWidth     =   14220
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   Begin ChamaleonButton.ChameleonBtn CmdExpand 
      Height          =   2895
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   5106
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   0
      Left            =   11760
      TabIndex        =   0
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":001C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   1
      Left            =   10920
      TabIndex        =   1
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0038
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   2
      Left            =   11760
      TabIndex        =   2
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0054
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   3
      Left            =   12600
      TabIndex        =   3
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0070
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   4
      Left            =   10920
      TabIndex        =   4
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":008C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   5
      Left            =   11760
      TabIndex        =   5
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":00A8
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   6
      Left            =   12600
      TabIndex        =   6
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":00C4
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   7
      Left            =   10920
      TabIndex        =   7
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":00E0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   8
      Left            =   11760
      TabIndex        =   8
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":00FC
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   9
      Left            =   12600
      TabIndex        =   9
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0118
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   10
      Left            =   12600
      TabIndex        =   10
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0134
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   11
      Left            =   10920
      TabIndex        =   11
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0150
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   12
      Left            =   13440
      TabIndex        =   12
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":016C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   13
      Left            =   13440
      TabIndex        =   13
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0188
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   14
      Left            =   13440
      TabIndex        =   14
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":01A4
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   15
      Left            =   13440
      TabIndex        =   15
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":01C0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   16
      Left            =   840
      TabIndex        =   16
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":01DC
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   17
      Left            =   1680
      TabIndex        =   17
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":01F8
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   18
      Left            =   2520
      TabIndex        =   18
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0214
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   19
      Left            =   3360
      TabIndex        =   19
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0230
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   20
      Left            =   4200
      TabIndex        =   20
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":024C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   21
      Left            =   5040
      TabIndex        =   21
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0268
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   22
      Left            =   5880
      TabIndex        =   22
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0284
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   23
      Left            =   6720
      TabIndex        =   23
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":02A0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   24
      Left            =   7560
      TabIndex        =   24
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":02BC
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   25
      Left            =   8400
      TabIndex        =   25
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":02D8
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   26
      Left            =   9240
      TabIndex        =   26
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":02F4
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   27
      Left            =   10080
      TabIndex        =   27
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0310
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   28
      Left            =   840
      TabIndex        =   28
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":032C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   29
      Left            =   1680
      TabIndex        =   29
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0348
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   30
      Left            =   2520
      TabIndex        =   30
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0364
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   31
      Left            =   3360
      TabIndex        =   31
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0380
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   32
      Left            =   4200
      TabIndex        =   32
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":039C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   33
      Left            =   5040
      TabIndex        =   33
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":03B8
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   34
      Left            =   5880
      TabIndex        =   34
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":03D4
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   35
      Left            =   6720
      TabIndex        =   35
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":03F0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   36
      Left            =   7560
      TabIndex        =   36
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":040C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   37
      Left            =   8400
      TabIndex        =   37
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0428
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   38
      Left            =   9240
      TabIndex        =   38
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0444
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   39
      Left            =   10080
      TabIndex        =   39
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0460
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   40
      Left            =   840
      TabIndex        =   40
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":047C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   41
      Left            =   1680
      TabIndex        =   41
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0498
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   42
      Left            =   2520
      TabIndex        =   42
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":04B4
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   43
      Left            =   3360
      TabIndex        =   43
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":04D0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   44
      Left            =   4200
      TabIndex        =   44
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":04EC
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   45
      Left            =   5040
      TabIndex        =   45
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0508
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   46
      Left            =   5880
      TabIndex        =   46
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0524
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   47
      Left            =   6720
      TabIndex        =   47
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0540
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   48
      Left            =   7560
      TabIndex        =   48
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":055C
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   49
      Left            =   8400
      TabIndex        =   49
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0578
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   50
      Left            =   9240
      TabIndex        =   50
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0594
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   51
      Left            =   10080
      TabIndex        =   51
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":05B0
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   52
      Left            =   840
      TabIndex        =   52
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":05CC
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   53
      Left            =   2520
      TabIndex        =   53
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":05E8
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   54
      Left            =   4200
      TabIndex        =   54
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0604
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Btn 
      CausesValidation=   0   'False
      Height          =   735
      Index           =   55
      Left            =   9240
      TabIndex        =   55
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMain.frx":0620
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nX As Long
Dim nY As Long

Private Sub Btn_Click(Index As Integer)
Dim nnX As Double
Dim nnY As Double
    nnX = Btn(Index).top
    nnY = Btn(Index).left
    Select Case Index
    Case 52, 53, 55
        cargaImg "_base2.png", Index
    Case 54
        cargaImg "_base6.png", Index
    Case Else
        cargaImg "_base.png", Index 'Set Btn(Index).PictureNormal = LoadPicture(App.Path & "\_base.jpg")  ', , , Btn(Index).Top + 30, Btn(Index).Left + 30)
    End Select
    Do_Events
    Do_Events
    Btn(Index).top = nnX + 30
    Btn(Index).left = nnY + 30
    Do_Events
    Select Case Index
    Case 52, 53, 55
        cargaImg "base2.png", Index
    Case 54
        cargaImg "base6.png", Index
    Case Else
        cargaImg "base.png", Index 'Set Btn(Index).PictureNormal = LoadPicture(App.Path & "\base.jpg") ', , , nX, nY)
    End Select
    Do_Events
    Do_Events
    Btn(Index).top = nnX
    Btn(Index).left = nnY
End Sub

Private Sub Btn_MouseOut(Index As Integer)
    Select Case Index
    Case 52, 53, 55
        cargaImg "base2.png", Index
    Case 54
        cargaImg "base6.png", Index
    Case Else
        cargaImg "base.png", Index
    End Select
End Sub

Private Sub Btn_MouseOver(Index As Integer)
    Select Case Index
    Case 52, 53, 55
        cargaImg "base_2.png", Index
    Case 54
        cargaImg "base_6.png", Index
    Case Else
        cargaImg "base_.png", Index
    End Select
End Sub

Private Sub CmdExpand_Click()
Dim i As Integer
Dim Token As Long

CmdExpand.Visible = False
If CmdExpand.Caption = ">" Then
    
    ActiveTransparency Me, False, False, 0

    For i = 255 To 0 Step -3
        ActiveTransparency Me, True, False, i
        Me.Refresh
    Next i
    
    
    Me.top = nY
    Me.left = Screen.Width - CmdExpand.Width
    Me.Width = CmdExpand.Width
    CmdExpand.Caption = "<"
    
    Token = InitGDIPlus
    Set FrmMain.CmdExpand.PictureNormal = LoadPictureGDIPlus(App.Path & "\" & "keyboard.png")
    FreeGDIPlus Token
    
    ActiveTransparency Me, False, False, 0

Else
    

    Me.top = nY
    Me.left = nX
    Me.Width = Screen.Width - Me.left
    CmdExpand.Caption = ">"
    
    ActiveTransparency Me, True, False, 0
    For i = 0 To 255 Step 3
        ActiveTransparency Me, True, False, i
        Me.Refresh
    Next i

    ActiveTransparency Me, False, False, 0

End If
CmdExpand.Visible = True
CmdExpand.Refresh

End Sub

Private Sub Form_Load()
'Dim nX As Long
'Dim nY As Long
'Dim Token As Long
'Dim sImg As String
'    sImg = "base.png"
'    If InStr(1, LCase(sImg), ".png") <> 0 Then
'        Token = InitGDIPlus
'        Set Btn(0).PictureNormal = LoadPictureGDIPlus(App.Path & "\" & sImg)
'        FreeGDIPlus Token
'    Else
'        Set Btn(0).PictureNormal = LoadPicture(App.Path & "\" & sImg)
'    End If

'Dim i As Integer
'    'Ex: all transparent at ratio 140/255
'    'ActiveTransparency Me, True, False, 140, Me.BackColor
'    'Ex: Form transparent, visible component at ratio 140/255
'    'ActiveTransparency Me, True, True, 140, Me.BackColor
'
'    'Example display the form transparency degradation
'    ActiveTransparency Me, True, False, 0
'    Me.Show
'    For i = 0 To 255 Step 3
'        ActiveTransparency Me, True, False, i
'        Me.Refresh
'    Next i
'
    nX = Screen.Width - Me.Width
    nY = Screen.Height - Me.Height
    Me.top = nY
    Me.left = nX
End Sub

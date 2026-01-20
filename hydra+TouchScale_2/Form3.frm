VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6227E780-EE80-4A81-B37C-ED4D8D3587EE}#1.0#0"; "HyperLbl.ocx"
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8775
   ClientLeft      =   3435
   ClientTop       =   1920
   ClientWidth     =   7200
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   ScaleHeight     =   8775
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":0000
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   5
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   6
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HyperLbl.HyperLabel RTB1 
      Height          =   495
      Index           =   7
      Left            =   4080
      TabIndex        =   7
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":00AE
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   2
      Left            =   3000
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":015C
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   3
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":020A
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":02B8
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   5
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":0366
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   6
      Left            =   3000
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":0414
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   735
      Index           =   7
      Left            =   4320
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":04C2
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

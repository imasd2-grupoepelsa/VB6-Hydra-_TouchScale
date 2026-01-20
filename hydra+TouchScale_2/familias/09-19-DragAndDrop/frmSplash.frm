VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   2880
   ClientTop       =   4530
   ClientWidth     =   8100
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   2000
      Left            =   5340
      Top             =   3780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Young Challenger"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   915
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   5685
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word Game"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   915
      Left            =   2145
      TabIndex        =   1
      Top             =   1320
      Width           =   3825
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   7080
      Picture         =   "frmSplash.frx":0000
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   7500
      Picture         =   "frmSplash.frx":1272
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4020
      Picture         =   "frmSplash.frx":24E4
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   900
      Picture         =   "frmSplash.frx":3756
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmSplash.frx":49C8
      Top             =   900
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 thevbprogrammer.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   4650
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub tmrSplash_Timer()
    frmWordGame.Show
    Unload Me
End Sub

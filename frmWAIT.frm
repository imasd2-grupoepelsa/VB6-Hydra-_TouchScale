VERSION 5.00
Begin VB.Form frmWAIT 
   Caption         =   "Wait a minute..."
   ClientHeight    =   1665
   ClientLeft      =   3150
   ClientTop       =   2025
   ClientWidth     =   4830
   Enabled         =   0   'False
   Icon            =   "frmWAIT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerWait 
      Enabled         =   0   'False
      Left            =   4440
      Top             =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmWAIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Do_Events
    TimerWait.Interval = 1000
    TimerWait.Enabled = True
    Me.Refresh
    Do_Events
    
End Sub

Private Sub Form_Terminate()
TimerWait.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
TimerWait.Enabled = False
End Sub

Private Sub TimerWait_Timer()
Do_Events
Me.Refresh
Do_Events
End Sub

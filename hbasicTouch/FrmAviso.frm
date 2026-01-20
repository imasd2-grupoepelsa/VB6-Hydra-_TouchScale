VERSION 5.00
Begin VB.Form FrmAviso 
   Caption         =   "Aviso Actividad TOUCH-SCALE."
   ClientHeight    =   1230
   ClientLeft      =   4080
   ClientTop       =   3495
   ClientWidth     =   7200
   Icon            =   "FrmAviso.frx":0000
   LinkTopic       =   "FrmAviso"
   ScaleHeight     =   1230
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblAviso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "FrmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

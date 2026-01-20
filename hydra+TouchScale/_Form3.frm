VERSION 5.00
Begin VB.Form FrmTextoECO 
   Caption         =   "Form3"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4944
   LinkTopic       =   "Form3"
   ScaleHeight     =   3012
   ScaleWidth      =   4944
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   372
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   2400
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   0
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1200
      Width           =   4932
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   0
      MaxLength       =   25
      TabIndex        =   1
      Top             =   960
      Width           =   4932
   End
   Begin VB.Image Image1 
      Height          =   612
      Left            =   0
      Top             =   2280
      Width           =   4932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   732
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4932
   End
End
Attribute VB_Name = "FrmTextoECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
    Unload Me
End Sub

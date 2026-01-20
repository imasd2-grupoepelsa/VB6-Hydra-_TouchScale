VERSION 5.00
Begin VB.Form frmorden 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura Msb200"
   ClientHeight    =   2148
   ClientLeft      =   3552
   ClientTop       =   1992
   ClientWidth     =   4740
   Enabled         =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "frmorden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2148
   ScaleWidth      =   4740
   Begin VB.TextBox textdes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox textcod 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox textfam 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox textsub 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox textsec 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000018&
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   4572
   End
   Begin VB.Label Label6 
      Caption         =   "Transforma Msb200 :"
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label labelmolino 
      Caption         =   "/"
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
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Descripción:"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Código:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Fam.:"
      Height          =   252
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   492
   End
   Begin VB.Label label2 
      Caption         =   "Sub.:"
      Height          =   252
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Secc.:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   492
   End
End
Attribute VB_Name = "frmorden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
End Sub


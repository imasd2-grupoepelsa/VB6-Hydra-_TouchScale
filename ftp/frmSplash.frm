VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3270
   ClientLeft      =   5850
   ClientTop       =   3465
   ClientWidth     =   3735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   1  'Dash
   FillColor       =   &H00FFFFC0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000013&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1320
      Picture         =   "frmSplash.frx":000C
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":1B4E
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   3015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProduct.Caption = App.Title
    
End Sub



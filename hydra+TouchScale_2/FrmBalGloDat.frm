VERSION 5.00
Begin VB.Form FrmBalGloDat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GLOBAL.DAT"
   ClientHeight    =   3105
   ClientLeft      =   2175
   ClientTop       =   2490
   ClientWidth     =   6570
   Enabled         =   0   'False
   FillColor       =   &H00C0E0FF&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3105
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2985
      ItemData        =   "FrmBalGloDat.frx":0000
      Left            =   0
      List            =   "FrmBalGloDat.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   6612
   End
End
Attribute VB_Name = "FrmBalGloDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyCont As Integer
Private Sub Form_Load()
    Me.Icon = Form2.Icon
End Sub
Public Sub dato(sDato As String)
    List1.AddItem sDato
    List1.TopIndex = List1.ListCount - 1
    If List1.ListCount > 8000 Then List1.Clear
    MyCont = MyCont + 1
    If MyCont > 9 Then
        Do_Events
        MyCont = 0
    End If
End Sub



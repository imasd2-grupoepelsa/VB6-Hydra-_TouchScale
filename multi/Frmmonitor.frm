VERSION 5.00
Begin VB.Form Frmmonitor 
   Caption         =   "Backup"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5172
   LinkTopic       =   "Form1"
   ScaleHeight     =   2544
   ScaleWidth      =   5172
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1392
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5172
   End
End
Attribute VB_Name = "Frmmonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Icon = FrmPrincipal.Icon
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
End Sub

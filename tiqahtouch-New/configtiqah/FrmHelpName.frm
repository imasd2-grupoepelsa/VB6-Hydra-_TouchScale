VERSION 5.00
Begin VB.Form FrmHelpName 
   Caption         =   "Ayuda nombre Fichero."
   ClientHeight    =   1455
   ClientLeft      =   1665
   ClientTop       =   3930
   ClientWidth     =   4485
   Icon            =   "FrmHelpName.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4485
   Begin VB.ListBox ListHelpName 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmHelpName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_LostFocus()
    Me.Hide
End Sub

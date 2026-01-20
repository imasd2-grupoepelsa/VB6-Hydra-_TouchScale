VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   600
   ClientLeft      =   6150
   ClientTop       =   2685
   ClientWidth     =   5580
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   5580
   Begin VB.Timer TmrExclusivo 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   120
   End
   Begin VB.Image Imagen 
      Height          =   585
      Left            =   0
      Picture         =   "Form2.frx":1E72
      Top             =   0
      Width           =   5565
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Exclusivo_Fin As Boolean


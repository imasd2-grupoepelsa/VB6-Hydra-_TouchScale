VERSION 5.00
Begin VB.Form FrmTools 
   Caption         =   "Form1"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2544
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1032
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   492
   End
   Begin VB.DirListBox Dir1 
      Height          =   720
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

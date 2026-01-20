VERSION 5.00
Begin VB.Form FrmTools 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   5820
   ClientTop       =   2775
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   3750
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   492
   End
   Begin VB.DirListBox Dir1 
      Height          =   504
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   852
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

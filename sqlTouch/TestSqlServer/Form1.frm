VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1560
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   Begin VB.CommandButton CmdLeeTickets 
      Caption         =   "LeeTickets"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdCapturaIVA 
      Caption         =   "Captura IVAS"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CmdCapturaArts 
      Caption         =   "Captura Artículos"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer TmrArranque 
      Interval        =   1000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCapturaArts_Click()
    CmdCapturaArts.Enabled = False
    CmdCapturaIVA.Enabled = False
    Call LeeSecciones
    Call LeeFamilias
    Call LeeImagenesFamilias
    Call LeeVendedores
    Call LeeImagenesVendedores
    Call LeeIVAS
    Call LeeArticulos
    Call LeeImagenesArticulos
    Call LeeTramos
    
    CmdCapturaArts.Enabled = True
    CmdCapturaIVA.Enabled = True

End Sub

Private Sub CmdCapturaIVA_Click()
    CmdCapturaArts.Enabled = False
    CmdCapturaIVA.Enabled = False
    Call LeeIVAS
    CmdCapturaArts.Enabled = True
    CmdCapturaIVA.Enabled = True

End Sub

Private Sub CmdLeeTickets_Click()
CmdLeeTickets.Enabled = False
Call LeeTickets
CmdLeeTickets.Enabled = True
End Sub

Private Sub TmrArranque_Timer()
    TmrArranque.Enabled = False
    'Call LeeImagenesArticulos
    'Call LeeImagenesFamilias
    'Call LeeImagenesVendedores
End Sub

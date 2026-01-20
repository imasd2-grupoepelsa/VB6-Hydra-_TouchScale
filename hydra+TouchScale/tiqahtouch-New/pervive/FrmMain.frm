VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Arranque"
   ClientHeight    =   465
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Pervive"
   ScaleHeight     =   465
   ScaleWidth      =   1560
   WindowState     =   1  'Minimized
   Begin VB.Timer TmrVer 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TmrVer_Timer()
    TmrVer.Enabled = False
    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\TIQAH.EXE", vbMinimizedFocus))
    TmrVer.Enabled = True
End Sub

VERSION 5.00
Begin VB.Form FormTQT 
   Caption         =   "FTQT"
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7200
End
Attribute VB_Name = "FormTQT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub timetqt_Timer()
    Dim nMy As Form
    
    timetqt.Enabled = False
    
    Set nMy = New frmControl
    nMy.timetqtimer
    Unload nMy
    Set nMy = Nothing
    'Unload FrmExportar
    'Set FrmExportar = Nothing
    'Unload Me
    'Load Me
    timetqt.Enabled = True
       
End Sub

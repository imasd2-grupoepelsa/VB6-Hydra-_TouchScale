VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1455
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   1965
   LinkTopic       =   "Form2"
   ScaleHeight     =   1455
   ScaleWidth      =   1965
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   240
      Top             =   360
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form2.frx":0000
      TrayTip         =   "VB 5 - SysTray Control."
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    cSysTray1.InTray = False
    Form1.Show
    
End Sub

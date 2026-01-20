VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSelectEuroscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub Show()
    FrmSelectEuroscale.Show
    FrmSelectEuroscale.LoadData
    Do While FrmSelectEuroscale.Visible = True
        Do_Events
    Loop
End Sub
Public Property Get SysNumber() As String
    SysNumber = FrmSelectEuroscale.SelectedSys
End Property
Private Sub Class_Initialize()
    Load FrmSelectEuroscale
End Sub
Private Sub Class_Terminate()
    Unload FrmSelectEuroscale
End Sub

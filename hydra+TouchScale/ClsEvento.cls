VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Clase genérica para comunicarse por eventos
'**********************************************
Option Explicit
Public Event Evento1(lData As Long)
Public Event Evento2(sData As String)
Public Event Evento3(lData As Long, sData As String)
Public Sub Raise1(ByVal lData As Long)
    RaiseEvent Evento1(lData)
End Sub
Public Sub Raise2(ByVal sData As String)
    RaiseEvent Evento2(sData)
End Sub
Public Sub Raise3(ByVal lData As Long, ByVal sData As String)
    RaiseEvent Evento3(lData, sData)
End Sub

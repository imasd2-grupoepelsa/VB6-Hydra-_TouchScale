VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventoBas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'********************************************
'  Clase auxiliar para poder disparar
'  Eventos desde módulos .BAS
'********************************************
Public Event BasTicketReceived(ByVal NTicket As Long, ByVal nSection As Long, ByVal NScale As Long, ByRef Cancel As Long)
Public Event BasItemReceived(ByVal TotalID As Long, ByVal lData As Long)
Public Event BasComError(ByVal Error_Code As Long)
Public Event BasComOK()
Public Sub Raise_TicketReceived(ByVal NTicket As Long, ByVal nSection As Long, ByVal NScale As Long, ByRef Cancel As Long)
    RaiseEvent BasTicketReceived(NTicket, nSection, NScale, Cancel)
End Sub
Public Sub Raise_ItemReceived(ByVal TotalID As Long, ByVal lData As Long)
    RaiseEvent BasItemReceived(TotalID, lData)
End Sub
Public Sub Raise_ComError(ByVal Error_Code As Long)
    RaiseEvent BasComError(Error_Code)
End Sub
Public Sub Raise_ComOK()
    RaiseEvent BasComOK
End Sub


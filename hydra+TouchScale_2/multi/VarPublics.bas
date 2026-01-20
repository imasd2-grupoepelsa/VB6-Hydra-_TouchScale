Attribute VB_Name = "VarPublics"
'***********************
' tipo :
' 0 --> Euroscale
' 1 --> V-12 SC10
' 2 --> V-12 C10
' 3 --> V-12 C9
' 4 --> V-8
' 5 --> ECO NET
'***********************
'***********************
' Recurso
' 0 --> TCP/IP directo
' 1 --> TCP/IP + acceso telefónico
' 2 --> TCP/IP + Proxy
' 3 --> RS-232 directo
' 4 --> RS-232 modem
' 5 --> Path de Red
'***********************
Public Type TipoTienda
    Numero As Integer
    Nombre As String
    Direccion As String
    Poblacion As String
    Provincia As String
    CodPostal As String
    TlfContacto As String
    Grp As String
    File As String
    Patron As String
    sPathS As String
    lScalePos As Boolean 'rehacer base de datos ScalePos 1.0.8
End Type
Public MiRuta As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public sPathTien As String
Public sPassG As String

Public nMyPer As Long
'''''''
Public D_TiendaActual() As TipoTienda
'''''''

Attribute VB_Name = "Module1"
Option Explicit
'Para trabajar con ADODB declaramos...
Public rs As Recordset
Public BD As Connection
'Para trabajar con DAO declaramos...
Public MyBase As dao.Database
Public MyRe As dao.Recordset
Public MyCli As dao.Recordset
' Se publica API para poder lanzar un link a una página web
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public clinom As String ' variable para concatenar código de cliente y nombre en Combobox4
Public codcli As Variant 'constante que nos indica si el valor de combo4 ha cambiado.
Public fecha1 As Variant 'constante que nos indica si el valor de combo1 y combo2 han cambiado.


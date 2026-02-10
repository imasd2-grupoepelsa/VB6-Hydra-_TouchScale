VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************
' Clase para compatibilidad Pantallas Locales
' - Pantallas Remotas
'***********************************************
Public Fields As Collection
Private Base As dao.Database
Private Registro As dao.Recordset
Private bEOF As Boolean
Private bBOF As Boolean
Private Sub FillData()
    Dim bucle As Integer
    If (Not Registro.EOF) And (Not Registro.BOF) Then
        Set Fields = New Collection
        For bucle = 0 To Registro.Fields.Count - 1
            Fields.Add Registro.Fields(bucle), Registro.Fields(bucle).name
        Next bucle
    Else
        Set Fields = Nothing
    End If
End Sub
Public Property Get EOF() As Boolean
    EOF = Registro.EOF
End Property
Public Property Get BOF() As Boolean
    BOF = Registro.BOF
End Property
Public Property Get Recordcount() As Long
    Recordcount = Registro.Recordcount
End Property
Public Sub Movenext()
    Registro.Movenext
    FillData
End Sub
Public Sub MovePrevious()
    Registro.MovePrevious
    FillData
End Sub
Public Sub MoveFirst()
    Registro.MoveFirst
    FillData
End Sub
Public Sub MoveLast()
    Registro.MoveLast
    FillData
End Sub

Public Sub OpenRecordset(cad As String, Optional ReadOnly As Boolean)
    Set Registro = Base.OpenRecordset(cad)
    If Not Registro.EOF Then
        Registro.MoveFirst
        FillData
    End If
End Sub

Private Sub Class_Initialize()

'On Error Resume Next
'If Base.Name <> "" Then
'    If Err.Number <> 0 Then
'        On Error GoTo 0
    Set Base = OpenDatabase(Base_General)
    '    End If
    'End If
    'On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Base.Close
End Sub

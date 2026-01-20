Attribute VB_Name = "fio"
Option Explicit
'//////////////
'//////////////
'//////////////
Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    Fichero = FreeFile
    On Error Resume Next
    Open PathLocal & "\loghydra.txt" For Append As Fichero
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    On Error GoTo 0
End Sub


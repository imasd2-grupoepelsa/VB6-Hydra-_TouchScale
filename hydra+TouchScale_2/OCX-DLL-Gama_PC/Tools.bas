Attribute VB_Name = "Tools"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public PathLog As String
Public EnviaEvento As CEventoBas

Public Function Do_Events()
    Sleep (1)
    DoEvents
End Function

Public Sub PonDato(Tipo As String, Valor As Long)
    Dim MiCad As String
    If Valor = 0 Then
        MiCad = "OK"
    Else
        MiCad = "Error : " & Valor
    End If
    CadenadeLog Tipo & " " & MiCad
End Sub


Attribute VB_Name = "Taquion_link"
Option Explicit
'*********************************************
' funciones referidas al enlace ASCII, originario
' de TAQUION / TQRETAIL
' que ha quedado como base del enlace estandar
'*********************************************
Public Type Tipo_Grupo_IVA
    porcentaje As Double
    Base As Double
    Importe As Double
End Type
Public Type Tipo_NombreTiquet
    Balanza As Integer
    seccion As Integer
    tiquet As Long
    Modo As Integer
    cliente As Integer
    vendedor As Integer
End Type
'c2f/caspiunza **
Public Function ObtenNombre(MiNombre As Tipo_NombreTiquet) As String
    Dim Resp As String
    Dim V1 As String
    Dim Cont As Integer
    Dim Buffer As String
    Dim Buf2 As String
    Dim Formato As String
    Dim b As Integer
    Dim B2 As Integer
    TqNombre = UCase(TqNombre)
    Resp = ""
    Cont = 1
    Buffer = Trim(TqNombre & "/")
    V1 = left(Buffer, 1)
    For b = 2 To Len(Buffer)
        If Mid(Buffer, b, 1) = V1 Then
            Cont = Cont + 1
        Else
            Formato = ""
            Buf2 = ""
            For B2 = 1 To Cont
                Formato = Formato & "0"
            Next B2
            Select Case V1
            Case "B"
                Buf2 = Format(MiNombre.Balanza, Formato)
            Case "S"
                Buf2 = Format(MiNombre.seccion, Formato)
            Case "T"
                Buf2 = Format(MiNombre.tiquet, Formato)
            Case "X"
                Buf2 = Format(MiNombre.Modo, Formato)
            Case "C"
                Buf2 = Format(MiNombre.cliente, Formato)
            Case "V"
                Buf2 = Format(MiNombre.vendedor, Formato)
            End Select
            If Len(Buf2) > Cont Then Buf2 = Right(Buf2, Cont)
            Resp = Resp & Buf2
            V1 = Mid(Buffer, b, 1)
            Cont = 1
        End If
    Next b
    ObtenNombre = Resp
End Function



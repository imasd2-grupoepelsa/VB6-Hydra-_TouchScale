Attribute VB_Name = "Preferencias"
Option Explicit
Private Sub Check_Preferencias()
    Dim Arch As Integer
    If Dir(Miruta & "\etc\preferencias.ini") = "" Then
        Arch = FreeFile()
        Open Miruta & "\etc\preferencias.ini" For Output As #Arch
        Print #Arch, "TECLAS    " & "01" & "0" & "0" & "056"
        Print #Arch, "CHGPRECIO1" & "01" & "0"
        Close #Arch
    End If
End Sub
Public Function RPreferencias(Tipo As String, tienda As Integer) As String
    Dim Arch As Integer
    Dim Buffer As String
    On Error GoTo fin
    If Dir(Miruta & "\etc\preferencias.ini") = "" Then
        Check_Preferencias
    End If
    Arch = FreeFile()
    Open Miruta & "\etc\preferencias.ini" For Input As #Arch
    Buffer = ""
    Do Until EOF(Arch)
        Line Input #Arch, Buffer
        If Len(Buffer) > 12 Then
            If UCase(left(Buffer, 10)) = UCase(Tipo) Then
                If Val(Mid(Buffer, 11, 2)) = tienda Then
                    Exit Do
                End If
            End If
        End If
    Loop
    Close #Arch
    If Len(Buffer) > 12 Then
        RPreferencias = Mid(Buffer, 13)
    Else
        RPreferencias = ""
    End If
    Exit Function
fin:
    On Error Resume Next
    Close #Arch
    On Error GoTo 0
    RPreferencias = ""
End Function
Public Sub WPreferencias(Tipo As String, tienda As Integer, texto As String)
    Dim Arch As Integer
    Dim ArchTemp As Integer
    Dim Buffer As String
    Dim eNCONTRADO As Boolean
    On Error GoTo fin
    If Dir(Miruta & "\etc\preferencias.ini") = "" Then
        Check_Preferencias
    End If
    Arch = FreeFile()
    Open Miruta & "\etc\preferencias.ini" For Input As #Arch
    ArchTemp = FreeFile()
    Open Miruta & "\etc\preferencias.tmp" For Output As #ArchTemp
    Buffer = ""
    Do Until EOF(Arch)
        Line Input #Arch, Buffer
        If Len(Buffer) > 12 Then
            If UCase(left(Buffer, 10)) = UCase(Tipo) Then
                If Val(Mid(Buffer, 11, 2)) = tienda Then
                    Print #ArchTemp, Tipo & Format(tienda, "00") & texto
                    eNCONTRADO = True
                Else
                    Print #ArchTemp, Buffer
                End If
            Else
                Print #ArchTemp, Buffer
            End If
        End If
    Loop
    If Not eNCONTRADO Then Print #ArchTemp, Tipo & Format(tienda, "00") & texto
    Close #Arch
    Close #ArchTemp
    On Error Resume Next
    If Dir(Miruta & "\etc\preferencias.tmp") <> "" Then
        FileCopy Miruta & "\etc\preferencias.tmp", Miruta & "\etc\preferencias.ini"
    End If
    On Error GoTo 0
    Exit Sub
fin:
    On Error Resume Next
    Close #Arch
    Close #ArchTemp
    On Error GoTo 0
End Sub

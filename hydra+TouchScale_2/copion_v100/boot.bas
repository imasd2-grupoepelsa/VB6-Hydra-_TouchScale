Attribute VB_Name = "boot"
Option Explicit
'''''''''''''''
'''''''''''''''
Sub Main()
    Dim Arch As Integer
    Dim orden As String
    Dim P1 As String
    Dim Buf1 As String
    Dim Buf2 As String
    Dim Buf3 As String
    Dim pntRuta As Long
    Dim bufCad As String
    On Error GoTo ret
    If Dir(App.Path & "\registro.txt") <> "" Then
        Kill App.Path & "\registro.txt"
    End If
    If Dir(App.Path & "\fifo") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\fifo" For Input As #Arch
        Line Input #Arch, DBASE_MDB
        Line Input #Arch, HYDRA_INI
        Line Input #Arch, orden
        Line Input #Arch, P1
        Line Input #Arch, Buf1
        Line Input #Arch, Buf2
        Line Input #Arch, Buf3
        Close (Arch)
    Else
        MsgBox "This program must be executed by Hydra+", vbCritical
        End
    End If
    If UCase(Buf1) = "Y" Then
        HaySeguridad = True
        UsuarioActual.Nombre = Buf2
        UsuarioActual.Asociado = Val(Buf3)
    Else
        HaySeguridad = False
        UsuarioActual.Nombre = "admin"
        UsuarioActual.Asociado = 0
    End If
    bufCad = DBASE_MDB
    Do
        pntRuta = InStr(1, bufCad, "\")
        If (pntRuta = 0) Then Exit Do
        bufCad = Right(bufCad, (Len(bufCad) - pntRuta))
    Loop
    var.rutaSYS = Left(DBASE_MDB, (Len(DBASE_MDB) - Len(bufCad) - 1))
    If (Dir(var.rutaSYS & "\S", vbDirectory) = "") Then MkDir (var.rutaSYS & "\S")
    LeerParametrosInicio
    Select Case LCase(orden)
        Case "proctotalhistorico"
           proc.backup False
        Case "procsiosi"
           proc.backup True
        Case Else
            Err.Raise 1000
    End Select
ret:
    If (Err.Number <> 0) Then MsgBox "Este programa sólo puede ser ejecutado desde Hydra+", vbCritical
    Err.Clear
    On Error GoTo 0
    End
End Sub
Private Sub Check_Especiales()
    If Dir(MiruTa & "\sabeco.sbc") <> "" Then Sabeco = True
    If Dir(MiruTa & "\ahorramas.cfg") <> "" Then AhorraMas = True
    If Dir(MiruTa & "\pathorde.ini") <> "" Then Alcampo = True
End Sub
Public Sub MostrarModal(Formulario As Form)
    Descarga = False
    Formulario.Show
    Do Until Descarga
        Do_Events
    Loop
End Sub
'''''''
'''''''
Public Sub MostrarModal2()
    Do Until Descarga
        Do_Events
    Loop
End Sub
'''''''
'''''''
Public Sub Do_Events()
    Sleep (1)
    DoEvents
End Sub
'''''''
'''''''

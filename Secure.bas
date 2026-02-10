Attribute VB_Name = "Secure"
Option Explicit
Public Function Encripta(texto As String) As String
    Dim Cl1 As String
    Dim Cl2 As String
    Dim MiSalida As String
    Dim bucle As Integer
    For bucle = 0 To 255
        Cl2 = Cl2 & Chr(bucle)
    Next bucle
    For bucle = 255 To 133 Step -1
        Cl1 = Cl1 & Chr(bucle)
    Next bucle
    For bucle = 0 To 132
        Cl1 = Cl1 & Chr(bucle)
    Next bucle
    For bucle = 1 To Len(texto)
        MiSalida = MiSalida & Mid(Cl1, InStr(Cl2, Mid(texto, bucle, 1)), 1)
    Next bucle
    Encripta = MiSalida
End Function
Public Function DesEncripta(texto As String) As String
    Dim Cl1 As String
    Dim Cl2 As String
    Dim MiSalida As String
    Dim bucle As Integer
    For bucle = 0 To 255
        Cl1 = Cl1 & Chr(bucle)
    Next bucle
    For bucle = 255 To 133 Step -1
        Cl2 = Cl2 & Chr(bucle)
    Next bucle
    For bucle = 0 To 132
        Cl2 = Cl2 & Chr(bucle)
    Next bucle
    For bucle = 1 To Len(texto)
        MiSalida = MiSalida & Mid(Cl1, InStr(Cl2, Mid(texto, bucle, 1)), 1)
    Next bucle
    DesEncripta = MiSalida
End Function
Public Function CargaPermisos(usuarios As Tipo_User, Buffer As String) As Boolean
    Dim MiCheck As String
    Dim MiCheck2 As String
    Dim Mipermiso As Byte
    Dim bucle As Integer
    MiCheck = Right(Buffer, 1)
    calcheck left(Buffer, Len(Buffer) - 1), MiCheck2
    If Asc(MiCheck) <> MiCheck2 Then
        CargaPermisos = False

        Exit Function
    End If
    usuarios.Nombre = StrConv(left(Buffer, 8), vbLowerCase)
    usuarios.Asociado = Val(Mid(Buffer, 25, 3))
    usuarios.password = Mid(Buffer, 9, 8)
    '*****************
    'obtiene Permisos
    '*****************
    For bucle = 1 To 8
        Mipermiso = Asc(Mid(Buffer, 16 + bucle, 1))
        usuarios.Permisos(bucle - 1) = Mipermiso
    Next bucle
End Function
Public Function Secure_Login(Mi_Usuario As Tipo_User, Optional NoCheckPass As Boolean) As Boolean
    Dim bucle As Integer
    Dim Archivo As Integer
    Dim Buffer As String
    Dim Nulos As String * 29
    Dim BufByte(29) As Byte
    Dim eNCONTRADO As Boolean
    Dim MiP As Byte
    Dim MiCheck
    Archivo = FreeFile()
    'Open MiruTa & "\pass.sec" For Binary As Archivo
    Open App.Path & "\pass.sec" For Binary As Archivo
    Do Until EOF(Archivo) Or eNCONTRADO
        For bucle = 0 To 28
            Get #Archivo, , BufByte(bucle)
        Next bucle
        'Buffer = Mid(Buffer, 2, Len(Buffer) - 1)
        Buffer = ""
        For bucle = 0 To 28
            Buffer = Buffer & Chr(BufByte(bucle))
        Next bucle

        If Buffer <> Nulos Then Buffer = DesEncripta(Buffer)
        If Buffer <> Nulos Then
            If Trim(left(Buffer, 8)) = Trim(Mi_Usuario.Nombre) Then
                eNCONTRADO = True
            End If
        End If
    Loop
    Close #Archivo

    If eNCONTRADO Then

        calcheck left(Buffer, Len(Buffer) - 1), MiCheck
        If Val(MiCheck) <> Asc(Right(Buffer, 1)) Then
            CadenadeLog "violación de seguridad : fichero pass.sec modificado"

            End
        End If
        Mi_Usuario.Permisos(0) = 0
        If Not NoCheckPass Then
            If Trim(Mi_Usuario.password) <> Trim(Mid(Buffer, 9, 8)) Then
                Secure_Login = False
                Exit Function
            End If
        End If
        Mi_Usuario.Asociado = Val(Mid(Buffer, 25, 8))
        MiP = Asc(Mid(Buffer, 17, 1))
        For bucle = 0 To 7
            If MiP And (2 ^ bucle) Then
                Mi_Usuario.Permisos(0) = Mi_Usuario.Permisos(0) + (2 ^ bucle)
            End If
        Next bucle
        Mi_Usuario.Permisos(1) = 0
        MiP = Asc(Mid(Buffer, 18, 1))
        For bucle = 0 To 7
            If MiP And (2 ^ bucle) Then
                Mi_Usuario.Permisos(1) = Mi_Usuario.Permisos(1) + (2 ^ bucle)
            End If
        Next bucle
        Mi_Usuario.Permisos(2) = 0
        MiP = Asc(Mid(Buffer, 19, 1))
        For bucle = 0 To 7
            If MiP And (2 ^ bucle) Then
                Mi_Usuario.Permisos(2) = Mi_Usuario.Permisos(2) + (2 ^ bucle)
            End If
        Next bucle
        Mi_Usuario.Permisos(3) = 0
        MiP = Asc(Mid(Buffer, 20, 1))
        For bucle = 0 To 7
            If MiP And (2 ^ bucle) Then
                Mi_Usuario.Permisos(3) = Mi_Usuario.Permisos(3) + (2 ^ bucle)
            End If
        Next bucle
        Mi_Usuario.Permisos(4) = 0
        MiP = Asc(Mid(Buffer, 21, 1))
        For bucle = 0 To 7
            If MiP And (2 ^ bucle) Then
                Mi_Usuario.Permisos(4) = Mi_Usuario.Permisos(4) + (2 ^ bucle)
            End If
        Next bucle
        Secure_Login = True
    Else
        Secure_Login = False
    End If

End Function
Public Function EncontrarUsuario() As Boolean
    Dim MiOK As Boolean
    Do Until MiOK
        If frmusuario.Visible = False Then
            Load frmusuario
            frmusuario.Show
            Do While frmusuario.Visible = True
                Do_Events
            Loop
            Unload frmusuario
        Else
            Exit Function
        End If
        If UsuarioActual.Nombre = "NONE" Then
            MiOK = False
            EncontrarUsuario = False
            Exit Function
        Else
            MiOK = Secure_Login(UsuarioActual)
            If MiOK = False Then
                MsgBox CargaCadena(769), vbCritical
                CadenadeLog CargaCadena(863) & " : LOGIN = " & UsuarioActual.Nombre
            Else
                'grventas
            End If
        End If
    Loop
    EncontrarUsuario = MiOK
End Function
Public Function ConsultarPermiso(ByVal Tipo As Integer) As Boolean
    Dim MiIndice As Integer
    '**********
    If Tipo < 8 Then
        MiIndice = 0
    Else
        If Tipo < 16 Then
            MiIndice = 1
            Tipo = Tipo - 8
        Else
            If Tipo < 24 Then
                MiIndice = 2
                Tipo = Tipo - 16
            Else
                If Tipo < 32 Then
                    MiIndice = 3
                    Tipo = Tipo - 24
                Else
                    MiIndice = 4
                    Tipo = Tipo - 32
                End If
            End If
        End If
    End If
    '**********
    If UsuarioActual.Permisos(MiIndice) And (2 ^ Tipo) Then
        ConsultarPermiso = True
    Else
        ConsultarPermiso = False
    End If
End Function
Public Sub CrearPassSec()
    Dim Fichero As Integer
    Dim Buffer As String
    Dim BufByte As Byte
    Dim Checksum
    Dim bucle As Integer
    Fichero = FreeFile()
    '1.8.9
    'Open MiruTa & "\pass.sec" For Output As #Fichero
    Open App.Path & "\pass.sec" For Output As #Fichero
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Close #Fichero
    Fichero = FreeFile()
    '1.8.9
    'Open MiruTa & "\pass.sec" For Binary As #Fichero
    Open App.Path & "\pass.sec" For Binary As #Fichero
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Buffer = "admin   " & "drake   "
    For bucle = 1 To 8
        Buffer = Buffer & Chr(255)
    Next bucle
    Buffer = Buffer & "000"
    Buffer = Buffer & Chr(1)
    calcheck Buffer, Checksum
    Buffer = Buffer & Chr(Checksum)
    Buffer = Encripta(Buffer)
    For bucle = 1 To 29
        BufByte = Asc(Mid(Buffer, bucle, 1))
        Put #Fichero, , BufByte
    Next bucle
    Buffer = "cmdline " & "cmdline "
    For bucle = 1 To 8
        Buffer = Buffer & Chr(255)
    Next bucle
    Buffer = Buffer & "000"
    Buffer = Buffer & Chr(2)
    calcheck Buffer, Checksum
    Buffer = Buffer & Chr(Checksum)
    Buffer = Encripta(Buffer)
    For bucle = 1 To 29
        BufByte = Asc(Mid(Buffer, bucle, 1))
        Put #Fichero, , BufByte
    Next bucle
    Close #Fichero
End Sub

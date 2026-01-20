Attribute VB_Name = "CDKEY"
'******************************************************
' EXA-FRANCE emplea HydraTouch protegido con
' un sistema de CD-KEY. Este es el módulo
' para implementar la protección.
' PAra que se compile HydraTouch con esta
' protección, hay que poner el valor de la
' constante CDKEY_USER (módulo arranque.bas)
' a true
'******************************************************
Option Explicit
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_CDROM = 5
Public ClavePublica As String
Public Sub Check_CDKEY()
    Dim Buffer As String
    Dim Resp As Integer
    Dim Resp2 As String
    Dim Continuar As Boolean
    Dim Arch As Integer
    Dim BufByte As Byte
    Dim BufClave As String
    Dim Contador As Integer
    CargaIdiomas
    Do Until Continuar
        Buffer = DameUnidad()
        If Buffer = "" Then
            Resp = MsgBox(CargaCadena(1121), vbOKCancel + vbInformation)
            If Resp = vbCancel Then End
        Else
            Continuar = True
        End If
    Loop
    Continuar = False
    Do Until Continuar
        Arch = FreeFile()
        On Error Resume Next
        Open Buffer For Binary As #Arch
        If Err.Number <> 0 Then
            MsgBox CargaCadena(1122), vbCritical
            End
        Else
            Contador = 0
            BufClave = ""
            Do Until EOF(Arch) Or Contador > 17
                Get #Arch, , BufByte
                BufClave = BufClave & Chr(BufByte)
                Contador = Contador + 1
            Loop
            Close #Arch
            If Contador = 17 Then
                BufClave = DesEncripta(left(BufClave, 16))
                On Error GoTo 0
                On Error Resume Next
                Buffer = GetSetting("Hydra Basic", "one", "User Key", "")
                If Err.Number <> 0 Or Buffer = "" Then
                    Continuar = False
                    Do Until Continuar
                        Resp2 = InputBox(CargaCadena(1123), "HydraTouch Activation")
                        If Len(Resp2) = 19 Then
                            Resp2 = left(Resp2, 4) & Mid(Resp2, 6, 4) & Mid(Resp2, 11, 4) & Right(Resp2, 4)
                            If Resp2 = ResuelvePar(BufClave) Then
                                ClavePublica = left(Resp2, 4) & "-" & Mid(Resp2, 5, 4) & "-" & Mid(Resp2, 9, 4) & "-" & Right(Resp2, 4)
                                Continuar = True
                                On Error GoTo 0
                                SaveSetting "Hydra Basic", "one", "User Key", Resp2
                            End If
                        End If
                        If Not Continuar Then
                            Resp = MsgBox(CargaCadena(1124), vbYesNo + vbExclamation)
                            If Resp = vbNo Then End
                        End If
                    Loop
                Else
                    If Buffer <> ResuelvePar(BufClave) Then
                        MsgBox CargaCadena(1122), vbCritical
                        End
                    Else
                        ClavePublica = left(Buffer, 4) & "-" & Mid(Buffer, 5, 4) & "-" & Mid(Buffer, 9, 4) & "-" & Right(Buffer, 4)
                        Continuar = True
                    End If
                End If
            Else
                MsgBox CargaCadena(1122), vbCritical
                End
            End If
        End If
    Loop
End Sub
Public Function ResuelvePar(P1 As String)
    Dim P2 As String
    Dim Bucle As Integer
    Dim MyAsc As Integer
    Dim Continuar As Boolean
    Dim MyVal As Integer
    P2 = ""
    For Bucle = 1 To 16
        P2 = P2 & Chr(91 - (CInt(Mid(P1, Bucle, 1)) + Bucle))
    Next Bucle
    ResuelvePar = P2
End Function
Private Function DameUnidad() As String
    Dim Bucle As Integer
    Dim Buffer As String
    Dim Buf2 As String
    For Bucle = 0 To 25
        Buffer = Chr(65 + Bucle) & ":\"
        If GetDriveType(Buffer) <> DRIVE_CDROM Then
            On Error Resume Next
            Buf2 = Dir(Buffer & "util\exa_fr.fsd")
            If Err.Number = 0 Then
                If Buf2 <> "" Then
                    Buffer = Buffer & "util\exa_fr.fsd"
                    Bucle = 30
                Else
                    Buffer = ""
                End If
            Else
                Buffer = ""
            End If
        End If
    Next Bucle
    If Bucle <> 31 Then Buffer = ""
    DameUnidad = Buffer
End Function

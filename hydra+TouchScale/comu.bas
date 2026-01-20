Attribute VB_Name = "comu"
'/////////////////////////////////
' módulo procedente de MNG,
' a quedado relegado para las
' funciones de envío de datos
' en bruto (que no se hacen
' a través de EPELCOM)
'/////////////////////////////////
Option Explicit

Public Sub enviar(que_enviar, ByRef que_llega As String, ByVal lEsEnvBin As Boolean)
    '////////////////////////////////
    ' envío a Euroscale
    Dim envio As String
    Dim respuesta As String
    Dim resp As Long
    If frmControl.EstadoBal <> 0 Then
        frmControl.Switch_mode 0
    End If
    envio = que_enviar
    If Not lEsEnvBin Then
        resp = frmControl.Balanza.EnviaASCII(envio, respuesta) '// Envío por protocolo ASCII
    Else
        resp = frmControl.Balanza.EnviaBin(envio, respuesta) '// Envío por protocolo Binario
    End If
    If resp >= 1000 Then
        canCelar = 1
        que_llega = ""
    Else
        If Not lEsEnvBin Then
            que_llega = respuesta
        Else
            globalrecibe = respuesta
            que_llega = "0"
        End If
    End If
End Sub
Public Sub enviarSC10(que_enviar, que_llega As String, Optional ByVal lTec)
    '/////////////
    ' envío a SC10
    Dim envio As String
    Dim respuesta As String
    Dim resp As Long
    Dim nMax As Long
    nMax = 0
    envio = que_enviar
    If frmControl.EstadoBal <> 1 And frmControl.EstadoBal <> 2 Then
        frmControl.Switch_mode 1
    End If
    If (Not TipoC9) And Not (SC10Rs232) Then envio = Mid(envio, 1, Len(envio) - 1)
    On Error Resume Next
    For nMax = 0 To 4
        resp = frmControl.Balanza.EnviaASCII(envio, respuesta)
        If (Not var.SC10TCPIP) And (Not var.SC10Rs232) Then
            If resp = 0 Then Exit For
        End If

        If respuesta = "" And (lTec = False) Then Exit For
        If (Asc(Left(respuesta, 1)) = 2) And (Asc(Right(respuesta, 1)) = 13) Then
            If lTec Then
                If Mid(respuesta, 2, 1) = Chr(6) Then
                    Exit For
                Else
                    respuesta = ""
                End If
            Else
                Exit For
            End If
        Else
            respuesta = ""
        End If
    
        If (lTec) Then Sleep (1500) '2.0.30
        
    Next nMax
    On Error GoTo 0
    If resp >= 1000 Then
        canCelar = 1
        que_llega = ""
    Else
        que_llega = respuesta
    End If
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   4530
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nFich1, nFichDes, nCont, nOrden As Integer
Dim sOrg As String
Dim sCompuesta As String
Dim sDes(99) As String
Dim sPes As String
Dim sPath As String
Dim nTime1, nTime2 As Date
Dim sCad As String
Dim lPendiente As Boolean
Dim lEsPrimera As Boolean
Dim lRestar As Boolean
Dim nResta As Integer
Dim lEsHydra As Boolean

If Dir(App.Path & "\hydra.exe") <> "" Then
    lEsHydra = True
Else
    lEsHydra = False
End If

If Dir(App.Path & "\pluasc.dir") = "" Then
    nFich1 = FreeFile()
    Open App.Path & "\pluasc.dir" For Output As #nFich1
    Print #nFich1, "C:\telecom\"
    Close #nFich1
Else
    nFich1 = FreeFile()
    Open App.Path & "\pluasc.dir" For Input As #nFich1
    Line Input #nFich1, sPath
    Close #nFich1
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    If Dir(sPath & "plu.asc") <> "" Then
        On Error Resume Next
        FileCopy sPath & "plu.asc", App.Path & "\plu.asc"
        On Error GoTo 0
        nTime1 = Time()
        nTime2 = Time()
        Do While (Second(nTime2)) < (Second(nTime1) + 3)
            nTime2 = Time()
        Loop
        
    End If
    
End If

On Error GoTo Problema

If Dir(App.Path & "\plu.asc") <> "" Then
    
    sCompuesta = ""
    lPendiente = False
    For nCont = 1 To 80
        sDes(nCont) = Space(24)
    Next nCont
    
    nFich1 = FreeFile()
    Open App.Path & "\plu.asc" For Input As #nFich1
    nFichDes = FreeFile()
    Open App.Path & "\balusera.dat" For Output As #nFichDes
    
    Do While Not EOF(nFich1) Or lPendiente = True
        If Not EOF(nFich1) Then
            Line Input #nFich1, sOrg
        End If
        If (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 2) = "00") Or (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 2) = "02") Or (EOF(nFich1) And lPendiente) Then
            If (EOF(nFich1) And lPendiente) Then
                If nOrden < 81 Then
                    nOrden = Val(Mid(sOrg, 15, 2))
                    sDes(nOrden) = Mid(sOrg, 18, 24)
                End If
            End If
            If sCompuesta <> "" Then
                For nCont = 1 To 80
                    sCompuesta = sCompuesta & sDes(nCont)
                Next nCont
                If Len(sCompuesta) > 1966 Then
                    'MsgBox sCompuesta
                    sCompuesta = Mid(sCompuesta, 1, 1966)
                End If
                If lEsHydra Then
                    sCompuesta = Replace(sCompuesta, "<b>", "|")
                    sCompuesta = Replace(sCompuesta, "</b>", "|")
                    If Len(sCompuesta) < 1966 Then
                        sCompuesta = sCompuesta & Space(1966 - Len(sCompuesta))
                    End If
                End If
                Print #nFichDes, sCompuesta & sPes & sCad
                sCompuesta = ""
                lPendiente = False
                If EOF(nFich1) Then Exit Do
            End If
            For nCont = 1 To 80
                sDes(nCont) = Space(24)
            Next nCont
            sCompuesta = Mid(sOrg, 55, 6) 'código
            lPendiente = True
            sCompuesta = sCompuesta & Mid(sOrg, 3, 2) 'sección
            sCompuesta = sCompuesta & "0000" 'Mid(sOrg, 57, 4) 'plu
            sCompuesta = sCompuesta & Mid(sOrg, 43, 6) 'precio
            sCompuesta = sCompuesta & Mid(sOrg, 50, 2) ' familia
            sCompuesta = sCompuesta & Mid(sOrg, 62, 1) 'IVA
            lEsPrimera = True
            lRestar = False
            nResta = 0
            sCompuesta = sCompuesta & Mid(sOrg, 18, 25) 'descriptivo principal
            If (Mid(sOrg, 16, 1) = "0") Then
                sPes = "W"
            Else
                sPes = "U"
            End If
        Else
            If Mid(sOrg, 13, 1) = "9" Then
                nOrden = Val(Mid(sOrg, 15, 2))
                If lRestar Then
                    If nOrden = 10 Then
                        nResta = 0 '8 '9
                    End If
                    lRestar = False
                End If
                nOrden = nOrden - nResta
                If nOrden < 81 Then
                    If (nOrden = 1) Or lEsPrimera Then
                        sCad = Mid(sOrg, 58, 3)
                        lEsPrimera = False
                        lRestar = True
                        'sDes(1) = Mid(sOrg, 18, 24) & " "
                        sDes(1) = Mid(sOrg, 18, 24)
                    End If
                    'If (2 * Int(nOrden / 2)) = nOrden Then
                    '    sDes(nOrden) = Mid(sOrg, 18, 24) & " "
                    'Else
                    '    sDes(nOrden) = " " & Mid(sOrg, 18, 24)
                    'End If
                    sDes(nOrden) = Mid(sOrg, 18, 24)
                Else
                nOrden = nOrden
                End If
            End If
        End If
    Loop
Problema:

If Err.Number <> 0 Then
    Err.Number = Err.Number
End If

    Close #nFich1
    Close #nFichDes
    On Error Resume Next
    Kill App.Path & "\plu.asc"
    Kill sPath & "plu.asc"
    On Error GoTo 0
    
End If

End

End Sub

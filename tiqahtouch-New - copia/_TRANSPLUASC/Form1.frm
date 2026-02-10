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
Dim sDes(10) As String
Dim sPes As String
Dim sPath As String
Dim nTime1, nTime2 As Date
Dim sCad As String

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


If Dir(App.Path & "\plu.asc") <> "" Then
    
    sCompuesta = ""
    
    nFich1 = FreeFile()
    Open App.Path & "\plu.asc" For Input As #nFich1
    nFichDes = FreeFile()
    Open App.Path & "\balusera.dat" For Output As #nFichDes
    
    Do While Not EOF(nFich1)
        
        Line Input #nFich1, sOrg
        
        If (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 1) = "0") Or (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 1) = "2") Then
            If sCompuesta <> "" Then
                For nCont = 1 To 10
                    sCompuesta = sCompuesta & sDes(nCont)
                Next nCont
                Print #nFichDes, sCompuesta & sPes & sCad
                sCompuesta = ""
            End If
            For nCont = 1 To 10
                sDes(nCont) = Space(25)
            Next nCont
            sCompuesta = Mid(sOrg, 54, 6) 'código
            sCompuesta = sCompuesta & Mid(sOrg, 3, 2) 'sección
            sCompuesta = sCompuesta & Mid(sOrg, 56, 4) 'plu
            sCompuesta = sCompuesta & Mid(sOrg, 42, 6) 'precio
            sCompuesta = sCompuesta & Mid(sOrg, 49, 2) ' familia
            sCompuesta = sCompuesta & Mid(sOrg, 61, 1) 'IVA
            
            sCompuesta = sCompuesta & Mid(sOrg, 17, 25) 'descriptivo principal
            If (Mid(sOrg, 15, 1) = "0") Then
                sPes = "W"
            Else
                sPes = "U"
            End If
            
        Else
            If Mid(sOrg, 13, 1) = "9" Then
                nOrden = Val(Mid(sOrg, 15, 1))
                If nOrden = 1 Then
                    sCad = Mid(sOrg, 57, 3)
                End If
                If (2 * Int(nOrden / 2)) = nOrden Then
                    sDes(nOrden) = Mid(sOrg, 17, 24) & " "
                Else
                    sDes(nOrden) = " " & Mid(sOrg, 17, 24)
                End If
                
            End If
        
        End If
    
    Loop

    Close #nFich1
    Close #nFichDes
    On Error Resume Next
    Kill App.Path & "\plu.asc"
    Kill sPath & "plu.asc"
    On Error GoTo 0
    
End If

End

End Sub

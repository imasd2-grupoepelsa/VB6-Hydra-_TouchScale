Attribute VB_Name = "Var"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOOWNERZORDER = &H200
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const Flags = SWP_NOOWNERZORDER Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
'\\\\\\\\\\\\\\\\\\\\\\\\
Public sIpAdd As String
Public sPort As String
Public nIpPort As Integer
'////////////////////////
Public lEspera As Boolean

Public sIpCajas1 As String
Public sHiper As String
Public sDirCajas1 As String
Public sUserCajas1 As String
Public sPassCajas1 As String
Public Const sPathOrg = "c:\balanzas\ia\entrada" '"C:\hydra+TouchScale\touchup" '
Public lastFFound As String

Public Function toIpAdd(ByVal sR As String)
    Dim Bytes(5) As Variant
    Dim Contador As Integer
    Dim Bucle As Integer
    Contador = 0
    
    For Bucle = 1 To Len(sR)
        Select Case Asc(Mid(sR, Bucle, 1))
            Case 48 To 57
                Bytes(Contador) = Bytes(Contador) & Mid(sR, Bucle, 1)
            Case Asc(".")
                Contador = Contador + 1
            Case Else
                Exit For
        End Select
        If Contador > 4 Then
            Exit For
        End If
    Next Bucle
    
    If Bytes(0) <> "" And Bytes(1) <> "" And Bytes(2) <> "" And Bytes(3) <> "" Then
        Bytes(0) = Val(Bytes(0))
        Bytes(1) = Val(Bytes(1))
        Bytes(2) = Val(Bytes(2))
        Bytes(3) = Val(Bytes(3))
    End If
    toIpAdd = CStr(Bytes(0)) & "." & CStr(Bytes(1)) & "." & CStr(Bytes(2)) & "." & CStr(Bytes(3))
   
End Function

Public Sub Do_Events()
    Sleep (3)
    DoEvents
End Sub

Public Function ElPing(ByVal ElHost As String) As Integer
    '///////////////////////////////////////////////
    '//  Lanza ping a la dirección que se indique //
    '//  llamando al OCX de ping                  //
    '///////////////////////////////////////////////
    Dim Contador As Integer
    Dim Echo As ICMP_ECHO_REPLY
    Dim Resp As Long
    Resp = ICMP_Ping(ElHost, "alive", Echo)
    If Resp = 0 Then
        ElPing = 1
    Else
        ElPing = 0
    End If
End Function
'resolver que puede haber más de un CRT...
Public Sub sendFiles(Optional ByVal mySD As String)
Dim Bucle As Integer
Dim nC As Integer
Dim siP, susR, spsS As String
Dim nFich As Integer
Dim ret As Long
Dim sOrg As String

''''
'Eliminar ficheros más antiguos de 45 dias de baarias --> Se lleva a main.
'On Error GoTo Si1
'    FrmAviso.Cls
'    FrmAviso.File1.Path = App.Path & "\baarias"
'    FrmAviso.File1.FileName = "*.*"
'    FrmAviso.File1.Refresh
'    If FrmAviso.File1.ListCount > 0 Then
'        For bucle = 0 To FrmAviso.File1.ListCount - 1
'            If Abs(DateDiff("d", CDate(Format(Now, "dd/mm/yyyy")), CDate(Format(FileDateTime(FrmAviso.File1.List(bucle)), "dd/mm/yyyy")))) > 45 Then
'                On Error Resume Next
'                Kill FrmAviso.File1.List(bucle)
'                On Error GoTo 0
'            End If
'        Next bucle
'    End If
'Si1:
'On Error GoTo 0
''''

    siP = toIpAdd(sIpAdd)
    susR = "pcscale"
    spsS = "epelsa"
    nFich = 0
    ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
    nRetConnFTP = ret
    lEspera = False
    If nRetConnFTP = 1 Then
        If Dir(App.Path & "\bal3.dat") <> "" Then
            ret = Enviar_Fichero_FTP(App.Path & "\bal3.dat", "/home/pcscale/scale/update/BAL2.dat", 0)
            If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                End
            End If
            FileCopy App.Path & "\bal3.dat", App.Path & "\baarias\bal3" & Format(Now(), "yymmddhhmmss") & ".dat"
            Sleep (5000)
            Kill App.Path & "\bal3.dat"
            Sleep (100)
        End If
        
        If mySD <> "" And (Not IsNull(mySD)) Then
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg 'spathorg
            FrmAviso.File1.FileName = "img" & sHiper & "???" & ".b" & mySD & ".z"
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & mySD & Format(Now(), "hhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "gpo" & sHiper & "???" & ".b" & mySD 'Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & mySD & Format(Now(), "hhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
    
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "tec" & sHiper & "???" & ".b" & mySD 'Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & mySD & Format(Now(), "hhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
    
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "crt" & sHiper & "???" & ".b" & mySD 'Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & mySD & Format(Now(), "hhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
                    
                Next Bucle
            End If
        
        Else
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg 'spathorg
            FrmAviso.File1.FileName = "img" & sHiper & "???" & ".b" & Format(Now(), "yymmdd") & ".z"
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "gpo" & sHiper & "???" & ".b" & Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
    
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "tec" & sHiper & "???" & ".b" & Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
    
                Next Bucle
            End If
            FrmAviso.Cls
            FrmAviso.File1.Path = sPathOrg
            FrmAviso.File1.FileName = "crt" & sHiper & "???" & ".b" & Format(Date, "yymmdd")
            FrmAviso.File1.Refresh
            If FrmAviso.File1.ListCount > 0 Then
                lEspera = True
                For Bucle = 0 To FrmAviso.File1.ListCount - 1
                    ret = Enviar_Fichero_FTP(sPathOrg & "\" & FrmAviso.File1.List(Bucle), "/home/pcscale/scale/update/" & UCase(FrmAviso.File1.List(Bucle)), 0)
                    If ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                        End
                    End If
                    'FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                    'Sleep (100)
                    'Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                    'Sleep (100)
                    On Error Resume Next
                    SetAttr sPathOrg & "\" & FrmAviso.File1.List(Bucle), vbNormal
                    Sleep (100)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                        On Error Resume Next
                        FileCopy sPathOrg & "\" & FrmAviso.File1.List(Bucle), App.Path & "\baarias\" & FrmAviso.File1.List(Bucle) & Format(Now(), "yymmddhhmmss")
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            On Error Resume Next
                            Sleep (100)
                            Kill sPathOrg & "\" & FrmAviso.File1.List(Bucle)
                            On Error GoTo 0
                            Sleep (100)
                        End If
                    End If
                    
                Next Bucle
            End If
        End If
        
        Sleep (1000)
    Else
        End
    End If
errorEnvioFtp:
    If Err.Number <> 0 And nRetConnFTP = 1 Then
        ret = Desconectar_FTP(0)
        FTP_hOpen(0) = 0
        FTP_HandleConect(0) = 0
        nRetConnFTP = 0
    End If
    
End Sub

Public Sub waitResul(Optional ByVal sD As String)
Dim sOrg As String
Dim lHorIni As Date
Dim lTOut As Boolean
Dim ret As Long
Dim nf As Integer
Dim nCont As Integer
Dim nMax As Integer

nf = FreeFile()
Open App.Path & "\resultadoftp.txt" For Output As #nf

    nMax = 0
    lHorIni = Now
    lTOut = False
    Sleep (400)
    'Do While getfilesize("/home/pcscale/scale/update/" & UCase("crt" & sHiper & "001" & ".b" & Format(Date, "yymmdd")), 0) <> 0 And lTOut = False
    For nCont = 99 To 1 Step -1
        sOrg = UCase("crt" & sHiper & Format(nCont, "000") & ".b" & sD)
        lTOut = False
        If getfilesize("/home/pcscale/scale/update/" & sOrg, 0) <> 0 Then
            nMax = nCont
            Print #nf, Now & " --> Existe:" & sOrg
            lTOut = True
            Exit For
        End If
        Do_Events
    Next nCont
    
    If lTOut Then
        lTOut = False
        'Do While ((getfilesize("/home/pcscale/scale/update/" & UCase("crt" & sHiper & Format(nMax, "000") & ".b" & sD), 0) <> 0) Or _
        '(getfilesize("/home/pcscale/scale/update/" & UCase("crt" & sHiper & Format(nMax, "000") & ".b" & Format(Date, "yymmdd")), 0) <> 0)) And lTOut = False
        Do While (getfilesize("/home/pcscale/scale/update/" & UCase("crt" & sHiper & Format(nMax, "000") & ".b" & sD), 0) <> 0) And lTOut = False
            If DateDiff("n", lHorIni, Now) > 120 Then lTOut = True
            Sleep (400)
            'Print #nf, Now & " --> Existe CRT...(Balanza)"
            Do_Events
        Loop
        If lTOut Then
            Print #nf, Now & " --> No se ha borrado CRT...(Balanza):" & CStr(nMax) & " MAX.120 sg"
            GoTo Sla2
        End If
        Print #nf, Now & " --> se ha borrado CRT...(Balanza):" & CStr(nMax)
        Do_Events
        Sleep (400)
        
        'sOrg = lastFFound
        If lTOut = False Then
        
        For nCont = 1 To nMax
        
            sOrg = UCase("log" & sHiper & Format(nCont, "000") & ".b" & sD) 'Format(Date, "yymmdd"))
            lTOut = False
            lHorIni = Now
            If sOrg <> "" Then
                Print #nf, Now & " --> Ver existencia:" & sOrg
                If (getfilesize("/home/pcscale/scale/update/" & sOrg, 0) <> 0) Then
                
                    ret = Recibir_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, App.Path & "\" & sOrg, 0)
                    Do_Events
                    If ret = 1 Then
                        Print #nf, Now & " --> Recibido: " & sOrg
                        Do_Events
                        'If lTOut = True Then End ' si no copiar a carpeta salida
                        ret = Borrar_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, 0)
                        If Dir(App.Path & "\" & sOrg) <> "" Then
                            FileCopy App.Path & "\" & sOrg, "c:\balanzas\ia\salida\" & sOrg
                            Print #nf, Now & " --> Copiado en Salida: " & sOrg
                        End If
                    'Else
                    '    sOrg = UCase("log" & sHiper & Format(nCont, "000") & ".b" & Format(Date, "yymmdd"))
                    '    ret = Recibir_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, App.Path & "\" & sOrg, 0)
                    '    Print #nf, Now & " --> Recibido: " & sOrg
                    '    Do_Events
                    '    'If lTOut = True Then End ' si no copiar a carpeta salida
                    '    ret = Borrar_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, 0)
                    '    If Dir(App.Path & "\" & sOrg) <> "" Then
                    '        FileCopy App.Path & "\" & sOrg, "c:\balanzas\ia\salida\" & sOrg
                    '        Print #nf, Now & " --> Copiado en Salida: " & sOrg
                    '    End If
                    End If
                    'Sleep (500)
                    'If lTOut = True Then End ' si no copiar a carpeta salida
                    'ret = Borrar_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, 0)
                    Do_Events
                    If ret = 1 Then
                        Print #nf, Now & " --> Borrado en Balanza: " & sOrg
                    Else
                        Print #nf, Now & " --> NO BORRADO EN BALANZA: " & sOrg
                    End If
                
                End If
                
                Do_Events
            End If
        
        Next nCont
    
        End If
    Else
        Print #nf, Now & " --> No se han encontrado ficheros traspasados..."
    End If
    
Sla2:

Close #nf
    
    Call GetEst
    
    End
End Sub

Public Sub GetEst()
Dim sOrg As String
Dim lHorIni As Date
Dim lTOut As Boolean
Dim ret, myret As Long
Dim nf As Integer
Dim siP, susR, spsS As String
Dim Bucle As Long
Dim pData As WIN32_FIND_DATA
Dim nCont As Integer
On Error GoTo salGetEst

    siP = toIpAdd(sIpAdd)
    susR = "pcscale"
    spsS = "epelsa"
    
    If nRetConnFTP = 1 Then
        ret = Desconectar_FTP(0)
        FTP_hOpen(0) = 0
        FTP_HandleConect(0) = 0
        nRetConnFTP = 0
    End If
    
    ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
    nRetConnFTP = ret
    
    'If Dir(App.Path & "\ESTRESUL", vbDirectory) = "" Then MkDir App.Path & "\ESTRESUL"
    'Sleep (100)
    'On Error Resume Next
    'Kill App.Path & "\ESTRESUL\*.*"
    'On Error GoTo 0
    'On Error GoTo salGetEst
    Sleep (100)
        
    'FrmAviso.Cls
    'FrmAviso.File1.Path = sPathOrg
    'FrmAviso.File1.FileName = "EST*.*"
    'FrmAviso.File1.Refresh
    'If FrmAviso.File1.ListCount > 0 Then
        
    '    For Bucle = 0 To FrmAviso.File1.ListCount - 1
nCont = 0
myret = 0
masFich:
Bucle = 0
ret = 0
    If nRetConnFTP = 1 Then
            pData.cFileName = String(260, 0)
            If nCont = 0 Then
                myret = FtpFindFirstFile(FTP_HandleConect(0), "/home/pcscale/scale/update/EST*.*", pData, 0, 0)
            Else
                ret = InternetFindNextFile(myret, pData)
            End If
            sOrg = Trim(Replace(pData.cFileName, Chr(0), ""))
            If sOrg <> "" Then
                ret = Recibir_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, "c:\balanzas\ia\salida\" & sOrg, 0)
                Do_Events
                Bucle = Bucle + ret + myret
                Sleep (500)
                ret = Borrar_Fichero_FTP("/home/pcscale/scale/update/" & sOrg, 0)
                Bucle = Bucle + ret
                Do_Events
                If Bucle > 1 Then
                    nCont = nCont + 1
                    GoTo masFich
                End If
            End If
    End If
    
    '        If Dir(App.Path & "\ESTRESUL\*.*") <> "" Then
    '            FileCopy App.Path & "\ESTRESUL\*.*", "c:\balanzas\ia\salida"
    '        End If

    '    Next Bucle
        
    
    'End If
    
salGetEst:
    If Err.Number <> 0 And nRetConnFTP = 1 Then
        ret = Desconectar_FTP(0)
        FTP_hOpen(0) = 0
        FTP_HandleConect(0) = 0
        nRetConnFTP = 0
    End If

    End
End Sub

Public Sub Main()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyBaseDES As DAO.Database
Dim MyRsDes As DAO.Recordset
Dim MyRsITM As DAO.Recordset
Dim nf As Integer
Dim nnF As Integer
Dim nnnF As Integer
Dim sPath() As String
Dim nCont As Integer
Dim sCfgSS(6) As String
Dim sGrpSS() As String
Dim sSubGrpSS() As String
Dim nGrp As Integer
Dim sS As String
Dim Token As Long
Dim myBase As DAO.Database
Dim myRS As DAO.Recordset
Dim cm As String
Dim ret As Long
'////////
Dim siP, susR, spsS As String
Dim nFich As Integer
Dim sOrg As String
'''''''''
Dim Bucle As Integer
Dim mySD As String

    If App.PrevInstance Then End
        
'Eliminar ficheros más antiguos de 45 dias de baarias
On Error GoTo Si1
    nf = FreeFile()
    Open App.Path & "\logtouchup.txt" For Output As #nf
    FrmAviso.Cls
    FrmAviso.File1.Path = App.Path & "\baarias"
    FrmAviso.File1.FileName = "*.*"
    FrmAviso.File1.Refresh
    If FrmAviso.File1.ListCount > 0 Then
        Print #nf, Format(Now, "dd/mm/yyyy hh:mm:ss")
        For Bucle = 0 To FrmAviso.File1.ListCount - 1
            Print #nf, FrmAviso.File1.List(Bucle) & " -- " & FileDateTime(App.Path & "\baarias\" & FrmAviso.File1.List(Bucle)) & " -- " & Abs(DateDiff("d", CDate(Format(Now, "dd/mm/yyyy")), Format(FileDateTime(App.Path & "\baarias\" & FrmAviso.File1.List(Bucle)), "dd/mm/yyyy")))
            
            
            If Abs(DateDiff("d", CDate(Format(Now, "dd/mm/yyyy")), Format(FileDateTime(App.Path & "\baarias\" & FrmAviso.File1.List(Bucle)), "dd/mm/yyyy"))) > 45 Then
                Print #nf, "Kill " & App.Path & "\baarias\" & FrmAviso.File1.List(Bucle)
                On Error Resume Next
                Kill App.Path & "\baarias\" & FrmAviso.File1.List(Bucle)
                On Error GoTo 0
            End If
        Next Bucle
    End If
Si1:
On Error GoTo 0
    Close #nf
    
    'Se supone única Balanza Servidora
    If Dir(App.Path & "\touchsrv.cfg") <> "" Then
        nf = FreeFile()
        Open App.Path & "\touchsrv.cfg" For Input As #nf
        Line Input #nf, sIpAdd
        Line Input #nf, sPort
        Close #nf
        sIpAdd = Trim(sIpAdd)
        sPort = Trim(sPort)
        nIpPort = Val(sPort)
    End If
    
    'Del fichero hiper se captura el número de Hiper Mercado (sHiper)
    If Dir(App.Path & "\hiper") <> "" Then
        nf = FreeFile()
        Open App.Path & "\hiper" For Input As #nf
        Line Input #nf, sHiper
        Line Input #nf, sIpCajas1
        Line Input #nf, sDirCajas1
        Line Input #nf, sUserCajas1
        Line Input #nf, sPassCajas1
        Close #nf
        sIpCajas1 = Trim(sIpCajas1)
        sHiper = Trim(sHiper)
        'sHiper = Format(Val(sHiper), "0000")
        sDirCajas1 = Trim(sDirCajas1)
        sUserCajas1 = Trim(sUserCajas1)
        sPassCajas1 = Trim(sPassCajas1)
    End If
        
'//////////// Test
    'siP = toIpAdd(sIpAdd)
    'susR = "pcscale"
    'spsS = "epelsa"
    'nFich = 0
    'LOG00FE001.B160913
    'ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
    'nRetConnFTP = ret
    'sOrg = UCase("log" & sHiper & "???" & ".b" & Format(Date, "yymmdd"))
    'ret = getfilesize("/home/pcscale/scale/update/" & "LOG00FE001.B160913", 0)
    'ret = Recibir_Fichero_FTP("/home/pcscale/scale/update/" & "LOG00FE001.B160913", App.Path & "\" & lastFFound, 0)
'Call waitResul
'////////////<----

    'Parámetros a tener en cuenta para los procesos.
    'If Trim(Command) <> "" Then
        
        'If LCase(Trim(Command)) = "update" Then
            'FrmAviso.LblAviso.Caption = "Envio Modificaciones a Touch-Scale"
            'FrmAviso.Show
            'FrmAviso.Refresh
            
'''''''''''''''''''''''''''''''''
        mySD = "990101"
        FrmAviso.Cls
        FrmAviso.File2.Path = sPathOrg
        FrmAviso.File2.FileName = "crt" & sHiper & "???" & ".b*" 'Format(Date, "yymmdd")
        FrmAviso.File2.Refresh
        If FrmAviso.File2.ListCount > 0 Then
            For Bucle = 0 To FrmAviso.File2.ListCount - 1
                'If CDate(Format("20" & Right(FrmAviso.File2.List(Bucle), 6), "yyyymmdd")) < CDate(Format(mySD, "yyyymmdd")) Then mySD = Right(FrmAviso.File2.List(Bucle), 6)
                'If FileDateTime(sPathOrg & "\" & FrmAviso.File2.List(Bucle)) < CDate(mySD) Then mySD = Right(FrmAviso.File2.List(Bucle), 6)
                If Val(Right(FrmAviso.File2.List(Bucle), 6)) < Val(mySD) Then mySD = Right(FrmAviso.File2.List(Bucle), 6)
            Next Bucle
        End If

'''''''''''''''''''''''''''''''''
            
            
            
            sendFiles mySD
            If lEspera Then
                'nf = FreeFile()
                'Open App.Path & UCase("\LOG" & sHiper & "001.B" & Format(Date, "yymmdd")) For Random As #nf
                'Close #nf
                'ret = Enviar_Fichero_FTP(App.Path & UCase("\LOG" & sHiper & "001.B" & Format(Date, "yymmdd")), "/home/pcscale/scale/update/" & UCase("LOG" & sHiper & "001.B" & Format(Date, "yymmdd")), 0)
                'Kill App.Path & UCase("\LOG" & sHiper & "001.B" & Format(Date, "yymmdd"))
                waitResul mySD
            Else
                Call GetEst
            End If
        'End If
        
        
        FrmAviso.Hide
        End
    'Else
    '
    '    End
    '
    'End If
    
End Sub

Public Function GeneraTextoIP(sData As String, Optional ByRef nPort As Long) As Boolean
    Dim Buf As String
    Dim Bucle As Long
    Dim nPuntos As Integer
    Dim esPunto As Boolean
    Dim MyOK As Boolean
    Dim nNum As Integer
    Dim Num(4) As String
    Dim sPort As String
    Dim lDosPuntos As Boolean
    
    Buf = sData
    MyOK = True
    lDosPuntos = False
    sPort = ""
    For Bucle = 1 To Len(Buf)
        Select Case Mid(Buf, Bucle, 1)
            Case "0" To "9"
                If lDosPuntos Then
                    sPort = sPort & Mid(Buf, Bucle, 1)
                Else
                    esPunto = False
                    Num(nNum) = Num(nNum) & Mid(Buf, Bucle, 1)
                End If
            Case "."
                If Not esPunto Then
                    nPuntos = nPuntos + 1
                    nNum = nNum + 1
                    esPunto = True
                Else
                    MyOK = False
                    Exit For
                End If
            Case ":"
                lDosPuntos = True
            Case Else
                MyOK = False
                Exit Function
        End Select
    Next Bucle
    If nPuntos <> 3 Then MyOK = False
    If IsNumeric(Num(0)) And IsNumeric(Num(1)) And IsNumeric(Num(2)) And IsNumeric(Num(3)) Then
        If (Val(Num(0)) < 256) And (Val(Num(1)) < 256) And (Val(Num(2)) < 256) And (Val(Num(3)) < 256) Then
            sData = Format(Val(Num(0)), "000") & "." & _
                    Format(Val(Num(1)), "000") & "." & _
                    Format(Val(Num(2)), "000") & "." & _
                    Format(Val(Num(3)), "000")
            
        Else
            MyOK = False
        End If
    Else
        MyOK = False
    End If
    
    If lDosPuntos Then
        nPort = (Val(sPort))
    End If
    GeneraTextoIP = MyOK
End Function


Attribute VB_Name = "Files"
Option Explicit
Private Do_Cancel As Long
'*****************************
'* Este módulo incorpora     *
'* las funciones de acceso   *
'* a ficheros del ocx        *
'*****************************
Public Function SuperLog(MiCad As String)
    '////////////////////////////////
    '// Esta función va generando  //
    '// el log de comunicaciones   //
    '// en un fichero llamado      //
    '// logepelcom.txt en la ruta  //
    '// donde esté situado el OCX  //
    '////////////////////////////////
    Dim Arch As Integer
    Dim Bucle As Integer
    Dim MiCad3 As String
    MiCad3 = ""
    For Bucle = 1 To Len(MiCad)
        MiCad3 = MiCad3 & Asc(Mid(MiCad, Bucle, 1)) & "|"
    Next Bucle

    Arch = FreeFile()
    On Error Resume Next
    Open PathLog For Append As #Arch
    If Err.Number = 0 Then
        Print #Arch, Now() & "-->" & MiCad3
        Close #Arch
    End If
    On Error GoTo 0
End Function
'////////////////////////////////
'// Esta función va generando  //
'// el log de comunicaciones   //
'// en un fichero llamado      //
'// logepelcom.txt en la ruta  //
'// donde esté situado el OCX  //
'////////////////////////////////
Public Function CadenadeLog(ByVal MiCad As String)
    Dim Arch As Integer
    Dim MyVAl As Long
    Dim MiCad2 As String
    Dim Bucle As Integer
    Dim MiCad3 As String
    Dim DoSuperLog As Boolean
    MiCad2 = ""
    MiCad3 = ""
    For Bucle = 1 To Len(MiCad)
        If Asc(Mid(MiCad, Bucle, 1)) < 32 Then
            DoSuperLog = True
            MiCad2 = MiCad2 & "·"
        Else
            MiCad2 = MiCad2 & Mid(MiCad, Bucle, 1)
        End If
    Next Bucle
    Arch = FreeFile()
    On Error Resume Next
    Open PathLog For Append As #Arch
    If Err.Number = 0 Then
        Print #Arch, Now() & "-->" & MiCad2
        Close #Arch
    End If
    If DoSuperLog Then SuperLog (Mid(MiCad, InStr(1, MiCad, Chr(2))))
    If Err.Number = 0 Then
        On Error Resume Next
        MyVAl = FileLen(PathLog)
        If MyVAl > 10048576 Then
            On Error Resume Next
            Kill PathLog
        End If
    End If
    On Error GoTo 0
End Function

'**********************************************************************************************
' Peticiones de datos de totales a ficheros
'**********************************************************************************************
'///////////////////////////////////////////////////////////////////////////////////////////////
Public Function F_Query_All_Tickets(lType As Long, Mode As Long, MyPath As String, Optional ByVal lPago As Boolean) As Long
    Dim Arch As Integer
    Dim Resp As Long
    Dim Resul As Long
    Dim MiTicket As Epel_CabTick
    Dim Buf1 As String
    Dim Bucle As Integer
    Dim MyFlag As Boolean
    Dim Berror As Long
    Dim NumTickGB As Long
    NumTickGB = 1
    If Dir(App.Path & "\cancel.tqt") <> "" Then
        On Error Resume Next
        Kill App.Path & "\cancel.tqt"
        On Error GoTo 0
    End If
    Do_Cancel = False
    Resp = 0
    Resul = 0
    '1 º checkea acceso al fichero
    If Mode = 1 Then
        MyFlag = True
    Else
        MyFlag = False
    End If
    Arch = FreeFile()
    On Error Resume Next
    Open MyPath For Append As #Arch
    If Err.Number = 0 Then
        On Error GoTo 0
        Do Until Resul <> 0
            For Berror = 1 To 3
                Select Case ScaleType
                    Case 10
                        '
                    Case 11
                        '
                    Case Else
                        Resul = 1010
                End Select
                If Resul = 0 Or Resul >= 1000 Then
                    If Resul >= 1000 Then Resp = 1000
                    Exit For
                End If
            Next Berror
            If Resul = 0 Then
                '*****************
                ' introduce datos de la cabecera
                '*****************
                NumTickGB = MiTicket.NTicket + 1
                Buf1 = "C"
                Buf1 = Buf1 & MiTicket.type & Format(MiTicket.NTicket, "00000")
                Buf1 = Buf1 & Format(MiTicket.Section, "00") & Format(MiTicket.Machine, "00")
                Buf1 = Buf1 & Format(MiTicket.Vendor, "0000") & Format(MiTicket.Client, "00000")
                Buf1 = Buf1 & MiTicket.date & MiTicket.Hour
                Buf1 = Buf1 & Format(MiTicket.nLines, "0000") & Format(MiTicket.Cancelled, "0")
                Buf1 = Buf1 & Format(MiTicket.positive, "0") & Format(MiTicket.amount, "000000000")
                Buf1 = Buf1 & Format(MiTicket.P1, "000000000") & Format(MiTicket.P2, "000000000")
                Buf1 = Buf1 & Format(MiTicket.Task, "0")
                On Error Resume Next
                Print #Arch, Buf1
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Resp = 1101
                    Exit Do
                Else
                    On Error GoTo 0
                    '*******************
                    ' introduce las líneas
                    '*******************
                    For Bucle = 0 To MiTicket.nLines - 1
                        Buf1 = "L"
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Number, "0000")
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).positive
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).Cancelled
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Code, "000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Price, "0000000")
                        Buf1 = Buf1 & MiTicket.Lines(Bucle).Weighted
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Pes, "000000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).amount, "000000000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).BovineCard, "00000")
                        Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).discount, "00000")
                        
                        On Error Resume Next
                        Print #Arch, Buf1
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            Resp = 1101
                            Exit For
                        Else
                            On Error GoTo 0
                            '******************
                            EnviaEvento.Raise_TicketReceived MiTicket.NTicket, MiTicket.Section, MiTicket.Machine, Do_Cancel
                            '*******************
                            If Dir(App.Path & "\cancel.tqt") <> "" Then
                                Do_Cancel = 4
                            End If
                            If Do_Cancel <> False Then
                                CadenadeLog "Recibida orden de cancelación de recepción de tiquets"
                                Resp = 1102
                                Exit For
                            End If
                        End If
                    Next Bucle
                    If Resp = 1101 Or Resp = 1102 Then Exit Do
                End If
            End If
        Loop
    Else
        On Error GoTo 0
        Resp = 1100
        Resul = 1100
    End If
    Close #Arch
    F_Query_All_Tickets = Resul
    If Resul >= 1000 Then
        EnviaEvento.Raise_ComError Resul
    Else
        EnviaEvento.Raise_ComOK
    End If
End Function

Public Function F_Query_All_Tickets2(lType As Long, MyPath As String, Optional ByVal lPago As Boolean) As Long
    '/////////////////////////////////////
    '// Esta función solicita todos     //
    '// los tiquets de SC10 para        //
    '// depositarlos en un fichero      //
    '// Funciona para numeración        //
    '// por sistema, sección o balanza  //
    '/////////////////////////////////////
    Dim Arch As Integer
    Dim Resp As Long
    Dim sBinSx As String
    Dim Resul As Long
    Dim MiTicket As Epel_CabTick
    Dim Buf1 As String
    Dim Bucle As Integer
    Dim MyOK As Boolean
    Dim MyModo As Long
    Dim MaxBucle As Integer
    Dim BGrande As Integer
    Dim BufNum As Long
    Dim Berror As Integer
    '************************
    Dim NSecs(99) As Long
    Dim NBals(99) As Long

    '************************
    If Dir(App.Path & "\cancel.tqt") <> "" Then
        On Error Resume Next
        Kill App.Path & "\cancel.tqt"
        On Error GoTo 0
    End If
    Do_Cancel = False
    Resp = 0
    Resul = 0
    Select Case TipoNumeracion
        Case 0 ' consecutiva
            MyModo = 0
        Case 1 ' por sección
            MyModo = 1
        Case 2 ' por balanza
            MyModo = 2
        Case Else
            Resp = 1010
    End Select
    '***************************************
    ' peticiones que se harán
    '***************************************
    Select Case MyModo
        Case 0
            MaxBucle = 1
    End Select
    '*********************************************
    '*********************************************
    '*********************************************
    '*********************************************
    Arch = FreeFile()
    On Error Resume Next
    Open MyPath For Append As #Arch
    If Err.Number = 0 Then
        On Error GoTo 0
        For BGrande = 1 To MaxBucle
            If (Resul <> 0 And Resul <> 4) Or Resp = 1101 Or Resp = 1102 Then
                Exit For
            End If
            Resul = 0
            MiTicket.NTicket = 0
            Do Until Resul <> 0
                BufNum = MiTicket.NTicket + 1
                For Berror = 1 To 3
                    If ScaleType > 2 Then
                        Resul = 1010
                    End If
                    If Resul = 0 Then
                        If MiTicket.NTicket < BufNum Then Resul = 4
                    End If
                    If Resul = 0 Or Resul >= 1000 Then Exit For
                Next Berror
                If Resul = 0 Then
                    '*****************
                    ' introduce datos de la cabecera
                    '*****************
                    Buf1 = "C"
                    Buf1 = Buf1 & MiTicket.type & Format(MiTicket.NTicket, "00000")
                    Buf1 = Buf1 & Format(MiTicket.Section, "00") & Format(MiTicket.Machine, "00")
                    Buf1 = Buf1 & Format(MiTicket.Vendor, "0000") & Format(MiTicket.Client, "00000")
                    Buf1 = Buf1 & MiTicket.date & MiTicket.Hour
                    Buf1 = Buf1 & Format(MiTicket.nLines, "0000") & Format(MiTicket.Cancelled, "0")
                    Buf1 = Buf1 & Format(MiTicket.positive, "0") & Format(MiTicket.amount, "000000000")
                    Buf1 = Buf1 & Format(MiTicket.P1, "000000000") & Format(MiTicket.P2, "000000000")
                    Buf1 = Buf1 & Format(MiTicket.Task, "0")
                    On Error Resume Next
                    Print #Arch, Buf1
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                        Resp = 1101
                        Exit Do
                    Else
                        On Error GoTo 0
                        '*******************
                        ' introduce las líneas
                        '*******************
                        For Bucle = 0 To MiTicket.nLines - 1
                            Buf1 = "L"
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Number, "0000")
                            Buf1 = Buf1 & MiTicket.Lines(Bucle).positive
                            Buf1 = Buf1 & MiTicket.Lines(Bucle).Cancelled
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Code, "000000")
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Price, "0000000")
                            Buf1 = Buf1 & MiTicket.Lines(Bucle).Weighted
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).Item_Pes, "000000000")
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).amount, "000000000")
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).BovineCard, "00000")
                            Buf1 = Buf1 & Format(MiTicket.Lines(Bucle).discount, "00000")
                            On Error Resume Next
                            Print #Arch, Buf1
                            If Err.Number <> 0 Then
                                On Error GoTo 0
                                Resp = 1101
                                Exit For
                            Else
                                On Error GoTo 0
                                EnviaEvento.Raise_TicketReceived MiTicket.NTicket, MiTicket.Section, MiTicket.Machine, Do_Cancel
                                '*******************
                                If Dir(App.Path & "\cancel.tqt") <> "" Then
                                    Do_Cancel = 4
                                End If
                                If Do_Cancel <> False Then
                                    CadenadeLog "Recibida orden de cancelación de recepción de tiquets"
                                    Resp = 1102
                                    Exit For
                                End If
                            End If
                        Next Bucle
                        If Resp = 1101 Or Resp = 1102 Then Exit Do
                    End If
                End If
            Loop
        Next BGrande
        Resp = Resul
    Else
        On Error GoTo 0
        Resp = 1100
    End If
    Close #Arch
    F_Query_All_Tickets2 = Resp
    If Resp >= 1000 Then
        EnviaEvento.Raise_ComError Resp
    Else
        EnviaEvento.Raise_ComOK
    End If
End Function

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

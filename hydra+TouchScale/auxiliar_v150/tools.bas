Attribute VB_Name = "tools"
Option Explicit
Public Type Tipo_TqCodBar
    Tipo As Integer
    NBalanza As Integer
    NMostrador As Integer
    NTicket As Long
    NVendedor As Long
    Importe As Double
End Type
Public DescargaReport As Boolean
Public Function AbreRecordset(sQuery As String) As ADODB.Recordset
    Dim MyReg As New ADODB.Recordset
    Dim sCadConn As String
    Dim Continuar As Boolean
    sCadConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
    Do Until Continuar
        'On Error Resume Next
        MyReg.Open sQuery, sCadConn
        If Err.Number = 0 Then Continuar = True
    Loop
    On Error GoTo 0
    Set AbreRecordset = MyReg
End Function
Public Function ToDouble(sData As String) As Double
    Dim Valor As Double
    Dim Valor2 As Double
    Dim Contador As Integer
    Dim Encontrado As Boolean
    Dim Bucle As Integer
    Dim EsNegativo As Boolean
    For Bucle = 1 To Len(sData)
        Select Case Mid(sData, Bucle, 1)
            Case 0 To 9
                If Not Encontrado Then
                    Valor = Valor * 10
                    Valor = Valor + Val(Mid(sData, Bucle, 1))
                Else
                    Contador = Contador + 1
                    Valor2 = Valor2 * 10
                    Valor2 = Valor2 + Val(Mid(sData, Bucle, 1))
                End If
            Case "."
                Encontrado = True
            Case ","
                Encontrado = True
            Case "-"
                EsNegativo = True
        End Select
    Next Bucle
    Valor2 = Valor2 / (10 ^ Contador)
    Valor = Valor + Valor2
    If EsNegativo Then Valor = Valor * (-1)
    ToDouble = Valor
End Function
Private Function Ado2Dao(Valor As Long) As Long
    Select Case Valor
        Case 200
            Ado2Dao = 10
        Case 202
            Ado2Dao = 10
        Case 11
            Ado2Dao = 1
        Case 7
            Ado2Dao = 8
        Case 2
            Ado2Dao = 3
        Case 3
            Ado2Dao = 4
        Case 5
            Ado2Dao = 7
        Case Else
            MsgBox Valor
    End Select
End Function
'\\\\\\\\\\\
'///////////
Public Sub CreaTablaPropia(Nombre As String, RegNet As ADODB.Recordset)
Dim Base As DAO.Database
Dim Bucle As Integer
Dim Tabla As TableDef
Dim Registro As DAO.Recordset
    Set Base = OpenDatabase(App.Path & "\dirtmptouch\tmp.mdb")
    With Base
        If .TableDefs.Count > 0 Then
            For Bucle = 0 To .TableDefs.Count - 1
                If UCase(.TableDefs(Bucle).Name) = UCase(Nombre) Then
                    .TableDefs.Delete (Nombre)
                    Exit For
                End If
            Next Bucle
        End If
        Set Tabla = .CreateTableDef(Nombre)
        For Bucle = 0 To RegNet.Fields.Count - 1
            Tabla.Fields.Append Tabla.CreateField(RegNet.Fields(Bucle).Name, Ado2Dao(RegNet.Fields(Bucle).type), 255)
        Next Bucle
        For Bucle = 0 To Tabla.Fields.Count - 1
                Tabla.Fields(Bucle).AllowZeroLength = True
        Next Bucle
        .TableDefs.Append Tabla
        .Close
    End With
    Set Base = Nothing
    Set Base = OpenDatabase(App.Path & "\dirtmptouch\tmp.mdb")
    Set Registro = Base.OpenRecordset(Nombre)
    With RegNet
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Registro.AddNew
                For Bucle = 0 To RegNet.Fields.Count - 1
                    Registro.Fields(Bucle) = .Fields(Bucle)
                Next Bucle
                Registro.Update
                .MoveNext
            Loop
        End If
    End With
    Base.Close
    Set Base = Nothing
End Sub
'\\\\\\\\\\\
'///////////
Public Sub CargaIdiomas()
Dim Registro As New ADODB.Recordset
Dim Bucle As Long
Dim MyID As Integer
Dim Arch As Integer
Dim Buf As String
Dim Buf2 As String
    If Dir(App.Path & "\hydramulti.ini") <> "" And Dir(App.Path & "\hydramulti.exe") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\hydramulti.ini" For Input As Arch
        Input #Arch, Buf, Buf2
        Input #Arch, Buf, Buf2
        MyID = Val(Buf) + 1
        Close #Arch
    Else
        MyID = (id / 1000)
        MyID = MyID + 1
    End If
    Registro.Open "select max(id) from texts", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\langtouch\langtouch.mdb" & ";Persist Security Info=False"
    With Registro
        nIdioma = Val(.Fields(0))
        ReDim Preserve sIdioma(nIdioma + 1)
    End With
    Registro.Close
    Registro.Open "select * from texts", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\langtouch\langtouch.mdb" & ";Persist Security Info=False"
    With Registro
        .MoveFirst
        Do Until .EOF
            If Not IsNull(.Fields("text" & Format(MyID, "0"))) Then
                sIdioma(.Fields("id")) = .Fields("text" & Format(MyID, "0"))
            End If
            .MoveNext
        Loop
    End With
End Sub
'//////
'//////
'//////
Public Function CargaCadena(ByVal NumCad As Integer) As String
    If NumCad <= nIdioma Then
        CargaCadena = sIdioma(NumCad)
    Else
        CargaCadena = "Cad : " & NumCad
    End If
End Function
'///////////
'///////////
'///////////
Public Function DrawGradient(TheObj As Object, ColorRed%, ColorGreen%, _
                               ColorBlue%, ColorStop%, ColorBandSize%, _
                               StartLine%, StopLine%, ModLine%, _
                               Optional AutoReDrawObj As Boolean = True, _
                               Optional MoreR% = 1, Optional MoreG% = 1, _
                               Optional MoreB% = 1)
  On Error Resume Next

    Dim sngBlueCur As Single, sngRedCur As Single, sngGreenCur As Single
    Dim sngBlueStep As Single, sngRedStep As Single, sngGreenStep As Single
    Dim intFormHeight As Integer, intFormWidth As Integer, intY As Integer
    If MoreR% <= 0 Then MoreR% = 1
    If MoreG% <= 0 Then MoreG% = 1
    If MoreB% <= 0 Then MoreB% = 1
    
    If MoreR% >= 20 Then MoreR% = 20
    If MoreG% >= 20 Then MoreG% = 20
    If MoreB% >= 20 Then MoreB% = 20
    
    Select Case AutoReDrawObj
      Case True
        TheObj.AutoRedraw = True
      Case False
        TheObj.AutoRedraw = False
    End Select
    intFormHeight = TheObj.ScaleHeight
    intFormWidth = TheObj.ScaleWidth

    sngRedStep = ColorBandSize% * (ColorStop% - ColorRed%) / intFormHeight
    sngRedCur = ColorRed%
    '
    sngGreenStep = ColorBandSize% * (ColorStop% - ColorGreen%) / intFormHeight
    sngGreenCur = ColorGreen%
    '
    sngBlueStep = ColorBandSize% * (ColorStop% - ColorBlue%) / intFormHeight
    sngBlueCur = ColorBlue%

    For intY = StopLine% To StartLine% Step -1 * ColorBandSize%
        TheObj.Line (-1, intY - 1)-(intFormWidth, intY + ColorBandSize% _
                    \ ModLine%), RGB(sngRedCur \ MoreR%, sngGreenCur \ MoreG% _
                    , sngBlueCur \ MoreB%), BF

        sngBlueCur = sngBlueCur + sngBlueStep
        sngRedCur = sngRedCur + sngRedStep
        sngGreenCur = sngGreenCur + sngGreenStep
    Next intY
    
End Function


Public Function calcheck(ByVal micade As String, check)

Dim i
Dim buffercheck(255)
Dim longitud_cade

longitud_cade = Len(micade)

check = 0

For i = 1 To longitud_cade Step 1
    buffercheck(i) = Mid(micade, i, 1)

    If buffercheck(i) = Chr$(126) Then
        buffercheck(i) = Chr$(0)
    End If
    check = check Xor Asc(buffercheck(i))
Next i

check = check Xor (longitud_cade + 2)
check = ((check And 63) Or 64)
End Function



Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    On Error GoTo fin
    Fichero = FreeFile
    Open MiruTa & "\loghydra.txt" For Append As Fichero
    On Error GoTo 0
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
fin:
    On Error GoTo 0
End Sub

Public Sub SALIRPROGRAMA(Optional Forzoso As Boolean)
   Dim Arch As Integer
   Dim logntq As Integer
   Dim Bucle As Integer
   On Error Resume Next
   If Dir(App.Path & "\laststart") <> "" Then
    Kill App.Path & "\laststart"
   End If
   
   If SoyReserva Then
        If Dir(App.Path & "\alarmas.exe") <> "" Then
            Arch = FreeFile()
            Open App.Path & "\dirtmptouch\alarmas.off" For Output As #Arch
            Close #Arch
        End If
   End If
   'If DataEnvironment1.Connection1.State <> 0 Then DataEnvironment1.Connection1.Close
   'If DataEnvironment1.Connection2.State <> 0 Then DataEnvironment1.Connection2.Close
   'If DataPlugin.Mode <> "" And DataPlugin.type <> "" Then
   '     If Dir(MiruTa & "\dirtmptouch") <> "" Then
   '         FrmHlink.STOPALL
   '     End If
   'End If
   If Dir(MiruTa & "\dirtmptouch\tmp.mdb") Then
     Kill MiruTa & "\dirtmptouch\tmp.mdb"
   End If
   
   
   On Error GoTo 0
   '********************************
   ' Compacta la base y crea copia
   ' de seguridad automática
   '********************************
   If Not Forzoso Then
    
    On Error GoTo siguiente
    'If Dir(MiruTa & "\autobackup.mdb") <> "" Then
    '     Kill MiruTa & "\autobackup.mdb"
    'End If
    'DBEngine.CompactDatabase MiruTa & "\" & "dbasetouch.mdb", MiruTa & "\" & "autobackup.mdb"
    'If Dir(MiruTa & "\" & "dbasetouch.mdb") <> "" Then
    '     FileCopy MiruTa & "\" & "autobackup.mdb", MiruTa & "\" & "dbasetouch.mdb"
    'End If
   End If
   '********************************
   ' Crea log de seguridad
   '********************************
siguiente:
   logntq = FreeFile
   If Dir(MiruTa & "\loghydra.txt") <> "" Then
        Open MiruTa & "\loghydra.txt" For Append As logntq
   Else
        Open MiruTa & "\loghydra.txt" For Output As logntq
   End If
   Print #logntq, Now() & " Program Finished ************************************"
   Close logntq
   
   If Forzoso Then MsgBox "ERROR CRÍTICO : NO SE PUDO ACCEDER A LA BASE DE DATOS. EL PROGRAMA ESTÁ DETENIDO", vbCritical, "HYDRATOUCH"
   End
End Sub
Public Sub Do_Events()
    Sleep (1)
    DoEvents
End Sub


Public Sub LogTXT(ByVal cadena As String)
Dim Bucle As Long
Dim Archivo As Integer
Archivo = FreeFile
Open "c:\log.txt" For Output As Archivo
Write #Archivo, cadena
For Bucle = 1 To Len(cadena)
    Print #Archivo, "(" & Bucle & ")" & Mid(cadena, Bucle, 1) & "," & Asc(Mid(cadena, Bucle, 1)) & "|"
Next Bucle
Close #Archivo
End Sub




Public Function versionar() As String
    versionar = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

Public Sub obtenPATHngt(ByVal s As String)
    Dim l As Integer
    Dim nCont As Integer
    Dim ss As String
    
    l = Len(s)
    For nCont = l To 1 Step -1
        If Mid(s, nCont, 1) = "\" Then Exit For
    
    Next nCont
    sPathNGT = Mid(s, 1, nCont)
    If UCase(Mid(sPathNGT, Len(sPathNGT) - 2)) = "\S\" Then
        sPathNGT = Mid(sPathNGT, 1, Len(sPathNGT) - 2)
    End If
End Sub

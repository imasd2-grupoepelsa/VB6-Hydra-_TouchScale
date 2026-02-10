Attribute VB_Name = "Interfaz"
Option Explicit
Public EstiloInterfaz As Integer

Private Declare Function GetSystemMenu Lib "user32" _
                                       (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" _
                                    (ByVal hMenu As Long, ByVal nPosition As Long, _
                                     ByVal wFlags As Long) As Long
Const SC_SIZE = &HF000
Const SC_MOVE = &HF010
Const SC_MINIMIZE = &HF020
Const SC_MAXIMIZE = &HF030
Const SC_CLOSE = &HF060
Const SC_RESTORE = &HF120
Const MF_SEPARATOR = &H800
Const MF_BYPOSITION = &H400
Const MF_BYCOMMAND = &H0

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
Public Sub MostrarModal(Formulario As Form)
    frmEpelsa.Enabled = False
    On Error Resume Next
    Formulario.Show
    On Error GoTo 0
End Sub
Public Sub MostrarFrmLabel(Formulario As Form)
    FrmArticulo2.Enabled = False
    Formulario.Show
End Sub

Public Sub MonitorBarra()
    On Error Resume Next
    frmMonitor.barra1.Value = frmMonitor.barra1.Value + 1
    On Error GoTo 0
End Sub
Public Sub Aux_Refresca_MostradoresTS(MICombo As ComboBox)
    Dim Base As Database
    Dim Regmostrador As Recordset
    Dim RegSeccion As Recordset
    Dim Buffer As String
    Dim numero As Integer
    Dim Contador As Long
    MICombo.Clear
    Set Base = AbrirBase()
    Set RegSeccion = Base.OpenRecordset _
                     ("select distinct secc_maqui,borrado from seccion where borrado=false order by secc_maqui")
    With RegSeccion
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                numero = 0
                Set Regmostrador = Base.OpenRecordset _
                                   ("select codi_ident,secc_maqui,borrado from seccion where borrado=false and secc_maqui=" & _
                                    !secc_maqui)
                With Regmostrador
                    If Not .EOF Then
                        Do Until .EOF
                            Buffer = Buffer & !codi_ident & " , "
                            Contador = Contador + 1
                            .Movenext
                        Loop
                        Buffer = left(Buffer, Len(Buffer) - 3)
                        MICombo.AddItem Buffer
                    End If
                End With
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub

Public Sub Aux_Refresca_equipos(MICombo As ComboBox, Optional forzar As Boolean)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    MICombo.Clear
    Set Base = OpenDatabase(Base_General)
    'If usaSC10 And TipoC9 And (Not forzar) Then
    '    Set Registro = Base.OpenRecordset("select codi_ident,numero_eqp from equipos where borrado=false and (modelo<100 or modelo=400) order by numero_eqp")
    'Else
    Set Registro = Base.OpenRecordset("select codi_ident,numero_eqp from equipos where borrado=false and (modelo<200 or modelo=400) order by numero_eqp")
    'End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If HaySeguridad And UsuarioActual.Asociado <> 0 Then
                    If .Fields("codi_ident") = UsuarioActual.Asociado Then
                        MICombo.AddItem .Fields("numero_eqp")
                    End If
                Else
                    MICombo.AddItem .Fields("numero_eqp")
                End If
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
Public Sub MostrarInterfaz()
    Dim hWnd As Long
    Dim hMenu As Long
    Dim nC As Integer
    hWnd = frmEpelsa.hWnd
    hMenu = GetSystemMenu(hWnd, 0)
    If EstiloInterfaz = 1 Then
        If frmEpelsa.Visible = False Then

            frmEpelsa.top = 0
            frmEpelsa.left = 0
            frmEpelsa.Width = Screen.Width
            frmEpelsa.Height = 540
            DeleteMenu hMenu, SC_SIZE, MF_BYCOMMAND
            DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND
            DeleteMenu hMenu, SC_MINIMIZE, MF_BYCOMMAND
        End If
    End If
    frmEpelsa.Show
    If IconificarInicio Then
        For nC = 1 To 100
            Do_Events
        Next nC
    End If
End Sub
Public Sub Aux_Refresca_Mostradores(MICombo As ComboBox)
    Dim Base As dao.Database
    Dim Regmostrador As dao.Recordset
    Dim RegSeccion As dao.Recordset
    Dim Buffer As String
    Dim numero As Integer
    Dim Contador As Long
    Dim strSeguridad As String
    strSeguridad = ""
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = " and codi_ident=" & UsuarioActual.Asociado
    End If
    MICombo.Clear
    Set Base = OpenDatabase(Base_General)
    Set RegSeccion = Base.OpenRecordset _
                     ("select distinct secc_maqui,borrado from seccion where borrado=false order by secc_maqui")
    With RegSeccion
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Buffer = ""
                numero = 0
                Set Regmostrador = Base.OpenRecordset _
                                   ("select codi_ident,secc_maqui,borrado from seccion where borrado=false " & _
                                  " and enviardatos<>" & Chr(34) & "ECOP" & Chr(34) & " and secc_maqui=" & _
                                    .Fields("secc_Maqui") & strSeguridad)
                With Regmostrador
                    If Not .EOF Then
                        Do Until .EOF
                            Buffer = Buffer & !codi_ident & " , "
                            Contador = Contador + 1
                            .Movenext
                        Loop
                        Buffer = left(Buffer, Len(Buffer) - 3)
                        MICombo.AddItem Buffer
                    End If
                End With
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
'''''''
Public Sub Aux_Refresca_Mostradores_Con_Familias(MICombo As ComboBox)
    Dim Base As dao.Database
    Dim RegSeccion As dao.Recordset
    Dim strSeguridad As String

    strSeguridad = ""
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = " and codi_ident=" & UsuarioActual.Asociado
    End If
    MICombo.Clear
    Set Base = OpenDatabase(Base_General)
    Set RegSeccion = Base.OpenRecordset _
                     ("select distinct secc_maqui,borrado from familias where borrado=false" & strSeguridad & " order by secc_maqui")
    With RegSeccion
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                MICombo.AddItem CStr(!secc_maqui)
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub
'''''''''''''''''
Public Sub Aux_Refresca_Familias_Mostrador(MICombo As ComboBox, ByVal sSec As String)
    Dim Base As dao.Database
    Dim RegSeccion As dao.Recordset
    Dim strSeguridad As String
    strSeguridad = ""
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        strSeguridad = " and codi_ident=" & UsuarioActual.Asociado
    End If
    MICombo.Clear
    Set Base = OpenDatabase(Base_General)
    Set RegSeccion = Base.OpenRecordset _
                     ("select codi_fam from familias where borrado=false and secc_maqui=" & Trim(sSec) & strSeguridad & " order by codi_fam")
    With RegSeccion
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                MICombo.AddItem CStr(!codi_fam)
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub

'''''''
Public Sub Aux_Ejecuta(Que_Ejecuta As String, Optional Parametros As String)
    Dim MyProc As Long
    Dim MyResul As Long
    Dim Arch As Integer
    Dim BufEnable As Boolean
    Dim NombrePrograma As String
    NombrePrograma = "auxiliarTouch.exe"
    'If LCase(Que_Ejecuta) = "cop5" Then NombrePrograma = "OLDCOP5.EXE"
    'If LCase(Que_Ejecuta) = "mng" Then NombrePrograma = "OLDCOP5.EXE"
    'If LCase(Que_Ejecuta) = "iv4com" Then NombrePrograma = "OLDCOP5.EXE"
    BufEnable = frmEpelsa.Enabled
    frmEpelsa.Enabled = False
    'If Dir(App.Path & "\bin", vbDirectory) <> "" Then
    If Dir(App.Path & "\" & NombrePrograma) <> "" Then
        Arch = FreeFile()
        Open App.Path & "\auxiliartouch.tmp" For Output As #Arch
        Print #Arch, Base_General
        Print #Arch, Hydra_INI
        Print #Arch, Que_Ejecuta
        Print #Arch, Parametros
        If HaySeguridad Then
            Print #Arch, "Y"
        Else
            Print #Arch, "N"
        End If
        Print #Arch, UsuarioActual.Nombre
        Print #Arch, UsuarioActual.Asociado
        Close #Arch
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Miruta & "\" & NombrePrograma, vbNormalFocus))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
    Else
        MsgBox "auxiliartouch.exe not found. Can't execute order", vbCritical
    End If
    'Else
    '    MsgBox "auxiliar.exe not found. Can't execute order", vbCritical
    'End If
    If Dir(App.Path & "\auxiliartouch.tmp") <> "" Then
        On Error Resume Next
        Kill App.Path & "\auxiliartouch.tmp"
        On Error GoTo 0
    End If
    frmEpelsa.Enabled = BufEnable
End Sub
'''''''
'''''''
Public Sub exe_copion(Que_Ejecuta As String, Optional Parametros As String)
    Dim MyProc As Long
    Dim MyResul As Long
    Dim Arch As Integer
    Dim BufEnable As Boolean
    Dim NombrePrograma As String
    NombrePrograma = "copiontouch.exe"
    BufEnable = frmEpelsa.Enabled
    frmEpelsa.Enabled = False
    If Dir(App.Path & "\" & NombrePrograma) <> "" Then
        Arch = FreeFile()
        Open App.Path & "\fifo" For Output As #Arch
        Print #Arch, Base_General
        Print #Arch, Hydra_INI
        Print #Arch, Que_Ejecuta
        Print #Arch, Parametros
        If HaySeguridad Then
            Print #Arch, "Y"
        Else
            Print #Arch, "N"
        End If
        Print #Arch, UsuarioActual.Nombre
        Print #Arch, UsuarioActual.Asociado
        Close #Arch
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Miruta & "\" & NombrePrograma, vbNormalFocus))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
        'Else
        'MsgBox "copiontouch.exe not found. Can't execute order", vbCritical
    End If
    If Dir(App.Path & "\fifo") <> "" Then
        On Error Resume Next
        Kill App.Path & "\fifo"
        On Error GoTo 0
    End If
    frmEpelsa.Enabled = BufEnable
End Sub
'''''''
'''''''
Public Sub aux_ejecuta_historico(Que_Ejecuta As String, Parametros As String, ByRef adb As String)
    Dim MyProc As Long
    Dim MyResul As Long
    Dim Arch As Integer
    Dim BufEnable As Boolean
    Dim NombrePrograma As String
    NombrePrograma = "auxiliartouch.exe"
    'If LCase(Que_Ejecuta) = "cop5" Then NombrePrograma = "OLDCOP5.EXE"
    'If LCase(Que_Ejecuta) = "mng" Then NombrePrograma = "OLDCOP5.EXE"
    'If LCase(Que_Ejecuta) = "iv4com" Then NombrePrograma = "OLDCOP5.EXE"
    BufEnable = frmEpelsa.Enabled
    frmEpelsa.Enabled = False
    If Dir(App.Path & "\" & NombrePrograma) <> "" Then
        Arch = FreeFile()
        Open App.Path & "\auxiliartouch.tmp" For Output As #Arch
        Print #Arch, adb
        Print #Arch, Hydra_INI
        Print #Arch, Que_Ejecuta
        Print #Arch, Parametros
        If HaySeguridad Then
            Print #Arch, "Y"
        Else
            Print #Arch, "N"
        End If
        Print #Arch, UsuarioActual.Nombre
        Print #Arch, UsuarioActual.Asociado
        Close #Arch
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Miruta & "\" & NombrePrograma, vbNormalFocus))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
    Else
        MsgBox "auxiliartouch.exe not found. Can't execute order", vbCritical
    End If
    If Dir(App.Path & "\auxiliartouch.tmp") <> "" Then
        On Error Resume Next
        Kill App.Path & "\auxiliartouch.tmp"
        On Error GoTo 0
    End If
    frmEpelsa.Enabled = BufEnable
End Sub
'''''''
'''''''


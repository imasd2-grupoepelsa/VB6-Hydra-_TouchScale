Attribute VB_Name = "Tools"
Option Explicit
Public DescargaForm As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public UsaEco2 As Boolean


Public Sub Do_Events()
    Sleep (1)
    DoEvents
End Sub
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
Public Function CheckTCPIP(ByVal LaIP As String, ByVal ElPuerto As Long) As Integer
    '************************
    ' devuelve
    ' 0 --> OK
    ' 1 --> IP incorrecta
    ' 2 --> Puerto Incorrecto
    ' 3 --> Dirección IP y Puerto incorrectos
    Dim bUF As String
    Dim Contador As String
    Dim Resp As Integer
    Dim Bucle As Integer
    '***********************************
    ' chequea dirección IP correcta
    '***********************************
    Resp = 0
    Contador = 0
    bUF = ""
    LaIP = LaIP & "."
    For Bucle = 1 To Len(LaIP)
        If IsNumeric(Mid(LaIP, Bucle, 1)) Then
            bUF = bUF & Mid(LaIP, Bucle, 1)
        Else
            If Mid(LaIP, Bucle, 1) = "." Or Bucle = Len(LaIP) Then
                If Len(bUF) > 0 Then
                    Contador = Contador + 1
                    If Val(bUF) >= 0 Or Val(bUF) <= 255 Then
                        bUF = ""
                    Else
                        Resp = 1
                        Exit For
                    End If
                Else
                    Resp = 1
                    Exit For
                End If
            Else
                Resp = 1
                Exit For
            End If
        End If
    Next Bucle
    If Contador <> 4 Then Resp = 1
    If ElPuerto < 1024 Or ElPuerto > 65535 Then Resp = Resp + 2
    CheckTCPIP = Resp
End Function

Public Sub Aux_Ejecuta(Que_Ejecuta As String, Base_Name As String, Optional Parametros As String, Optional lScalePos As Boolean)
    Dim MyProc As Long
    Dim MyResul As Long
    Dim Arch As Integer
    Dim BufEnable As Boolean
    '****************************
    'de momento
    Dim HaySeguridad As Boolean
    
    'Atención Modificar para AuxiliarTouch.exe ScalePos 1.0.8
    
    '*****************************
    'If lScalePos = False Then
    ''If Dir(App.Path & "\bin", vbDirectory) <> "" Then
    '    If Dir(App.Path & "\auxiliartouch.exe") <> "" Then
    '        Arch = FreeFile()
    '        Open App.Path & "\auxiliartouch.tmp" For Output As #Arch
    '        Print #Arch, Base_Name & "\dbasetouch.mdb"
    '        Print #Arch, Base_Name & "\hydratouch.ini"
    '        Print #Arch, Que_Ejecuta
    '        Print #Arch, Parametros
    '        If HaySeguridad Then
    '            Print #Arch, "Y"
    '        Else
    '            Print #Arch, "N"
    '        End If
    '        Print #Arch, UsuarioActual.Nombre
     '       Print #Arch, UsuarioActual.Asociado
    '        Close #Arch
    '        MyResul = STILL_ACTIVE
    '        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiRuta & "\auxiliar.exe", vbNormalFocus))
    '        Do While MyResul = STILL_ACTIVE
    '            GetExitCodeProcess MyProc, MyResul
    '            Do_Events
    '        Loop
    '    Else
    '        MsgBox "auxiliar.exe not found. Can't execute order", vbCritical
    '    End If
    ''Else
    ''    MsgBox "auxiliar.exe not found. Can't execute order", vbCritical
    ''End If
    'Else
        If Dir(App.Path & "\auxiliartouch.exe") <> "" Then
            Arch = FreeFile()
            Open App.Path & "\auxiliartouch.tmp" For Output As #Arch
            Print #Arch, Base_Name & "\dbasetouch.mdb"
            Print #Arch, Base_Name & "\hydratouch.ini"
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
            MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiRuta & "\auxiliartouch.exe", vbNormalFocus))
            Do While MyResul = STILL_ACTIVE
                GetExitCodeProcess MyProc, MyResul
                Do_Events
            Loop
        Else
            MsgBox "auxiliartouch.exe not found. Can't execute order", vbCritical
        End If
    
    'End If
    
End Sub
Public Sub MostrarModal(Padre As Form, Hijo As Form)
    DescargaForm = False
    Padre.Enabled = False
    Hijo.Show
    Do Until DescargaForm = True
        Do_Events
    Loop
    Padre.Enabled = True
    Padre.SetFocus
End Sub

Public Function Crea_Grp()
Dim dB As Database
Dim tdf As TableDef
Dim Rst As Recordset
Dim Base_General As String

    If Dir(App.Path & "\tiendas.mdb") = "" Then
        Set dB = CreateDatabase(App.Path & "\tiendas.mdb", dbLangGeneral, dbVersion30)
        dB.Close
        Set dB = OpenDatabase(App.Path & "\tiendas.mdb")
        Set tdf = dB.CreateTableDef("grupos")
        With tdf
            .Fields.Append .CreateField("Tienda", dbLong)
            .Fields.Append .CreateField("Grp", dbLong)
            .Fields.Append .CreateField("Fichero", dbText)
            .Fields.Append .CreateField("Patron", dbText)
            .Fields.Append .CreateField("Path", dbText)
            '.Fields.Append .CreateField("ScalePos", dbBoolean) 'scalepos 1.0.8
            dB.TableDefs.Append tdf
        End With
        dB.Close
    
    End If
End Function

'ver si la función que se llama tiene permisos antes de añadir al menú
' Comunicaciones
Public Function mnuCTest(Index As Integer, ByRef nCont As Integer)
    Dim Permiso As Integer
    
    Select Case Index
        Case 0 'envio de modificaciones
            Permiso = 13
        Case 12 'grandes totales.
            Permiso = 14
        Case 13 'base maestra.
            Permiso = 11
        Case 16 'recibir tiques.
            Permiso = 14
        Case 17 'rec totales + env. modificaciones
            Permiso = 14
        Case 18 'recoger pedido e inventario.
            Permiso = 14
    End Select
    If nMyPer = 65535 Then
        If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
            If ConsultarPermiso(Permiso) = True Then
                nMyPer = Permiso
                nCont = nCont + 1
            End If
        Else
            nMyPer = Permiso
            nCont = nCont + 1
        End If
    End If
    mnuCTest = nCont
End Function

'ver si la función que se llama tiene permisos antes de añadir al menú
' ARTÍCULOS
Public Function mnuTTest(Index As Integer, ByRef nCont As Integer)
    Dim Permiso As Integer
    Select Case Index
        Case 0 'mostradores y equipos.
            Permiso = 0
        Case 1 'códigos de Barras.
            Permiso = 7
        Case 2 'Cabeceras y leyendas.
            Permiso = 6
        Case 3 'Publicidad y ofertas.
            Permiso = 8
        Case 4
            Permiso = 6
            
        Case 6 'entrada en artículos...
            Permiso = 1
        Case 7 'cambios de precio...
            Permiso = 2
        Case 8 'entrada en teclas...
            Permiso = 10
        Case 9 'envio de modificaciones...
        'If Dir(App.Path & "\gigante.txt") <> "" Then
            Permiso = 13
        'End If
        Case 12 'Vendedores
            Permiso = 9
        Case 18 'Familias
            Permiso = 0
        Case 10
            Permiso = 5
        Case 11
            Permiso = 9
        Case 14
            Permiso = 3
        Case 15
            Permiso = 3
        Case 16
            If (Not UsaEco2) Then
                nMyPer = 65535
                mnuTTest = nCont
                Exit Function
            End If
        Case 17
            If (Not UsaEco2) Then
                nMyPer = 65535
                mnuTTest = nCont
                Exit Function
            End If
    
    End Select
    If nMyPer = 65535 Then
        If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
            If ConsultarPermiso(Permiso) = True Then
                nMyPer = Permiso
                nCont = nCont + 1
            End If
        Else
            nMyPer = Permiso
            nCont = nCont + 1
        End If
    End If
    mnuTTest = nCont
End Function


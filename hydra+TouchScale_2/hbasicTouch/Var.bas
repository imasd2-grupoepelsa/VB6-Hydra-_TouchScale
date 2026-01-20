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
Public DbMysqli As ADODB.Connection
Public DbTotales As ADODB.Connection
Public lMySqlAbierta As Boolean
Public lTotalesAbierta As Boolean
Public lComunica As Boolean
Public sIpAdd As String
Public nIpPort As Integer
Public sPathHydra As String
Public lAllItem As Boolean
Public nSecSer As Integer
Public lSalir As Boolean
Public lTodo As Boolean
'////////////////////////
Public Type RegArt
    AB As String
    BOT As String
    FAM As String
    POS1 As String
    POS2 As String
    EAN As String
    DES As String
    CON As String
    Img As String
    COD As String
    PLU As String
End Type

Public Type teclas
    Code As Double
    Counter As Integer
    Section As Integer
    PLU As Integer
    Group As Integer
    SubGroup As Integer
    ShortName As String
    Name As String
    Key As Integer
    icon_name As String
    qLevel As Integer
    qLevelName As String
End Type
Public Type SubGrupo
    Grupo As Integer
    Indice As Integer
    ImgSelect As String
    ImgNoSelect As String
    lblTxt As String
    IniKey As Integer
    EndKey As Integer
    teclas(20) As teclas
End Type
Public Type Grupo
    Indice As Integer
    ImgSelect As String
    ImgNoSelect As String
    lblTxt As String
    SubGroupnr As Integer
    IniKey As Integer
    EndKey As Integer
    SubGrupo() As SubGrupo
End Type
Public Type TecGrupos
    Indice() As Grupo
End Type

Public TTecGrupos As TecGrupos
Public nGrupos As Integer
Public nGrupoActual As Integer 'ídice del grupo
Public nSubGrupoActual As Integer 'indice subgrupo
Public nCodiIdent As Integer
Public nSeccMaqui As Integer
Public TeclaTmp As teclas
Public lAcT As Boolean
Public MyFam As Integer
Public sIpCajas1 As String
Public sHiper As String
Public sDirCajas1 As String
Public sUserCajas1 As String
Public sPassCajas1 As String

Public Sub RefrescaTeclaTmp()
Dim Token As Long
'Token = InitGDIPlus()
'ImgFam(n).Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(TTecGrupos.Indice(Index).SubGrupo(n).ImgSelect))
'FreeGDIPlus Token
    FrmSelect.LblTec1(0).Caption = TeclaTmp.ShortName
    FrmSelect.LblTec2(0).Caption = CStr(TeclaTmp.Key)
    Token = InitGDIPlus()
    FrmSelect.ImgTec(0).Picture = LoadPictureGDIPlus(Trim(TeclaTmp.icon_name))
    FreeGDIPlus Token
    'Token = InitGDIPlus()
    'FrmSelect.ImgQ(0).Picture = LoadPictureGDIPlus(Trim(TeclaTmp.qLevelName))
    'FreeGDIPlus Token
    FrmSelect.Refresh
End Sub

Public Function extraeBMP(ByVal sS As String) As String
'C:\hbasicTouch\images\02 Zanahoria tecla 174.bmp
Dim nL As Integer
Dim nCont As Integer
Dim sDes As String
    If InStr(1, sS, "\") <> 0 Then
        nL = Len(sS)
        For nCont = nL To 1 Step -1
            If Mid(sS, nCont, 1) = "\" Then
                sDes = Mid(sS, nCont + 1)
                Exit For
            End If
        Next nCont
        sDes = Format(nCont, "000") & sDes
    Else
        sDes = "000" & sS
    End If
    extraeBMP = sDes
End Function
Public Function bmpTopng(ByVal sS As String)
Dim sDes As String
Dim nL As Integer
    nL = InStr(1, sS, ".bmp")
    If nL <> 0 Then
        sDes = Mid(sS, 1, nL - 1) & ".png"
    
    Else
        sDes = sS
    End If
    
    bmpTopng = sDes
End Function
Public Function comaTopunto(ByVal sS As String)
Dim sDes As String
Dim nL As Integer
    nL = InStr(1, sS, ",")
    If nL <> 0 Then
        sDes = Mid(sS, 1, nL - 1) & "." & Mid(sS, nL + 1)
    
    Else
        sDes = sS
    End If

    comaTopunto = sDes
End Function
Public Sub AbrirMySQL(Optional ByVal ssIp As String)
Dim MyFich As Integer
'Dim sserver As String
    ''''''''''''''''''''MySql
    On Error GoTo ErrorConexion
    ''''''''''''''''''''MySQL
    Set DbMysqli = New ADODB.Connection
    If IsNull(ssIp) Or ssIp = "" Then
        DbMysqli.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sIpAdd & "; Port=" & CStr(nIpPort) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; connection timeout=5; OPTION=3"
    Else
        DbMysqli.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & ssIp & "; Port=" & CStr(nIpPort) & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; connection timeout=5; OPTION=3"
    End If
    DbMysqli.Open
ErrorConexion:
    
    If Err.Number <> 0 Then
        lMySqlAbierta = False
        If DbMysqli.State <> 0 Then
            DbMysqli.Close
            Set DbMysqli = Nothing
        End If
        On Error GoTo 0
    Else
        lMySqlAbierta = True
    End If
'ErrorConexion:
    If Err.Number <> 0 Then
        'On Error GoTo 0
        
        CadenadeLog "No es posible la conexión con Touch-Scale... Error:" & CStr(Err.Number) & " " & Err.Description
        lMySqlAbierta = False
    End If
    On Error GoTo 0
    ''''''''''''''''''''
End Sub
Public Sub AbrirTotales()
Dim MyFich As Integer
    On Error GoTo ErrorConexion
    Set DbTotales = New ADODB.Connection
    DbTotales.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase.mdb"
    DbTotales.Open
    If Err.Number <> 0 Then
        lTotalesAbierta = False
        DbTotales.Close
        Set DbTotales = Nothing
        On Error GoTo 0
    Else
        lTotalesAbierta = True
    End If
ErrorConexion:
    If Err.Number <> 0 Then
        'On Error GoTo 0
        
        CadenadeLog "Touch-Scale... No es posible conexión con BDASE.MDB" & " Error:" & CStr(Err.Number) & " " & Err.Description
        lTotalesAbierta = False
    End If
    On Error GoTo 0
    ''''''''''''''''''''
End Sub

Public Function sToUTF8(ByVal sS As String) As String
Dim objStream As ADODB.Stream
Dim Data() As Byte
Dim sDes As String
Dim nCont As Integer
    If sS <> "" Then
    
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
    
    objStream.WriteText sS
    objStream.Flush
    
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3
    
    Data = objStream.Read()
    objStream.Close
    
    On Error GoTo Salir
    sDes = ""
    For nCont = 0 To 1000
        sDes = sDes & Chr(Data(nCont))
    Next nCont
    
Salir:
    On Error GoTo 0
    Else
        sDes = " "
    End If
    sToUTF8 = sDes

End Function

Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    On Error GoTo fin
    Fichero = FreeFile
    Open App.Path & "\loghydra.txt" For Append As Fichero
    On Error GoTo 0
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
fin:
    On Error GoTo 0
End Sub

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

Public Sub ExtractImages()
Dim MyResul As Long
Dim MyProc As Long
Dim Bucle As Integer
Dim nC As Integer
Dim siP, susR, spsS As String
Dim nFich As Integer
Dim Ret As Long

    siP = toIpAdd(sIpAdd)
    susR = "pcscale"
    spsS = "epelsa"
    nFich = 0
    Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
    nRetConnFTP = Ret
    FrmAviso.File1.Path = "c:\balanzas\ia\entrada"
    FrmAviso.File1.FileName = "img" & sHiper & "???" & ".b" & Format(Date, "yymmdd") & ".z"
    FrmAviso.File1.Refresh
    For Bucle = 0 To FrmAviso.File1.ListCount - 1
        FileCopy "c:\balanzas\ia\entrada\" & FrmAviso.File1.List(Bucle), App.Path & "\" & FrmAviso.File1.List(Bucle)
        Do_Events
        Do_Events
        Kill "c:\balanzas\ia\entrada\" & FrmAviso.File1.List(Bucle)
        Do_Events
        Do_Events
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\7z e " & Trim(FrmAviso.File1.List(Bucle)) & " -y", vbMinimized))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
        MyResul = STILL_ACTIVE
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\tartool -x " & Mid(Trim(FrmAviso.File1.List(Bucle)), 1, Len(Trim(FrmAviso.File1.List(Bucle))) - 2) & " images", vbMinimized))
        Do While MyResul = STILL_ACTIVE
            GetExitCodeProcess MyProc, MyResul
            Do_Events
        Loop
        If Dir(App.Path & "\" & Mid(Trim(FrmAviso.File1.List(Bucle)), 1, Len(Trim(FrmAviso.File1.List(Bucle))) - 2)) <> "" Then
            Kill App.Path & "\" & Mid(Trim(FrmAviso.File1.List(Bucle)), 1, Len(Trim(FrmAviso.File1.List(Bucle))) - 2)
            Do_Events
        End If
        If Dir(App.Path & "\" & Trim(FrmAviso.File1.List(Bucle))) <> "" Then
            FileCopy App.Path & "\" & Trim(FrmAviso.File1.List(Bucle)), App.Path & "\baarias\" & Trim(FrmAviso.File1.List(Bucle))
            Do_Events
            Kill App.Path & "\" & Trim(FrmAviso.File1.List(Bucle))
        End If
        nFich = nFich + 1
    Next Bucle
    
    If nRetConnFTP = 1 And nFich <> 0 Then
        FrmAviso.File1.Path = App.Path & "\images"
        FrmAviso.File1.FileName = "*.*"
        FrmAviso.File1.Refresh
        For Bucle = 0 To FrmAviso.File1.ListCount - 1
            If InStr(1, Trim(FrmAviso.File1.List(Bucle)), ".bmp") <> "" Or InStr(1, Trim(FrmAviso.File1.List(Bucle)), ".jpg") <> "" Or InStr(1, Trim(FrmAviso.File1.List(Bucle)), ".png") <> "" Then
                Ret = Enviar_Fichero_FTP(App.Path & "\images\" & Trim(FrmAviso.File1.List(Bucle)), "./scale/resources/icons/item/" & Trim(FrmAviso.File1.List(Bucle)), 0)
                'Ret = Err.LastDllError
                If Ret <> 1 Or (Err.LastDllError) Then 'error en DLL
                    GoTo errorEnvioFtp
                End If
            
            End If
        Next Bucle
    End If
    
errorEnvioFtp:
    If nRetConnFTP <> 0 Then
        Ret = Desconectar_FTP(0)
        FTP_hOpen(0) = 0
        FTP_HandleConect(0) = 0
        nRetConnFTP = 0
    End If
        
End Sub

Public Sub ExtractItems()
Dim Bucle As Integer
Dim nC As Integer
Dim siP, susR, spsS As String
Dim nFich As Integer
Dim Ret As Long
Dim sOrg As String
Dim sArt As RegArt
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyRsFam As DAO.Recordset
Dim cm As String

    On Error Resume Next
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    If Err.Number = 0 Then
        FrmAviso.File1.Path = "c:\balanzas\ia\entrada"
        FrmAviso.File1.FileName = "tec" & sHiper & "???" & ".b" & Format(Date, "yymmdd")
        FrmAviso.File1.Refresh
        For Bucle = 0 To FrmAviso.File1.ListCount - 1
            FileCopy "c:\balanzas\ia\entrada\" & FrmAviso.File1.List(Bucle), App.Path & "\" & FrmAviso.File1.List(Bucle)
            Do_Events
            Do_Events
            Kill "c:\balanzas\ia\entrada\" & FrmAviso.File1.List(Bucle)
            Do_Events
            Do_Events
        
            nFich = FreeFile()
            Open App.Path & "\" & FrmAviso.File1.List(Bucle) For Input As #nFich
            Line Input #nFich, cm
            Close #nFich
            cm = Replace(cm, Chr(10), Chr(13) & Chr(10))
            nFich = FreeFile()
            Open App.Path & "\" & FrmAviso.File1.List(Bucle) For Output As #nFich
            Print #nFich, cm
            Close #nFich
            nFich = FreeFile()
            Open App.Path & "\" & FrmAviso.File1.List(Bucle) For Input As #nFich
            Do While Not EOF(nFich)
                Line Input #nFich, sOrg
                sArt.AB = Mid(sOrg, 1, 1)
                sArt.BOT = Mid(sOrg, 2, 2)
                sArt.FAM = Mid(sOrg, 4, 2)
                sArt.POS1 = Mid(sOrg, 6, 2)
                sArt.POS2 = Mid(sOrg, 8, 2)
                sArt.EAN = Mid(sOrg, 10, 14)
                sArt.DES = Mid(sOrg, 24, 15)
                sArt.CON = Mid(sOrg, 39, 1)
                sArt.Img = Mid(sOrg, 40, 12)
                sArt.COD = Mid(sOrg, 52, 6)
                sArt.PLU = Mid(sOrg, 58, 3)
                If Val(sArt.COD) <> 0 And Mid(sArt.EAN, 2, 2) = "29" Then
                    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where code=" & CStr(Val(sArt.COD)))
                    If sArt.AB = "A" Then
                        If MyRsORG.EOF Then
                            MyRsORG.AddNew
                        Else
                            MyRsORG.Edit
                        End If
                        MyRsORG.Fields("code") = Val(sArt.COD)
                        MyRsORG.Fields("shortname") = sArt.DES
                        MyRsORG.Fields("plu") = Val(sArt.PLU)
                        MyRsORG.Fields("counter") = nSecSer
                        MyRsORG.Fields("section") = nSecSer
                        MyRsORG.Fields("ean13") = Mid(sArt.EAN, 2)
                        MyRsORG.Fields("icon_name") = Trim(sArt.Img)
                        MyRsORG.Fields("icon_path") = App.Path & "\images"
                        MyRsORG.Fields("grp") = Val(sArt.FAM)
                        MyRsORG.Fields("family") = Val(sArt.FAM)
                        If sArt.CON = "N" Then
                            MyRsORG.Fields("set") = 1
                        Else
                            MyRsORG.Fields("set") = 0
                        End If
                        MyRsORG.Fields("key") = Val(sArt.POS1)
                        'Ver si POS2 se pone en subgrupo
                        MyRsORG.Fields("modificacion") = ""
                        MyRsORG.Update
                    Else
                        cm = "delete from items where code=" & CStr(Val(sArt.COD))
                        MyBaseORG.Execute cm
                    End If
                    MyRsORG.Close
                    Set MyRsORG = Nothing
                
                Else 'pueden ser familias
                    If Val(sArt.EAN) = 0 Then
                        Set MyRsFam = MyBaseORG.OpenRecordset("select * from grp where grp=" & CStr(Val(sArt.FAM)))
                        If Not MyRsFam.EOF Then
                            MyRsFam.Edit
                            MyRsFam.Fields("text") = Trim(sArt.DES)
                            MyRsFam.Fields("imgnoselect") = Trim(sArt.Img)
                            MyRsFam.Fields("imgselect") = Trim(sArt.Img)
                            MyRsFam.Update
                            cm = "update subgroup set imgnoselect='" & Trim(sArt.Img) & "', imgselect='" & Trim(sArt.Img) & "' where grp=" & CStr(Val(sArt.FAM))
                            MyBaseORG.Execute cm
                        End If
                        MyRsFam.Close
                        Set MyRsFam = Nothing
                    End If
                End If
            Loop
            Close #nFich
            
            If Dir(App.Path & "\" & Trim(FrmAviso.File1.List(Bucle))) <> "" Then
                FileCopy App.Path & "\" & Trim(FrmAviso.File1.List(Bucle)), App.Path & "\baarias\" & Trim(FrmAviso.File1.List(Bucle))
                Do_Events
                Kill App.Path & "\" & Trim(FrmAviso.File1.List(Bucle))
            End If
            
        Next Bucle
        
    End If
    
    On Error GoTo 0
End Sub


' accept a byte array containing utf-8 data
' and convert it to a string
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As String

    Dim objStream As ADODB.Stream
    Dim Data() As Byte
    Dim nC, nL As Long
    Dim s As String
    
    If Dir(App.Path & "\linux.txt") <> "" Then '5.4.3
    
    If strText = "" Or IsNull(strText) Then
        s = ""
        GoTo salconver
    End If
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
    
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3 ' skip first 3 bytes as this is the utf8 detection header
    Data = objStream.Read()
    nL = objStream.size
    ' close up and return
    objStream.Close
    On Error GoTo salconver
    s = ""
    For nC = 0 To nL - 1
        s = s & Chr(Data(nC))
    Next nC
salconver:
    On Error GoTo 0
    Else
        s = strText
    End If
    ConvertStringToUtf8Bytes = s

End Function

Public Sub mainEnvCambios()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyBaseORG1 As DAO.Database
Dim MyRsORG1 As DAO.Recordset
Dim MyBaseBAL As DAO.Database
Dim MyRsBAL As DAO.Recordset
Dim siP As String
Dim MyRsDes As New ADODB.Recordset
Dim myRsKEYS As New ADODB.Recordset
Dim myRsTMP As New ADODB.Recordset
Dim cm As String
Dim nF As Integer



    lComunica = True
    FrmKeyTouch.CmdEnvCambios.Enabled = False
    FrmKeyTouch.CmdExit.Enabled = False
    
    siP = toIpAdd(sIpAdd)
    
    FrmKeyTouch.TxtVer.Text = FrmKeyTouch.TxtVer.Text & Now() & " --> Iniciando conexión Servidora:" & siP & vbCrLf
    FrmKeyTouch.TxtVer.Refresh
    
    If ElPing(siP) = 1 Then
    
        Call AbrirMySQL(siP)
        
        If lMySqlAbierta Then
            
            Call CapturaDB
            
            FrmKeyTouch.TxtVer.Text = FrmKeyTouch.TxtVer.Text & Now() & " --> Conexión Correcta." & vbCrLf
            FrmKeyTouch.TxtVer.Refresh
                        
            Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
            
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from grp order by grp")
            If Not MyRsORG.EOF Then
                MyRsORG.MoveFirst
                Do While Not MyRsORG.EOF
                    cm = "delete from pcscale.families where code=" & CStr(MyRsORG.Fields("grp"))
                    DbMysqli.Execute cm
                    MyRsDes.Open "select * from pcscale.families where code=" & CStr(MyRsORG.Fields("grp")), DbMysqli, adOpenKeyset, adLockOptimistic
                    MyRsDes.AddNew
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(54) NOT NULL DEFAULT '',
'  `Icon` varchar(127) DEFAULT '',
'  `DefaultIcon` varchar(127) DEFAULT '',
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
                    MyRsDes.Fields("code") = MyRsORG.Fields("grp")
                    MyRsDes.Fields("name") = MyRsORG.Fields("text")
                    MyRsDes.Fields("icon") = MyRsORG.Fields("imgselect")
                    MyRsDes.Fields("DefaultIcon") = ""
                    MyRsDes.Fields("label") = ""
                    MyRsDes.Fields("isbovine") = 0
                    MyRsDes.Update
                    MyRsDes.Close
                    Set MyRsDes = Nothing
                    
                    'asignación de indice
                    cm = "delete from pcscale.counterfamilies where familycode=" & CStr(MyRsORG.Fields("grp"))
                    DbMysqli.Execute cm
                    cm = "delete from pcscale.counterfamilies where familyindex=" & CStr(MyRsORG.Fields("grp"))
                    DbMysqli.Execute cm
                    
                    MyRsDes.Open "select * from pcscale.counterfamilies where familycode=" & CStr(MyRsORG.Fields("grp")) & " and countercode=" & CStr(nSecSer), DbMysqli, adOpenKeyset, adLockOptimistic
                    MyRsDes.AddNew
'DROP TABLE IF EXISTS `pcscale`.`counterfamilies`;
'CREATE TABLE  `pcscale`.`counterfamilies` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `CounterCode` int(11) NOT NULL,
'  `FamilyIndex` int(11) NOT NULL DEFAULT '1',
'  `FamilyCode` int(11) NOT NULL,
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `idcofam` (`CounterCode`,`FamilyIndex`),
'  CONSTRAINT `fkcofam` FOREIGN KEY (`CounterCode`) REFERENCES `counters` (`Code`) ON DELETE CASCADE
') ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=utf8;
                    MyRsDes.Fields("countercode") = nSecSer
                    MyRsDes.Fields("familyindex") = MyRsORG.Fields("grp")
                    MyRsDes.Fields("familycode") = MyRsORG.Fields("grp")
                    MyRsDes.Update
                    MyRsDes.Close
                    Set MyRsDes = Nothing
                    
                    
                    MyRsORG.MoveNext
                Loop
            End If
            MyRsORG.Close
            Set MyRsORG = Nothing
            
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from items where code<>0 and price<>0 order by key")
            ''''''''''''''''''''''''''''''''
            cm = "delete from pcscale.items"
            DbMysqli.Execute cm
            ''''''''''''''''''''''''''''''''
            
            If Not MyRsORG.EOF Then
                MyRsORG.MoveFirst
                Do While Not MyRsORG.EOF
                    
                    cm = "select * from pcscale.items where code=" & CStr(Trim(MyRsORG.Fields("code")))
                    MyRsDes.Open cm, DbMysqli, adOpenDynamic, adLockOptimistic
                    
                    If MyRsDes.EOF Then
                        MyRsDes.AddNew
                    End If
                    
                    MyRsDes!Code = MyRsORG.Fields("code")
                    MyRsDes!Name = (UCase(Trim(MyRsORG.Fields("shortname")))) 'ConvertStringToUtf8Bytes
                    MyRsDes!ShortName = (UCase(Trim(MyRsORG.Fields("shortname")))) 'ConvertStringToUtf8Bytes
                    MyRsDes!family = MyRsORG.Fields("family")
                    MyRsDes!price = MyRsORG.Fields("price")
                    MyRsDes!saleform = MyRsORG.Fields("saleform")
                    MyRsDes!PLU = MyRsORG.Fields("plu")
                    MyRsDes!Counter = MyRsORG.Fields("counter")
                    MyRsDes!vat = MyRsORG.Fields("vat_value")
                    'MyRsDes!ean13 = sToUTF8(Trim(MyRsORG.Fields("ean13")))
                    MyRsDes!expiring = MyRsORG.Fields("expiring")
                    MyRsDes!bestbefore = MyRsORG.Fields("bestbefore")
                    MyRsDes!tare = MyRsORG.Fields("tare")
                    MyRsDes!Label = ""
                    MyRsDes!Icon = (Trim(MyRsORG.Fields("icon_name"))) 'ConvertStringToUtf8Bytes
                    MyRsDes!discount = 0
                    MyRsDes!onkey = 1
                    MyRsDes!familyorder = MyRsORG.Fields("key")
                    
                    MyRsDes.Update
                    MyRsDes.Close
                    Set MyRsDes = Nothing
                    
                    
                    MyRsORG.MoveNext
                Loop
                
                cm = "select * from pcscale.barcodeean13"
                myRsKEYS.Open cm, DbMysqli, adOpenDynamic, adLockOptimistic
                If Not myRsKEYS.EOF Then
                    cm = "update pcscale.barcodeean13 set number=" & CStr(nSecSer)
                    DbMysqli.Execute cm
                    cm = "update pcscale.barcodeean13 set networkss=" & "'2CCCCCCiiiii'" & ", localss=" & "'2CCCCCCiiiii'"
                    DbMysqli.Execute cm
        'NetworkPacket1 CHAR(12) NOT NULL, NetworkPacket2
                    cm = "update pcscale.barcodeean13 set NetworkPacket1=" & "'2CCCCCCIIIII'" & ", NetworkPacket2=" & "'2CCCCCCIIIII'"
                    DbMysqli.Execute cm
        'LocalPacket1 CHAR(12) NOT NULL,  LocalPacket2 CHAR(12) NOT NULL
                    cm = "update pcscale.barcodeean13 set LocalPacket1=" & "'2CCCCCCIIIII'" & ", LocalPacket2=" & "'2CCCCCCIIIII'"
                    DbMysqli.Execute cm
                End If
                myRsKEYS.Close
                Set myRsKEYS = Nothing
                
            End If
            
            
            MyRsORG.Close
            Set MyRsORG = Nothing
            MyBaseORG.Close
            Set MyBaseORG = Nothing
            
            If lMySqlAbierta Then
                DbMysqli.Close
                Set DbMysqli = Nothing
            End If

            FrmKeyTouch.TxtVer.Text = FrmKeyTouch.TxtVer.Text & Now() & " --> Datos enviados." & vbCrLf
            FrmKeyTouch.TxtVer.Refresh
        
        
        Else
            FrmKeyTouch.TxtVer.Text = FrmKeyTouch.TxtVer.Text & Now() & " --> No ha sido posible la Conexión." & vbCrLf
            FrmKeyTouch.TxtVer.Refresh
        
        End If
    
    Else
        
        FrmKeyTouch.TxtVer.Text = FrmKeyTouch.TxtVer.Text & Now() & " --> No ha sido posible la Conexión.(sin ping)" & vbCrLf
        FrmKeyTouch.TxtVer.Refresh
    
    End If
    
    
    FrmKeyTouch.CmdEnvCambios.Enabled = True
    FrmKeyTouch.CmdExit.Enabled = True
    lComunica = False
    

End Sub

Public Sub CapturaDB()
Dim MyBaseORG As DAO.Database
Dim MyBaseDES As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyRsDes As DAO.Recordset
Dim myRsTMP As DAO.Recordset
Dim sS As String
Dim lEsAlta As Boolean
Dim nPlu As Integer

    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyBaseDES = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    
    Set MyRsDes = MyBaseDES.OpenRecordset("select * from items where code<>0")
    If Not MyRsDes.EOF Then
        
        MyRsDes.MoveFirst
        Do While Not MyRsDes.EOF
            Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes.Fields("code")))
            If Not MyRsORG.EOF Then
                MyRsDes.Edit
                MyRsDes.Fields("price") = MyRsORG.Fields("euros")
                If MyRsORG.Fields("codi_pes") = "W" Then
                    MyRsDes.Fields("saleform") = 1
                Else
                    MyRsDes.Fields("saleform") = 0
                End If
                MyRsDes.Fields("vat_type") = MyRsORG.Fields("tipo_iva")
                MyRsDes.Fields("vat_value") = 0
                MyRsDes.Fields("expiring") = MyRsORG.Fields("caducidad")
                MyRsDes.Fields("bestbefore") = 0 'no tenemos dato
                MyRsDes.Fields("tare") = MyRsORG.Fields("tara")
                MyRsDes.Fields("label_value") = 0
                MyRsDes.Fields("label_text") = ""
                ''''''''''
                MyRsDes.Update
                ''''''''''
            End If
            MyRsDes.MoveNext
        Loop
        
    End If

Salir:
    On Error Resume Next
    MyRsDes.Close
    Set MyRsDes = Nothing
    MyRsORG.Close
    Set MyRsDes = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    MyBaseDES.Close
    Set MyBaseDES = Nothing
    On Error GoTo 0
End Sub

Public Sub Main()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyBaseDES As DAO.Database
Dim MyRsDes As DAO.Recordset
Dim MyRsITM As DAO.Recordset
Dim nF As Integer
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

    If App.PrevInstance Then End
    'Se supone que la única Balanza con Port asignado es la Servidora
    'Se recogen datos de IP, Port, Mostrador al que asignar todos los artículos que se envíen
    ' y si se pueden asignar oo no a teclas todos los artículos capturados de Baaria (Dbase.mdb)
    On Error Resume Next
    Set myBase = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    If Err.Number = 0 Then
        Set myRS = myBase.OpenRecordset("select * from balanzas where port<>0")
        If Not myRS.EOF Then
            myRS.Edit
            sIpAdd = Trim(myRS.Fields("ip"))
            nIpPort = CStr(myRS.Fields("port"))
            nSecSer = myRS.Fields("counter")
            If myRS.Fields("allitems") Then
                lAllItem = True
            Else
                lAllItem = False
            End If
        Else
            myRS.Close
            lSalir = True
            Load FrmConfiguracion
            FrmConfiguracion.Show
            Do While FrmConfiguracion.Visible
                Do_Events
            Loop
            cm = "insert into cfg values(" & CStr(nSecSer) & "," & CStr(nSecSer) & "," & "480,2,6,40)"
            myBase.Execute cm
            cm = "insert into grp values(1,'Grp-Fam 1',6,1,240,'LCarrefour.png','LCarrefour.png')"
            myBase.Execute cm
            cm = "insert into grp values(2,'Grp-Fam 2',6,1,240,'LCarrefour.png','LCarrefour.png')"
            myBase.Execute cm
            For nF = 1 To 2
                For nnF = 1 To 6
                
                    cm = "insert into subgroup values(" & CStr(nF) & "," & CStr(nnF) & ",'Posiciones:" & CStr((nnF - 1) * 40 + 1) & " a " & CStr((nnF - 1) * 40 + 40) & "'," & CStr((nnF - 1) * 40 + 1) & "," & CStr((nnF - 1) * 40 + 40) & ",'LCarrefour.png'" & ",'LCarrefour.png')"
                    myBase.Execute cm
                    
                Next nnF
            
            Next nF
            MsgBox "Se saldrá de la utilidad... Por favor reinicie.", vbExclamation
        End If
        myRS.Close
    End If
    myBase.Close
    If lSalir Then End
    
    lAcT = True
    'Del fichero hiper se captura el número de Hiper Mercado (sHiper)
    If Dir(App.Path & "\hiper") <> "" Then
        nF = FreeFile()
        Open App.Path & "\hiper" For Input As #nF
        Line Input #nF, sHiper
        Line Input #nF, sIpCajas1
        Line Input #nF, sDirCajas1
        Line Input #nF, sUserCajas1
        Line Input #nF, sPassCajas1
        Close #nF
        sIpCajas1 = Trim(sIpCajas1)
        sHiper = Trim(sHiper)
        sHiper = Format(Val(sHiper), "0000")
        sDirCajas1 = Trim(sDirCajas1)
        sUserCajas1 = Trim(sUserCajas1)
        sPassCajas1 = Trim(sPassCajas1)
    End If
        
        
'prueba
'FrmAviso.LblAviso.Caption = "Proceso/Envio Modificaciones a Touch-Scale"
'FrmAviso.Show
'FrmAviso.Refresh
'
'Call ExtractImages
'Call ExtractItems
'''''''
    
    Call CapturaDB
    
    'Parámetros a tener en cuenta para los procesos.
    If Trim(Command) <> "" And Trim(LCase(Command)) <> "hydra" Then
        
        If LCase(Trim(Command)) = "update" Then
            FrmAviso.LblAviso.Caption = "Proceso/Envio Modificaciones a Touch-Scale"
            FrmAviso.Show
            FrmAviso.Refresh
            
            Call ExtractImages
            Call ExtractItems
            
            'Call mainEnvCambios
        End If
        
        
        FrmAviso.Hide
        End
    Else
        
        FrmKeyTouch.Lista.Visible = False
        
        Load FrmKeyTouch
        FrmKeyTouch.Show
    
    End If
    
End Sub

Public Function DigitoControl(ByVal Valor As String) As Integer
   '**************************
   ' Calcula el Checksum de
   ' un código de barras
   '***************************
   Dim Longitud As Integer
   Dim sumapar As Integer
   Dim sumaimpar As Integer
   Dim suma As Integer
   Dim digito, i As Integer
   sumapar = 0
   sumaimpar = 0
   Longitud = Len(Valor)
   For i = 1 To Longitud
      If i Mod 2 = 0 Then
         sumapar = sumapar + Val(Mid(Valor, i, 1))
      Else
         sumaimpar = sumaimpar + Val(Mid(Valor, i, 1))
      End If
   Next
   If Longitud Mod 2 = 0 Then
      suma = sumapar * 3 + sumaimpar
   Else
      suma = sumapar + sumaimpar * 3
   End If

   digito = 10 - suma Mod 10
   If digito = 10 Then digito = 0
   DigitoControl = digito
End Function

''''¡¡¡¡¡¡¡¡¡¡¡¡¡ Solo utilizar en final día!!!!!!!!!!!!!!!!!!!'''''''

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


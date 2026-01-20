Attribute VB_Name = "Var"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
'////////////////////////
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


Public Sub RefrescaTeclaTmp()
    FrmSelect.LblTec1(0).Caption = TeclaTmp.ShortName
    FrmSelect.LblTec2(0).Caption = CStr(TeclaTmp.Key)
    FrmSelect.ImgTec(0).Picture = LoadPicture(Trim(TeclaTmp.icon_name))
    FrmSelect.ImgQ(0).Picture = LoadPicture(Trim(TeclaTmp.qLevelName))
    FrmSelect.Refresh
End Sub

Public Function extraeBMP(ByVal sS As String) As String
'C:\hbasicTouch\items\02 Zanahoria tecla 174.bmp
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
Public Sub AbrirMySQL(ByVal sserver As String)
Dim MyFich As Integer
'Dim sserver As String
    ''''''''''''''''''''MySql
    On Error GoTo ErrorConexion
    ''''''''''''''''''''MySQL
    Set DbMysqli = New ADODB.Connection
    
    'If Dir(App.Path & "\myserver.cfg") <> "" Then
    '    MyFich = FreeFile()
    '    Open App.Path & "\myserver.cfg" For Input As #MyFich
    '    Line Input #MyFich, sServer
    '    Close #MyFich
    '    sServer = sServer & ";"
    'Else
    '    sServer = "localhost;"
    'End If
    'DbMysqli.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=localhost; DATABASE=balanzas; UID=root; PWD=; OPTION=3"
    'DbMysqli.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.50.2.102; DATABASE=balanzas; UID=root; PWD=; OPTION=3"
    DbMysqli.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & sserver & " DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
    DbMysqli.Open
    If Err.Number <> 0 Then
        lMySqlAbierta = False
        DbMysqli.Close
        Set DbMysqli = Nothing
        On Error GoTo 0
    Else
        lMySqlAbierta = True
    End If
ErrorConexion:
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
    
    On Error GoTo salir
    sDes = ""
    For nCont = 0 To 1000
        sDes = sDes & Chr(Data(nCont))
    Next nCont
    
salir:
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
   'Dim tmp As String
   'Dim nlen As Integer
   'Dim Cont As Integer
   'Dim s As String
   
   'nlen = Len(sR)
   'For Cont = 1 To nlen
   '   s = Mid(sR, Cont, 1)
   '   If Cont > 1 Then
   '     If (Asc(s) < 49 Or Asc(s) > 57) And Mid(sR, Cont - 1, 1) = "." Then
   '     Else
   '      tmp = tmp + s
   '     End If
   '   Else
   '      If Asc(s) < 49 Or Asc(s) > 57 Then
   '      Else
   '         tmp = tmp + s
   '      End If
   '   End If
   '
   'Next
   
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
    
    If App.PrevInstance Then End
    
    If Trim(Command) = "" Then
        MsgBox "Este ejecutable debe ser Utilizado por HydraBasic.", vbCritical
        End
    End If
    'PATH's TOUCHSCALE
    If Dir(App.Path & "\touchpath.cfg") = "" Then
        MsgBox "LOS PATH's de TOUCH-SCALE's NO ESTÁN DEFINIDOS...", vbCritical
        End
    End If
    nF = FreeFile()
    nCont = 0
    Open App.Path & "\touchpath.cfg" For Input As #nF
    Do While Not EOF(nF)
        nCont = nCont + 1
        ReDim Preserve sPath(nCont)
        Line Input #nF, sPath(nCont)
    Loop
    Close #nF
    '''''''''''''''''''''''
    'CFG SS
    If Dir(App.Path & "\touchss.cfg") = "" Then
        MsgBox "LAS CARACTERÍSTICAS BÁSICAS AUTOSERVICIO NO ESTÁN DEFINIDAS...", vbCritical
        End
    End If
    nF = FreeFile()
    nCont = 0
    Open App.Path & "\touchss.cfg" For Input As #nF
    Do While Not EOF(nF)
        nCont = nCont + 1
        Line Input #nF, sCfgSS(nCont)
    Loop
    Close #nF
    'Recarga de grupos...
    If Dir(App.Path & "\touchgrp.cfg") = "" Then
        MsgBox "LAS CARACTERÍSTICAS DE LOS GRUPOS NO ESTÁN DEFINIDAS...", vbCritical
        End
    End If
    nF = FreeFile()
    nCont = 0
    ReDim sGrpSS((Val(sCfgSS(4))) * 5)
    Open App.Path & "\touchgrp.cfg" For Input As #nF
    Do While Not EOF(nF)
        nCont = nCont + 1
        If nCont > (Val(sCfgSS(4))) * 5 Then Exit Do 'solo se cogen los primeros hasta máximo de grupos indicado en touchss.cfg
        Line Input #nF, sGrpSS(nCont)
    Loop
    nGrp = (nCont / 5)
    Close #nF
    '''''''''''''''''''''
    'Recarga de subgrupos...
    If Dir(App.Path & "\touchsubgrp.cfg") = "" Then
        MsgBox "LAS CARACTERÍSTICAS DE LOS SUB-GRUPOS NO ESTÁN DEFINIDAS...", vbCritical
        End
    End If
    nF = FreeFile()
    nCont = 0
    Open App.Path & "\touchsubgrp.cfg" For Input As #nF
    Do While Not EOF(nF)
        nCont = nCont + 1
        ReDim Preserve sSubGrpSS(nCont)
        Line Input #nF, sSubGrpSS(nCont)
    Loop
    Close #nF
    '''''''''''''''''''''
    'compactar base
    If Dir(App.Path & "\ctouchscale.mdb") <> "" Then
        FileCopy App.Path & "\ctouchscale.mdb", App.Path & "\dirtmp\ctouchscale.mdb"
        Kill App.Path & "\ctouchscale.mdb"
        Do While Dir(App.Path & "\ctouchscale.mdb") <> ""
            Do_Events
        Loop
    End If
    DBEngine.CompactDatabase App.Path & "\touchscale.mdb", App.Path & "\ctouchscale.mdb"
    Do While Dir(App.Path & "\ctouchscale.mdb") = ""
        Do_Events
    Loop
    FileCopy App.Path & "\touchscale.mdb", App.Path & "\dirtmp\touchscale.mdb"
    Kill App.Path & "\touchscale.mdb"
    Do While Dir(App.Path & "\touchscale.mdb") <> ""
        Do_Events
    Loop
    FileCopy App.Path & "\ctouchscale.mdb", App.Path & "\touchscale.mdb"
    Do While Dir(App.Path & "\touchscale.mdb") = ""
        Do_Events
    Loop
    '''''''''''''''''''''
    'Recarga de Equipos
    Set MyBaseDES = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsDes = MyBaseDES.OpenRecordset("select * from balanzas")
    If Not MyRsDes.EOF Then
        MyRsDes.MoveFirst
        Do While Not MyRsDes.EOF
            MyRsDes.Delete
            MyRsDes.MoveNext
        Loop
    End If
    
    FrmKeyTouch.Lista.Visible = False
    
    
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from equipos where modelo>=100")
    If Not MyRsORG.EOF Then 'Si no existen equipos TOUCH en DBASE.MDB no se permite carga formulario y no se continua carga
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF 'EQUIPOS
            MyRsDes.AddNew
            MyRsDes.Fields("scalenr") = MyRsORG.Fields("numero_eqp")
            MyRsDes.Fields("counter") = MyRsORG.Fields("codi_ident")
            MyRsDes.Fields("section") = MyRsORG.Fields("secc_maqui")
            MyRsDes.Fields("ip") = MyRsORG.Fields("prog_eqp")
            MyRsDes.Fields("port") = 12000 + MyRsORG.Fields("numero_eqp")
            MyRsDes.Fields("itemimagepath") = sPath(1)
            MyRsDes.Fields("publicityimagepath") = sPath(2)
            MyRsDes.Fields("qualityimagepath") = sPath(3)
            MyRsDes.Update
            MyRsORG.MoveNext
        Loop
        'Recarga Tabla CFG
        MyRsDes.Close
        Set MyRsDes = MyBaseDES.OpenRecordset("select * from cfg")
        If MyRsDes.EOF Then
            MyRsDes.AddNew
        Else
            MyRsDes.Edit
        End If
        'Límites actuales...
        If Val(sCfgSS(4)) > 3 Then sCfgSS(4) = "3"
        If Val(sCfgSS(5)) > 6 Then sCfgSS(5) = "6"
        If Val(sCfgSS(6)) > 20 Then sCfgSS(6) = "20"
        '''''''''''''''''''''
        MyRsDes.Fields("counter") = Val(sCfgSS(1))
        MyRsDes.Fields("scalesection") = Val(sCfgSS(2))
        nCodiIdent = Val(sCfgSS(1))
        nSeccMaqui = Val(sCfgSS(2))
        MyRsDes.Fields("keynr") = Val(sCfgSS(4)) * Val(sCfgSS(5)) * Val(sCfgSS(6))
        MyRsDes.Fields("groupnr") = Val(sCfgSS(4))
        MyRsDes.Fields("subgroupnr") = Val(sCfgSS(5))
        MyRsDes.Fields("keynrbysg") = Val(sCfgSS(6))
        MyRsDes.Update
        'Recarga Tabla Group
        MyRsDes.Close
        Set MyRsDes = MyBaseDES.OpenRecordset("select * from grp")
        If Not MyRsDes.EOF Then
            MyRsDes.MoveFirst
            Do While Not MyRsDes.EOF
                MyRsDes.Delete
                MyRsDes.MoveNext
            Loop
        End If
        MyRsORG.Close
        Set MyRsORG = MyBaseDES.OpenRecordset("select * from subgroup")
        If Not MyRsORG.EOF Then
            MyRsORG.MoveFirst
            Do While Not MyRsORG.EOF
                MyRsORG.Delete
                MyRsORG.MoveNext
            Loop
        End If
        nCont = 0
        For nnF = 1 To nGrp
            MyRsDes.AddNew
            If Val(sGrpSS(1 + ((nnF - 1) * 5))) > Val(sCfgSS(4)) Then sGrpSS(1 + ((nnF - 1) * 5)) = CStr(nnF)
            MyRsDes.Fields("grp") = Val(sGrpSS(1 + ((nnF - 1) * 5)))
            If Len(sGrpSS(2 + ((nnF - 1) * 5))) > 70 Then sGrpSS(2 + ((nnF - 1) * 5)) = Mid(sGrpSS(2 + ((nnF - 1) * 5)), 1, 70)
            MyRsDes.Fields("text") = sGrpSS(2 + ((nnF - 1) * 5))
            If Val(sGrpSS(3 + ((nnF - 1) * 5))) > Val(sCfgSS(5)) Then sGrpSS(3 + ((nnF - 1) * 5)) = "6"
            MyRsDes.Fields("subgroupnr") = Val(sGrpSS(3 + ((nnF - 1) * 5)))
            MyRsDes.Fields("inikey") = Val(sGrpSS(4 + ((nnF - 1) * 5)))
            MyRsDes.Fields("endkey") = Val(sGrpSS(5 + ((nnF - 1) * 5)))
            MyRsDes.Update
            MyRsORG.Close
            Set MyRsORG = MyBaseDES.OpenRecordset("select * from subgroup")
            '''''''''ver aquí
            MyRsDes.MoveLast
            MyRsDes.Edit
            For nnnF = 1 To MyRsDes.Fields("subgroupnr")
                MyRsORG.AddNew
                nCont = nCont + 1
                MyRsORG.Fields("grp") = Val(sSubGrpSS(nCont))
                nCont = nCont + 1
                MyRsORG.Fields("subgroup") = Val(sSubGrpSS(nCont))
                nCont = nCont + 1
                MyRsORG.Fields("text") = sSubGrpSS(nCont)
                nCont = nCont + 1
                MyRsORG.Fields("inikey") = Val(sSubGrpSS(nCont))
                nCont = nCont + 1
                MyRsORG.Fields("endkey") = Val(sSubGrpSS(nCont))
                MyRsORG.Update
            Next nnnF
            '''''''''''''''''
       Next nnF
       'Lectura imágenes grupos y subgrupos.
       nF = FreeFile()
       Open App.Path & "\touchgrpfamimg.cfg" For Input As #nF
       Set MyRsDes = MyBaseDES.OpenRecordset("select * from grp order by grp")
       MyRsDes.MoveFirst
       Do While Not MyRsDes.EOF
           MyRsDes.Edit
           Line Input #nF, sS
           MyRsDes.Fields("imgnoselect") = Trim(sS)
           Line Input #nF, sS
           MyRsDes.Fields("imgselect") = Trim(sS)
           MyRsDes.Update
           MyRsDes.MoveNext
       Loop
       Set MyRsDes = MyBaseDES.OpenRecordset("select * from subgroup order by grp,subgroup")
       MyRsDes.MoveFirst
       Do While Not MyRsDes.EOF
           MyRsDes.Edit
           Line Input #nF, sS
           MyRsDes.Fields("imgnoselect") = Trim(sS)
           Line Input #nF, sS
           MyRsDes.Fields("imgselect") = Trim(sS)
           MyRsDes.Update
           MyRsDes.MoveNext
       Loop
       Close #nF
       '''''''''''''''''''''''''''''''''''''''
       'Lectura Qlevels.
       Set MyRsDes = MyBaseDES.OpenRecordset("select * from qlevels")
       If Not MyRsDes.EOF Then MyRsDes.MoveFirst
       Do While Not MyRsDes.EOF
            MyRsDes.Delete
            MyRsDes.MoveNext
       Loop
       MyRsDes.Close
       Set MyRsDes = Nothing
       nF = FreeFile()
       Open App.Path & "\touchquality.cfg" For Input As #nF
       Set MyRsDes = MyBaseDES.OpenRecordset("select * from qlevels")
       Do While Not EOF(nF)
            MyRsDes.AddNew
            Line Input #nF, sS
            MyRsDes.Fields("code") = Trim(sS)
            Line Input #nF, sS
            MyRsDes.Fields("icon") = Trim(sS)
            Line Input #nF, sS
            MyRsDes.Fields("description") = Trim(sS)
            MyRsDes.Update
       Loop
       Close #nF
       
    Else
        MyRsDes.Close
        MyRsORG.Close
        MyBaseDES.Close
        Set MyBaseDES = Nothing
        MyBaseORG.Close
        Set MyBaseORG = Nothing
        Set MyRsDes = Nothing
        Set MyRsORG = Nothing
        End
    End If
       
    'carga de variables...
    'Set MyRsDES = MyBaseDES.OpenRecordset("select * from grp order by group")
    'MyRsDES.MoveLast
    'ReDim TTecGrupos.Indice(MyRsDES.RecordCount)
    'MyRsDES.MoveFirst
    'Do While Not MyRsDES.EOF
    '    nnF = MyRsDES.Fields("group") - 1
    '    TTecGrupos.Indice(nnF).Indice = nnF
    '    TTecGrupos.Indice(nnF).lblTxt = MyRsDES.Fields("text")
    '    TTecGrupos.Indice(nnF).SubGroupnr = MyRsDES.Fields("subgroupnr")
    '    TTecGrupos.Indice(nnF).IniKey = MyRsDES.Fields("inikey")
    '    TTecGrupos.Indice(nnF).EndKey = MyRsDES.Fields("endkey")
    '    TTecGrupos.Indice(nnF).ImgNoSelect = MyRsDES.Fields("imgnoselect")
    '    TTecGrupos.Indice(nnF).ImgSelect = MyRsDES.Fields("imgselect")
    '    If nnF = 0 Then
    '        ImgGrp(nnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsDES.Fields("imgselect"))
    '    Else
    '        ImgGrp(nnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsDES.Fields("imgnoselect"))
    '    End If
    '    LblGrp(nnF).Caption = MyRsDES.Fields("text")
    '    MyRsDES.MoveNext
    'Loop
    ''''''''''''''''''''''''
    Set MyRsDes = MyBaseDES.OpenRecordset("select * from grp order by grp")
    MyRsDes.MoveLast
    ReDim TTecGrupos.Indice(MyRsDes.RecordCount)
    nGrupos = MyRsDes.RecordCount
    MyRsDes.MoveFirst
    Do While Not MyRsDes.EOF
        nnF = MyRsDes.Fields("grp") - 1
        TTecGrupos.Indice(nnF).Indice = nnF
        TTecGrupos.Indice(nnF).lblTxt = MyRsDes.Fields("text")
        TTecGrupos.Indice(nnF).SubGroupnr = MyRsDes.Fields("subgroupnr")
        TTecGrupos.Indice(nnF).IniKey = MyRsDes.Fields("inikey")
        TTecGrupos.Indice(nnF).EndKey = MyRsDes.Fields("endkey")
        TTecGrupos.Indice(nnF).ImgNoSelect = MyRsDes.Fields("imgnoselect")
        TTecGrupos.Indice(nnF).ImgSelect = MyRsDes.Fields("imgselect")
        ReDim TTecGrupos.Indice(nnF).SubGrupo(MyRsDes.Fields("subgroupnr"))
        If nnF = 0 Then
            FrmKeyTouch.ImgGrp(nnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsDes.Fields("imgselect"))
        Else
            FrmKeyTouch.ImgGrp(nnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsDes.Fields("imgnoselect"))
        End If
        FrmKeyTouch.LblGrp(nnF).Caption = MyRsDes.Fields("text")
        Set MyRsORG = MyBaseDES.OpenRecordset("select * from subgroup where grp=" & CStr(nnF + 1) & " order by subgroup")
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            nnnF = MyRsORG.Fields("subgroup") - 1
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).Grupo = MyRsORG.Fields("grp")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).Indice = MyRsORG.Fields("subgroup") - 1
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).lblTxt = MyRsORG.Fields("text")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey = MyRsORG.Fields("inikey")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey = MyRsORG.Fields("endkey")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgNoSelect = MyRsORG.Fields("imgnoselect")
            TTecGrupos.Indice(nnF).SubGrupo(nnnF).ImgSelect = MyRsORG.Fields("imgselect")
            If nnF = 0 Then
                If nnnF = 0 Then
                    FrmKeyTouch.ImgFam(nnnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsORG.Fields("imgselect"))
                    'carga items primera pantalla
                    For nCont = 0 To 19
                        FrmKeyTouch.ImgTec(nCont).Visible = False
                        FrmKeyTouch.ImgQ(nCont).Visible = False
                        FrmKeyTouch.LblTec1(nCont).Visible = False
                        FrmKeyTouch.LblTec2(nCont).Visible = False
                    Next nCont
                    nCont = 0
                    For nF = TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey To TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey
                        FrmKeyTouch.ImgTec(nCont).Visible = True
                        FrmKeyTouch.LblTec1(nCont).Visible = True
                        FrmKeyTouch.LblTec2(nCont).Visible = True
                        FrmKeyTouch.LblTec1(nCont).Caption = ""
                        FrmKeyTouch.LblTec2(nCont).Caption = CStr(nF)
                        Set MyRsITM = MyBaseDES.OpenRecordset("select * from items where key=" & CStr(nF))
                        If Not MyRsITM.EOF Then
                            FrmKeyTouch.ImgTec(nCont).Picture = LoadPicture(App.Path & "\items\" & Trim(MyRsITM.Fields("icon_name")))
                            FrmKeyTouch.LblTec1(nCont).Caption = Trim(MyRsITM.Fields("shortname"))
                            If MyRsITM.Fields("qualitynr") <> 0 Then
                                FrmKeyTouch.ImgQ(nCont).Picture = LoadPicture(App.Path & "\quality\" & Trim(MyRsITM.Fields("qualityname")))
                                FrmKeyTouch.ImgQ(nCont).Visible = True
                            End If
                        Else
                            FrmKeyTouch.ImgTec(nCont).Picture = LoadPicture(App.Path & "\items\" & "pralcampo.bmp")
                        End If
                        nCont = nCont + 1
                    Next nF
                Else
                    FrmKeyTouch.ImgFam(nnnF).Picture = LoadPicture(App.Path & "\imggrpfam\" & MyRsORG.Fields("imgnoselect"))
                End If
                FrmKeyTouch.LblFam(nnnF).Caption = MyRsORG.Fields("text")
            End If
            
            
            MyRsORG.MoveNext
        Loop
        MyRsDes.MoveNext
    Loop
    
    nGrupoActual = 0
    nSubGrupoActual = 0
    
    FrmKeyTouch.Refresh
    
    MyRsDes.Close
    MyRsORG.Close
    MyBaseDES.Close
    Set MyBaseDES = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    Set MyRsDes = Nothing
    Set MyRsORG = Nothing

    '''Igualar TouchScale.mdb a Dbase.Mdb, Artículos por código (Sustituir PLU si necesario)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    FrmInactividad.TmrInactivo.Enabled = True
    
    'en final dia
    'nF = Tiques(False)
    '''''''''''''
    'Call FrmKeyTouch.CmdEnvTodo_Click
    'Load FrmKeyTouch
    'FrmKeyTouch.Show
    '''''''''''''
    
    If Trim(Command) <> "" And Trim(LCase(Command)) <> "hydra" Then
        
        If LCase(Trim(Command)) = "modificaciones" Then
            FrmAviso.LblAviso.Caption = "Transferencia Modificaciones a Touch-Scale"
            FrmAviso.Show
            FrmAviso.Refresh
            
            Call FrmKeyTouch.CmdEnvCambios_Click
        End If
        If LCase(Trim(Command)) = "todo" Then
            FrmAviso.LblAviso.Caption = "Transferencia Completa a Touch-Scale"
            FrmAviso.Show
            FrmAviso.Refresh
            
            Call FrmKeyTouch.CmdEnvTodo_Click
        End If
        If LCase(Trim(Command)) = "totales" Then
            FrmAviso.LblAviso.Caption = "Recogida Totales Touch-Scale"
            FrmAviso.Show
            FrmAviso.Refresh
            
            nF = Tiques(False)
            
        End If
        
        If LCase(Trim(Command)) = "limpiamdb" Then
            FrmAviso.LblAviso.Caption = "Eliminación datos TouchScale.Mdb"
            FrmAviso.Show
            FrmAviso.Refresh
            
            Call LimpiaMDB
            
        End If
        
        FrmAviso.Hide
        End
    Else
        
        'FrmAviso.LblAviso.Caption = "Carga Pantalla de datos...."
        
        'FrmAviso.Show
        'FrmAviso.Refresh
        
        Load FrmKeyTouch
        FrmKeyTouch.Show
    
    
    End If
    
End Sub
Public Sub LimpiaMDB()
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from balanzas")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from cfg")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from grp")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from htickets")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from items")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from ltickets")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from qlevels")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from subgroup")
    If Not MyRsORG.EOF Then
        MyRsORG.MoveFirst
        Do While Not MyRsORG.EOF
            MyRsORG.Delete
            MyRsORG.MoveNext
        Loop
    End If
    MyRsORG.Close
    Set MyRsORG = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    
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
'   Se borran datos de la Balanza
'   Se deben introducir tiques en DBASE.MDB con typtic=4
'   En principio la información de cabecera no es válida...
'   Se pensaba en utilizar las tablas de TouchScale.MDB pero se va a traspasar directamente a DBASE.MDB
Public Function Tiques(ByVal lLocal As Boolean) As Integer
Dim MyBaseORG As DAO.Database
Dim MyRsORG As DAO.Recordset

Dim MyBaseORG1 As DAO.Database
Dim MyRsORG1 As DAO.Recordset
Dim MyrsART As DAO.Recordset
Dim MyRsVEND As DAO.Recordset

Dim MyBaseBAL As DAO.Database
Dim MyRsBAL As DAO.Recordset
Dim sIp As String
Dim MyRsDes As New ADODB.Recordset
Dim myRsKEYS As New ADODB.Recordset
Dim myRsTMP As New ADODB.Recordset
Dim Cm As String
Dim nF As Integer
Dim nCont As Integer
Dim sS As String
Dim nImporte As Long
Dim sHour As String
Dim sFecha As String
'Dim nF As Integer


    lComunica = True
nF = FreeFile()
Open App.Path & "\LOGTOUCH.TXT" For Output As #nF
Print #nF, "Recogida Tiques Touch-Scale..."
    
    sHour = Format(Time, "hh:mm:ss")
    sFecha = Format(Date, "dd/mm/yy")
    
    Set MyBaseBAL = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsBAL = MyBaseBAL.OpenRecordset("select * from balanzas where ip<>'127-0-0-1' order by scalenr")
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where borrado=false")
    'Call AbrirTotales
    'If lTotalesAbierta Then
    '    nF = nF
    'End If
    'Cm = "select * from articulo where borrado=false"
    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic


    If Not MyRsBAL.EOF Then
        MyRsBAL.MoveFirst
        Do While Not MyRsBAL.EOF
        
            sIp = toIpAdd(Trim(MyRsBAL.Fields("ip")))
            
            If ElPing(sIp) = 1 Then
    
                Call AbrirMySQL(sIp)
                
                If lMySqlAbierta Then
                    For nCont = 0 To 1 '0=local, 1=Red
                        If nCont = 0 Then
Print #nF, "Recogida Tiques local Touch-Scale..."
                            
                            
                            Cm = "select * from pcscale.LocalLTickets"
                            MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF

                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG!seccion = CStr(8)
                                    MyRsORG!balanza = MyRsBAL!scalenr
                                    MyRsORG!vended = CStr(6099)
                                    MyRsORG!numlin = CStr(1)
                                    MyRsORG!lincan = CStr(0)
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG!importe = MyRsDes!price
                                        nImporte = (MyRsDes!price) * 100
                                    Else
                                        MyRsORG!importe = Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                        nImporte = 100 * (Round((MyRsDes!price) * (MyRsDes!Weight), 2))
                                    End If
                                    MyRsORG!hora = sHour
                                    MyRsORG!onoff = "Local"
                                    MyRsORG!stpetic = CStr(16)
                                    MyRsORG!codcli = CStr(0)
                                    MyRsORG!Fecha = sFecha
                                    MyRsORG!typtic = CStr(4)
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
                                    sS = sS & CStr(DigitoControl(sS))
                                    MyRsORG!codbar = sS
                                    MyRsORG!reserv1 = "No Enviado"
                                    MyRsORG!reserv2 = "*"
                                    MyRsORG!reserv3 = "SI"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sFecha
                                    MyRsORG1!Hour = sHour
                                    MyRsORG1!numbal = MyRsBAL!scalenr
                                    MyRsORG1!vendor = CStr(6099)
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG1!numlin = CStr(1)
                                    MyRsORG1!secc = CStr(8)
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    
                                    If MyrsART.EOF Then
                                        MyRsORG1!PLU = CStr(0)
                                        MyRsORG1!Item = ""
                                    Else
                                        MyRsORG1!PLU = CStr(MyrsART!PLU)
                                        MyRsORG1!Item = MyrsART!des_plu1
                                    End If
                                    MyRsORG1!price = MyRsDes!price
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG1!amount = MyRsDes!price
                                        MyRsORG1!weigth = 0
                                        MyRsORG1!units = 1
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                        MyRsORG1!Weight = MyRsDes!Weight
                                        MyRsORG1!units = 0
                                    End If
                                    MyRsORG1!stpetic = CStr(16)
                                    MyRsORG1!lincan = CStr(0)
                                    MyRsORG1!iva = 0
                                    MyRsORG1!typtic = CStr(4)
                                    
                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sFecha, "dd/mm/yy")
                                            !codi_ident = 188
                                            !ident_vend = 6099
                                            !secc_maqui = 8
                                            !Modo = 4
                                            !tiquets = 0
                                            !operacion = 0
                                            !cancela = 0
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !peso = 0
                                            !unidad = 0
                                            ' ****
                                            !pesetas = 0
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
                                            !supertot = 0
                                            !total = 0
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
                                            !operacion = !operacion + 1
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!price)
                                                !unidad = !unidad + 1
                                                !total = !total + (MyRsDes!price)
                                                !supertot = !total + MyRsDes!price
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                                !peso = !peso + (MyRsDes!Weight)
                                                !total = !total + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                                !supertot = !total + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    
                                    MyRsDes.MoveNext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                        
                            MyRsVEND.Clone
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        Else
                            
Print #nF, "Recogida Tiques Red Touch-Scale..."
                            
                            Cm = "select * from pcscale.LTickets"
                            MyRsDes.Open Cm, DbMysqli, adOpenDynamic, adLockOptimistic
                            Set MyRsORG = MyBaseORG.OpenRecordset("select * from cabecera")
                            Set MyRsORG1 = MyBaseORG.OpenRecordset("select * from tickets")
                            'Cm = "select * from cabecera"
                            'MyRsORG.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            'Cm = "select * from tickets"
                            'MyRsORG1.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                            
                            If Not MyRsDes.EOF Then
                                MyRsDes.MoveFirst
                                Do While Not MyRsDes.EOF
                                    
                                    MyRsORG.AddNew
                                    MyRsORG1.AddNew
                                    '''''''Registro Cabecera
                                    MyRsORG!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG!seccion = CStr(8)
                                    MyRsORG!balanza = MyRsBAL!scalenr
                                    MyRsORG!vended = CStr(6099)
                                    MyRsORG!numlin = CStr(1)
                                    MyRsORG!lincan = CStr(0)
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG!importe = MyRsDes!price
                                        nImporte = (MyRsDes!price) * 100
                                    Else
                                        MyRsORG!importe = Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                        nImporte = 100 * (Round((MyRsDes!price) * (MyRsDes!Weight), 2))
                                    End If
                                    MyRsORG!hora = sHour
                                    MyRsORG!onoff = "Red"
                                    MyRsORG!stpetic = CStr(16)
                                    MyRsORG!codcli = CStr(0)
                                    MyRsORG!Fecha = sFecha
                                    MyRsORG!typtic = CStr(4)
                                    sS = "20" & Format(MyRsDes!Item, "00000") & Format(nImporte, "00000")
                                    sS = sS & CStr(DigitoControl(sS))
                                    MyRsORG!codbar = sS
                                    MyRsORG!reserv1 = "No Enviado"
                                    MyRsORG!reserv2 = "*"
                                    MyRsORG!reserv3 = "SI"
                                    '''''''Registro línea
                                    MyRsORG1!Date = sFecha
                                    MyRsORG1!Hour = sHour
                                    MyRsORG1!numbal = MyRsBAL!scalenr
                                    MyRsORG1!vendor = CStr(6099)
                                    MyRsORG1!nume = CStr(MyRsDes!idhticket)
                                    MyRsORG1!numlin = CStr(1)
                                    MyRsORG1!secc = CStr(8)
                                    MyRsORG1!Code = CStr(MyRsDes!Item)
                                    Set MyrsART = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false")
                                    'MyrsART.FindFirst "codigo=" & CStr(MyRsDes!Item)
                                    'Cm = "select * from articulo where codigo=" & CStr(MyRsDes!Item) & " and borrado=false"
                                    'MyrsART.Open Cm, DbTotales, adOpenDynamic, adLockOptimistic
                                    If MyrsART.EOF Then
                                        MyRsORG1!PLU = CStr(0)
                                        MyRsORG1!Item = ""
                                    Else
                                        MyRsORG1!PLU = CStr(MyrsART!PLU)
                                        MyRsORG1!Item = MyrsART!des_plu1
                                    End If
                                    MyRsORG1!price = MyRsDes!price
                                    If MyRsDes!saleform = 0 Then
                                        MyRsORG1!amount = MyRsDes!price
                                        MyRsORG1!weigth = 0
                                        MyRsORG1!units = 1
                                    Else
                                        MyRsORG1!amount = Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                        MyRsORG1!Weight = MyRsDes!Weight
                                        MyRsORG1!units = 0
                                    End If
                                    MyRsORG1!stpetic = CStr(16)
                                    MyRsORG1!lincan = CStr(0)
                                    MyRsORG1!iva = 0
                                    MyRsORG1!typtic = CStr(4)
                                    

                                    MyRsORG.Update
                                    MyRsORG1.Update
                                    
                                    'Total vendedor 6099
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                        If .EOF Then
                                            .AddNew
                                            !Fecha = Format(sFecha, "dd/mm/yy")
                                            !codi_ident = 188
                                            !ident_vend = 6099
                                            !secc_maqui = 8
                                            !Modo = 4
                                            !tiquets = 0
                                            !operacion = 0
                                            !cancela = 0
                                            '**********
                                            ' Nombre del vendedor
                                            '***********
                                            !Nombre = "SS-Touch"
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            !peso = 0
                                            !unidad = 0
                                            ' ****
                                            !pesetas = 0
                                            ' ****
                                            ' Cancelada o no
                                            ' ****
                                            !totcance = 0
                                            !supertot = 0
                                            !total = 0
                                            .Update
                                        End If
                                    End With
                                    Set MyRsVEND = MyBaseORG.OpenRecordset _
                                    ("select * from gtvend where ident_vend=" & CStr(6099) & _
                                    " and codi_ident=" & CStr(188) & " and modo=" & CStr(4) & _
                                    " and fecha=#" & Format(sFecha, "mm/dd/yy") & "#")
                                    With MyRsVEND
                                           
                                            .Edit
                                            !tiquets = !tiquets + 1
                                            !operacion = !operacion + 1
                                            ' ****
                                            'pesado o unidades
                                            ' ****
                                            If MyRsDes!saleform = 0 Then
                                                !pesetas = !pesetas + (MyRsDes!price)
                                                !unidad = !unidad + 1
                                                !total = !total + (MyRsDes!price)
                                                !supertot = !total + MyRsDes!price
                                            Else
                                                !pesetas = !pesetas + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                                !peso = !peso + (MyRsDes!Weight)
                                                !total = !total + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                                !supertot = !total + Round((MyRsDes!price) * (MyRsDes!Weight), 2)
                                            End If

                                            .Update
                                    End With
                                    ''''''''''''''''''''
                                    
                                    MyRsDes.MoveNext
                                Loop
                            End If
                            
                            On Error Resume Next
                            MyRsDes.Close
                            Set MyRsDes = Nothing
                            
                            MyRsORG.Close
                            MyRsORG1.Close
                            Set MyRsORG = Nothing
                            Set MyRsORG1 = Nothing
                            MyRsVEND.Close
                            Set MyRsVEND = Nothing
                            On Error GoTo 0
                            
                        End If
                    
                    Next nCont
                            
                    'Borrado de los datos de Totales en Balanza
Print #nF, "Borrado Tiques Touch-Scale..."
                    On Error Resume Next
                    Cm = "truncate pcscale.paytickets"
                    DbMysqli.Execute Cm
                    Cm = "truncate pcscale.HTickets"
                    DbMysqli.Execute Cm
                    Cm = "truncate pcscale.LTickets"
                    DbMysqli.Execute Cm
                    
                    'Cm = "delete from pcscale.LocalHTickets"
                    'DbMysqli.Execute Cm
                    'Cm = "delete from pcscale.LocalLTickets"
                    'DbMysqli.Execute Cm
                    'Cm = "delete from pcscale.localnumtickets"
                    'DbMysqli.Execute Cm
                    
                    'Intento reiniciar numeración
                    Cm = "update pcscale.status set resetticketnumber=1 where id=1"
                    DbMysqli.Execute Cm
                    Cm = "update pcscale.erasetickets set dattim=now()"
                    DbMysqli.Execute Cm
                    Cm = "truncate pcscale.numtickets"
                    DbMysqli.Execute Cm
                    ''''''''''''''''''''''''''''''
                    On Error GoTo 0
                End If
            
            End If
        
            If lMySqlAbierta Then
                DbMysqli.Close
                Set DbMysqli = Nothing
            End If
            
            MyRsBAL.MoveNext
        
        Loop
        
    End If

    MyRsBAL.Close
    Set MyRsBAL = Nothing
    
    MyBaseBAL.Close
    Set MyBaseBAL = Nothing
    
    MyrsART.Close
    Set MyrsART = Nothing

    MyBaseORG.Close
    Set MyBaseORG = Nothing
    'If lTotalesAbierta Then
    '    DbTotales.Close
    '    Set DbTotales = Nothing
    'End If
    
    Close #nF
    
    lComunica = False
    
End Function


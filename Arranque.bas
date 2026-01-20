Attribute VB_Name = "Arranque"
Option Explicit
Public Multi_Salir As Boolean
Public HayMulti As Boolean
Public Hydra_INI As String
Public TiendaActual As Integer
Public Const CDKEY_USER = False
Public nQueImagen As Integer
Public sQuePathImagen As String
Public sQueNombreImagen As String
Public sMyTienda As String
Public lJpg As Boolean
Public lFornes As Boolean
Public lBoka As Boolean
Public lBelRos As Boolean
Public lAgora As Boolean
Public nGTOrd As Long

Private Const adReadAll = -1
Private Const adSaveCreateOverWrite = 2
Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adWriteChar = 0

Private Sub UTF8toANSI(ByVal UTF8FName, ByVal ANSIFName)
    Dim strText

    With CreateObject("ADODB.Stream")
        .Open
        .type = adTypeBinary
        .LoadFromFile UTF8FName
        .type = adTypeText
        .Charset = "utf-8"
        strText = .ReadText(adReadAll)
        .Position = 0
        .SetEOS
        .Charset = "_autodetect" 'Use current ANSI codepage.
        .WriteText strText, adWriteChar
        .SaveToFile ANSIFName, adSaveCreateOverWrite
        .Close
    End With
End Sub

Public Function GetShopPath(ByVal nT As Integer) As String
    Dim s As String
    Dim db2 As dao.Database
    Dim rstdb2 As dao.Recordset
    
    Set db2 = OpenDatabase(App.Path & "\tiendas.mdb")
    Set rstdb2 = db2.OpenRecordset("select * from grupos where tienda=" & CStr(nT))
    If Not rstdb2.EOF Then
       
        s = rstdb2.Fields(4)
    Else
        s = ""
    End If
    Set rstdb2 = Nothing
    db2.Close
    
    GetShopPath = s
End Function
Private Sub pasosLog(ByVal n As Integer)
    CadenadeLog CStr(n)
End Sub
Public Sub transferSHOP(Optional ByVal nShop As Integer)
    'transfer
    Dim bucle As Long
    Dim dB1 As dao.Database
    Dim db2 As dao.Database
    Dim db3 As dao.Database
    Dim RstdB1 As dao.Recordset
    Dim rstdb2 As dao.Recordset
    Dim rStdB21 As dao.Recordset
    Dim sMiBase As String
    Dim nMiGrupo As Integer
    Dim nFieldsCount As Integer
    Dim nBucle As Integer
    Dim snonF As String
    Dim nT, nC As Integer
    Dim snT() As String
    Dim sntIni() As String
    Dim snTCFG() As String
    Dim snTienda() As String
    Dim MiProceso As Long
    Dim MiResultado As Long
    Dim sQueEjecutar As String
    
    If Not IsNull(nShop) Then
        TiendaActual = nShop
        CadenadeLog "función TransferShop:" & CStr(nShop)
        Set db2 = OpenDatabase(App.Path & "\tiendas.mdb")
        Set rstdb2 = db2.OpenRecordset("select * from grupos where tienda=" & CStr(nShop))
        If Not rstdb2.EOF Then
            CadenadeLog "TransferShop llamada con parámetro...Localizada tienda Origen.:" & CStr(nShop)
            Base_General = Trim(rstdb2.Fields("path")) & "dbasetouch.mdb"
            Hydra_INI = Trim(rstdb2.Fields("path")) & "hydratouch.ini"
            Call LeerParametrosInicio
        Else
            Exit Sub
        End If
        rstdb2.Close
        db2.Close
    End If
    
    nT = 0

    If Dir(App.Path & "\transfertable.cfg") = "" Then
        Set dB1 = AbrirBase 'abrirbase 'OpenDatabase(Base_General)
        CadenadeLog "Se Genera Global.Dat con:" & Base_General
        Call Exporta_GlobalDat_Art(dB1)
        dB1.Close
    End If
    
    Set db2 = OpenDatabase(App.Path & "\tiendas.mdb")
    If Not IsNull(nShop) Then
        TiendaActual = nShop
    End If
    Set rstdb2 = db2.OpenRecordset("select * from grupos where tienda=" & CStr(TiendaActual))

'Call pasosLog(1)
    
    If Not rstdb2.EOF Then
       
       nMiGrupo = rstdb2.Fields("grp")
       rstdb2.Close
       If TiendaActual <> 999 Then '1.7.23 c2f
        Set rstdb2 = db2.OpenRecordset("select * from grupos where grp<>0 order by tienda")
       Else
        Set rstdb2 = db2.OpenRecordset("select * from grupos order by tienda")
       End If
       If Not rstdb2.EOF Then
       
'Call pasosLog(2)

             rstdb2.MoveFirst
             
             Do While Not rstdb2.EOF
                
'Call pasosLog(3)

                If rstdb2.Fields("tienda") <> TiendaActual And (rstdb2.Fields("grp") = nMiGrupo Or (TiendaActual = 999)) Then '1.7.23 c2f si es la tienda 999 siempre se transfiere.
                    nT = nT + 1
                    ReDim Preserve snT(nT)
                    ReDim Preserve sntIni(nT)
                    ReDim Preserve snTCFG(nT)
                    ReDim Preserve snTienda(nT)
                    snT(nT) = rstdb2.Fields(4) & "dbasetouch.mdb"
                    sntIni(nT) = rstdb2.Fields(4) & "hydratouch.ini"
                    snTCFG(nT) = rstdb2.Fields(4) & "tienda.cfg"
                    If rstdb2.Fields("tienda") < 100 Then
                        snTienda(nT) = Format(rstdb2.Fields("tienda"), "00")
                    Else
                        snTienda(nT) = Format(rstdb2.Fields("tienda"), "000")
                    End If
                    'Base_General = rStdB2.Fields(4) & "dbase.mdb"
                    
                    'Call Importa_Global_Dat
                
                End If
          
                rstdb2.Movenext
             Loop
          
          
       End If
       
    End If
    
    rstdb2.Close
    db2.Close

'Call pasosLog(4)

    For nC = 1 To nT
'Call pasosLog(5)
        sQueEjecutar = sTipoDestino(snTCFG(nC))
        If LCase(WHOIAM) = LCase(sQueEjecutar) Then
            If Dir(App.Path & "\transfertable.cfg") = "" Then
                Base_General = snT(nC)
                Hydra_INI = sntIni(nC)
                Call LeerParametrosInicio
                CadenadeLog "Se capturan datos en:" & Base_General
                sMyTienda = snTienda(nC)
                Call Importa_Global_Dat
            Else
                Call transferirTABLA(Base_General, snT(nC))
                
            End If
        Else
            MiResultado = STILL_ACTIVE
            MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\" & sQueEjecutar & " /" & snTienda(nC) & "GLOBAL", vbHide))
            Do While MiResultado = STILL_ACTIVE
                GetExitCodeProcess MiProceso, MiResultado
                Do_Events
            Loop
        End If
    Next nC

    If Dir(App.Path & "\transfertable.cfg") <> "" Then
        Set dB1 = AbrirBase 'abrirbase 'OpenDatabase(Base_General)
        CadenadeLog "Se Genera Global.Dat con:" & Base_General
        Call Exporta_GlobalDat_Art(dB1)
        dB1.Close
    End If


End Sub
Public Function sTipoDestino(ByVal sCfg As String) As String
    Dim nFich As Integer
    Dim s As String
    
    If LCase(WHOIAM) = "hydra.exe" And Dir(App.Path & "\hydratouch.exe") = "" Then
        sTipoDestino = WHOIAM
        Exit Function
    End If
    If LCase(WHOIAM) = "hydratouch.exe" And Dir(App.Path & "\hydra.exe") = "" Then
        sTipoDestino = WHOIAM
        Exit Function
    End If
    
    nFich = FreeFile()
    Open sCfg For Input As #nFich
    Do While Not EOF(nFich)
        Line Input #nFich, s
    Loop
    Close #nFich
    If CBool(s) = True Then
        sTipoDestino = "hydratouch.exe"
    Else
        sTipoDestino = "hydra.exe"
    End If
End Function




Sub Main()
Dim logntq As Integer
Dim FServer As Integer
Dim Buffer As String
Dim Archivo As Integer
Dim Elmaximo As String
Dim Hiper As String
Dim bucle As Long
Dim dB1 As dao.Database
Dim db2 As dao.Database
Dim db3 As dao.Database
Dim RstdB1 As dao.Recordset
Dim rstdb2 As dao.Recordset
Dim rStdB21 As dao.Recordset
Dim sMiBase As String
Dim nMiGrupo As Integer
Dim nFieldsCount As Integer
Dim nBucle As Integer
Dim snonF As String
Dim nMyFich As Integer
Dim sNumdec As String
Dim sNum100 As String
Dim myRec As dao.Recordset
Dim nf As Integer
Dim cm As String
'''''''''''''''''''
'cm = verLF("|d    d  d  |  hhhsq| dddd |...||||[^]")
'cm = cm
'Base_General = App.Path & "\t999\dbasetouch.mdb"
'TiendaActual = 999
'Call transferSHOP

'TiendaActual = 0
'Base_General = App.Path & "\dbasetouch.mdb"
'Call cmcSchemaADO_Click

''Set dB1 = AbrirBase
'lFornes = True
''Call export_BL(True)
'Call Bal_Epelsa_txt

'cm = "CREATE TABLE nutrition (card long,portion_weight double,energy_kj_100 double," & _
'  "energy_kj_portion double,energy_kcal_100 double,energy_kcal_portion double," & _
'  "fat_100 double,fat_portion double,saturates_100 double,saturates_portion double," & _
'  "mono_unsaturates_100 double,mono_unsaturates_portion double,polyunsaturates_100 double," & _
'  "polyunsaturates_portion double,carbohydrate_100 double,carbohydrate_portion double," & _
'  "sugars_100 double,sugars_portion double,polyols_100 double," & _
'  "polyols_portion double,starch_100 double,starch_portion double," & _
'  "fibre_100 double,fibre_portion double,protein_100 double,protein_portion double," & _
'  "salt_100 double,salt_portion double)"

'dB1.Execute cm
'dB1.Close

'dB1.Execute "drop table text15"
'dB1.Close
'''''''''''''''''''
'Call transferirTABLA(App.Path & "\dbasetouch.mdb", App.Path & "\t02\dbasetouch.mdb")
'''''''''''''''''''
   '*************************
   ' detecta instancia previa
   '*************************
   Detecta_Instancia
    'c2f 1.1.1
    'If Dir(App.Path & "\nover1.txt") <> "" Then
    '    For Bucle = 1 To 3000
    '        Sleep (2)
    '        DoEvents
    '    Next Bucle
    'End If
    
    'If Dir(App.Path & "\mysql-connector-odbc-5.1.13-win32.msi") <> "" Then
    '    'msiexec.exe /i mysql-connector-odbc-5.1.13-win32.msi /quiet
    '    Shell App.Path & "\mysql-connector-odbc-5.1.13-win32.msi /quiet", vbNormalFocus
    '    Sleep (2000)
    '    Kill App.Path & "\mysql-connector-odbc-5.1.13-win32.msi"
    'End If
    
    'UTF8toANSI App.Path & "\artcon.dat", App.Path & "\artcon.new"
    'nMyFich = FreeFile()
    'Open App.Path & "\artcon.n" For Output As #nMyFich
    '
    'nf = FreeFile()
    'Open App.Path & "\artcon.dat" For Input As #nf
    'Do While Not EOF(nf)
    '    Line Input #nf, snonF
    '    snonF = ConvertUtf8BytesToStringnW(DecodeBase64(snonF))
    '    snonF = snonF
    '    Print #nMyFich, snonF
        
    '
    '
    'Loop
    'Close #nf
    'Close #nMyFich
    
    '''''
    '3.8.3-5 se borra al arrancar...
    If Dir(App.Path & "\loghydra.txt") <> "" Then
        Kill App.Path & "\loghydra.txt"
    End If
    '''''''
    var.miver = "Release" & " " & tools.version_app
    HayComandos = False
    Miruta = App.Path
    Hydra_INI = App.Path & "\hydratouch.ini"
    LockBase = False
    
    Call carga_data_Vitamin
    
    'If Dir(App.Path & "\upper.txt") <> "" And Dir(App.Path & "\test.tst") <> "" Then
    '    lUpperNW = True
    'End If
    If Dir(App.Path & "\upper.txt") <> "" Then
        '1.8.3
        'nf = FreeFile()
        'Open App.Path & "\test.tst" For Output As #nf
        'Close #nf
        ''''''
        lUpper = True
    End If
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        lTransPlu = True
    End If
    lhydraexe = False
    If Dir(App.Path & "\hydra.exe") <> "" Then
        lhydraexe = True
    End If
    
    If (Dir(App.Path & "\cocobio.cfg") <> "") Or (Dir(App.Path & "\asigfam.cfg") <> "") Then
        lCOCOBIO = True
    End If
    If Dir(App.Path & "\alba.txt") <> "" Then
        lAlba = True
    End If
    If Dir(App.Path & "\upperjpg.txt") <> "" Then
        lJpg = True
    End If
    lAcumulaSS = False
    lAcumulaET = False
    If Dir(App.Path & "\acumulass") <> "" Then
        lAcumulaSS = True
    End If
    If Dir(App.Path & "\acumulaet") <> "" Then
        lAcumulaET = True
    End If
    lFornes = False
    If Dir(App.Path & "\etc\Bal_Epelsa.txt.fia") <> "" Then
        lFornes = True
    End If
    lBoka = False
    If Dir(App.Path & "\etc\boka.txt.fet") <> "" Then
        lBoka = True
    End If
    If Dir(App.Path & "\etc\fin_dbal.fet") <> "" Then lBelRos = True
    If Dir(App.Path & "\agora.txt") <> "" Then lAgora = True
    
    '*******************************
    ' Inicia / Borra log del sistema
    '*******************************
    On Error Resume Next
    If Dir(Miruta & "\loghydratotales.txt") <> "" Then
        If FileLen(Miruta & "\loghydratotales.txt") > 48638000 Then
            
            FileCopy App.Path & "\loghydratotales.txt", App.Path & "\" & Format(Date, "ddmmyyhhmmss") & "loghydratotales.txt"
            Sleep (100)
            logntq = FreeFile()
            Open Miruta & "\loghydratotales.txt" For Output As logntq
        Else
            logntq = FreeFile()
            Open Miruta & "\loghydratotales.txt" For Append As logntq
        End If
    Else
        logntq = FreeFile()
        Open Miruta & "\loghydratotales.txt" For Output As logntq
    End If
    Print #logntq, Now() & " Program HYDRATOUCH Started *************************************"
    Close logntq
    
    If Dir(Miruta & "\loghydra.txt") <> "" Then
        If FileLen(Miruta & "\loghydra.txt") > 1638000 Then
            ReduceLogHydra
            logntq = FreeFile()
            Open Miruta & "\loghydra.txt" For Output As logntq
        Else
            logntq = FreeFile()
            Open Miruta & "\loghydra.txt" For Append As logntq
        End If
    Else
        logntq = FreeFile()
        Open Miruta & "\loghydra.txt" For Output As logntq
    End If
    Print #logntq, Now() & " Program HYDRATOUCH Started *************************************"
    Close logntq
    logntq = FreeFile()
    Open App.Path & "\laststart" For Output As #logntq
    Close #logntq
    On Error GoTo 0
    '******************
    ' Fin de log del sistema
    '******************
    If Dir(App.Path & "\40l.txt") = "" And Dir(App.Path & "\39l.txt") = "" And Dir(App.Path & "\5l.txt") = "" Then
        logntq = FreeFile()
        Open App.Path & "\40l.txt" For Output As #logntq
        Close #logntq
    End If
    
    '1.0.3 --> Leclerc
    lClR = False
    sPathMaj = ""
    '''''''''''''
    
    lEsExport = False
    lesUpdate = False
    
    If Dir(App.Path & "\cmdlinetouch.cfg") = "" Then
        CrearParametrosDefecto
    End If
    
    If Trim(Command) = "" And Dir(App.Path & "\hydrarun.txt") <> "" Then Kill App.Path & "\hydrarun.txt"
    On Error GoTo 0
    If Dir(App.Path & "\servidoratouch.ini") <> "" Then
        On Error Resume Next
        Kill App.Path & "\servidoratouch.ini"
        If Err.Number <> 0 Then
            MsgBox App.Path & CargaCadena(1381), vbCritical
            End
        End If
        On Error GoTo 0
    End If
   
   '*****************************
   ' detecta programa multitienda
   '*****************************
    If Trim(Command) = "" Then 'restaurar
        If Dir(App.Path & "\hydramulti.exe") <> "" And Dir(App.Path & "\multistart.ord") = "" Then
            CadenadeLog "Se detecta en MAIN() hydramulti.exe. Se arranca."
            Shell App.Path & "\hydramulti.exe " & Command, vbNormalFocus
            End
        End If
    End If
   
    On Error Resume Next
    logntq = FreeFile()
    Open App.Path & "\hydrarun.txt" For Output As #logntq
    Close #logntq

    'If CDKEY_USER Then Check_CDKEY
    OrdenMulti = "00"
    
        '***************************
        ' Crea carpetas del programa
        '***************************
        CreaCarpetas
        '***************************
        ' comprueba si hay ficheros
        ' de plugin en etc
        '***************************
        frmControl.File1.Path = App.Path & "\etc"
        frmControl.File1.FileName = "autoplugin.*"
        frmControl.File1.Refresh
        For bucle = 0 To frmControl.File1.ListCount - 1
             If Len(frmControl.File1.List(bucle)) = 17 Then
                 If IsNumeric(Right(frmControl.File1.List(bucle), 6)) Then
                     If Val(Right(frmControl.File1.List(bucle), 6)) < Val(Format(Now, "yymmdd")) Then
                         Kill App.Path & "\etc\" & frmControl.File1.List(bucle)
                     End If
                 End If
             End If
        Next bucle
        If Dir(App.Path & "\bar.jpg") <> "" Then
             On Error Resume Next
             Form2.Imagen.Picture = LoadPicture(App.Path & "\bar.jpg")
             On Error GoTo 0
        End If

   '***********
   ' De momento
   '***********
   Miruta = App.Path
   Base_General = Miruta & "\dbasetouch.mdb"
   
   
    'Set dB1 = dao.OpenDatabase(Base_General)
    'Set myRec = dB1.OpenRecordset("select * from fam_code where codi_fam=1")
    'If myRec.EOF Then
    '    myRec.AddNew
    '    myRec.Fields("codi_fam") = 1
    '    myRec.Fields("codi_ident") = 0
    '    myRec.Fields("secc_maqui") = 0
    '    myRec.Fields("posicion") = 0
    '    myRec.Fields("txt_fam") = "DPT. 1"
    '    myRec.Fields("imagen") = ""
    '    myRec.Update
    'End If
    'myRec.Close
    'Set myRec = Nothing
    'dB1.Close
    'Set dB1 = Nothing
   
   'TouchScale ... prueba captura paises...
   'Call capturaPaises
   ''''''''''''''''''''''''''''''''''''''''
   
   
   If Dir(App.Path & "\cocobio.cfg") Then
    Alta_Familias_Counter_Fichero ("cocobio.cfg")
   End If
   If Dir(App.Path & "\asigfam.cfg") Then
    Alta_Familias_Counter_Fichero ("asigfam.cfg")
   End If
   
   Check_Multi
   
   'Set db3 = AbrirBase
   'cm = "delete from text15 where codigo not in (select codigo from articulo)"
   'db3.Execute cm
   'cm = "delete from lintxt2040 where codigo not in (select codigo from articulo)"
   'db3.Execute cm
   'cm = "delete from text15 where"
   'db3.Close
   'Set db3 = Nothing
   
   '*******************************
   ' Parametros generales de inicio
   'CurDir
   cuenTaTqt = 1
   '************
   ' nro. máximo de tiquets a recoger AH / Path de Red / ASCII
   '**********************************************************
   NumMaxTiquets = 0
   On Error GoTo siguiente
   If Dir(Miruta & "\asorden\maxtiq.ord") <> "" Then
        Archivo = FreeFile()
        Open Miruta & "\asorden\maxtiq.ord" For Input As #Archivo
        If Not EOF(Archivo) Then Line Input #Archivo, Elmaximo
        Close #Archivo
        If Trim(Elmaximo) <> "" Then
            If IsNumeric(Elmaximo) Then
                NumMaxTiquets = Val(Elmaximo)
            End If
        End If
    End If
siguiente:
   On Error GoTo 0
   ' chequea configuraciones especiales ... Si procede
   Clientes_Especiales
   '******************
   ' Comprueba si existe el fichero de configuración general, y si no es así lo crea
   If (Command <> "") Then
    If UCase(Mid(Command, 1, 5)) = "TRANS" Then
        LeerLineaComandos
    Else
        Crea_Fichero_INI
    End If
   Else
       Crea_Fichero_INI
   End If

    '***************
    ' Si el programa se cerró de forma abrupta trata de reparar la base de datos
    '***********************************
    'c2f 1.1.1
    'If (Not HayMulti) Or (HayMulti And Base_General <> "") Then
    '     If Dir(App.Path & "\laststart") <> "" Then Repara_Base
    'End If
    '*****
    
    
    ' Crea la base si no existe (o la retoma de la dejada por el master/slave)
    If (Not HayMulti) Or (HayMulti And Base_General <> "") Then
         If Dir(Base_General) = "" Then
              If Dir(Miruta & "\new.dbasetouch.mdb") <> "" Then
                  On Error Resume Next
                  Name Miruta & "\new.dbasetouch.mdb" As Miruta & "\dbasetouch.mdb"
                  On Error GoTo 0
              Else
                  Form2.Show
                  Do_Events
                  If HayMulti And TiendaActual = 0 Then
                  Else
                     CadenadeLog "Creación Base General. Tienda:" & CStr(TiendaActual) & " Nombre:" & Base_General
                     Crear_Base_Maestra
                  End If
                  Form2.Hide
              End If
         End If
    End If
    
    
    '*********************************************************
    ' comprueba si existe una base de datos antigua de formato
    ' incompatible MNG
    '*****************
    If (Not HayMulti) Or (HayMulti And Base_General <> "") Then
         CadenadeLog "Se revisa si necesario Actualizar: " & Base_General
         FrmCambiaFormato.CambiarFormatoBase
         If Dir(App.Path & "\soloformato") <> "" Then
            Kill App.Path & "\soloformato"
            salir_programa
         End If
    End If
    '*****
    
    If Dir(Miruta & "\new.dbasetouch.mdb") <> "" Then
         On Error Resume Next
         Kill Miruta & "\new.dbasetouch.mdb"
         On Error GoTo 0
    End If
    '************************************
    ' Si procede borra totales acumulados
    '************************************
    'If ((Not HayMulti) Or (HayMulti And Base_General <> "")) And Dir(App.Path & "\nover1.txt") = "" And Command = "" Then
    If ((Not HayMulti) Or (HayMulti And Base_General <> "")) And Command = "" Then
        If DiasBorrado > 0 Then
            'c2f 1.1.1
            'If Dir(App.Path & "\nover1.txt") = "" Then
                Load frmWAIT
                frmWAIT.Label1.Caption = "Volcado a Histórico..."
                frmWAIT.Label2.Caption = ""
                frmWAIT.Visible = True
            'End If
            Do_Events
            'If Dir(App.Path & "\nover1.txt") = "" Then
                Interfaz.exe_copion "procTotalHistorico" '''''''''''''''''''''''''''''''''''''''cas.v170
                Unload frmWAIT
                Set frmWAIT = Nothing
            'End If
            Load frmMantenimiento
            frmMantenimiento.MostrarMsg = False
            frmMantenimiento.Opttotal(0).Value = True
            If Dir(App.Path & "\etc\l_date.txt") = "" Then
                frmMantenimiento.TxtDesde(0).TexT = "01/01/80"
                frmMantenimiento.TxtDesde(1).TexT = AdaptaFecha(Format((Now - DiasBorrado), "dd/mm/yy"))
            Else
                frmMantenimiento.TxtFecha(0).TexT = "01"
                frmMantenimiento.TxtFecha(1).TexT = "01"
                frmMantenimiento.TxtFecha(2).TexT = "80"
                frmMantenimiento.TxtFecha(3).TexT = left(Format((Now - DiasBorrado), "dd/mm/yy"), 2)
                frmMantenimiento.TxtFecha(4).TexT = Mid(Format((Now - DiasBorrado), "dd/mm/yy"), 4, 2)
                frmMantenimiento.TxtFecha(5).TexT = Right(Format((Now - DiasBorrado), "dd/mm/yy"), 2)
            End If
            
            lNoSi = True
            frmMantenimiento.cmdespere_Click
            lNoSi = True
            
            frmMantenimiento.MostrarMsg = True
            'c2f 1.7.2
            Unload frmMantenimiento
            Set frmMantenimiento = Nothing
            If UCase(Trim(Command)) = "HISTORICO" Then
                If AhorraMas Then
                    If Dir(App.Path & "\laststart") <> "" Then
                        Kill App.Path & "\laststart"
                    End If
                    If Dir(App.Path & "\hydrarun.txt") <> "" Then
                        Kill App.Path & "\hydrarun.txt"
                    End If
                    End
                Else
                    salir_programa
                End If
            End If
        End If
        '***************************************************
        ' Si procede borra bases de datos antiguas de Backup
        '***************************************************
        If BorrarBackup > 0 Then
           FrmResturarBackup.Borrar_Copias BorrarBackup
           Unload FrmResturarBackup
        End If
    End If
    
    
    '****************
    ' Lee los posibles parámetros de Linea de comandos
    '********************
    'c2f/caspiunza
    If OrdenMulti <> "00" Then
         FrmHlink.RealizaAccion OrdenMulti
         
         
         'c2f DIA se pretende traspasar lo cambios en artículos
         'a tiendas del mismo Grupo
         
         If (OrdenMulti = "21" Or OrdenMulti = "31") And (Dir(App.Path & "\tiendas.mdb") <> "") And _
             (Dir(App.Path & "\shoptras.txt") <> "") Then
            
            Call transferSHOP(TiendaActual)
            
            
            salir_programa
         
            
         End If
         
         salir_programa
    
    End If
   

    'Call CargaPaises
    'If lUpperNW = False Then
    If lUpper = False And lCOCOBIO = False Then
        Set dB1 = dao.OpenDatabase(Base_General)
        Set myRec = dB1.OpenRecordset("select * from fam_code where codi_fam=1")
        If myRec.EOF Then
            myRec.AddNew
            myRec.Fields("codi_fam") = 1
            myRec.Fields("codi_ident") = 0
            myRec.Fields("secc_maqui") = 0
            myRec.Fields("posicion") = 0
            myRec.Fields("txt_fam") = "DPT. 1"
            myRec.Fields("imagen") = ""
            myRec.Fields("etiqueta") = ""
            myRec.Update
        End If
        myRec.Close
        Set myRec = Nothing
        dB1.Execute "ALTER TABLE codbar ALTER COLUMN CB_EV1 TEXT(24)"
        dB1.Execute "ALTER TABLE codbar ALTER COLUMN CB_EV2 TEXT(24)"
        dB1.Close
        Set dB1 = Nothing
    End If
    
    If Dir(App.Path & "\adaptafam") <> "" Then
        If Dir(App.Path & "\adaptado") = "" Then
            Dim nMymax
            Set dB1 = dao.OpenDatabase(Base_General)
            Set myRec = dB1.OpenRecordset("select max(index) as mymax from fam_code")
            If Not myRec.EOF Then
                nMymax = myRec.Fields("mymax")
            Else
                nMymax = 0
            End If
            myRec.Close
            Set myRec = Nothing
            If IsNull(nMymax) Then nMymax = 0
            Set myRec = dB1.OpenRecordset("select * from fam_code where index<>0 order by codi_fam")
            
            If Not myRec.EOF Then
                Do While Not myRec.EOF
                    If Not IsNull(myRec.Fields("index")) Then
                        dB1.Execute "update familias set index=" & CStr(myRec.Fields("index")) & " where codi_fam=" & CStr(myRec.Fields("codi_fam")) & " and index=0"
                    Else
                        nMymax = nMymax + 1
                        dB1.Execute "update familias set index=" & CStr(nMymax) & " where codi_fam=" & CStr(myRec.Fields("codi_fam")) & " and index=0"
                    End If
                    myRec.Movenext
                Loop
            End If
            myRec.Close
            Set myRec = Nothing
            dB1.Close
            Set dB1 = Nothing
            FileCopy App.Path & "\adaptafam", App.Path & "\adaptado"
        End If
        Kill App.Path & "\adaptafam"
    End If
    
   
   
''''''
'Test comandos
''''''
'LeerLineaComandos "/01END_DAY"
''''''
   
   '****************
   ' Lee los posibles parámetros de Linea de comandos
   '********************
   LeerLineaComandos
   
   
   
  '******************
  ' Activa Recogida contínua de tiquets
  '******************
  'If descAuto = True Then
  '    frmControl.timetqt.Enabled = True
  '    '1.0.4 --> activar
  '    PausaTiquets = False
  '    frmEpelsa.BloquearMenues
  '    frmEpelsa.CmdComunicaciones(3).Caption = CargaCadena(704)
  '    'c2f 1.7.2
  '    frmEpelsa.CmdComunicaciones(3).Enabled = True
  '    '''''''''''''''''''
  'End If
 
 UsuarioActual.Nombre = ""
 lCogeTiquet = False

 If Not HayMulti Then
      If HaySeguridad And Not IconificarInicio Then
          If EncontrarUsuario = False Then
              '4.5.28
                 UsuarioActual.Nombre = "opcen"
                 UsuarioActual.password = ""
                 Call Secure_Login(UsuarioActual, True)
              '''''''
              
              If Dir(App.Path & "\hydrarun.txt") <> "" Then
                  Kill App.Path & "\hydrarun.txt"
                  End
              End If
          End If
      End If
    Dim spa As String
    Dim Arch As Integer
    spa = App.Path & "\"
    nGTOrd = 1
    If Dir(spa & "ngtord") <> "" Then
        Arch = FreeFile()
        Open spa & "ngtord" For Input As #Arch
        Line Input #Arch, cm
        If Val(cm) > 9999 Then
            cm = "0001"
            nGTOrd = Val(cm)
            Close #Arch
            Arch = FreeFile()
            Open spa & "ngtord" For Output As #Arch
            Print #Arch, cm
        End If
        Close #Arch
    End If
      
 End If
 '//////////////////////////////////////////////////////////////////cas.v130
  If (Dir(Miruta & "\marcaTqt.txt") <> "") Then
      marca_tqt_exportacion
      Kill Miruta & "\marcaTqt.txt"
  End If
 '//////////////////////////////////////////////////////////////////////////
 'If Dir(App.Path & "\ata34") = "" Then
 '   Set dB1 = dao.OpenDatabase(Base_General)
 '   For nF = 21 To 40
 '       cm = "alter table lintxt2040 alter column txt_" & CStr(nF) & " TEXT(40)"
 '       dB1.Execute cm
 '   Next nF
 '   dB1.Close
 '   Set dB1 = Nothing
 '   nF = FreeFile()
 '   Open App.Path & "\ata34" For Output As #nF
 '   Close #nF
 'End If
  Load frmEpelsa
  'If Not IconificarInicio Then
      MostrarInterfaz
  'End If
 '**************************************
 ' Iconificación
 If IconificarInicio Then
    frmControl.cSysTray1.InTray = True
    frmEpelsa.Visible = False
 End If
    
    '******************
    ' Activa Recogida contínua de tiquets
    '******************
    If descAuto = True Then
        frmControl.timetqt.Enabled = True
        '1.0.4 --> activar
        PausaTiquets = False
        frmEpelsa.BloquearMenues
        frmEpelsa.CmdComunicaciones(3).Caption = CargaCadena(704)
        'c2f 1.7.2
        frmEpelsa.CmdComunicaciones(3).Enabled = True
        '''''''''''''''''''
        lCogeTiquet = True
        cgdtiquet = False
    End If
    
End Sub


Private Function marca_tqt_exportacion() As Integer
Dim rst As dao.Recordset
Dim db As dao.Database
Dim sSQL As String
Dim cntT As Long
Dim cntR As Long
Dim sTabla As String
    Set db = AbrirBase
    For cntT = 0 To 1
        If cntT = 0 Then sTabla = "cabecera"
        If cntT = 1 Then sTabla = "tickets"
        'If cntT = 2 Then sTabla = "cabeceraSC10"
        'If cntT = 3 Then sTabla = "ticketsSC10"
        'If cntT = 4 Then sTabla = "cabeceraTQ"
        'If cntT = 5 Then sTabla = "ticketsTQ"
        sSQL = "SELECT RESERV3 " & _
               "FROM " & sTabla & " " & _
               "WHERE RESERV3=" & "'" & "NO" & "'"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                .MoveLast
                .MoveFirst
                For cntR = 0 To .Recordcount - 1
                    .Edit
                    .Fields("RESERV3") = "SI"
                    .Update
                    .Movenext
                Next cntR
            End If
            .Close
        End With
        Set rst = Nothing
    Next cntT
    CerrarBase db
    Set db = Nothing
End Function
Private Sub Detecta_Instancia()
    Dim nMyFich As Integer
    Dim nMyFich2 As Integer
    Dim sTest As String
    If App.PrevInstance = True Then
        If (Dir(App.Path & "\reinit.slave") <> "") Then
            On Error Resume Next
            Kill App.Path & "\reinit.slave"
            On Error GoTo 0
        Else
            CadenadeLog "Detectada Instancia Previa..."
            End
        End If
    End If
End Sub

Public Sub CreaCarpetas()
    Dim Base As dao.Database
    Dim Tabla As TableDef
    Dim Reintentos As Integer
    Dim b As Integer
    Dim i As Integer
    Dim Conseguido As Boolean
    '********************************
    ' crea carpeta temporal y tmp.mdb
    '********************************
    'If Dir(Miruta & "\dirtmptouch", vbArchive) <> "" Then
    '    On Error Resume Next
    '    Kill Miruta & "\dirtmptouch"
    '    On Error GoTo 0
    'End If
    If Dir(Miruta & "\dirtmptouch", vbDirectory) = "" Then
        On Error Resume Next
        MkDir (Miruta & "\dirtmptouch")
        On Error GoTo 0
    End If
    On Error GoTo fin
    
    If Dir(Miruta & "\images", vbDirectory) = "" Then
        On Error Resume Next
        MkDir (Miruta & "\images")
        On Error GoTo 0
    End If
    'If Dir(Miruta & "\images\families", vbDirectory) = "" Then
    '    On Error Resume Next
    '    MkDir (Miruta & "\images\families")
    '    On Error GoTo 0
    'End If
    'If Dir(Miruta & "\images\vendors", vbDirectory) = "" Then
    '    On Error Resume Next
    '    MkDir (Miruta & "\images\vendors")
    '    On Error GoTo 0
    'End If
    'If Dir(Miruta & "\images\items", vbDirectory) = "" Then
    '    On Error Resume Next
    '    MkDir (Miruta & "\images\items")
    '    On Error GoTo 0
    'End If
    
    If Dir(Miruta & "\dirtmptouch\tmp.mdb", vbArchive) <> "" Then
        Kill Miruta & "\dirtmptouch\tmp.mdb"
    End If
    On Error GoTo 0
    Conseguido = False
    Reintentos = 0
    
    'MsgBox Miruta & "\dirtmptouch\tmp.mdb"
    
    Do Until Conseguido Or Reintentos > 200
        On Error Resume Next
        Set Base = CreateDatabase(Miruta & "\dirtmptouch\tmp.mdb", dbLangGeneral, dbVersion40)
        If Err.Number <> 0 Then
            Sleep (100)
            Reintentos = Reintentos + 1
        Else
            Conseguido = True
        End If
    Loop
    
    'MsgBox CStr(Conseguido)
    
    
    If Conseguido Then
        Set Tabla = Base.CreateTableDef("consart")
        With Tabla
            .Fields.Append .CreateField("Euroscale", dbBoolean, 6)
            .Fields.Append .CreateField("SC10", dbBoolean, 9)
            .Fields.Append .CreateField("Cod", dbDouble, 9)
            .Fields.Append .CreateField("Plu", dbDouble, 9)
            .Fields.Append .CreateField("Sec", dbDouble, 9)
            .Fields.Append .CreateField("Sub", dbDouble, 9)
            .Fields.Append .CreateField("Fam", dbDouble, 9)
            .Fields.Append .CreateField("Prc", dbDouble, 1)
            .Fields.Append .CreateField("Cad", dbDouble, 1)
            .Fields.Append .CreateField("Tara", dbDouble, 1)
            .Fields.Append .CreateField("Pref", dbDouble, 1)
            .Fields.Append .CreateField("Iva", dbDouble, 1)
            .Fields.Append .CreateField("wgh", dbText, 1)
            .Fields.Append .CreateField("etq", dbDouble, 1)
            .Fields.Append .CreateField("ean", dbText, 13)
            For b = 0 To 21
                .Fields.Append .CreateField("desc" & b, dbText, 25)
            Next b
            For b = 0 To 9
                .Fields.Append .CreateField("tlt" & b, dbDouble)
            Next b
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Set Tabla = Base.CreateTableDef("Totales")
        With Tabla
            For b = 1 To 4
                .Fields.Append .CreateField("D" & b, dbDate)
            Next b
            For b = 1 To 20
                .Fields.Append .CreateField("N" & b, dbDouble)
            Next b
            For b = 1 To 15
                .Fields.Append .CreateField("T" & b, dbText)
            Next b
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Set Tabla = Base.CreateTableDef("Pedido")
        With Tabla
            .Fields.Append .CreateField("Codigo", dbDouble)
            .Fields.Append .CreateField("Descriptivo", dbText)
            .Fields.Append .CreateField("Tipo", dbText)
            .Fields.Append .CreateField("Pedido", dbDouble)
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Set Tabla = Base.CreateTableDef("PedInv")
        With Tabla
            .Fields.Append .CreateField("Mostrador", dbDouble)
            .Fields.Append .CreateField("Codigo", dbDouble)
            .Fields.Append .CreateField("Plu", dbDouble)
            .Fields.Append .CreateField("Inventario", dbDouble)
            .Fields.Append .CreateField("Pedido", dbDouble)
            .Fields.Append .CreateField("Tipo", dbText)
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Set Tabla = Base.CreateTableDef("Inventario")
        With Tabla
            .Fields.Append .CreateField("Codigo", dbDouble)
            .Fields.Append .CreateField("Descriptivo", dbText)
            .Fields.Append .CreateField("Tipo", dbText)
            .Fields.Append .CreateField("Inventario", dbDouble)
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Set Tabla = Base.CreateTableDef("Tiendas")
        With Tabla
            .Fields.Append .CreateField("Numero", dbText)
            .Fields.Append .CreateField("Nombre", dbText)
            .Fields.Append .CreateField("Provincia", dbText)
            .Fields.Append .CreateField("Poblacion", dbText)
            .Fields.Append .CreateField("CP", dbText)
            .Fields.Append .CreateField("Direccion", dbText)
            .Fields.Append .CreateField("Telefono", dbText)
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
        Base.Close
    Else
        MsgBox CargaCadena(1382), vbCritical
        End
    End If
fin:
    On Error GoTo 0
    '***********************************
    ' crea directorio de configuraciones
    '***********************************
    If Dir(Miruta & "\etc", vbDirectory) = "" Then
        ChDir Miruta
        MkDir (Miruta & "\etc")
    End If
    frmControl.File1.Path = App.Path
    frmControl.File1.Pattern = "*.fia"
    frmControl.File1.Refresh
    For b = 0 To frmControl.File1.ListCount - 1
        On Error Resume Next
        FileCopy Miruta & "\" & frmControl.File1.List(b), Miruta & "\etc\" & frmControl.File1.List(b)
        Kill Miruta & "\" & frmControl.File1.List(b)
        On Error GoTo 0
    Next b
    frmControl.File1.Pattern = "*.fet"
    frmControl.File1.Refresh
    For b = 0 To frmControl.File1.ListCount - 1
        On Error Resume Next
        FileCopy Miruta & "\" & frmControl.File1.List(b), Miruta & "\etc\" & frmControl.File1.List(b)
        Kill Miruta & "\" & frmControl.File1.List(b)
        On Error GoTo 0
    Next b
End Sub
Private Sub Check_Multi()
    Dim Arch As Integer
    Dim Buffer As String
    Dim bucle As Integer
    CadenadeLog "Check_Multi()"
    
    
    If Trim(Command) = "" Then
        CadenadeLog "Check_Multi() --> Sin parámetro de arranque..."
        If Dir(App.Path & "\multistart.ord") <> "" Then
            
            CadenadeLog "Check_multi() --> Detectado multistart.ord"
            
            FileCopy App.Path & "\multistart.ord", App.Path & "\multistart.cop"
            
            HayMulti = True
            Arch = FreeFile()
            Open App.Path & "\multistart.ord" For Input As #Arch
            Line Input #Arch, Buffer
            id = Buffer
            Line Input #Arch, Buffer
            Hydra_INI = Buffer
            
            CadenadeLog "Leido de multistart.ord, Hydra_INI=" & Hydra_INI
            
            Line Input #Arch, Buffer
            Base_General = Buffer
            
            CadenadeLog "Leido de multistart.ord, Base_General=" & Base_General
            
            Line Input #Arch, Buffer
           
            NombreTienda = Buffer
            
            Line Input #Arch, Buffer
            OrdenMulti = Buffer
            Line Input #Arch, Buffer
            TiendaActual = Val(Buffer)
            Line Input #Arch, UsuarioActual.Nombre
            If UsuarioActual.Nombre <> "" Then
                Secure_Login UsuarioActual, True
                HaySeguridad = True
            Else
                UsuarioActual.Nombre = "admin"
                HaySeguridad = False
            End If
            Close #Arch
            Sleep (1000)
            Kill App.Path & "\multistart.ord"
            Call hazGS
        End If
    Else
        '"TRANS"
        If UCase(Mid(Command, 1, 5)) <> "TRANS" Then
            
            If Dir(App.Path & "\hydramulti.exe") <> "" Then Check_TiendaCommand
        End If
    End If
    CargaIdiomas
End Sub
Public Sub hazGS()
    Dim rst As dao.Recordset
    Dim basTi As dao.Database
    Dim nArch As Integer
    
    Set basTi = OpenDatabase(App.Path & "\tiendas.mdb")
    Set rst = basTi.OpenRecordset("select * from grupos where tienda=" & CStr(TiendaActual))
    If Not rst.EOF Then
        GrupoActual = rst.Fields("grp")
    Else
        GrupoActual = 0
    End If
    rst.Close
    basTi.Close
    Set rst = Nothing
    Set basTi = Nothing
    nArch = FreeFile()
    Open App.Path & "\gianshop.txt" For Output As #nArch
    If TiendaActual < 100 Then
        Print #nArch, Format(TiendaActual, "00")
        Print #nArch, Format(GrupoActual, "00")
    Else
        Print #nArch, Format(TiendaActual, "000")
        Print #nArch, Format(GrupoActual, "000")
    End If
    Close #nArch
End Sub
Private Sub Repara_Base()
    CadenadeLog "Detectado fin abrupto la última vez que arrancó HydraTouch, tratando de reparar base..."
    On Error Resume Next
    DBEngine.RepairDatabase Base_General
    If Err.Number = 0 Then
        Compactar_Base
        CadenadeLog "base reparada"
    Else
        CadenadeLog "La base no pudo ser reparada"
    End If
    On Error GoTo 0
End Sub
Public Sub Check_TiendaCommand()
    Dim Definiciones() As TipoDefinicion
    Dim CuentaDefiniciones As Integer
    Dim Fichero As Integer
    Dim Buffer As String
    Dim Buffer2 As String
    Dim BufDef As String
    Dim BufDef2 As String
    Dim MiTienda As String
    Dim bucle As Integer
    Dim b As Integer
    Dim eNCONTRADO As Boolean
    Dim sCaptura As String
    Dim Base As dao.Database

    If (Dir(App.Path & "\cmdlinetouch.cfg") = "") And (Not HaySeguridad) Then
       MsgBox CargaCadena(1383), vbCritical
       End
    End If
    Buffer = Trim(Command)
    If UCase(Mid(Buffer, 5)) = "KILLTRA" Or UCase(Mid(Buffer, 2)) = "KILLTRA" Then
        MiTienda = ""
        If Len(Buffer) = 11 Then MiTienda = Mid(Buffer, 2, 3)
        If Len(Buffer) = 10 Then MiTienda = Mid(Buffer, 2, 2)
        If Not IsNumeric(MiTienda) Then
            salir_programa
        End If
        If Val(MiTienda) < 100 Then
            Base_General = App.Path & "\t" & Format(Val(MiTienda), "00") & "\dbasetouch.mdb"
            Hydra_INI = App.Path & "\t" & Format(Val(MiTienda), "00") & "\hydratouch.ini"
        Else
            Base_General = App.Path & "\t" & Format(Val(MiTienda), "000") & "\dbasetouch.mdb"
            Hydra_INI = App.Path & "\t" & Format(Val(MiTienda), "000") & "\hydratouch.ini"
        End If
        If Dir(Base_General) <> "" Then
            Set Base = AbrirBase
                Base.Execute "delete * from cabecera where typtic='12'"
                Base.Execute "delete * from tickets where typtic='12'"
                Base.Execute "delete * from gtsecs where modo=12"
                Base.Execute "delete * from gtarti where modo=12"
                Base.Execute "delete * from gtinfven where modo=12"
                Base.Execute "delete * from gtvend where modo=12"
                Base.Execute "delete * from gthora"
                CadenadeLog "Execute: delete * from gthora"
            'End If
            Base.Close
            Set Base = Nothing
        End If
        'salir_programa
        End
    End If
    
    '2.0.40
       '2.0.37 adaptación parámetro a 3 dígitos tienda...
       If left(Buffer, 1) = "/" And IsNumeric(Mid(Buffer, 2, 1)) Then
           If IsNumeric(Mid(Buffer, 3, 1)) Then
               CadenadeLog "Adaptación parámetro:" & Buffer
               If Not IsNumeric(Mid(Buffer, 4, 1)) Then
                   Buffer = "/0" & Mid(Buffer, 2)
               End If
               CadenadeLog "               Queda:" & Buffer
           End If
           
       End If
       ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''
    CuentaDefiniciones = 1
    CadenadeLog "MIRUTA: " & Miruta
    CadenadeLog "APP PATH:" & App.Path
    CadenadeLog "BUFFER:" & Buffer
    CadenadeLog "Mid(Buffer, 1, 5)=" & Mid(Buffer, 1, 5)
    CadenadeLog ""
    If (Buffer = "" Or UCase(Mid(Buffer, 1, 5)) = "TRANS") Then Exit Sub
       
       CadenadeLog "LECTURA DE DEFINICIONES EXISTENTES EN CMDLINETOUCH.CFG... check_tiendacommand"
       
       Fichero = FreeFile()
       Open App.Path & "\cmdlinetouch.cfg" For Input As Fichero
       If Not EOF(Fichero) Then
            Do Until EOF(Fichero)
                ReDim Preserve Definiciones(CuentaDefiniciones + 1)
                Line Input #Fichero, Definiciones(CuentaDefiniciones).texto
                CadenadeLog "Definición:" & CStr(CuentaDefiniciones) & " --> " & Definiciones(CuentaDefiniciones).texto
                Line Input #Fichero, Definiciones(CuentaDefiniciones).accion
                Line Input #Fichero, Definiciones(CuentaDefiniciones).Adicional
                CuentaDefiniciones = CuentaDefiniciones + 1
            Loop
        End If
        Close #Fichero
        If Mid(Buffer, 1, 1) = "/" Then
            Buffer2 = "/"
            For bucle = 2 To Len(Buffer)
                If Mid(Buffer, bucle, 1) = " " Or Mid(Buffer, bucle, 1) = "/" Then
                    Exit For
                Else
                    Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
                End If
            Next bucle
        End If
        Buffer2 = StrConv(Buffer2, vbUpperCase)
        
        CadenadeLog "Conversión a mayúsculas parámetro:" & Buffer2
        
'CadenadeLog Buffer2

        eNCONTRADO = False
        For bucle = 1 To CuentaDefiniciones
            If InStr(1, Definiciones(bucle).texto, "?") > 0 Then
                BufDef = ""
                For b = 1 To Len(Definiciones(bucle).texto)
                    If b <= Len(Buffer2) Then
                        If Mid(Definiciones(bucle).texto, b, 1) = "?" Then
                            BufDef = BufDef & Mid(Buffer2, b, 1)
                        Else
                            BufDef = BufDef & Mid(Definiciones(bucle).texto, b, 1)
                        End If
                    End If
                Next b
            Else
                BufDef = Definiciones(bucle).texto
            End If
            If InStr(1, BufDef, "#") > 0 Then
                BufDef2 = ""
                MiTienda = ""
                For b = 1 To Len(BufDef)
                    If b <= Len(Buffer2) Then
                        If Mid(BufDef, b, 1) = "#" Then
                            BufDef2 = BufDef2 & Mid(Buffer2, b, 1)
                            MiTienda = MiTienda & Mid(Buffer2, b, 1)
                        Else
                            BufDef2 = BufDef2 & Mid(BufDef, b, 1)
                        End If
                    End If
                Next b
            Else
                BufDef2 = BufDef
            End If
            
            CadenadeLog "Comparación Buffer y definición: " & Buffer2 & " --- " & BufDef2
            If Buffer2 = BufDef2 Then
                'accion = Val(Left(Definiciones(Bucle).accion, 3))
                eNCONTRADO = True
                If Val(MiTienda) < 100 Then
                    If Dir(App.Path & "\t" & Format(Val(MiTienda), "00"), vbDirectory) = "" Then
                        
                        If Val(left(Definiciones(bucle).accion, 3)) <> 40 Then
                            MsgBox CargaCadena(1384), vbCritical
                            End
                        End If
                        
                    End If
                Else
                    If Dir(App.Path & "\t" & Format(Val(MiTienda), "000"), vbDirectory) = "" Then
                        
                        If Val(left(Definiciones(bucle).accion, 3)) <> 40 Then
                            MsgBox "Selected Store Does Not Exists. You have to Create It", vbCritical
                            End
                        End If
                        
                    End If
                
                End If
                TiendaActual = Val(MiTienda)
                
                CadenadeLog "Tienda a arrancar:" & CStr(TiendaActual)
                
                If TiendaActual <> 0 Then
                    '1.7.19 gigante
                    sCaptura = GetShopPath(TiendaActual)
                    
                    CadenadeLog "Captura de path en Tiendas.mdb" & ": " & sCaptura
                    
                    If sCaptura = "" Then
                        If Val(MiTienda) < 100 Then
                            Base_General = App.Path & "\t" & Format(Val(MiTienda), "00") & "\dbasetouch.mdb"
                            Hydra_INI = App.Path & "\t" & Format(Val(MiTienda), "00") & "\hydratouch.ini"
                        Else
                            Base_General = App.Path & "\t" & Format(Val(MiTienda), "000") & "\dbasetouch.mdb"
                            Hydra_INI = App.Path & "\t" & Format(Val(MiTienda), "000") & "\hydratouch.ini"
                        
                        End If
                    Else
                        
                        Base_General = sCaptura & "dbasetouch.mdb"
                        Hydra_INI = sCaptura & "hydratouch.ini"
                        
                        CadenadeLog "Base_General:" & Base_General
                        CadenadeLog "Hydra_ini:" & Hydra_INI
                    End If
                    
                Else
                    Base_General = ""
                    Hydra_INI = ""
                End If
                HayMulti = True
                Call hazGS
                Exit For
            End If
            
            
        Next bucle
        If Not eNCONTRADO Then
            MsgBox "Command Line : Unknown Parameter", vbCritical
            End
        End If
End Sub
Private Sub Clientes_Especiales()
    Dim Arch As Integer
    '/////////////////////////////////////
    '// chequea si procede activar      //
    '// configuraciones especiales      //
    '/////////////////////////////////////
    PathOrden = ""

    If Dir(App.Path & "\pathmaj.txt") <> "" Then
        lClR = True
        Arch = FreeFile()
        Open App.Path & "\pathmaj.txt" For Input As #Arch
        On Error Resume Next
        Line Input #Arch, sPathMaj
        Close #Arch
        sPathMaj = Trim(sPathMaj)
        On Error GoTo 0
        
        
    End If

End Sub
Private Sub Crea_Fichero_INI()
    CadenadeLog "Función Crea_fichero_ini"
    If (Not HayMulti) Or (HayMulti And TiendaActual <> 0) Then
        If HayMulti Then
            CadenadeLog "Hay multi..."
        End If
        If TiendaActual <> 0 Then
            CadenadeLog "Tienda actual <> 0"
        End If
        CadenadeLog "hydratouch.ini = " & Hydra_INI
        'CadenadeLog "pass.sec = " & MiruTa & "\pass.sec"
        CadenadeLog "pass.sec = " & App.Path & "\pass.sec"
        If Dir(Hydra_INI) = "" Then
            'If Dir(MiruTa & "\pass.sec") <> "" Then
            'c2f 1.1.1
            'If Dir(App.Path & "\pass.sec") <> "" Then
            '    MsgBox "SECURITY VIOLATION", vbCritical
            '    CadenadeLog CargaCadena(852)
            '    End
            'Else
                If Dir(App.Path & "\hydraserver.exe") <> "" Then StoreFlow = True
                LeerParametrosInicio
                Load FrmConfiguracion2
                FrmConfiguracion2.CmdCancelar.Enabled = False
                FrmConfiguracion2.Show
                Do While FrmConfiguracion2.Visible
                    Do_Events
                Loop
                Unload FrmConfiguracion2
            'End If
        Else
            'If Dir(MiruTa & "\pass.sec") = "" Then
            'c2f 1.1.1
            'If Dir(App.Path & "\pass.sec") = "" Then
            '    MsgBox "SECURITY VIOLATION", vbCritical
            '    CadenadeLog CargaCadena(853)
            '    End
            'Else
                If Hydra_INI <> "" Then LeerParametrosInicio
            'End If
        End If
    End If
End Sub
Public Sub ReduceLogHydra()
    Dim Arch1 As Integer, Arch2 As Integer
    Dim Buf As String
    Dim bucle As Long
    Dim SumaTotal As Long
    Dim TotalActual As Long
    TotalActual = FileLen(App.Path & "\loghydra.txt")
    On Error Resume Next
    Arch1 = FreeFile()
    Open App.Path & "\loghydra.txt" For Input As #Arch1
    On Error Resume Next
    Arch2 = FreeFile()
    Open App.Path & "\loghydra.tmp" For Output As #Arch2
    If Err.Number <> 0 Then Exit Sub
    Do Until EOF(Arch1)
        Line Input #Arch1, Buf
        SumaTotal = SumaTotal + Len(Buf) + 2
        If (TotalActual - SumaTotal) <= 1000000 Then Exit Do
    Loop
    Do Until EOF(Arch1)
        Line Input #Arch1, Buf
        Print #Arch2, Buf
    Loop
    Close #Arch1
    Close #Arch2
    On Error Resume Next
    Kill App.Path & "\loghydra.txt"
    Name App.Path & "\loghydra.tmp" As App.Path & "\loghydra.txt"
End Sub



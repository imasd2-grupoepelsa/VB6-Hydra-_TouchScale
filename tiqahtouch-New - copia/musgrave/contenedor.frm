VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form frmControl 
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   2685
   ClientTop       =   3015
   ClientWidth     =   2730
   Icon            =   "contenedor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   2730
   Begin SysTrayCtl.cSysTray cSysTray3 
      Left            =   1560
      Top             =   840
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":08CA
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin SysTrayCtl.cSysTray cSysTray2 
      Left            =   960
      Top             =   840
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":11A4
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   360
      Top             =   840
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":1A7E
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Timer tmrpluasc 
      Interval        =   20000
      Left            =   2280
      Top             =   120
   End
   Begin VB.Timer TmRet 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer TimeTQT 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   120
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////
'\\\\\\\\\\\\\\
Private MyContador As Long
Private CuentaGRande As Long
Public Ret As Boolean
'////////////////////
'\\\\\\\\\\\\\\\\\\\\
Private Sub cSysTray1_MouseDblClick(Button As Integer, id As Long)
    frmMonitor.Show
End Sub
'//////
'\\\\\\
Private Sub cSysTray2_MouseDblClick(Button As Integer, id As Long)
    frmMonitor.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
   If UnloadMode = 2 Then
       CadenadeLog "Salida: Detectado IPL..."
   Else
       If Cancel = 0 And UnloadMode = 1 Then
            CadenadeLog "MINIMIZADO Por Usuario"
       Else
            CadenadeLog "Salida: Detectado Kill por ejecución KILL( ó administrador Tareas)..."
       End If
   End If
   
   CadenadeLog "Solicitada Salida:" & Format(Cancel, "0000000000") & "     " & Format(UnloadMode, "0000000000")

   
End Sub

'\\\\\\
'//////
Public Sub TimeTQT_Timer()
    Dim MyResul As Long
    Dim MyProc As Long
    Dim MyFich As Integer
    Dim nFlg As Integer
    Dim lRecOK As Boolean
    Dim nL1, nL2, nL3 As Double
    
    Dim nnn, Fich1, Resp, nC As Integer
    Dim MiResultado, MiProceso As Long
    Dim sPath As String
    Dim sCommand As String
    Dim nContTmp As Integer
    Dim nTMP As Integer
    Dim nLG As Double
    Dim nTienda As Integer
    Dim sTienda As String
    Dim Mybase As dao.Database
    Dim MyRs As dao.Recordset
    
    TimeTQT.Enabled = False
    
    Call LeeINI
    
    If Dir(App.Path & "\pluasc.dir") <> "" Then
        MyFich = FreeFile()
        Open App.Path & "\pluasc.dir" For Input As #MyFich
        Line Input #MyFich, sPath
        Close #MyFich
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    Else
        sPath = App.Path & "\"
        MyFich = FreeFile()
        Open App.Path & "\pluasc.dir" For Output As #MyFich
        Print #MyFich, sPath
        Close #MyFich
    End If
    
    If Dir(App.Path & "\tienda.cfg") <> "" Then
        MyFich = FreeFile()
        Open App.Path & "\tienda.cfg" For Input As #MyFich
        Line Input #MyFich, sTienda
        Close #MyFich
        nTienda = Val(sTienda)
    Else
        nTienda = 1
    End If
    sTienda = Format(nTienda, "000")
    
    'CuentaGRande = CuentaGRande + 1
    
    
    'If (Dir(App.Path & "\hydrarun.txt") <> "") And (Dir(App.Path & "\transpluasc.exe") = "") Then
    '    CadenadeLog "Detectado HydraRun.txt"
    '    '''''''''''''''''''''''''
    '    reset '.balanza
    '    '''''''''''''''''''''''''
    '   TimeTQT.Enabled = True
    '   Exit Sub
    'Else
        If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
            Kill App.Path & "\dirtmp\tqpausa"
        End If
        If Dir(App.Path & "\dirtmp\tqreanuda") <> "" Then
            Kill App.Path & "\dirtmp\tqreanuda"
        End If
    
    'End If
    'If Dir(App.Path & "\bal.dat") <> "" Or Dir(App.Path & "\ballss.dat") <> "" Or Dir(App.Path & "\balenv.dat") <> "" Then
    '    Exit Sub
    'End If
    
    If nTocaLocal > nContLocal Then nTocaLocal = 0
    
    If (CuentaGRande >= 30 And TimeAuto <= 5 And (TimeAutoSeg = 0)) Or (TimeAuto > 5 And (CuentaGRande * 2) >= TimeAuto And (TimeAutoSeg = 0)) Or Trim(Command) = "/01PEEKTOT" Or ((CuentaGRande * 2) >= TimeAutoSeg And (TimeAutoSeg <> 0)) Then
        MyContador = MyContador + 1
        If (MyContador >= TimeAuto And TimeAuto <= 5) Or (TimeAuto > 5 And (CuentaGRande * 2) >= TimeAuto) Or Trim(Command) = "/01PEEKTOT" Or ((CuentaGRande * 2) >= TimeAutoSeg And (TimeAutoSeg <> 0)) Then
            
                TimeTQT.Enabled = False
                
                nTocaLocal = nTocaLocal + 1
                
                CuentaGRande = 0
            
                MyContador = 0
                
                If Dir(App.Path & "\pervive.exe") <> "" Then
                    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\pervive.exe", vbMinimizedFocus))
                End If
                
                'If Dir(App.Path & "\hydrarun.txt") <> "" And Dir(App.Path & "\transpluasc.exe") = "" Then
                '    frmControl.cSysTray2.InTray = False
                '    frmControl.cSysTray1.InTray = True
                '    'Do_Pausa
                '    TimeTQT.Interval = MyInterval
                '    'TimeTQT.Enabled = True
                '    '''''''''''''''''''''''''
                '    reset
                '    '''''''''''''''''''''''''
                '    TimeTQT.Enabled = True
                '    Exit Sub
                'End If
                
                lRecOK = False
                
                If Dir(App.Path & "\tqretail.cfg") <> "" Then
                    nTMP = FreeFile()
                    Open App.Path & "\nosdh.txt" For Output As #nTMP
                    Close #nTMP
                    Sleep (500)
                End If
                
                If Not (lNoCoger) Or Dir(sPath & "reco.txt") <> "" Then
                    lNoConnect = False
                    'Musgrave se deja solo tiques de Venta
                    LeeTickets 0
                    'If Dir(App.Path & "\nolabel.txt") = "" Then
                    '    LeeTickets 1
                    '    If Dir(App.Path & "\nopack.txt") = "" Then
                    '        LeeTickets 2
                    '    End If
                    'End If
                    CadenadeLog Now() & " --> Fin Recepción todos los tipos de tique."
                    If lBoka Then
                        Call Trata_Boka
                    End If
                    
                    If Not (lNoConnect) Then
                        If Dir(sPath & "reco.txt") <> "" Then
                            Kill sPath & "reco.txt"
                        End If
                        lRecOK = True
                    Else
                        lNoConnect = False
                        nFlg = FreeFile()
                        Open sPath & "noreco.txt" For Output As #nFlg
                        Close #nFlg
                        If Dir(sPath & "reco.txt") <> "" Then
                            Kill sPath & "reco.txt"
                        End If
                    End If
                End If
                
                'CadenadeLog "Recogida 0003"
                
                If Dir(App.Path & "\recosine.txt") <> "" Then 'si existe recosine.txt se recogen tiques pero
                                                              ' no se exportan hasta borrado de totales...
                                                              ' Si la hora de borrado de totales es <> "00:00:00"
                    LeeTickets 0
                    
                    'If Dir(App.Path & "\nolabel.txt") = "" Then
                    '    LeeTickets 1
                    '    If Dir(App.Path & "\nopack.txt") = "" Then
                    '        LeeTickets 2
                    '    End If
                    'End If
                    CadenadeLog Now() & " --> Fin Recepción todos los tipos de tique."
                    If lBoka Then
                        Call Trata_Boka
                    End If
                    
                Else
                
                    LeeTickets 0
                    'Musgrave, solo tiques venta
                    'If Dir(App.Path & "\nolabel.txt") = "" Then
                    '    LeeTickets 1
                    '    If Dir(App.Path & "\nopack.txt") = "" Then
                    '        LeeTickets 2
                    '    End If
                    'End If
                    CadenadeLog Now() & " --> Fin Recepción todos los tipos de tique."
                    If lBoka Then
                        Call Trata_Boka
                    End If
                    
                    If lFornes Then
                        Call export_BL_Fornes(True)
                    Else
                    If lUpper Then
                        Call export_CABLIN(True)
                    Else
                        If Dir(App.Path & "\codisys.txt") <> "" Or Dir(App.Path & "\tqretail.cfg") <> "" Or Dir(App.Path & "\libravahy.txt") <> "" Or Dir(App.Path & "\ahorramas.cfg") <> "" Then

                            If Dir(App.Path & "\ahorramas.cfg") <> "" Then
                                CadenadeLog "Exportar Tiques..."
                                Principal.Taquion_EnviaTiquetsPendientes_ah False 'sc10
                            Else
                                If nMaqType = 0 Then
                                    Principal.Taquion_EnviaTiquetsPendientes_ah True 'euroscale
                                Else
                                    Principal.Taquion_EnviaTiquetsPendientes_ah False 'sc10
                                End If
                            End If
                        Else
                            If Dir(App.Path & "\infoser.txt") <> "" Then
                                If nMaqType = 0 Then 'EuroScale
                                    Call export_BL(True)
                                Else
                                    Call export_BL(False) 'sc10
                                End If
                            
                            Else
                                If Dir(App.Path & "\bdp.txt") <> "" Then
                                    Call export_BDP(True)
                                Else
                                    If Dir(App.Path & "\tiendamt.txt") <> "" Then
                                        CadenadeLog "detectado tiendamt.txt"
                                        Call export_MT(True)
                                    Else
                                        If nMaqType = 10 Then
                                            Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                                        Else
                                            CadenadeLog "Recogida 0007"
                                            Principal.Taquion_EnviaTiquetsPendientes False 'sc10
                                            CadenadeLog "Recogida 0008"
                                        End If
                                    End If
                                    
                                    'Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                                End If
                            End If
                        End If
                    End If
                    End If
                    CadenadeLog Now() & " --> Fin exportación Tique."
                    
                End If
                
                
                If Dir(App.Path & "\tqretail.cfg") <> "" Then
                    If Dir(App.Path & "\nosdh.txt") <> "" Then
                        Kill App.Path & "\nosdh.txt"
                        Sleep (500)
                    End If
                End If
                
                'borrado de totales
                If (Time >= dHoraFin And (dHoraFin <> CDate("00:00:00"))) And Dir(App.Path & "\" & Format(Date, "ddmmyy") & ".fin") = "" Then
                    Fich1 = FreeFile()
                    Open App.Path & "\" & Format(Date, "ddmmyy") & ".fin" For Output As #Fich1
                    Print #Fich1, Now
                    Close #Fich1
                    Switch_mode nMaqType
                    
                    LeeTickets 0
                    If Dir(App.Path & "\nolabel.txt") = "" Then
                        LeeTickets 1
                        If Dir(App.Path & "\nopack.txt") = "" Then
                            LeeTickets 2
                        End If
                    End If
                    CadenadeLog Now() & " --> Fin Recepción todos los tipos de tique."
                    If lBoka Then
                        Call Trata_Boka
                    End If
                    
                    If lUpper Then
                        Resp = Erase_Totals(0) '.balanza se reinicia numeración
                    Else
                        Resp = Erase_Totals(1) '.balanza no se reinicia numeración
                    End If
                    If PathSC10 <> "NOEXPORTAR" Then
                        If lFornes Then
                            Call export_BL_Fornes(True)
                        Else
                        If lUpper Then
                            Call export_CABLIN(True)
                        Else
                            If Dir(App.Path & "\codisys.txt") <> "" Or Dir(App.Path & "\tqretail.cfg") <> "" Then
                                If nMaqType = 0 Then
                                    Principal.Taquion_EnviaTiquetsPendientes_ah True 'euroscale
                                Else
                                    Principal.Taquion_EnviaTiquetsPendientes_ah False 'sc10
                                End If
                            Else
                                If Dir(App.Path & "\infoser.txt") <> "" Then
                                    '1.4.3
                                    'Call export_BL(True)
                                    If nMaqType = 0 Then 'EuroScale
                                        Call export_BL(True)
                                    Else
                                        Call export_BL(False) 'sc10
                                    End If
                                    
                                Else
                                    If Dir(App.Path & "\bdp.txt") <> "" Then
                                        Call export_BDP(True)
                                    Else
                                        If Dir(App.Path & "\tiendamt.txt") <> "" Then
                                            CadenadeLog "detectado tiendamt.txt"
                                            Call export_MT(True)
                                        Else
                                            '1.4.3
                                            'Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                                            If nMaqType = 0 Then
                                                Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                                            Else
                                                Principal.Taquion_EnviaTiquetsPendientes False 'sc10
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        End If
                    End If
                    
                    'arranque hydra+ exporta end_day.ini
                    If Dir(App.Path & "\end_day.ini") <> "" Then
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01END_DAY1", vbMinimizedFocus))
                        Do While MiResultado = STILL_ACTIVE
                            GetExitCodeProcess MiProceso, MiResultado
                            Do_Events
                        Loop
                        Sleep (1000)
                        If Dir(App.Path & "\totart.dat") <> "" Then
                            FileCopy App.Path & "\totart.dat", App.Path & "\" & Format(Date, "ddmmyy") & "TA" & ".dat"
                            Sleep (500)
                            Kill App.Path & "\totart.dat"
                        End If
                    End If
                    
                End If
                
                Sleep (500)
                If lRecOK Then
                    nFlg = FreeFile()
                    Open sPath & "recok.txt" For Output As #nFlg
                    Close #nFlg
                End If
                
                'problema con Hydra+
                If Dir(App.Path & "\vermonitor.txt") = "" Then
                    If Dir(App.Path & "\transpluasc.exe") = "" And Dir(App.Path & "\nsmon.txt") = "" Then
                        nFlg = FreeFile()
                        Open App.Path & "\nover1.txt" For Output As #nFlg
                        Close #nFlg
                    End If
                End If
                
                
                If Dir(App.Path & "\transpluasc.exe") <> "" Then
                    
                    If Dir(sPath & "plu.asc") <> "" Then
                        If nDelayPluAsc <> 0 Then
                            For nL1 = 1 To nDelayPluAsc - 1
                                Sleep (1000)
                            Next nL1
                        End If
                        nL1 = FileLen(sPath & "plu.asc")
                        Sleep (1500)
                        nL2 = FileLen(sPath & "plu.asc")
                        If nL1 = nL2 Then
                        
                            Sleep (500)
                            
                            nL3 = FileLen(sPath & "plu.asc")
                            
                            If nL1 = nL3 Then
                            
                                If Trim(LCase((sPath & "plu.asc"))) <> Trim(LCase((App.Path & "\plu.asc"))) Then FileCopy sPath & "plu.asc", App.Path & "\plu.asc"
                                Sleep (1000)
                                'MiResultado = STILL_ACTIVE
                                'MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("transpluasc", vbMinimizedFocus))
                                'Do While MiResultado = STILL_ACTIVE
                                '    GetExitCodeProcess MiProceso, MiResultado
                                '    Do_Events
                                'Loop
                                'Call CambiaPluASC
                                
                                MiResultado = STILL_ACTIVE
                                MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("transpluasc.exe", vbMinimizedFocus))
                                Do While MiResultado = STILL_ACTIVE
                                    GetExitCodeProcess MiProceso, MiResultado
                                    Do_Events
                                Loop
                                Sleep (1000)
                                'If Trim(LCase((sPath & "plu.asc"))) <> Trim(LCase((App.Path & "\plu.asc"))) Then
                                '    Kill sPath & "plu.asc"
                                '    Kill App.Path & "\plu.asc"
                                'Else
                                '    Kill App.Path & "\plu.asc"
                                'End If
                                MiResultado = STILL_ACTIVE
                                MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                                Do While MiResultado = STILL_ACTIVE
                                    GetExitCodeProcess MiProceso, MiResultado
                                    Do_Events
                                Loop
                        
                            End If
                            
                        End If
                    
                    Else
                        If Dir(sPath & "plupesca.dat") <> "" Then
                          On Error Resume Next
                          FileCopy sPath & "plupesca.dat", App.Path & "\plupesca.dat"
                          On Error GoTo 0
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          
                          On Error Resume Next
                          Kill sPath & "plupesca.dat"
                          Sleep (200)
                          Kill App.Path & "\plupesca.dat"
                          On Error GoTo 0

                          
                        End If
                    End If
                    
                Else
                    'c2f 2.0.21 --> ahorramas no importa datos desde tiqah sino desde Hydra+
                    If Dir(App.Path & "\ahorramas.cfg") = "" Then
                    
                      If Dir(sPath & "bascula.asc") <> "" And Dir(sPath & "\pbasiva.asc") <> "" And Dir(App.Path & "\datisa.exe") <> "" Then
                          On Error Resume Next
                          FileCopy sPath & "bascula.asc", App.Path & "\bascula.asc"
                          FileCopy sPath & "\pbasiva.asc", App.Path & "\pbasiva.asc"
                          Sleep (200)
                          Kill sPath & "bascula.asc"
                          Kill sPath & "pbasiva.asc"
                          Sleep (200)
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("datisa.exe", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (200)
                          On Error GoTo 0
                      End If
                      
                      If Dir(App.Path & "\panxtra.exe") <> "" Then
                          On Error Resume Next
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("panxtra.exe", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          On Error GoTo 0
                      End If
                      
                      'If ((Dir(App.Path & "\bal.dat") <> "" Or Dir(App.Path & "\tec.dat") <> "") And Dir(App.Path & "\bdp.txt") = "") Then
                      '    If lSerial Then
                      '        '''''''''''''''''''''''''
                      '        reset
                      '        '''''''''''''''''''''''''
                      '    End If
                      '
                      '    MiResultado = STILL_ACTIVE
                      '    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydratouch /01EZ", vbMinimizedFocus))
                      '    Do While MiResultado = STILL_ACTIVE
                      '        GetExitCodeProcess MiProceso, MiResultado
                      '        Do_Events
                      '    Loop
                      '    Sleep (1000)
                      '    If Dir(App.Path & "\bal.dat") <> "" Then
                      '        Kill App.Path & "\bal.dat"
                      '    End If
                      '    If Dir(App.Path & "\tec.dat") <> "" Then
                      '        Kill App.Path & "\tec.dat"
                      '    End If
                      'End If
                      
                    
                      'TQRETAIL
                      If (Dir(sPath & "balglo.dat") <> "") Then
                          On Error Resume Next
                          FileCopy sPath & "balglo.dat", App.Path & "\balglo.dat"
                          On Error GoTo 0
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          
                          If Dir(sPath & "balglo.dat") <> "" Then
                              Kill sPath & "balglo.dat"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\balglo.dat") <> "" Then
                              Kill App.Path & "\balglo.dat"
                          End If
                      End If
                      
                      '''''''''''
                        'ver ficheros belros
                        'ARTICULO.bal
                        'MODIFICA.bal
                      If (Dir(sPath & "articulo.bal") <> "" Or Dir(sPath & "modifica.bal") <> "") Then
                          
                          On Error Resume Next
                          FileCopy sPath & "articulo.bal", App.Path & "\articulo.bal"
                          FileCopy sPath & "modifica.bal", App.Path & "\modifica.bal"
                          On Error GoTo 0
                          
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "modifica.bal") <> "" Then
                              Kill sPath & "modifica.bal"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\modifica.bal") <> "" Then
                              Kill App.Path & "\modifica.bal"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(sPath & "articulo.bal") <> "" Then
                              Kill sPath & "articulo.bal"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\articulo.bal") <> "" Then
                              Kill App.Path & "\articulo.bal"
                              Sleep (1000)
                          End If
                      End If
                        
                      If (Dir(sPath & "bal.dat") <> "" Or Dir(sPath & "tec.dat") <> "") And Dir(App.Path & "\bdp.txt") = "" Then
                          
                          On Error Resume Next
                          FileCopy sPath & "bal.dat", App.Path & "\bal.dat"
                          FileCopy sPath & "tec.dat", App.Path & "\tec.dat"
                          On Error GoTo 0
                          
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "bal.dat") <> "" Then
                              Kill sPath & "bal.dat"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\bal.dat") <> "" Then
                              Kill App.Path & "\bal.dat"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(sPath & "tec.dat") <> "" Then
                              Kill sPath & "tec.dat"
                              Sleep (1000)
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\tec.dat") <> "" Then
                              Kill App.Path & "\tec.dat"
                              Sleep (1000)
                          End If
                      End If
                      
                      If ((Dir(sPath & "in.txt") <> "") And Dir(App.Path & "\tiendamt.txt") <> "") Then
                          If FileLen(sPath & "in.txt") >= 201 Then
                              On Error Resume Next
                              FileCopy sPath & "in.txt", App.Path & "\in.txt"
                              On Error GoTo 0
                              Sleep (1000)
                              If lSerial Then
                                  '''''''''''''''''''''''''
                                  reset
                                  '''''''''''''''''''''''''
                              End If
                              
                              MiResultado = STILL_ACTIVE
                              MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                              Do While MiResultado = STILL_ACTIVE
                                  GetExitCodeProcess MiProceso, MiResultado
                                  Do_Events
                              Loop
                              Sleep (1000)
                              If Dir(sPath & "in.txt") <> "" Then
                                  Kill sPath & "in.txt"
                              End If
                              Sleep (1000)
                              If Dir(App.Path & "\in.txt") <> "" Then
                                  Kill App.Path & "\in.txt"
                              End If
                          End If
                      End If
                      
                      If Dir(sPath & "orchestra.dat") <> "" Then
                          On Error Resume Next
                          FileCopy sPath & "orchestra.dat", App.Path & "\orchestra.dat"
                          On Error GoTo 0
                          On Error Resume Next
                          If Dir(App.Path & "\orchestra.dat") <> "" Then
                              If Dir(App.Path & "\orchestra.bak") <> "" Then
                                  Kill App.Path & "\orchestra.bak"
                              End If
                              FileCopy App.Path & "\orchestra.dat", App.Path & "\orchestra.bak"
                          End If
                          On Error GoTo 0
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          Sleep (1000)
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          
                          If Dir(sPath & "orchestra.dat") <> "" Then
                              Kill sPath & "orchestra.dat"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\orchestra.dat") <> "" Then
                              Kill App.Path & "\orchestra.dat"
                          End If
                      End If
                      
                      If ((Dir(sPath & "bal.dat") <> "") And Dir(App.Path & "\bdp.txt") <> "") Then
                          On Error Resume Next
                          FileCopy sPath & "bal.dat", App.Path & "\bal.dat"
                          On Error GoTo 0
                          Call TrataBal_BDP
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "bal.dat") <> "" Then
                              Kill sPath & "bal.dat"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\bal.dat") <> "" Then
                              Kill App.Path & "\bal.dat"
                          End If
                      
                      End If
                      
                      If ((Dir(sPath & "artcon.dat") <> "")) Then
                          On Error Resume Next
                          FileCopy sPath & "artcon.dat", App.Path & "\artcon.dat"
                          On Error GoTo 0
                          'Call TrataBal_BDP
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "artcon.dat") <> "" Then
                              Kill sPath & "artcon.dat"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\artcon.dat") <> "" Then
                              Kill App.Path & "\artcon.dat"
                          End If
                      
                      End If
                      
                        If Dir(sPath & "Bal_Epelsa.txt") <> "" Then
                            Call transforma_Fornes
                        End If
                        If Dir(sPath & "plu.csv") <> "" Then
                            Call transforma_Parkers
                        End If
                      
                      If Dir(sPath & "global.dat") <> "" Or Dir(App.Path & "\global.dat") <> "" Then
                          If Dir(App.Path & "\global.dat") <> "" Then
                            CadenadeLog "Detectado: " & App.Path & "\global.dat"
                          Else
                            CadenadeLog "Detectado: " & sPath & "global.dat"
                            Do_Events
                            nLG = FileLen(sPath & "global.dat")
                            Do_Events
                            Sleep (3000)
                            Do_Events
                            Do While nLG <> FileLen(sPath & "global.dat")
                                Do_Events
                                nLG = FileLen(sPath & "global.dat")
                                Sleep (3000)
                                Do_Events
                            Loop
                            On Error Resume Next
                            FileCopy sPath & "\global.dat", App.Path & "\global.dat"
                            On Error GoTo 0
                          End If
                          
                          CadenadeLog "LLamada Hydra proceso: " & "global.dat"
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01GLOBAL", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          CadenadeLog "Procesado: " & "global.dat" & " Por parte de Hydra."
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          CadenadeLog "Envia de Modificaciones desde Hydra..."
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          If Dir(App.Path & "\envok.dat") <> "" Then
                              CadenadeLog "Modificaciones Correctamente Enviadas Por parte de Hydra."
                          Else
                              CadenadeLog "Error en envio Modificaciones Por parte de Hydra."
                          End If
                          'If lUpper And lHydra Then
                          '  CadenadeLog "LLamada Hydra+ proceso: " & "global.dat"
                          '  MiResultado = STILL_ACTIVE
                          '  MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01GLOBAL", vbMinimizedFocus))
                          '  Do While MiResultado = STILL_ACTIVE
                          '      GetExitCodeProcess MiProceso, MiResultado
                          '      Do_Events
                          '  Loop
                          '  CadenadeLog "Procesado: " & "global.dat" & " Por parte de Hydra+."
                          '  If lSerial Then
                          '      '''''''''''''''''''''''''
                          '      reset
                          '      '''''''''''''''''''''''''
                          '  End If
                          '  CadenadeLog "Envia de Modificaciones desde Hydra+..."
                          '  MiResultado = STILL_ACTIVE
                          '  MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                          '  Do While MiResultado = STILL_ACTIVE
                          '      GetExitCodeProcess MiProceso, MiResultado
                          '      Do_Events
                          '  Loop
                          '  If Dir(App.Path & "\envok.dat") <> "" Then
                          '      CadenadeLog "Modificaciones Correctamente Enviadas Por parte de Hydra+."
                          '  Else
                          '      CadenadeLog "Error en envio Modificaciones Por parte de Hydra+."
                          '  End If
                          'End If
                          
                          'Sleep (1000)
                          If Dir(sPath & "global.dat") <> "" Then
                              Kill sPath & "global.dat"
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\global.dat") <> "" Then
                              Kill App.Path & "\global.dat"
                          End If
                      End If
                      
                      If Dir(sPath & "fichas.dat") <> "" Then
                          CadenadeLog "Detectado: " & sPath & "fichas.dat"
                          Do_Events
                          nLG = FileLen(sPath & "fichas.dat")
                          Do_Events
                          Sleep (3000)
                          Do_Events
                          Do While nLG <> FileLen(sPath & "fichas.dat")
                              Do_Events
                              nLG = FileLen(sPath & "fichas.dat")
                              Sleep (3000)
                              Do_Events
                          Loop
                          On Error Resume Next
                          FileCopy sPath & "\fichas.dat", App.Path & "\fichas.dat"
                          On Error GoTo 0
                          CadenadeLog "LLamada Hydra+ proceso: " & "fichas.dat"
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /VAC", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          CadenadeLog "Procesado: " & "fichas.dat"
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          CadenadeLog "Envia de Modificaciones desde Hydra+..."
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          If Dir(App.Path & "\envok.dat") <> "" Then
                              CadenadeLog "Modificaciones Correctamente Enviadas."
                          Else
                              CadenadeLog "Error en envio Modificaciones."
                          End If
                          'Sleep (1000)
                          If Dir(sPath & "fichas.dat") <> "" Then
                              Kill sPath & "fichas.dat"
                          End If
                          'Sleep (1000)
                          If Dir(App.Path & "\fichas.dat") <> "" Then
                              Kill App.Path & "\fichas.dat"
                          End If
                      End If
                      
                      If Dir(sPath & "articulo.txt") <> "" Then
                          On Error Resume Next
                          FileCopy sPath & "\articulo.txt", App.Path & "\articulo.txt"
                          On Error GoTo 0
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "articulo.txt") <> "" Then
                              Kill sPath & "articulo.txt"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\articulo.txt") <> "" Then
                              Kill App.Path & "\articulo.txt"
                          End If
                      End If
                      
                      If Dir(sPath & "orden.dat") <> "" Or Dir(sPath & "articulo.dat") <> "" Then
                          On Error Resume Next
                          If Dir(sPath & "orden.dat") <> "" Then
                              FileCopy sPath & "\orden.dat", App.Path & "\orden.dat"
                          End If
                          If Dir(sPath & "articulo.dat") <> "" Then
                              FileCopy sPath & "\articulo.dat", App.Path & "\articulo.dat"
                          End If
                          On Error GoTo 0
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(sPath & "orden.dat") <> "" Then
                              Kill sPath & "orden.dat"
                          End If
                          If Dir(App.Path & "\orden.dat") <> "" Then
                              Kill App.Path & "\orden.dat"
                          End If
                          If Dir(sPath & "articulo.dat") <> "" Then
                              Kill sPath & "articulo.dat"
                          End If
                          If Dir(App.Path & "\articulo.dat") <> "" Then
                              Kill App.Path & "\articulo.dat"
                          End If
                      
                      End If
                      
                      
                      'If (Dir(App.Path & "\adicional.dat") <> "" And Dir(App.Path & "\bdp.txt") = "") Then
                      '    If lSerial Then
                      '        '''''''''''''''''''''''''
                      '        reset
                      '        '''''''''''''''''''''''''
                      '    End If
                      '
                      '    MiResultado = STILL_ACTIVE
                      '    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01ADICIONALZ", vbMinimizedFocus))
                      '    Do While MiResultado = STILL_ACTIVE
                      '        GetExitCodeProcess MiProceso, MiResultado
                      '        Do_Events
                      '    Loop
                      '    Sleep (1000)
                      '    If Dir(App.Path & "\adicional.dat") <> "" Then
                      '        Kill App.Path & "\adicional.dat"
                      '    End If
                      'End If
                      
                      If Dir(sPath & "adicional.dat") <> "" And Dir(App.Path & "\bdp.txt") = "" Then
                          
                          On Error Resume Next
                          FileCopy sPath & "adicional.dat", App.Path & "\adicional.dat"
                          On Error GoTo 0
                          
                          Sleep (1000)
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01ADICIONALZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          Sleep (1000)
                          If Dir(App.Path & "\adicional.dat") <> "" Then
                              Kill App.Path & "\adicional.dat"
                              Sleep (1000)
                          End If
                          If Dir(sPath & "adicional.dat") <> "" Then
                              Kill sPath & "adicional.dat"
                              'Sleep (1000)
                          End If
                      End If
                      
                      'seidor
                      
                      If ((Dir(sPath & "cargabalanzas" & sTienda & ".txt") <> "") And Dir(App.Path & "\seidor.txt") <> "") Then
                          On Error Resume Next
                          CadenadeLog "Copiando:" & sPath & "cargabalanzas" & sTienda & ".txt" & " COMO: " & "BAL.DAT"
                          FileCopy sPath & "cargabalanzas" & sTienda & ".txt", App.Path & "\bal.dat"
                          On Error GoTo 0
                          Sleep (4000)
                          If Dir(App.Path & "\bal.dat") <> "" Then
                              CadenadeLog "Copiado BAL.DAT..."
                          Else
                            CadenadeLog "No se ha podido copiar BAL.DAT..."
                          End If
                          
                          If lSerial Then
                              '''''''''''''''''''''''''
                              reset
                              '''''''''''''''''''''''''
                          End If
                          
                          CadenadeLog "EJECUTANDO IMPORTACIÓN..."
                          MiResultado = STILL_ACTIVE
                          MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01EZ", vbMinimizedFocus))
                          Do While MiResultado = STILL_ACTIVE
                              GetExitCodeProcess MiProceso, MiResultado
                              Do_Events
                          Loop
                          CadenadeLog "IMPORTACIÓN EFECTUADA... Borrado de ficheros."
                          
                          Sleep (1000)
                          If Dir(sPath & "cargabalanzas" & sTienda & ".txt") <> "" Then
                              Kill sPath & "cargabalanzas" & sTienda & ".txt"
                          End If
                          Sleep (1000)
                          If Dir(App.Path & "\bal.dat") <> "" Then
                              Kill App.Path & "\bal.dat"
                          End If
                      
                      End If
                      
                      '''''''
                      
                      'If Dir(App.Path & "\envmod.txt") <> "" Then
                      '    If lMod() = True Then
                      '        Switch_mode 0
                      '        MiResultado = STILL_ACTIVE
                      '        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                      '        Do While MiResultado = STILL_ACTIVE
                      '            GetExitCodeProcess MiProceso, MiResultado
                      '            Do_Events
                      '        Loop
                      '        Sleep (1000)
                      '
                      '    End If
                      '
                      'End If

                    
                    End If 'c2f 2.0.21
                    
                    
                End If
                
                'problema con hydra+
                'If Dir(App.Path & "\nover1.txt") <> "" Then
                '    Kill App.Path & "\nover1.txt"
                'End If
                
                Switch_mode nMaqType

            TimeTQT.Interval = MyInterval
            'TimeTQT.Enabled = True
            
        Else
            
            
            'If Dir(App.Path & "\envmod.txt") <> "" And Dir(App.Path & "\hydrarun.txt") <> "" Then
            '        '''''''''''''''''''''''''
            '        reset
            '        '''''''''''''''''''''''''
            'End If
            
            'If Dir(App.Path & "\envmod.txt") <> "" And Dir(App.Path & "\hydrarun.txt") = "" Then
                
                If lMod() = True Then
                    
                    '''''''''''''''''''''''''
                    'reset
                    '''''''''''''''''''''''''
                    
                    MiResultado = STILL_ACTIVE
                    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                    Do While MiResultado = STILL_ACTIVE
                        GetExitCodeProcess MiProceso, MiResultado
                        Do_Events
                    Loop
                    Sleep (1000)
            
                End If
            
            'End If
        
        
        End If
    Else
        
        'If Dir(App.Path & "\envmod.txt") <> "" And Dir(App.Path & "\hydrarun.txt") <> "" Then
        '        '''''''''''''''''''''''''
        '        reset
        '        '''''''''''''''''''''''''
        'End If
        
        'If Dir(App.Path & "\envmod.txt") <> "" And Dir(App.Path & "\hydrarun.txt") = "" Then
            
            If lMod() = True Then
                
                '''''''''''''''''''''''''
                'reset
                '''''''''''''''''''''''''
                
                MiResultado = STILL_ACTIVE
                MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("hydra /01TZ", vbMinimizedFocus))
                Do While MiResultado = STILL_ACTIVE
                    GetExitCodeProcess MiProceso, MiResultado
                    Do_Events
                Loop
                Sleep (1000)
        
            End If
        
        'End If
        
        If lBelRos Then
        
            LeeTickets 0
            If Dir(App.Path & "\nolabel.txt") = "" Then
                LeeTickets 1
                If Dir(App.Path & "\nopack.txt") = "" Then
                    LeeTickets 2
                End If
            End If
        
            Call MiraTotPend
            Set Mybase = dao.OpenDatabase(Base_General)
            Set MyRs = Mybase.OpenRecordset("select * from gtpend where exported=false")
            If Not MyRs.EOF Then
                MyRs.MoveFirst
                Do While Not MyRs.EOF
                    On Error GoTo sigue_T
'FrmExportar.ExportarFichero 0, "boka.txt", 0, True, True
CadenadeLog "Exportaciones BelRos..."
                    FrmExportar.ExportaGTarti "fin_dbal", MyRs.Fields("fecha"), , MyRs.Fields("ntotal")
                    FrmExportar.ExportaLinea "tick_art", MyRs.Fields("fecha")
                    FrmExportar.ExportaCabecera "tick_tot", MyRs.Fields("fecha")
                    Genera_Tot_Grup MyRs.Fields("fecha"), MyRs.Fields("ntotal")
                    
                    MyRs.Edit
                    MyRs.Fields("exported") = True
                    MyRs.Update
sigue_T:
If Err.Number <> 0 Then
                    CadenadeLog "Error:" & CStr(Err.Number) & " --> " & Err.Description
End If
                    On Error GoTo 0
                    MyRs.MoveNext
                Loop
            End If
            MyRs.Close
            Mybase.Close
            Set MyRs = Nothing
            Set Mybase = Nothing
        End If
        
        CuentaGRande = CuentaGRande + 1
        
        'c2f se anula
        'If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
        '    TimeTQT.Enabled = False
        '    Do_Pausa
        '    TimeTQT.Interval = MyInterval
        '    TimeTQT.Enabled = True
        'End If
    End If
    
    'If Dir(App.Path & "\musgrave.txt") <> "" Then
    '    'musgrave --> descarga formulario.
    '    reset
   '
    '    Switch_mode nMaqType
    '
    '    TimeTQT.Enabled = True
    '    ''''''''''''''''''''''''''''''''''
    'Else
        TimeTQT.Enabled = True
    'End If
    
End Sub
'//////
'\\\\\\
Private Sub TmRet_Timer()
    Ret = True
    TmRet.Enabled = False
End Sub

'''''''''¿¿¿
Private Sub CambiaPluASC()
Dim nFich1, nFichDes, ncont, nOrden As Integer
Dim sOrg As String
Dim sCompuesta As String
Dim sDes(10) As String
Dim sPes As String
Dim sPath As String
Dim nTime1, nTime2 As Date
Dim sCad As String

On Error GoTo Vete

lFinPluAsc = False
tmrpluasc.Enabled = True

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
        
        If (sPath & "plu.asc") <> App.Path & "\plu.asc" Then
            FileCopy sPath & "plu.asc", App.Path & "\plu.asc"
        End If
        
        'nTime1 = Time()
        'nTime2 = Time()
        'Do While (Second(nTime2)) < (Second(nTime1) + 3)
        '    nTime2 = Time()
        'Loop
        Sleep (3000)
        
    End If
    
End If


If Dir(App.Path & "\plu.asc") <> "" Then
    
    sCompuesta = ""
    
    nFich1 = FreeFile()
    Open App.Path & "\plu.asc" For Input As #nFich1
    nFichDes = FreeFile()
    Open App.Path & "\balusera.dat" For Output As #nFichDes
    
    Do While Not EOF(nFich1) And Not (lFinPluAsc)
        
        Line Input #nFich1, sOrg
        
        If (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 1) = "0") Or (Mid(sOrg, 13, 1) = "0" And Mid(sOrg, 15, 1) = "2") Then
            If sCompuesta <> "" Then
                For ncont = 1 To 10
                    sCompuesta = sCompuesta & sDes(ncont)
                Next ncont
                Print #nFichDes, sCompuesta & sPes & sCad
                sCompuesta = ""
            End If
            For ncont = 1 To 10
                sDes(ncont) = Space(25)
            Next ncont
            sCompuesta = Mid(sOrg, 54, 6) 'código
            sCompuesta = sCompuesta & Mid(sOrg, 3, 2) 'sección
            If Dir(App.Path & "\plucero.txt") = "" Then
                sCompuesta = sCompuesta & Mid(sOrg, 56, 4) 'plu
            Else
                sCompuesta = sCompuesta & "0000"
            End If
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

'nTime1 = Time()
'nTime2 = Time()
'Do While (Second(nTime2)) < (Second(nTime1) + 2)
'    nTime2 = Time()
'Loop
Sleep (2000)

Vete:
    On Error GoTo 0
    tmrpluasc.Enabled = False
    lFinPluAsc = False
    
End Sub

Private Sub tmrpluasc_Timer()
    lFinPluAsc = True
End Sub
Public Sub Trata_Boka()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Contador = 0
    Load FrmExportar
    FrmExportar.Hide

    CadenadeLog "Se exporta:" & var.PathExpor & "\boka.txt"
    'FrmExportar.MostrarMensajes = False
    FrmExportar.ExportarFichero 0, "boka.txt", 0, True, True
    Set FrmExportar = Nothing
    
End Sub


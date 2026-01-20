VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Object = "{534B7428-6015-4AB1-B994-3F0D09018523}#1.0#0"; "epelcom.ocx"
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
   Begin vb6projectepelcomA.EPELCOM100 Balanza 
      Left            =   2160
      Top             =   720
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SysTrayCtl.cSysTray cSysTray3 
      Left            =   1560
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":08CA
      TrayTip         =   ""
   End
   Begin SysTrayCtl.cSysTray cSysTray2 
      Left            =   840
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":11A4
      TrayTip         =   ""
   End
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   120
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "contenedor.frx":1A7E
      TrayTip         =   ""
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
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    frmMonitor.Show
End Sub
'//////
'\\\\\\
Private Sub cSysTray2_MouseDblClick(Button As Integer, Id As Long)
    frmMonitor.Show
End Sub

'\\\\\\
'//////
Public Sub TimeTQT_Timer()
    Dim MyResul As Long
    Dim MyProc As Long
    Dim MyFich As Integer
    Dim nFlg As Integer
    Dim lRecOK As Boolean
    
    
    Dim nnn, Fich1, resp As Integer
    Dim MiResultado, MiProceso As Long
    Dim sPath As String
    Dim sCommand As String
    
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
    
    CuentaGRande = CuentaGRande + 1
    
    If Dir(App.Path & "\hydrarun.txt") <> "" Then
       Exit Sub
    Else
        If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
            Kill App.Path & "\dirtmp\tqpausa"
        End If
        If Dir(App.Path & "\dirtmp\tqreanuda") <> "" Then
            Kill App.Path & "\dirtmp\tqreanuda"
        End If
    
    End If
    'If Dir(App.Path & "\bal.dat") <> "" Or Dir(App.Path & "\ballss.dat") <> "" Or Dir(App.Path & "\balenv.dat") <> "" Then
    '    Exit Sub
    'End If
    
    If (CuentaGRande >= 30 And TimeAuto <= 5) Or (TimeAuto > 5 And (CuentaGRande * 2) >= TimeAuto Or Trim(Command) = "/01PEEKTOT") Then
        MyContador = MyContador + 1
        If (MyContador >= TimeAuto And TimeAuto <= 5) Or (TimeAuto > 5 And (CuentaGRande * 2) >= TimeAuto Or Trim(Command) = "/01PEEKTOT") Then
            
                TimeTQT.Enabled = False
                
                CuentaGRande = 0
            
                MyContador = 0
                
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    frmControl.cSysTray2.InTray = False
                    frmControl.cSysTray1.InTray = True
                    'Do_Pausa
                    TimeTQT.Interval = MyInterval
                    TimeTQT.Enabled = True
                    
                    Exit Sub
                End If
                
                lRecOK = False
                
                If Not (lNoCoger) Or Dir(sPath & "reco.txt") <> "" Then
                    lNoConnect = False
                    CG_PideTicket 0, 0
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
                
                
                If Dir(App.Path & "\recosine.txt") <> "" Then 'si existe recosine.txt se recogen tiques pero
                                                              ' no se exportan hasta borrado de totales...
                                                              ' Si la hora de borrado de totales es <> "00:00:00"
                    CG_PideTicket 0, 0
                Else
                    Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                End If
                
                'borrado de totales
                If (Time >= dHoraFin And (dHoraFin <> CDate("00:00:00"))) And Dir(App.Path & "\" & Format(Date, "ddmmyy") & ".fin") = "" Then
                    Fich1 = FreeFile()
                    Open App.Path & "\" & Format(Date, "ddmmyy") & ".fin" For Output As #Fich1
                    Print #Fich1, Now
                    Close #Fich1
                    
                    CG_PideTicket 0, 0
                    
                    frmControl.Balanza.Set_Timeout (15)
                    
                    'resp = frmControl.Balanza.Erase_Totals(0) ' sin reinicio.
                    Switch_mode 0

                    resp = frmControl.Balanza.Erase_Totals(1) 'se reinicia numeración
                    
                    frmControl.Balanza.Set_Timeout (5)
                    
                    Principal.Taquion_EnviaTiquetsPendientes True 'euroscale
                    
                End If
                
                Sleep (500)
                If lRecOK Then
                    nFlg = FreeFile()
                    Open sPath & "recok.txt" For Output As #nFlg
                    Close #nFlg
                End If
                
                If Dir(App.Path & "\transpluasc.exe") <> "" Then
                    
                    If Dir(sPath & "plu.asc") <> "" Then
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("transpluasc", vbNormalFocus))
                        Do While MiResultado = STILL_ACTIVE
                            GetExitCodeProcess MiProceso, MiResultado
                            Do_Events
                        Loop
                        
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("cop5 /01EZ", vbNormalFocus))
                        Do While MiResultado = STILL_ACTIVE
                            GetExitCodeProcess MiProceso, MiResultado
                            Do_Events
                        Loop
                    End If
                    
                Else
                    If Dir(sPath & "bal.dat") <> "" Then
                        On Error Resume Next
                        FileCopy sPath & "bal.dat", App.Path & "\bal.dat"
                        On Error GoTo 0
                        Sleep (1000)
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("cop5 /01EZ", vbNormalFocus))
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
                    If Dir(sPath & "global.dat") <> "" Then
                        On Error Resume Next
                        FileCopy sPath & "\global.dat", App.Path & "\global.dat"
                        On Error GoTo 0
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("cop5 /01GLOBAL2", vbNormalFocus))
                        Do While MiResultado = STILL_ACTIVE
                            GetExitCodeProcess MiProceso, MiResultado
                            Do_Events
                        Loop
                        Sleep (1000)
                        If Dir(sPath & "global.dat") <> "" Then
                            Kill sPath & "global.dat"
                        End If
                        Sleep (1000)
                        If Dir(App.Path & "\global.dat") <> "" Then
                            Kill App.Path & "\global.dat"
                        End If
                    End If
                
                End If
                

            TimeTQT.Interval = MyInterval
            TimeTQT.Enabled = True
            
        End If
    Else
        CuentaGRande = CuentaGRande + 1
        'c2f se anula
        'If Dir(App.Path & "\dirtmp\tqpausa") <> "" Then
        '    TimeTQT.Enabled = False
        '    Do_Pausa
        '    TimeTQT.Interval = MyInterval
        '    TimeTQT.Enabled = True
        'End If
    End If
    
End Sub
'//////
'\\\\\\
Private Sub TmRet_Timer()
    Ret = True
    TmRet.Enabled = False
End Sub

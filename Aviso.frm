VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Object = "{1FB9BF8C-0F2A-4F63-8FA9-06976E585D29}#1.0#0"; "ctrlimpatouch.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   0  'None
   Caption         =   "Exportación"
   ClientHeight    =   1875
   ClientLeft      =   3735
   ClientTop       =   3930
   ClientWidth     =   3705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Aviso.frx":0000
   ScaleHeight     =   1875
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin ImportadorTOUCH.ControlImportTouch ControlImport1 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SysTrayCtl.cSysTray cSysTray2 
      Left            =   2040
      Top             =   960
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Aviso.frx":0342
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   2040
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Aviso.frx":21C4
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Timer tmrFtp 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   720
      Top             =   1320
   End
   Begin VB.Timer timetqt 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TmrExclusivo 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer TmrGranTotal 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer TmrCierre 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   600
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const BALSAM = False

Public Sub Balanza_ComError(ByVal Error_Code As Long)
    Dim MyCad As String
    Dim CierraEcoLabel As Boolean
    If Not CancelHumano Then MyCad = CargaCadena(544)
    Select Case EstadoBal
    Case 0
        MyCad = MyCad & " - TouchScale"
    End Select
    frmMonitor.MostrarDato MyCad
    frmMonitor.Label1.Caption = MyCad
    EstadoBal = 1000
    If Not CancelHumano Then frmMonitor.LblError.Visible = True
    canCelar = 1
End Sub


Public Sub Balanza_ComOK()
    If frmMonitor.LuzEnvio.BackColor = vbBlack Then frmMonitor.LuzEnvio.BackColor = vbBlue
    If frmMonitor.LuzEnvio.BackColor = vbBlack Then frmMonitor.LuzEnvio.BackColor = vbRed
    If frmMonitor.LuzEnvio.BackColor <> vbBlue Then
        frmMonitor.LuzEnvio.BackColor = vbBlue
    Else
        frmMonitor.LuzEnvio.BackColor = vbRed
    End If
    If frmMonitor.LuzRecibe.BackColor <> vbBlue Then
        frmMonitor.LuzRecibe.BackColor = vbBlue
    Else
        frmMonitor.LuzRecibe.BackColor = vbRed
    End If
    Select Case EstadoBal
    Case 0
        frmEpelsa.LblGA.ForeColor = vbGreen
        frmEpelsa.LblGA.Caption = "TouchScale -> " & Format(Now, "hh:mm:ss")
    End Select
End Sub


Private Sub Balanza_InventoryReceived(ByVal plu_code As Long, ByVal plu_order As Long, ByVal plu_name As String)
    frmMonitor.MostrarDato "Cod : " & Format(plu_code, "000000") & " " & Format(plu_order, "00000000") & " " & plu_name
End Sub

Private Sub Balanza_ItemReceived(ByVal mach_code As Long, ByVal plu_code As Long)
    frmMonitor.MostrarDato "Item : " & plu_code
End Sub

Private Sub Balanza_OrderInventoryReceived(ByVal plu_code As Long, ByVal plu_order As Long, ByVal plu_inventory As Long)
    frmMonitor.MostrarDato "Cod : " & Format(plu_code, "000000") & " " & Format(plu_order, "00000000") & " " & Format(plu_inventory, "00000000")
End Sub

Private Sub Balanza_OrderReceived(ByVal plu_code As Long, ByVal plu_order As Long, ByVal plu_name As String)
    frmMonitor.MostrarDato "Cod : " & Format(plu_code, "000000") & " " & Format(plu_order, "00000000") & " " & plu_name
End Sub

Private Sub balanza_TicketReceived(ByVal NTicket As Long, ByVal NSection As Long, ByVal NScale As Long, Cancel As Long)
    frmMonitor.MostrarDato CargaCadena(866) & " : " & NTicket & " Sec. : " & NSection & " Eqp. : " & NScale
    If canCelar = 1 Then Cancel = 1
End Sub


Private Sub cSysTray1_MouseDblClick(Button As Integer, id As Long)
    If frmEpelsa.Visible = False Then
        If HaySeguridad Then
            If frmusuario.Visible = True Then Exit Sub
            cSysTray1.InTray = False
            If EncontrarUsuario = False Then
                cSysTray1.InTray = True
                Exit Sub
            End If
        End If
        frmEpelsa.Visible = True
        cSysTray1.InTray = True
        Do_Events
    Else
        frmEpelsa.Visible = False
    End If
End Sub

#If BALSAM = True Then
Public Sub Switch_M(ByVal MiP As String)

    Dim MyFich As Integer
    Dim s As String
    Dim BufDetiene As Variant
    Dim valorip As String
    Dim sTi As String
    Dim nResp As Integer

    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    EstadoBal = 1000
    On Error GoTo 0
    'nutricion
    'carga_data_Vitamin
    ''''''''''

    MiP = Val(left(MiP, 3)) & "." & Val(Mid(MiP, 5, 3)) & "." & Val(Mid(MiP, 9, 3)) & "." & Val(Right(MiP, 3))

    BufDetiene = frmMonitor.cmdmonitorstop.Enabled
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.TmrEnableCancel.Interval = 5000
    frmMonitor.TmrEnableCancel.Enabled = True
    frmMonitor.Image2.Visible = True
    frmMonitor.TmrConecta.Interval = 100
    frmMonitor.TmrConecta.Enabled = True

    '''''''''''''''''''''''''''''''
    'Conectar con sAddIp y sPortBal
    '''''''''''''''''''''''''''''''
    On Error Resume Next
    frmMonitor.MostrarDato CargaCadena(986) & "-->" & MiP & "," & sPortBal
    CadenadeLog CargaCadena(986) & "-->" & MiP & "," & sPortBal
    Me.Refresh
    Do_Events

    nResp = ElPing(MiP)
    If nResp = 1 Then

        If Dir(App.Path & "\odbc53.txt") <> "" Then
            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
        Else
            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
        End If

        CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
        CadenadeLog "Error:" & CStr(Err.Number)

        If Err.Number <> 0 Then

            On Error GoTo 0
            EstadoBal = 1000
            GoTo finMSWITCH
        Else
            EstadoBal = 0
        End If

    Else
        frmMonitor.MostrarDato "No Ping:" & MiP
        EstadoBal = 1000
        GoTo finMSWITCH
    End If

    frmMonitor.cmdmonitorstop.Enabled = BufDetiene
    frmMonitor.Image2.Visible = False
    frmMonitor.TmrConecta.Enabled = False

finMSWITCH:
    If EstadoBal <> 0 Then
        canCelar = 1
        If Not CancelHumano Then frmMonitor.LblError.Visible = True
        frmMonitor.Label1.Caption = CargaCadena(544)
    Else
        canCelar = 0
    End If

End Sub


#Else
Public Sub Switch_M(ByVal MiP As String)

    Dim MyFich As Integer
    Dim s As String
    Dim BufDetiene As Variant
    Dim valorip As String
    Dim sTi As String
    Dim nResp As Integer

    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    EstadoBal = 1000
    On Error GoTo 0
    'nutricion
    'carga_data_Vitamin
    ''''''''''

    MiP = Val(left(MiP, 3)) & "." & Val(Mid(MiP, 5, 3)) & "." & Val(Mid(MiP, 9, 3)) & "." & Val(Right(MiP, 3))

    BufDetiene = frmMonitor.cmdmonitorstop.Enabled
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.TmrEnableCancel.Interval = 5000
    frmMonitor.TmrEnableCancel.Enabled = True
    frmMonitor.Image2.Visible = True
    frmMonitor.TmrConecta.Interval = 100
    frmMonitor.TmrConecta.Enabled = True

    '''''''''''''''''''''''''''''''
    'Conectar con sAddIp y sPortBal
    '''''''''''''''''''''''''''''''
    On Error Resume Next
    frmMonitor.MostrarDato CargaCadena(986) & "-->" & MiP & "," & sPortBal
    CadenadeLog CargaCadena(986) & "-->" & MiP & "," & sPortBal
    Me.Refresh
    Do_Events

    nResp = ElPing(MiP)
    If nResp = 1 Then

        If Dir(App.Path & "\odbc53.txt") <> "" Then
            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        Else
            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        End If

        CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
        CadenadeLog "Error:" & CStr(Err.Number)

        If Err.Number <> 0 Then

            On Error GoTo 0
            EstadoBal = 1000
            GoTo finMSWITCH
        Else
            EstadoBal = 0
        End If

    Else
        frmMonitor.MostrarDato "No Ping:" & MiP
        EstadoBal = 1000
        GoTo finMSWITCH
    End If

    frmMonitor.cmdmonitorstop.Enabled = BufDetiene
    frmMonitor.Image2.Visible = False
    frmMonitor.TmrConecta.Enabled = False

finMSWITCH:
    If EstadoBal <> 0 Then
        canCelar = 1
        If Not CancelHumano Then frmMonitor.LblError.Visible = True
        frmMonitor.Label1.Caption = CargaCadena(544)
    Else
        canCelar = 0
    End If

End Sub
#End If

#If BALSAM = True Then
Public Sub Switch_mode(ByVal Scaletype As Integer)
'comprobar conexión...con la base de datos
    Dim MyFich As Integer
    Dim s As String
    Dim BufDetiene As Variant
    Dim valorip As String
    Dim sTi As String
    Dim nResp As Integer
    'Dim Cm As String
    'Dim Sm As String

    'lastIpconn
    ''''tqretail
    On Error GoTo sigueSw
    If MyConnObj.State <> 0 Or MyConnObj.Errors.Count <> 0 Then
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
    Do_Events
    'nutricion
    'carga_data_Vitamin
    ''''''''''
    Do_Events

    'If MyConnObj.State = 1 And MyConnObj.Errors.Count = 0 Then
    '    On Error GoTo 0
    '    Do_Events
    '    MyConnObj.Errors.Refresh
    '    Do_Events
    '    Sleep (50)
    '    Do_Events
    '    MyConnObj.Errors.Refresh
    '    Do_Events
    '    CadenadeLog "Conexión ya realizada..."
    '    Exit Sub
    'End If
    '''
sigueSw:
    'if Err.Number <> 0 Then
    On Error GoTo 0
    'End If

    On Error GoTo finSWITCH

    ' Gama Alta Multiservidora (NetID)
    If MultiEuroscale Then
        MUE_Sistemas
        Scaletype = 700
    End If
    '''''''''

    If Scaletype >= 700 And Scaletype < 800 Then
        If MultiEuroscale Then
            MyIP = SisEur.Sistemas(Scaletype - 700).IP
        Else
            If Dir(App.Path & "\eqpip.txt") <> "" Then
                MyFich = FreeFile
                Open App.Path & "\eqpip.txt" For Input As #MyFich
                Line Input #MyFich, s
                Close #MyFich
                MyIP = s
            End If
        End If
        Scaletype = 0
    End If

    If MultiEuroscale Then
        If MyIP <> MultiIP Then
            MultiIP = MyIP
        End If
    End If

    'If Not Taquion Then
    BufDetiene = frmMonitor.cmdmonitorstop.Enabled
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.TmrEnableCancel.Interval = 5000
    frmMonitor.TmrEnableCancel.Enabled = True
    frmMonitor.Image2.Visible = True
    frmMonitor.TmrConecta.Interval = 100
    frmMonitor.TmrConecta.Enabled = True
    'End If


    If MultiEuroscale Then
        '''''''''''''''''''''''''''''
        'Conectar con MyIP y sPortBal
        ' Si hay conexión MultiIP=MyIP, si no hay conexión MultiIP=""
        '''''''''''''''''''''''''''''
        frmMonitor.MostrarDato CargaCadena(986) & "(SAM)-->" & MyIP & "," & sPortBal
        On Error Resume Next

        'MyConnObj.CommandTimeout = 1
        nResp = ElPing(MyIP)
        If nResp = 1 Then
            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source=" & MyIP & "," & sPortBal
            If Dir(App.Path & "\odbc53.txt") <> "" Then
                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
            Else
                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
            End If
            If Err.Number <> 0 Then
                On Error GoTo 0
                MultiIP = ""
                EstadoBal = 1000
                Exit Sub
            Else
                EstadoBal = 0
                MultiIP = MyIP
            End If
        Else
            On Error GoTo 0
            frmMonitor.MostrarDato "No Ping:" & MyIP
            MultiIP = ""
            EstadoBal = 1000
            Exit Sub
        End If
        On Error GoTo finSWITCH
    Else

        If leeIp Then
            valorip = Check_Dns_File()
            If valorip <> "" Then
                ''''''''''''''''''''''''''''''''
                'Conectar con Valorip y sportbal
                ''''''''''''''''''''''''''''''''
                On Error Resume Next
                frmMonitor.MostrarDato CargaCadena(986) & "(SAM)-->" & valorip & "," & sPortBal
                Me.Refresh
                Do_Events
                nResp = ElPing(sAddIp)
                If nResp = 1 Then

                    'MyConnObj.CommandTimeout = 1
                    MyConnObj.Errors.Clear
                    'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source=" & valorip & "," & sPortBal
                    If Dir(App.Path & "\odbc53.txt") <> "" Then
                        MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & valorip & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                    Else
                        MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & valorip & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                    End If
                    If Err.Number <> 0 Or MyConnObj.State <> 1 Or MyConnObj.Errors.Count <> 0 Then
                        On Error GoTo 0
                        EstadoBal = 1000
                        Exit Sub
                    Else
                        EstadoBal = 0
                    End If

                Else
                    EstadoBal = 1000
                    Exit Sub
                End If
                On Error GoTo finSWITCH
            Else
                '''''''''''''''''''''''''''''''
                'Conectar con sAddIp y sPortBal
                '''''''''''''''''''''''''''''''
                On Error Resume Next
                frmMonitor.MostrarDato CargaCadena(986) & "-->" & sAddIp & "," & sPortBal
                CadenadeLog CargaCadena(986) & "-->" & sAddIp & "," & sPortBal
                Me.Refresh
                Do_Events

                nResp = ElPing(sAddIp)
                If nResp = 1 Then

                    'MyConnObj.CommandTimeout = 1
                    If Dir(App.Path & "\office.txt") <> "" Then
                        If TiendaActual = 0 Then
                            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale01;Data Source=" & sAddIp & "," & sPortBal
                            If Dir(App.Path & "\odbc53.txt") <> "" Then
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                            Else
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                            End If
                        Else
                            If TiendaActual < 100 Then
                                sTi = Format(TiendaActual, "00")
                            Else
                                sTi = Format(TiendaActual, "000")
                            End If
                            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale" & sTi & ";Data Source=" & sAddIp & "," & sPortBal
                            If Dir(App.Path & "\odbc53.txt") <> "" Then
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                            Else
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                            End If
                        End If
                    Else
                        'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source = " & sAddIp & ", " & sPortBal
                        If Dir(App.Path & "\odbc53.txt") <> "" Then
                            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                        Else
                            'MyConnObj.Open "DRIVER={MariaDB ODBC 3.0 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=3"
                            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=epelMaster; PWD=epelsa; OPTION=34"
                            'restaurar XS
                            'MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=mbd; UID=pcscale; PWD=epelsa; OPTION=34"
                        End If
                        'CadenadeLog "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                        'MyConnObj.Open "Provider=SQLNCLI10;Uid=sa;password=masterkey;Initial Catalog=TouchScale;DataTypeCompatibility=80;Data Source=" & sAddIp & "," & sPortBal
                    End If

                    CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
                    CadenadeLog "Error:" & CStr(Err.Number)

                    If Err.Number <> 0 Or MyConnObj.State = 0 Then

                        On Error GoTo 0
                        EstadoBal = 1000
                        GoTo finSWITCH
                    Else
                        EstadoBal = 0
                    End If
                    On Error GoTo finSWITCH
                Else
                    frmMonitor.MostrarDato "No Ping:" & sAddIp
                    EstadoBal = 1000
                    GoTo finSWITCH
                End If

            End If
        End If

    End If

    frmMonitor.TmrEnableCancel.Enabled = False

    'If Not Taquion Then
    frmMonitor.cmdmonitorstop.Enabled = BufDetiene
    frmMonitor.Image2.Visible = False
    frmMonitor.TmrConecta.Enabled = False
    'End If
finSWITCH:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Switch_mode--> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        On Error GoTo 0
        EstadoBal = 1000
    End If
    If EstadoBal <> 0 Then
        canCelar = 1
        If Not CancelHumano Then frmMonitor.LblError.Visible = True
        frmMonitor.Label1.Caption = CargaCadena(544)
    Else
        'Call LeeIVAS
        canCelar = 0
    End If
    On Error GoTo 0
End Sub
#Else
Public Sub Switch_mode(ByVal Scaletype As Integer)
'comprobar conexión...con la base de datos
    Dim MyFich As Integer
    Dim s As String
    Dim BufDetiene As Variant
    Dim valorip As String
    Dim sTi As String
    Dim nResp As Integer
    'Dim Cm As String
    'Dim Sm As String
    Dim MyRecAux As New ADODB.Recordset

    CadenadeLog "Comprobar conexión con Base de datos..."

    'lastIpconn
    ''''tqretail
    On Error GoTo sigueSw
    If MyConnObj.State <> 0 Or MyConnObj.Errors.Count <> 0 Then
        CadenadeLog "Se detecta error en conexión establecida previamente... Se cierra conexión previa"
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
    Do_Events
    'nutricion
    'carga_data_Vitamin
    ''''''''''
    Do_Events

    'If MyConnObj.State = 1 And MyConnObj.Errors.Count = 0 Then
    '    On Error GoTo 0
    '    Do_Events
    '    MyConnObj.Errors.Refresh
    '    Do_Events
    '    Sleep (50)
    '    Do_Events
    '    MyConnObj.Errors.Refresh
    '    Do_Events
    '    CadenadeLog "Conexión ya realizada..."
    '    Exit Sub
    'End If
    '''
sigueSw:
    'if Err.Number <> 0 Then
    On Error GoTo 0
    'End If

    On Error GoTo finSWITCH

    ' Gama Alta Multiservidora (NetID)
    If MultiEuroscale Then
        MUE_Sistemas
        Scaletype = 700
    End If
    '''''''''

    If Scaletype >= 700 And Scaletype < 800 Then
        If MultiEuroscale Then
            MyIP = SisEur.Sistemas(Scaletype - 700).IP
        Else
            If Dir(App.Path & "\eqpip.txt") <> "" Then
                MyFich = FreeFile
                Open App.Path & "\eqpip.txt" For Input As #MyFich
                Line Input #MyFich, s
                Close #MyFich
                MyIP = s
            End If
        End If
        Scaletype = 0
    End If

    If MultiEuroscale Then
        If MyIP <> MultiIP Then
            MultiIP = MyIP
        End If
    End If

    'If Not Taquion Then
    BufDetiene = frmMonitor.cmdmonitorstop.Enabled
    frmMonitor.cmdmonitorstop.Enabled = False
    frmMonitor.TmrEnableCancel.Interval = 5000
    frmMonitor.TmrEnableCancel.Enabled = True
    frmMonitor.Image2.Visible = True
    frmMonitor.TmrConecta.Interval = 100
    frmMonitor.TmrConecta.Enabled = True
    'End If


    If MultiEuroscale Then
        '''''''''''''''''''''''''''''
        'Conectar con MyIP y sPortBal
        ' Si hay conexión MultiIP=MyIP, si no hay conexión MultiIP=""
        '''''''''''''''''''''''''''''
        frmMonitor.MostrarDato CargaCadena(986) & "-->" & MyIP & "," & sPortBal & " --> Multi."
        On Error Resume Next
        CadenadeLog "Verficación PING..."
        'MyConnObj.CommandTimeout = 1
        nResp = ElPing(MyIP)
        If nResp = 1 Then
            CadenadeLog "Se detecta ping con Servidora."
            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source=" & MyIP & "," & sPortBal
            If Dir(App.Path & "\odbc53.txt") <> "" Then
                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                CadenadeLog "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 1"
            Else
                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                CadenadeLog "DRIVER={MySQL ODBC 5.1 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 2"
            End If
            If Err.Number <> 0 Then
                On Error GoTo 0
                MultiIP = ""
                EstadoBal = 1000
                Exit Sub
            Else
                EstadoBal = 0
                MultiIP = MyIP
            End If
        Else
            CadenadeLog "Sin ping con Servidora..."
            On Error GoTo 0
            frmMonitor.MostrarDato "No Ping:" & MyIP
            MultiIP = ""
            EstadoBal = 1000
            Exit Sub
        End If
        On Error GoTo finSWITCH
    Else

        If leeIp Then
            CadenadeLog "Leida dirección IP"
            valorip = Check_Dns_File()
            If valorip <> "" Then
                CadenadeLog "Se detecta DNS..."
                ''''''''''''''''''''''''''''''''
                'Conectar con Valorip y sportbal
                ''''''''''''''''''''''''''''''''
                On Error Resume Next
                frmMonitor.MostrarDato CargaCadena(986) & "-->" & valorip & "," & sPortBal
                Me.Refresh
                Do_Events
                CadenadeLog "Se realiza Ping DNS..."
                nResp = ElPing(valorip)
                If nResp = 1 Then
                    CadenadeLog "Ping con Servidora DNS..."
                    'MyConnObj.CommandTimeout = 1
                    MyConnObj.Errors.Clear
                    'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source=" & valorip & "," & sPortBal
                    If Dir(App.Path & "\odbc53.txt") <> "" Then
                        MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & valorip & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                        'MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & valorip & ":" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                        CadenadeLog "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 3"
                    Else
                        MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & valorip & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                        'MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & valorip & ":" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                        CadenadeLog "DRIVER={MySQL ODBC 5.1 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 4"
                    End If
                    If Err.Number <> 0 Or MyConnObj.State <> 1 Or MyConnObj.Errors.Count <> 0 Then
                        On Error GoTo 0
                        EstadoBal = 1000
                        Exit Sub
                    Else
                        EstadoBal = 0
                    End If

                Else
                    CadenadeLog "Sin Ping DNS..."
                    EstadoBal = 1000
                    Exit Sub
                End If
                On Error GoTo finSWITCH
            Else
                '''''''''''''''''''''''''''''''
                'Conectar con sAddIp y sPortBal
                '''''''''''''''''''''''''''''''
                On Error Resume Next
                frmMonitor.MostrarDato CargaCadena(986) & "-->" & sAddIp & "," & sPortBal
                CadenadeLog CargaCadena(986) & "-->" & sAddIp & "," & sPortBal
                Me.Refresh
                Do_Events
                CadenadeLog "Realizar ping a Servidora..."
                nResp = ElPing(sAddIp)
                If nResp = 1 Then
                    CadenadeLog "Con ping a Servidora."
                    'MyConnObj.CommandTimeout = 1
                    If Dir(App.Path & "\office.txt") <> "" Then
                        If TiendaActual = 0 Then
                            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale01;Data Source=" & sAddIp & "," & sPortBal
                            If Dir(App.Path & "\odbc53.txt") <> "" Then

                                CadenadeLog "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 5"
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                            Else

                                CadenadeLog "DRIVER={MySQL ODBC 5.1 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 6"
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                            End If
                        Else
                            If TiendaActual < 100 Then
                                sTi = Format(TiendaActual, "00")
                            Else
                                sTi = Format(TiendaActual, "000")
                            End If
                            'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale" & sTi & ";Data Source=" & sAddIp & "," & sPortBal
                            If Dir(App.Path & "\odbc53.txt") <> "" Then
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                                CadenadeLog "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 7"
                            Else
                                MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
                                CadenadeLog "DRIVER={MySQL ODBC 5.1 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 8"
                            End If
                        End If
                    Else
                        'MyConnObj.Open "Provider=SQLOLEDB;Uid=sa;password=masterkey;Initial Catalog=TouchScale;Data Source = " & sAddIp & ", " & sPortBal
                        If Dir(App.Path & "\odbc53.txt") <> "" Then

                            CadenadeLog "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 9"
                            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                        Else

                            CadenadeLog "DRIVER={MySQL ODBC 5.1 ANSI Driver}; SERVER=" & MyIP & "; Port=" & sPortBal & "; DATABASE=pcscale; OPTION=3 --> 10"
                            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                        End If
                        'CadenadeLog "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & sAddIp & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"

                        'MyConnObj.Open "Provider=SQLNCLI10;Uid=sa;password=masterkey;Initial Catalog=TouchScale;DataTypeCompatibility=80;Data Source=" & sAddIp & "," & sPortBal
                    End If

                    CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
                    CadenadeLog "Error:" & CStr(Err.Number)

                    If Err.Number <> 0 Or MyConnObj.State = 0 Then
                        CadenadeLog "Detectado error enn conexión:" & CStr(Err.Number) & " Desc.:" & Err.Description
                        On Error GoTo 0
                        EstadoBal = 1000
                        GoTo finSWITCH
                    Else
                        EstadoBal = 0
                    End If
                    On Error GoTo finSWITCH
                Else
                    CadenadeLog "Sin ping a Servidora."
                    frmMonitor.MostrarDato "No Ping:" & sAddIp
                    EstadoBal = 1000
                    GoTo finSWITCH
                End If

            End If
        End If

    End If

    frmMonitor.TmrEnableCancel.Enabled = False

    'If Not Taquion Then
    frmMonitor.cmdmonitorstop.Enabled = BufDetiene
    frmMonitor.Image2.Visible = False
    frmMonitor.TmrConecta.Enabled = False
    'End If
finSWITCH:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Switch_mode--> Err:" & CStr(Err.Number) & " Desc.:" & Err.Description
        On Error GoTo 0
        EstadoBal = 1000
    End If
    If EstadoBal <> 0 Then
        canCelar = 1
        If Not CancelHumano Then frmMonitor.LblError.Visible = True
        frmMonitor.Label1.Caption = CargaCadena(544)
    Else
        'Call LeeIVAS
        canCelar = 0
        'On Error GoTo 0
        'On Error GoTo interno123
        MyRecAux.Open "select * from pcscale.status", MyConnObj, adOpenKeyset, adLockOptimistic
        If Not MyRecAux.EOF Then
            bbDDver = MyRecAux.Fields("version")
        End If
        MyRecAux.Close
        Set MyRecAux = Nothing
        'interno123:
        'On Error GoTo 0
    End If
    On Error GoTo 0
End Sub
#End If


Public Sub timetqt_timer()
    Dim bucle As Long
    Dim C11 As Integer
    Dim Origen As Integer
    Dim Buffer As String
    Dim SalvaTexto As String
    Dim Archivo As Integer
    Dim MiHora As Integer
    Dim MiMinuto As Integer
    Dim BufEpelsa As Boolean
    Dim BufEtiqueta As String
    Dim AAA As Variant
    Dim hHandle As Integer
    Dim nMy As Form
    Dim nFich As Integer
    Dim MyProc As Long
    Dim MyResul As Long
    Dim Ret As Long
    Dim nInt As Long
    Dim sTaqui As String

    timetqt.Enabled = False

    lAutoRecEnv = True

    'frmControl.Switch_mode 0
    'If EstadoBal <> 0 Then
    '    frmEpelsa.LblGA.Caption = "Touch. Con. Error -> " & Format(Now, "hh:mm:ss")
    'Else
    frmEpelsa.LblGA.Caption = "Touch. Con.:-> " & Format(Now, "hh:mm:ss")

    'End If
    'Workspaces(0).close

    '//////////////////////////////////////////////////
    '// Recogida automática de totales
    '//////////////////////////////////////////////////
    If RE_TOTAL Then
        BufEpelsa = frmEpelsa.Enabled
        BufEtiqueta = frmEpelsa.LblComunicaciones.Caption
        frmEpelsa.Enabled = False
        'timetqt.Enabled = False
        frmEpelsa.LblComunicaciones.Caption = "Realizando Cierre , espere..."

        If Dir(App.Path & "\borrartotales.txt") <> "" Then
            Set nMy = New FrmExportar
            nMy.AN_FinDeDia True
            Unload nMy
            Set nMy = Nothing
        Else
            Set nMy = New FrmExportar
            nMy.AN_RecogerTotales (False)
            Unload nMy
            Set nMy = Nothing

            FrmResturarBackup.Realizar_Backup False

            Call Trata_Fin_Dia_2

            If Dir(Miruta & "\end_day.bat") <> "" Then
                ChDir Miruta
                Shell Miruta & "\end_day.bat"
            End If

            On Error Resume Next
            Kill App.Path & "\exporta\*.*"
            On Error GoTo 0

        End If

        If lFornes Then Call export_BL(True)

        CadenadeLog "Proceso de Fin de día realizado"
        frmEpelsa.LblGA.Caption = "Proceso de Fin de día realizado"
        For bucle = 1 To 58
            frmEpelsa.LblSC10.Caption = "Reset. Pausa ..." & Format((59 - bucle), "00")
            Sleep (1000)
            Do_Events
        Next bucle
        frmEpelsa.LblGA.Caption = ""
        frmEpelsa.LblSC10.Caption = ""
        'If Dir(App.Path & "\ticket.dat") <> "" Then
        '    Kill App.Path & "\ticket.dat"
        'End If
        RE_TOTAL = False
        frmEpelsa.LblComunicaciones.Caption = BufEtiqueta
        frmEpelsa.Enabled = BufEpelsa
        'timetqt.Enabled = True
    End If

    'Workspaces(0).close

    '////////////////////////////////////////////////
    '// cierre automático del programa
    '////////////////////////////////////////////////
    If CI_Salir Or Dir(App.Path & "\killhydra") <> "" Then
        'timetqt.Enabled = False
        frmEpelsa.Enabled = False
        frmEpelsa.LblComunicaciones.Caption = "Cierre Automático , espere..."
        If Dir(App.Path & "\killhydra") = "" Then
            For bucle = 1 To 48
                frmEpelsa.LblSC10.Caption = "Reset. Pausa ..." & Format((49 - bucle), "00")
                Sleep (1000)
                Do_Events
            Next bucle
        End If
        salir_programa
    End If
    MiHora = Hour(Now)
    MiMinuto = Minute(Now)

    'Check_Plugin_Auto

    '*************************
    ' no se lanza si se hace una pausa en recogida
    '*************************
    'Workspaces(0).close
    If PausaTiquets Then
        frmEpelsa.HabilitarMenues
        'timetqt.Enabled = True
        Exit Sub
    End If

    'Workspaces(0).close

    If lCogeTiquet = True And cgdtiquet = False And descAuto Then
        If cuenTaTqt >= Val(left(descautotime, 3)) Then

            'If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 0 Then
            '    Dim siP, susR, spsS
            '    nFich = FreeFile()
            '    Open App.Path & "\getputfileftp.dat" For Input As #nFich
            '    'Do While Not EOF(nFich)
            '        Line Input #nFich, siP
            '        Line Input #nFich, susR
            '        Line Input #nFich, spsS
            '    'Loop
            '    Close #nFich
            '    CadenadeLog "conexión FTP..."
            '    Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
            '    nRetConnFTP = Ret
            'End If

            Ret = conecta_FTP()

            '***************************************
            ' desactiva los menues de comunicaciones
            '***************************************
            cgdtiquet = True
            'timetqt.Enabled = False
            frmEpelsa.BloquearMenues
            '1.0.4 se modifica para recoger modificaciones...
            '**********************************
            ' TAQUION : ' mira si existe(n) fichero(s) de importación, y lo(s) procesa
            '**********************************
            Silencio = True
            If Taquion Then    'c2f caspiunza
                On Error Resume Next
                If Dir(App.Path & "\orden.dat") <> "" Then
                    If FileLen(App.Path & "\orden.dat") < 20 Then
                        Kill App.Path & "\orden.dat"
                    End If
                End If
                If Dir(App.Path & "\ordtec.dat") <> "" Then
                    If FileLen(App.Path & "\ordtec.dat") < 20 Then
                        Kill App.Path & "\ordtec.dat"
                    End If
                End If
                ''Carrefour_NW --> Ejemplo decompress images, para artículos se usará BAL3.dat, es decir, transorden.exe copiará bal2.dat a bal3.dat
                'If Dir(App.Path & "\IMG001W001.T151013.Z") Then
                '    MyResul = STILL_ACTIVE
                '    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\7z e IMG001W001.T151013.Z -y", vbMinimized))
                '    Do While MyResul = STILL_ACTIVE
                '        GetExitCodeProcess MyProc, MyResul
                '        Do_Events
                '    Loop
                '    MyResul = STILL_ACTIVE
                '    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(App.Path & "\tartool -x IMG001W001.T151013 images", vbMinimized))
                '    Do While MyResul = STILL_ACTIVE
                '        GetExitCodeProcess MyProc, MyResul
                '        Do_Events
                '    Loop
                '
                'End If
                ''''''''''''''
                If Dir(App.Path & "\getputfileftp.dat") <> "" And nRetConnFTP = 1 Then

                    CadenadeLog "comienzo..."
                    Ret = Trae_FTP("fampos.dat")
                    If FileLen(App.Path & "\fampos.dat") = 0 Then Kill App.Path & "\fampos.dat"
                    Ret = Trae_FTP("venpos.dat")
                    If FileLen(App.Path & "\venpos.dat") = 0 Then Kill App.Path & "\venpos.dat"
                    Ret = Trae_FTP("artcon.dat")
                    If FileLen(App.Path & "\artcon.dat") = 0 Then Kill App.Path & "\artcon.dat"
                    Ret = Trae_FTP("clientes.dat")
                    If FileLen(App.Path & "\clientes.dat") = 0 Then Kill App.Path & "\clientes.dat"
                    Ret = Trae_FTP("clientesx.dat")
                    If FileLen(App.Path & "\clientesx.dat") = 0 Then Kill App.Path & "\clientesx.dat"

                    If LCase(TaquionFichero) <> LCase("artcon.dat") And LCase(TaquionFichero) <> LCase("fampos.dat") And LCase(TaquionFichero) <> LCase("venpos.dat") And LCase(TaquionFichero) <> LCase("clientes.dat") And LCase(TaquionFichero) <> LCase("clientesx.dat") Then
                        Ret = Trae_FTP(TaquionFichero)
                        If FileLen(App.Path & "\" & TaquionFichero) = 0 Then Kill App.Path & "\" & TaquionFichero
                    End If

                    sTaqui = ""
                    nInt = InStr(1, TaquionFichero, ".")
                    If nInt <> 0 Then
                        sTaqui = Mid(TaquionFichero, 1, nInt - 1) & ".rar"
                    End If
                    Ret = Trae_FTP("fampos.rar")
                    If FileLen(App.Path & "\fampos.rar") = 0 Then Kill App.Path & "\fampos.rar"
                    Ret = Trae_FTP("venpos.rar")
                    If FileLen(App.Path & "\venpos.rar") = 0 Then Kill App.Path & "\venpos.rar"
                    Ret = Trae_FTP("artcon.rar")
                    If FileLen(App.Path & "\artcon.rar") = 0 Then Kill App.Path & "\artcon.rar"
                    Ret = Trae_FTP("clientes.rar")
                    If FileLen(App.Path & "\clientes.rar") = 0 Then Kill App.Path & "\clientes.rar"
                    If sTaqui <> "" Then
                        If LCase(sTaqui) <> LCase("artcon.rar") And LCase(sTaqui) <> LCase("staqui") And LCase(sTaqui) <> LCase("venpos.rar") And LCase(sTaqui) <> LCase("clientes.rar") Then
                            Ret = Trae_FTP(sTaqui)
                            If FileLen(App.Path & "\" & sTaqui) = 0 Then Kill App.Path & "\" & sTaqui
                        End If
                    End If
                    If Dir(App.Path & "\fampos.rar") <> "" Then
                        nInt = Extrae(0, "fampos.rar", "")
                    End If
                    If Dir(App.Path & "\venpos.rar") <> "" Then
                        nInt = Extrae(0, "venpos.rar", "")
                    End If
                    If Dir(App.Path & "\artcon.rar") <> "" Then
                        nInt = Extrae(0, "artcon.rar", "")
                    End If
                    If Dir(App.Path & "\clientes.rar") <> "" Then
                        nInt = Extrae(0, "clientes.rar", "")
                    End If
                    If Dir(App.Path & "\" & sTaqui) <> "" Then
                        nInt = Extrae(0, sTaqui, "")
                    End If

                    If Dir(App.Path & "\fampos.rar") <> "" Then
                        Kill App.Path & "\fampos.rar"
                    End If
                    If Dir(App.Path & "\venpos.rar") <> "" Then
                        Kill App.Path & "\venpos.rar"
                    End If
                    If Dir(App.Path & "\artcon.rar") <> "" Then
                        Kill App.Path & "\artcon.rar"
                    End If
                    If Dir(App.Path & "\clientes.rar") <> "" Then
                        Kill App.Path & "\clientes.rar"
                    End If
                    If Dir(App.Path & "\" & sTaqui) <> "" Then
                        Kill App.Path & "\" & sTaqui
                    End If

                    CadenadeLog "FIN."

                    'CadenadeLog "Leer Ficheros FTP..."
                    'If nRetConnFTP = 1 Then 'ha sido posible conectar
                    '    Ret = Recibir_Fichero_FTP("fampos.dat", App.Path & "\fampos.dat", 0)
                    '    'Ret = Recibir_Fichero_FTP("asciihydratouch.doc", App.Path & "\ash.doc", 0)
                    '    If Ret = 1 Then
                    '        If FileLen(App.Path & "\fampos.dat") > 0 Then
                    '            Ret = Borrar_Fichero_FTP("fampos.dat", 0)
                    '        Else
                    '            Kill App.Path & "\fampos.dat"
                    '        End If
                    '    End If
                    '    Ret = Recibir_Fichero_FTP("venpos.dat", App.Path & "\venpos.dat", 0)
                    '    If Ret = 1 Then
                    '        If FileLen(App.Path & "\venpos.dat") > 0 Then
                    '            Ret = Borrar_Fichero_FTP("venpos.dat", 0)
                    '        Else
                    '            Kill App.Path & "\venpos.dat"
                    '        End If
                    '    End If
                    '    Ret = Recibir_Fichero_FTP("artcon.dat", App.Path & "\artcon.dat", 0)
                    '    If Ret = 1 Then
                    '        If FileLen(App.Path & "\artcon.dat") > 0 Then
                    '            Ret = Borrar_Fichero_FTP("artcon.dat", 0)
                    '        Else
                    '            Kill App.Path & "\artcon.dat"
                    '        End If
                    '    End If
                    '    Ret = Recibir_Fichero_FTP("clientes.dat", App.Path & "\clientes.dat", 0)
                    '    If Ret = 1 Then
                    '        If FileLen(App.Path & "\clientes.dat") > 0 Then
                    '            Ret = Borrar_Fichero_FTP("clientes.dat", 0)
                    '        Else
                    '            Kill App.Path & "\clientes.dat"
                    '        End If
                    '    End If
                    '    'TaquionFichero
                    '    Ret = Recibir_Fichero_FTP(TaquionFichero, App.Path & "\" & TaquionFichero, 0)
                    '    If Ret = 1 Then
                    '        If FileLen(App.Path & "\" & TaquionFichero) > 0 Then
                    '            Ret = Borrar_Fichero_FTP(TaquionFichero, 0)
                    '        Else
                    '            Kill App.Path & "\" & TaquionFichero
                    '        End If
                    '    End If
                    '
                    '    'Ret = Err.LastDllError
                    '    'If Ret <> 0 Then 'error en DLL
                    '    '    If Ret <> 12003 Then
                    '    '        Ret = Desconectar_FTP(0)
                    '    '        FTP_hOpen(0) = 0
                    '    '        FTP_HandleConect(0) = 0
                    '    '        nRetConnFTP = 0
                    '    '    End If
                    '    'End If
                    '
                    ''End If
                    ''Ret = Desconectar_FTP(0)
                    ''CadenadeLog "Leidos Ficheros FTP..."

                End If
                '1.0.4 revisar
                'If Dir(App.Path & "\enlascii.exe") <> "" Then
                '    nFich = FreeFile()
                '    Open App.Path & "\enlascii.cfg" For Output As #nFich
                '    Print #nFich, CStr(UsaFichasVacuno)
                '    Print #nFich, TaquionFichero
                '    Print #nFich, MiruTa
                '    Print #nFich, CStr(Supeco)
                '    Print #nFich, CStr(FamiliasFijas)
                '    Print #nFich, CStr(SubSeccionesFijas)
                '    Print #nFich, CStr(RechazarPlu0)
                '    Print #nFich, Base_General
                '    Print #nFich, FamiliaEspecial
                '    Print #nFich, CStr(DebugActivo)
                '    Print #nFich, CStr(decimales)
                '    Print #nFich, CStr(UsaEuro)
                '    Print #nFich, EtiquetaEspecial
                '    Close #nFich
                '    MyResul = STILL_ACTIVE
                '    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiruTa & "\" & "enlascii.exe", vbNormalFocus))
                '    Do While MyResul = STILL_ACTIVE
                '        GetExitCodeProcess MyProc, MyResul
                '        Do_Events
                '    Loop
                '
                '    'c2f para que el enlace sea completo se debería implementar en enlascii.exe
                '    ' lo que está anulado aquí...
                '
                '    If PausaTiquets = True Then frmEpelsa.HabilitarMenues
                '    frmEpelsa.CmdComunicaciones(3).Enabled = True
                '
                '
                'Else
                Set nMy = New FrmExportar
                If Dir(App.Path & "\raro") <> "" Then
                    Call Absorver_Estructura
                    Call nMy.TQ_ImportaFicheros
                    AN_BaseMaestra False
                Else
                    If lFornes Then
                        Call transforma_Fornes
                        If Dir(App.Path & "\global.dat") <> "" Then
                            Importa_Global_Dat
                            Kill App.Path & "\global.dat"
                        End If
                    Else
                        Call nMy.TQ_ImportaFicheros
                    End If
                End If
                Unload nMy
                Set nMy = Nothing
                'End If

                If Dir(App.Path & "\memoria.vbe") <> "" Then
                    Shell (App.Path & "\memoria.vbe")
                End If
                If Dir(App.Path & "\orden.dat") <> "" Then
                    Kill App.Path & "\orden.dat"
                End If
                If Dir(App.Path & "\ordtec.dat") <> "" Then
                    Kill App.Path & "\ordtec.dat"
                End If
                On Error GoTo 0
            End If
            '************ desde aquí
            ' /TAQUION
            '************
            ' Lanza la recogida de totales sin borrado de fin de dia
            '***********************************************************
            cgdtiquet = True
            CI_EstadoSistema = 10

            'If Taquion Then Silencio = True
            'Silencio = True
            Set nMy = New FrmExportar
            nMy.AN_RecogerTotales False
            Unload nMy
            Set nMy = Nothing
            'If Taquion Then Silencio = False



            '1.7.2 grventas
            If EnvioContinuo Then
                'Set nMy = New FrmExportar
                AN_BaseMaestra True
                'Unload nMy
                'Set nMy = Nothing
            End If
            'Silencio = False
            'If Taquion Then Silencio = False

            '****** hasta aquí se origina consumo memoria...
            ' /Recogida
            '************
            ' Taquion / Storeflow, envía
            ' los tiquets pendientes, por Socket o ASCII
            '***********************************************
            If (Taquion And TaquionTiquets) Or Dir(App.Path & "\ahorramas.cfg") <> "" Then
                frmEpelsa.Enabled = False
                frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905)
                If lFornes Then
                    Call export_BL(True)
                Else
                    If UsaGamaAlta Then
                        Set nMy = New FrmExportar
                        Call nMy.Taquion_EnviaTiquetsPendientes(True)
                        Unload nMy
                        Set nMy = Nothing
                    End If
                End If
                frmEpelsa.LblGA.Caption = CargaCadena(904)
                frmEpelsa.Enabled = True

            End If
            If Not CI_Salir And Not RE_TOTAL Then CI_EstadoSistema = 1
            cuenTaTqt = 0
            cgdtiquet = False
            'Silencio = False
        Else
            'Silencio = True
            cuenTaTqt = cuenTaTqt + 1
            If cgdtiquet = False And descAuto Then
                If Taquion Then
                    If TqDirecto = False Then
                        If (ModificacionesPendientes(0) And 128) = 128 Then ModificacionesPendientes(0) = ModificacionesPendientes(0) - 128
                    End If
                End If
                If (ModificacionesPendientes(0) And 128) = 128 Then
                    Silencio = True

                    AN_BaseMaestra True, , 7
                    Silencio = False
                    If canCelar = 0 Then
                        If (ModificacionesPendientes(0) And 128) = 128 Then ModificacionesPendientes(0) = ModificacionesPendientes(0) - 128
                    End If
                End If
            End If
            'Silencio = False
        End If
    End If

    'Workspaces(0).close

    '*********************************************************
    ' reactiva los menues de comunicaciones si procede
    '*********************************************************
    If Right(descautotime, 3) = "min" Then
        frmEpelsa.HabilitarMenues
    End If

    If Dir(App.Path & "\nover1.txt") <> "" Then
        On Error Resume Next
        Unload frmMonitor
        Set frmMonitor = Nothing
        Unload FrmExportar
        Set FrmExportar = Nothing
        On Error GoTo 0
    End If

    Silencio = False
    'Workspaces(0).close


    lAutoRecEnv = False

    timetqt.Enabled = True
    ''''''''''''''
    'CadenadeLog "pasa por recogida..."
    'Unload frmControl
    'Set frmControl = Nothing
    'Load frmControl
End Sub



Private Sub TmrCierre_Timer()
    Dim Activacion As Boolean
    Dim Archivo As Integer
    Dim Buffer As String
    '**************************
    ' comprueba si tiene que exportar
    ' un fichero
    '**************************
    If StoreFlow Then
        TmrCierre.Enabled = False
        Exit Sub
    End If
    On Error GoTo siguiente
    If Dir(Miruta & "\asorden\export.ord") <> "" And cgdtiquet = False Then
        Archivo = FreeFile()
        Open Miruta & "\asorden\export.ord" For Input As #Archivo
        Do Until EOF(Archivo)
            Line Input #Archivo, Buffer
            Buffer = Trim(Buffer)
            If Buffer <> "" Then
                If Dir(Miruta & "\etc\" & Buffer & ".fet") <> "" Then
                    Load FrmExportar
                    FrmExportar.MostrarMensajes = False
                    FrmExportar.ExportarFichero 0, Buffer
                    Unload FrmExportar
                End If
            End If
        Loop
        Close #Archivo
        Kill Miruta & "\asorden\export.ord"
    End If
    '**************************
    ' comprueba si tiene que cerrar
    '**************************
siguiente:
    On Error GoTo 0
    If Dir(Miruta & "\asorden\off.ord") <> "" Then
        If DebugActivo Then CadenadeLog ("ord.off encontrado --> Salir Programa")
        CI_Salir = True
        Sleep (500)
        Kill App.Path & "\asorden\off.ord"
    End If
    If (Val(Hour(Now)) = Val(CI_HORA) And Val(Minute(Now)) = Val(CI_MINUTO)) And CI_Salir = False Then
        CI_Salir = True
    End If
End Sub



Private Sub TmrExclusivo_Timer()
    Exclusivo_Fin = True
End Sub

Private Sub tmrFtp_Timer()
    tmrFtp.Enabled = False
    CadenadeLog "Se cierra conexión FTP tras 30 segundos..."
    Call Cierra_FTP
End Sub

Private Sub TmrGranTotal_Timer()
    Dim Activacion As Boolean
    If StoreFlow Then
        TmrGranTotal.Enabled = False
        Exit Sub
    End If
    If Dir(Miruta & "\asorden\gtotal.ord") <> "" Then
        RE_TOTAL = True
        Sleep (500)
        Kill Miruta & "\asorden\gtotal.ord"
    End If
    If ((Val(Hour(Now)) = Val(RE_HORA) And Val(Minute(Now)) = Val(RE_MINUTO))) And (Not RE_TOTAL) Then
        RE_TOTAL = True
    End If
End Sub
Private Sub Check_Plugin_Auto()
    Dim Buf As String
    Dim hora As Integer
    Dim Minuto As Integer
    Dim Arch As Integer
    Dim Reint As Integer
    Dim OKLanzar As Boolean
    If PlugActivo Then Exit Sub
    If AutoLanzado Then Exit Sub
    If cgdtiquet = True Then Exit Sub
    OKLanzar = True
    If Dir(App.Path & "\etc\autoplugin.prog") <> "" Then
        If Dir(App.Path) & "\hydplugin.exe" <> "" Then
            Reint = 0
            Do Until Reint >= 10
                Arch = FreeFile()
                On Error Resume Next
                Open App.Path & "\etc\autoplugin.prog" For Input As #Arch
                If Err.Number = 0 Then
                    Exit Do
                Else
                    Reint = Reint + 1
                    Sleep (50)
                End If
            Loop
            If Reint < 10 Then
                On Error Resume Next
                Line Input #Arch, Buf
                If Err.Number = 0 Then
                    If IsNumeric(Buf) Then
                        hora = Val(Buf)
                    Else
                        OKLanzar = False
                    End If
                Else
                    OKLanzar = False
                End If
                On Error Resume Next
                Line Input #Arch, Buf
                If Err.Number = 0 Then
                    If IsNumeric(Buf) Then
                        Minuto = Val(Buf)
                    Else
                        OKLanzar = False
                    End If
                Else
                    OKLanzar = False
                End If
                Line Input #Arch, Buf
                If Err.Number <> 0 Then
                    OKLanzar = False
                End If
                Close #Arch
                If OKLanzar Then
                    If Hour(Now) >= hora And Minute(Now) >= Minuto Then
                        If Dir(App.Path & "\etc\autoplugin." & Format(Now, "yymmdd")) = "" Then
                            '******************************
                            ' lanzamiento del plugin
                            '******************************
                            On Error Resume Next
                            Arch = FreeFile()
                            Reint = 0
                            Do Until Reint >= 10
                                On Error Resume Next
                                Open App.Path & "\etc\autoplugin." & Format(Now, "yymmdd") For Output As #Arch
                                If Err.Number = 0 Then
                                    Close #Arch
                                    Exit Do
                                Else
                                    Reint = Reint + 1
                                    Sleep (50)
                                End If
                            Loop
                            AutoLanzado = True
                            lCogeTiquet = False
                            frmEpelsa.Enabled = False
                            PlugActivo = True
                            FrmHlink.Start "000"
                            FrmHlink.WPlugin Buf
                            If PlugOculta Then frmEpelsa.Hide
                            If Iconificar Or IconificarInicio Then frmControl.cSysTray1.InTray = False
                            Do While PlugActivo = True
                                lCogeTiquet = False
                                Do_Events
                            Loop
                            If Iconificar Or IconificarInicio Then frmControl.cSysTray1.InTray = True
                            If PlugOculta Then MostrarInterfaz
                            frmEpelsa.Enabled = True
                            lCogeTiquet = True
                            AutoLanzado = False
                        End If
                    End If
                End If
            Else
                On Error Resume Next
                Close #Arch
                On Error GoTo 0
                If DebugActivo Then CadenadeLog "Imposible acceder a \etc\autoplugin.prog"
            End If
        End If
    End If
End Sub
'''''''
Public Sub Trae_Ficheros_Supeco()
    Dim Arch As Integer
    Dim Arch2 As Integer
    Dim Buffer As String
    Dim bucle As Integer
    If Dir(Supeco_Art, vbDirectory) <> "" Then
        For bucle = 1 To 2
            If Dir(Supeco_Art & "\artibal" & Format(bucle, "0")) <> "" Then
                Arch = FreeFile()
                On Error Resume Next
                Open Supeco_Art & "\artibal" & Format(bucle, "0") For Input As #Arch
                If Err.Number <> 0 Then
                    CadenadeLog "SUPECO : Sin acceso de lectura a artibal" & Format(bucle, "0")
                    Exit Sub
                End If
                Arch2 = FreeFile()
                Open App.Path & "\artibal" For Append As #Arch2
                Do Until EOF(Arch)
                    Line Input #Arch, Buffer
                    Print #Arch2, Buffer
                Loop
                Close #Arch
                Close #Arch2
                On Error Resume Next
                Kill Supeco_Art & "\artibal" & Format(bucle, "0")
                If Err.Number <> 0 Then
                    CadenadeLog "SUPECO : Archivo bloqueado --> No puedo borrar ARTIBAL" & Format(bucle, "0")
                End If
            End If
        Next bucle
    End If
    On Error GoTo 0
End Sub
'\\\\\\\\\\\
'///////////
Public Function ini_import()
    frmControl.ControlImport1.FamFijas = FamiliasFijas
    frmControl.ControlImport1.SubFijas = SubSeccionesFijas
    ControlImport1.No_modWGH = False
    ControlImport1.Idioma = id
    If UsaEuro Then
        ControlImport1.Euros = True
    Else
        ControlImport1.Euros = False
    End If
    ControlImport1.NumeroDecimales = decimales
    ControlImport1.FactorEuro = 1
End Function
'\\\\\\\\\\\
'///////////

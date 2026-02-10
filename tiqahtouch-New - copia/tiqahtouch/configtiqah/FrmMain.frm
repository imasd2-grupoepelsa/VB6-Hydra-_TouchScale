VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Configuración TiqAH"
   ClientHeight    =   6720
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   7200
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   7200
   Begin VB.CheckBox Check2 
      Caption         =   "Recoger Tiques sin Exportar."
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CheckBox CheckMonitor 
      Caption         =   "Visualizar Monitores."
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   3240
      Width           =   1815
   End
   Begin VB.ComboBox CmbLocal 
      Height          =   315
      Left            =   3360
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Auto-Servicio Local."
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Envasado Local."
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Super Local."
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Venta Local."
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Auto-Servicio RED."
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Envasado RED."
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recepción Tique Super RED."
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CheckBox ChkTiques 
      Caption         =   "Recepción Tique Venta RED."
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   2535
   End
   Begin MSMask.MaskEdBox MaskBor 
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtGn 
      Height          =   285
      Left            =   2880
      MaxLength       =   12
      TabIndex        =   13
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox TxtPort 
      Height          =   285
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin MSMask.MaskEdBox MaskIp 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   15
      Mask            =   "###.###.###.###"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox CmbT 
      Height          =   315
      ItemData        =   "FrmMain.frx":08CA
      Left            =   2760
      List            =   "FrmMain.frx":08CC
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label LblLocal 
      Caption         =   "Peticiones Red antes de Peticiones Local:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label LblBor 
      Caption         =   "Hora de Borrado de Totales:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label LblGn 
      Caption         =   "Nombre Fichero de Tiques General:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label LblPort 
      Caption         =   "Puerto TCP Servidora:"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label LblIP 
      Caption         =   "Dirección IP Servidora:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label LblN 
      Caption         =   "Nombre Fichero Tiques Individual:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label LblTRe 
      Caption         =   "Tiempo Aprox. Recogida Tiques:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label LblEx 
      Caption         =   "Path Exportación Tiques:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label LblIm 
      Caption         =   "Path Importación Datos:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check2_Click()
Dim nFich As Integer

    If Check1.Value = vbChecked Then
        nFich = FreeFile()
        Open App.Path & "\recosine.txt" For Output As #nFich
        Close #nFich
    Else
        If Dir(App.Path & "\recosine.txt") <> "" Then
            Kill App.Path & "\recosine.txt"
        End If
    End If
    

End Sub

Private Sub Form_Load()
Dim nC As Integer
Dim sS As String

    FrmHelpName.ListHelpName.AddItem "B = Dígito número de Balanza."
    FrmHelpName.ListHelpName.AddItem "S = Dígito Número de Sección."
    FrmHelpName.ListHelpName.AddItem "T = Dígito número Tique."
    FrmHelpName.ListHelpName.AddItem "X = Dígito Modo de Venta."
    FrmHelpName.ListHelpName.AddItem "C = Dígito Número de Cliente."
    FrmHelpName.ListHelpName.AddItem "V = Dígito Número de Vendedor."
    
    For nC = 0 To 58 Step 2
        CmbT.AddItem CStr(nC) & " seg." ', (6 + nC - 1)
    Next nC
    For nC = 1 To 5
        CmbT.AddItem CStr(nC) & " min." ', nC - 1
    Next nC
    For nC = 10 To 200 Step 10
        CmbLocal.AddItem CStr(nC)
    Next nC
    CmbT.ListIndex = 0
    CmbLocal.ListIndex = 0
    'Lectura o creación TiqAh.ini por defecto...
    Call LeeINI
    
    
End Sub

Private Sub CheckMonitor_Click()
Dim n As Integer
If CheckMonitor.Value Then
    n = FreeFile()
    Open App.Path & "\nsmon.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\nsmon.txt") <> "" Then
        Kill App.Path & "\nsmon.txt"
    End If
End If
End Sub

Private Sub LeeINI()
    Dim Arch As Integer
    Dim Buf1 As String, Buf2 As String
    Dim nCTQ As Integer
    
    If Dir(App.Path & "\tiqah.ini") = "" Then
        On Error GoTo ErrorCritico1
        Arch = FreeFile()
        Open App.Path & "\tiqah.ini" For Output As #Arch
        If Dir(App.Path & "\infoser.txt") <> "" Then
            Write #Arch, "Path de SC10, ruta de red al concentrador              ", "z:\balanzas"
            Write #Arch, "Intervalo de recogida en minutos (mayor que cero   )   ", "62" '2
            Write #Arch, "Máximo número de tiquets por Recogida (0 no hay límite)", "0"
            Write #Arch, "Intervalo de retardo en milisegundos, entre copias     ", "5000"
            Write #Arch, "Nombre del Fichero de Tiquets Individual               ", "BBTTTT" '5
            Write #Arch, "Ruta del directorio almacen de ficheros de exportacion ", "c:\Balanzas" '6
            Write #Arch, "Dirección IP / Port Balanza Servidora 386              ", "127.000.000.001:6000" '7
            Write #Arch, "Nombre del Fichero de Tiquets General                  ", "tqgen" '8
            Write #Arch, "Hora de Borrado de Totales                             ", "22:30:00" '9
            '************
            ' de que tipo
            ' 0 --> Venta
            ' 1 --> Super
            ' 2 --> Envasado
            ' 3 --> Autoservicio
            ' 4 --> Venta Local
            ' 5 --> Super Local
            ' 6 --> Envasado Local
            ' 7 --> Autoservicio Local
            ' 9 --> anulados....
            '*************************
            Write #Arch, "Recepción Tique Venta RED                              ", "1"
            Write #Arch, "Recepción Tique Super RED                              ", "0"
            Write #Arch, "Recepción Tique Envasado RED                           ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio RED                      ", "0"
            Write #Arch, "Recepción Tique Venta Local                            ", "1"
            Write #Arch, "Recepción Tique Super Local                            ", "0"
            Write #Arch, "Recepción Tique Envasado Local                         ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio Local                    ", "0"
            Write #Arch, "Número peticiones Red antes de Locales                 ", "100"
        Else
            Write #Arch, "Path de SC10, ruta de red al concentrador              ", "z:\balanzas"
            Write #Arch, "Intervalo de recogida en minutos (mayor que cero   )   ", "6" '2
            Write #Arch, "Máximo número de tiquets por Recogida (0 no hay límite)", "0"
            Write #Arch, "Intervalo de retardo en milisegundos, entre copias     ", "5000"
            Write #Arch, "Nombre del Fichero de Tiquets Individual               ", "BBTTTT" '5
            Write #Arch, "Ruta del directorio almacen de ficheros de exportacion ", "" '6
            Write #Arch, "Dirección IP / Port Balanza Servidora 386              ", "127.000.000.001:6000" '7
            Write #Arch, "Nombre del Fichero de Tiquets General                  ", "tqgen" '8
            Write #Arch, "Hora de Borrado de Totales                             ", "22:30:00" '9
        
            Write #Arch, "Recepción Tique Venta RED                              ", "1"
            Write #Arch, "Recepción Tique Super RED                              ", "0"
            Write #Arch, "Recepción Tique Envasado RED                           ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio RED                      ", "0"
            Write #Arch, "Recepción Tique Venta Local                            ", "1"
            Write #Arch, "Recepción Tique Super Local                            ", "0"
            Write #Arch, "Recepción Tique Envasado Local                         ", "0"
            Write #Arch, "Recepción Tique Auto-Servicio Local                    ", "0"
            Write #Arch, "Número peticiones Red antes de Locales                 ", "100"
            
        End If
        Close #Arch
        On Error GoTo 0
    End If
    
    On Error GoTo ErrorCritico2
    Arch = FreeFile()
    Open App.Path & "\tiqah.ini" For Input As #Arch
    Input #Arch, Buf1, Buf2 '1
    PathSc10 = Buf2
    Input #Arch, Buf1, Buf2 '2
    If IsNumeric(Buf2) Then
        TimeAuto = Val(Buf2)
        TimeAutoSeg = 0
        'sInftxt1 = CStr(TimeAuto)
        If TimeAuto <= 5 Then
            'sInftxt1 = sInftxt1 & " Minutos."
        Else
            If TimeAuto > 60 Then
                TimeAuto = TimeAuto - 60
                TimeAutoSeg = TimeAuto
                'sInftxt1 = CStr(TimeAutoSeg) & " Segundos."
            Else
                'sInftxt1 = sInftxt1 & " Segundos."
            End If

        End If
        If TimeAuto < 1 Then
            TimeAuto = 10
            'lNoCoger = True
            'sInftxt1 = "10 segundos (Si Reco.txt)"
            'CadenadeLog "Tiempo de recogida automática incorrecto"
            'GoTo ErrorCritico2
        End If
    Else
        'CadenadeLog "Tiempo de recogida automática incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '3
    If IsNumeric(Buf2) Then
        NumMax = Val(Buf2)
        If NumMax < 0 Then
            CadenadeLog "Número máximo de tiquets incorrecto"
            GoTo ErrorCritico2
        End If
    Else
        CadenadeLog "Número máximo de tiquets incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '4
    If IsNumeric(Buf2) Then
        MiliTiquet = Val(Buf2)
        If MiliTiquet < 0 Then
            CadenadeLog "Número máximo de tiquets incorrecto"
            GoTo ErrorCritico2
        End If
    Else
        CadenadeLog "Número máximo de tiquets incorrecto"
        GoTo ErrorCritico2
    End If
    Input #Arch, Buf1, Buf2 '5
    Var.TqNombre = Buf2
    sInftxt2 = Buf2
    
    Input #Arch, Buf1, Buf2 '6
    If Buf2 = "" Then
        Buf2 = App.Path
    End If
    Var.PathExpor = Buf2
    sInftxt3 = Buf2
    '
    Input #Arch, Buf1, Buf2 '7
    sAddIP = Buf2
    sInftxt4 = Buf2
    '
    Input #Arch, Buf1, Buf2 '8
    sTiqGen = Buf2
    sInftxt5 = Buf2
    
    sTiqGen = sTiqGen & Format(Day(Now), "00")
    If Month(Now) < 10 Then
        sTiqGen = sTiqGen & CStr(Month(Now))
    Else
        sTiqGen = sTiqGen & Chr(Asc("A") + Month(Now) - 10)
    End If
    Input #Arch, Buf1, Buf2 '9
    sHoraFin = Buf2
    sInftxt6 = Buf2
    dHoraFin = CDate(sHoraFin)
    
    On Error GoTo INIERR
    For nCTQ = 0 To 7
        Input #Arch, Buf1, Buf2
        TipoTQT(nCTQ) = Buf2
    Next nCTQ
    
    Input #Arch, Buf1, Buf2
    
    nContLocal = Val(Buf2)
    
INIERR:
    If Err.Number <> 0 Then
        For nCTQ = 0 To 7
        Select Case nCTQ
            Case 0
            TipoTQT(nCTQ) = "1"
            Case 1
            TipoTQT(nCTQ) = "0"
            Case 2
            TipoTQT(nCTQ) = "0"
            Case 3
            TipoTQT(nCTQ) = "0"
            Case 4
            TipoTQT(nCTQ) = "1"
            Case 5
            TipoTQT(nCTQ) = "0"
            Case 6
            TipoTQT(nCTQ) = "0"
            Case 7
            TipoTQT(nCTQ) = "0"
        End Select
        
        Next nCTQ
        nContLocal = 200
        
    End If
    
    Close #Arch
    On Error GoTo 0
    Exit Sub
ErrorCritico1:
    MsgBox "No se puede escribir el fichero de configuración TIQAH.INI", vbCritical, "Recogida Contínua de Tiquets"
    End
ErrorCritico2:
    MsgBox "fichero de configuración TIQAH.INI incorrecto", vbCritical, "Recogida Contínua de Tiquets"
    End
End Sub

Private Sub TxtName_DblClick()
    FrmHelpName.Visible = True
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If Chr(KeyAscii) <> "B" And Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "T" And Chr(KeyAscii) <> "X" And Chr(KeyAscii) <> "C" And Chr(KeyAscii) <> "V" Then
            KeyAscii = 0
        End If
    
    End If

End Sub


Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

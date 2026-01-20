VERSION 5.00
Begin VB.Form FrmAsistente 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente para la creación de la Tienda"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7770
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Siguiente >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "< Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   3015
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   -240
      Width           =   7575
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   15
         Top             =   2400
         Width           =   5055
      End
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   13
         Top             =   2040
         Width           =   5055
      End
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   12
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox TxtTienda 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Teléfono de Contacto "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Código Postal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Provincia "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Población"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Dirección"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Nombre de la tienda "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Indique los datos identificativos de la tienda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   7452
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3372
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   4800
      Width           =   7575
      Begin VB.TextBox TxtTCP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         MaxLength       =   5
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Puerto TCP de la balanza Servidora"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Indique los Parámetros de Comunicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   7575
      Begin VB.TextBox TxtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Número de la Tienda :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Indique número de Tienda TouchScale/KScale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7455
      End
   End
End
Attribute VB_Name = "FrmAsistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Indice As Integer
Private Sub CambiarIdioma()
    Label4.Caption = CargaCadena(1138)
    Label1(0).Caption = CargaCadena(1456)
    Label1(1).Caption = CargaCadena(1135)
    'Label1(2).Caption = CargaCadena(1136)
    Label1(3).Caption = CargaCadena(1137)
    Command1(1).Caption = "< " & CargaCadena(1168)
    Command1(2).Caption = CargaCadena(1169) & " >"
    Command1(0).Caption = CargaCadena(288)
    'Option2(0).Caption = CargaCadena(1170)
    'Option2(1).Caption = CargaCadena(1171)
    'Option2(2).Caption = CargaCadena(1172)
    'Option2(3).Caption = CargaCadena(1173)
    Label2(0).Caption = CargaCadena(552)
    Label2(1).Caption = CargaCadena(1127)
    Label2(2).Caption = CargaCadena(1128)
    Label2(3).Caption = CargaCadena(1129)
    Label2(4).Caption = CargaCadena(1130)
    Label2(5).Caption = CargaCadena(1131)
    Label3(1).Caption = CargaCadena(260)
    'Label3(2).Caption = CargaCadena(260)
    'Label3(3).Caption = CargaCadena(261)
    'Label3(4).Caption = CargaCadena(262)
    'Label3(5).Caption = CargaCadena(263)
    'Label3(6).Caption = CargaCadena(1066)
    'Label3(7).Caption = CargaCadena(1174)
    'Label5(0).Caption = CargaCadena(951)
    'Label5(1).Caption = CargaCadena(952)
End Sub
'Private Sub CmbC9_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            Select Case CmbC9(0).ListIndex
'                Case 0
'                    CmbSerie(1).ListIndex = 3
'                    CmbSerie(2).ListIndex = 0
'                    CmbSerie(3).ListIndex = 1
'                Case 1
'                    CmbSerie(1).ListIndex = 3
'                    CmbSerie(2).ListIndex = 2
'                    CmbSerie(3).ListIndex = 0
'                Case 2
'                    CmbSerie(1).ListIndex = 1
'                    CmbSerie(2).ListIndex = 2
'                    CmbSerie(3).ListIndex = 0
'                Case 3
'                    CmbSerie(1).ListIndex = 1
'                    CmbSerie(2).ListIndex = 0
'                    CmbSerie(3).ListIndex = 0
'                Case 4
'                    CmbSerie(1).ListIndex = 0
'                    CmbSerie(2).ListIndex = 2
'                    CmbSerie(3).ListIndex = 0
'            End Select
'    End Select
'End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Resp
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Indice > 0 Then
                Indice = Indice - 1
                If Indice = 0 Then Command1(1).Enabled = False
                Pon_Frame
            End If
            Command1(2).Caption = CargaCadena(1169) & " >"
        Case 2
            If Indice < 2 Then
                Command1(1).Enabled = True
                Indice = Indice + 1
                If Indice = 3 Then
                    'If Option1(6).Value = True Then
                        TxtTCP.Text = "3306"
                    'Else
                    '    TxtTCP.Text = "6000"
                    'End If
                End If
                Pon_Frame
           
            Else
                Resp = Check_Valores()
                Resp = MsgBox(CargaCadena(1008), vbYesNo + vbDefaultButton2 + vbQuestion)
                If Resp = vbYes Then
                    Crea_la_Tienda
                    Unload Me
                End If
            End If
            If Indice = 2 Then
                Command1(2).Caption = CargaCadena(287)
            Else
                Command1(2).Caption = CargaCadena(1169) & " >"
            End If
    End Select
           
End Sub
Private Function Check_Valores() As Long
    Dim Bucle As Integer
    Dim MiRecurso As Integer
    Dim Resp As Long
    Resp = 0
    'For Bucle = 0 To 3
    '    If Option2(Bucle).Value Then MiRecurso = Bucle
    'Next Bucle
    Select Case MiRecurso
        Case 0
            If Val(TxtTCP.Text) < 1024 Or Val(TxtTCP.Text) > 65535 Then
                MsgBox CargaCadena(958), vbCritical
                Resp = 1
            End If
        Case 1
        Case 2
        Case 3
            'If Trim(TxtTelefono.Text) = "" Then
            '    MsgBox CargaCadena(1139), vbCritical
            '    Resp = 1
            'End If
    End Select
End Function
Private Sub Form_Load()
    Dim Bucle As Integer
    CambiarIdioma
    'Option1(0).Value = True
    Me.Height = 4440
    Me.Icon = FrmPrincipal.Icon
    Me.Caption = CargaCadena(1178)
    Indice = 0
    Command1(1).Enabled = False
    Command1(2).Enabled = True
    Command1(2).Caption = CargaCadena(1169) & " >"
    'CmbC9(0).ListIndex = 0
    'CmbC9(1).ListIndex = 0
    Pon_Frame
    TxtNumero.Text = Val(Consulta_NTiendas + 1)
        
    'If Option1(6).Value = True Then
        TxtTCP.Text = "3306"
    'Else
    '    TxtTCP.Text = "6000"
    'End If
    'For Bucle = 0 To 3
    '    CmbSerie(Bucle).ListIndex = 0
    'Next Bucle
End Sub


Private Sub Form_Unload(Cancel As Integer)
    DescargaForm = True
End Sub

'Private Sub Option1_Click(Index As Integer)
'    If Option1(0).Value = True Then
'        Option2(0).Enabled = True
'        Option2(1).Enabled = True
'        Option2(2).Enabled = False
'        Option2(3).Enabled = True
'        Option2(0).Value = True
'        Label5(0).Visible = False
'        Label5(1).Visible = False
'        CmbC9(0).Visible = False
'        CmbC9(1).Visible = False
'        CmbSerie(1).Enabled = True
'        CmbSerie(2).Enabled = True
'        CmbSerie(3).Enabled = True
'    End If
'    If Option1(6).Value = True Then 'ScalePos 1.0.8
'        Option2(0).Enabled = True
'        Option2(1).Enabled = False
'        Option2(2).Enabled = False
'        Option2(3).Enabled = False
'        Option2(0).Value = True
'        Label5(0).Visible = False
'        Label5(1).Visible = False
'        CmbC9(0).Visible = False
'        CmbC9(1).Visible = False
'        CmbSerie(1).Enabled = False
'        CmbSerie(2).Enabled = False
'        CmbSerie(3).Enabled = False
'    End If
    
'    If Option1(1).Value = True Then
'        Option2(0).Enabled = False
'        Option2(1).Enabled = True
'        Option2(2).Enabled = True
'        Option2(3).Enabled = True
'        Option2(1).Value = True
'        Label5(0).Visible = False
'        Label5(1).Visible = False
'        CmbC9(0).Visible = False
'        CmbC9(1).Visible = False
'        CmbSerie(1).Enabled = True
'        CmbSerie(2).Enabled = True
'        CmbSerie(3).Enabled = True
'    End If
'    If Option1(2).Value = True Then
'        Option2(0).Enabled = False
'        Option2(1).Enabled = True
'        Option2(2).Enabled = False
'        Option2(3).Enabled = True
'        Option2(1).Value = True
'        Label5(0).Visible = True
'        Label5(1).Visible = True
'        CmbC9(0).Visible = True
'        CmbC9(1).Visible = True
'        CmbSerie(1).Enabled = False
'        CmbSerie(2).Enabled = False
'        CmbSerie(3).Enabled = False
'        CmbC9(0).ListIndex = 2
'    End If
'    If Option1(3).Value Or Option1(4).Value Then
'        Label1(3).Caption = CargaCadena(1165)
'
'    Else
'        Label1(3).Caption = CargaCadena(1137)
'
'    End If
'    If Option1(3).Value Or Option1(4).Value Or Option1(5).Value Then
'        Option2(0).Enabled = False
'        Option2(1).Enabled = True
'        Option2(2).Enabled = False
'        Option2(3).Enabled = True
'        Option2(1).Value = True
'        Option2_Click (1)
'    End If
'    If Option1(5).Value Then
'        Label3(2).Visible = True
'        CmbSerie(0).Visible = True
'    End If
'End Sub
Private Sub Pon_Frame()
    Dim Bucle As Integer
    For Bucle = 0 To 2
        If Indice = Bucle Then
            Frame1(Bucle).Top = 0
            Frame1(Bucle).Left = 0
            Frame1(Bucle).Visible = True
        Else
            Frame1(Bucle).Visible = False
        End If
    Next Bucle
End Sub

'Private Sub Option2_Click(Index As Integer)
'    Dim Bucle As Integer
'    For Bucle = 1 To 7
'        Label3(Bucle).Visible = False
'    Next Bucle
'
'    TxtTCP.Visible = False
'    TxtTelefono.Visible = False
'    TxtCadena.Visible = False
'    For Bucle = 0 To 3
'        CmbSerie(Bucle).Visible = False
'    Next Bucle
'    Select Case Index
'        Case 0
'
'            TxtTCP.Visible = True
'
'            Label3(1).Visible = True
'        Case 1
'            If Not (Option1(3) Or Option1(4)) Then
'                For Bucle = 0 To 3
'                    CmbSerie(Bucle).Visible = True
'                Next Bucle
'                For Bucle = 2 To 5
'                    Label3(Bucle).Visible = True
'                Next Bucle
'            End If
'            If Option1(5) Then
'                CmbSerie(1).Visible = False
'                CmbSerie(2).Visible = False
'                CmbSerie(3).Visible = False
'                Label3(3).Visible = False
'                Label3(4).Visible = False
'                Label3(5).Visible = False
'            End If
'        Case 2
'            Label3(6).Caption = CargaCadena(266)
'            Label3(6).Visible = True
'            TxtTelefono.Visible = True
'        Case 3
'            CmbSerie(0).Visible = True
'            CmbSerie(1).Visible = True
'            For Bucle = 2 To 3
'                CmbSerie(Bucle).Visible = False
'            Next Bucle
'            Label3(2).Visible = True
'            Label3(3).Visible = True
'            Label3(6).Visible = True
'            Label3(7).Visible = True
'            For Bucle = 4 To 5
'                Label3(Bucle).Visible = False
'            Next Bucle
'            TxtTelefono.Visible = True
'            Label3(6).Caption = CargaCadena(1066)
'            TxtCadena.Visible = True
'
'    End Select
'End Sub





'Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 48 To 59
'        Case vbKeyDelete
'        Case vbKeyBack
'        Case Else
'            KeyAscii = 0
'            Beep
'    End Select
'End Sub

'Private Sub TxtNumero_LostFocus()
'    If IsNumeric(TxtNumero.Text) Then
'        Select Case Val(TxtNumero.Text)
'            Case 0 To 255
'            Case Else
'                If Val(TxtNumero.Text) <> 999 Then '1.1.4 c2f
'                    MsgBox CargaCadena(1140), vbExclamation
'                    Indice = 0
'                    Pon_Frame
'                    TxtNumero.SetFocus
'                End If
'        End Select
'    End If
'End Sub
Private Sub Crea_la_Tienda()
    Dim DatosConfig As tDatosConfig
    Dim MiTienda As TipoTienda
    Dim Bucle As Integer
    '
    If Val(TxtNumero.Text) < 1 Or Val(TxtNumero.Text) > 250 Then
        'c2f 2.0
        'If Val(TxtNumero.Text) <> 999 Then '1.1.4 c2f
        '    MsgBox CargaCadena(1140), vbCritical
        '    Exit Sub
        'End If
    End If
    If Trim(TxtTienda(0).Text) = "" Then
        MsgBox CargaCadena(1141), vbCritical
        Exit Sub
    End If
    'If Option2(0).Value Then
        If Not IsNumeric(TxtTCP.Text) Or Val(TxtTCP.Text) < 1024 Or Val(TxtTCP.Text) > 65535 Then
            MsgBox CargaCadena(958), vbCritical
            Exit Sub
        End If
    'End If
    'If Option2(2).Value Then
    '    TxtTelefono.Text = Trim(TxtTelefono.Text)
    '    If TxtTelefono.Text = "" Then
    '        MsgBox CargaCadena(1142), vbCritical
    '        Exit Sub
    '    Else
    '        If Right(TxtTelefono.Text, 1) = "\" Then TxtTelefono.Text = Left(TxtTelefono.Text, Len(TxtTelefono.Text) - 1)
    '    End If
    'End If
    'If Option2(3).Value Then
    '    If Trim(TxtTelefono.Text) = "" Then
    '        MsgBox CargaCadena(1139), vbCritical
    '        Exit Sub
    '    End If
    'End If
    
    '
    If Val(TxtNumero.Text) = 999 Then
        
        If Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "000"), vbDirectory) = "" Then
            MkDir App.Path & "\t" & Format(Val(TxtNumero.Text), "000")
        Else
            If Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "000") & "\dbasetouch.mdb") <> "" _
            And Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "000") & "\hydratouch.ini") <> "" _
            Then
                MsgBox CargaCadena(1143), vbCritical
                Exit Sub
            End If
        End If
    
    Else
    
        If Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "00"), vbDirectory) = "" Then
            MkDir App.Path & "\t" & Format(Val(TxtNumero.Text), "00")
        Else
            If Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "00") & "\dbasetouch.mdb") <> "" _
            And Dir(App.Path & "\t" & Format(Val(TxtNumero.Text), "00") & "\hydratouch.ini") <> "" _
            Then
                MsgBox CargaCadena(1143), vbCritical
                Exit Sub
            End If
        End If
    
    End If
    '
    MiTienda.Numero = Val(TxtNumero.Text)
    MiTienda.CodPostal = TxtTienda(4).Text
    MiTienda.Direccion = TxtTienda(1).Text
    MiTienda.Nombre = TxtTienda(0).Text
    MiTienda.Poblacion = TxtTienda(2).Text
    MiTienda.Provincia = TxtTienda(3).Text
    MiTienda.TlfContacto = TxtTienda(5).Text
    'aquí
    'On Error Resume Next
    
    FrmDatosTienda.TxtTienda(7).Text = ""
    FrmDatosTienda.TxtTienda(6).Text = ""
    
    If IsNull(FrmDatosTienda.TxtTienda(6).Text) Then FrmDatosTienda.TxtTienda(6).Text = "00"
    MiTienda.Grp = Val(FrmDatosTienda.TxtTienda(6).Text)
    If IsNull(FrmDatosTienda.TxtTienda(7).Text) Then FrmDatosTienda.TxtTienda(7).Text = ""
    
    If FrmDatosTienda.TxtTienda(7).Text = "" Or IsNull(FrmDatosTienda.TxtTienda(7).Text) Then
        If MiTienda.Numero < 100 Then
            FrmDatosTienda.TxtTienda(7).Text = App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
        Else
            FrmDatosTienda.TxtTienda(7).Text = App.Path & "\t" & Format(MiTienda.Numero, "000") & "\"
        End If
    End If
    
    MiTienda.sPathS = FrmDatosTienda.TxtTienda(7).Text
    FrmDatosTienda.TxtTienda(7).Text = ""
    FrmDatosTienda.TxtTienda(6).Text = ""
    
    'On Error GoTo 0
    DatosConfig.Numero = Val(TxtNumero.Text)
    DatosConfig.TipoC9 = False
    DatosConfig.UsaGamaAlta = False
    DatosConfig.UsaModem = False
    DatosConfig.UsaSC10 = False
    DatosConfig.UsaEco2 = False
    DatosConfig.UsaEco1 = False
    DatosConfig.UsaEcoPrint = False
    DatosConfig.GA_RS232 = "19200,E,8,1,1"
    'If Option1(6).Value = True Then
        DatosConfig.PuertoGA = 3306
    'Else
    '    DatosConfig.PuertoGA = 6000
    'End If
    DatosConfig.SC10_PATH = "z:\sc10"
    DatosConfig.SC10_RS232 = "19200,E,8,1,1"
    DatosConfig.GA_TCPIP = False
    DatosConfig.SC10_NETPATH = False
    DatosConfig.protocolo = 0
    DatosConfig.modoc9 = 0
    DatosConfig.Modem_Cadena = ""
    DatosConfig.Modem_Puerto = 1
    DatosConfig.Modem_Telefono = ""
    DatosConfig.Modem_Velocidad = 1
    'If Option1(0).Value Then
        DatosConfig.UsaGamaAlta = True
        'If Option2(0).Value Then
            DatosConfig.GA_TCPIP = True
            DatosConfig.PuertoGA = Val(TxtTCP.Text)
        'Else
        '    If Option2(1).Value Then
        '        DatosConfig.GA_RS232 = CmbSerie(1).Text & "," & CmbSerie(2).Text & "," & CmbSerie(3).Text & "," & (1 + CmbSerie(0).ListIndex) & ",1"
        '    End If
        'End If
    'End If
    
    'If Option1(6).Value Then
        DatosConfig.lScalePos = True
        MiTienda.lScalePos = True
        DatosConfig.GA_TCPIP = True
        DatosConfig.PuertoGA = Val(TxtTCP.Text)
    'Else
    '    DatosConfig.lScalePos = False
    '    MiTienda.lScalePos = False
    '    'DatosConfig.GA_TCPIP = True
    '    'DatosConfig.PuertoGA = Val(TxtTCP.Text)
    '
    'End If
    ''''Prueba
    On Error Resume Next
    D_TiendaActual(0).lScalePos = MiTienda.lScalePos
    If Err.Number <> 0 Then
        On Error GoTo 0
        ReDim Preserve D_TiendaActual(0)
    End If
    D_TiendaActual(0).lScalePos = MiTienda.lScalePos
    
    ''''''''''
    'If Option1(1).Value Then
    '    DatosConfig.UsaSC10 = True
    '    If Option2(1).Value Then
    '        DatosConfig.SC10_RS232 = CmbSerie(1).Text & "," & CmbSerie(2).Text & "," & CmbSerie(3).Text & ",1," & (1 + CmbSerie(0).ListIndex)
    '    Else
    '        If Option2(2).Value Then
    '            DatosConfig.SC10_NETPATH = True
    '            DatosConfig.SC10_PATH = TxtTelefono.Text
    '        End If
    '    End If
    'End If
    
    'If Option1(2).Value Then
    '    DatosConfig.UsaSC10 = True
    '    DatosConfig.TipoC9 = True
    '    DatosConfig.modoc9 = CmbC9(1).ListIndex
    '    DatosConfig.protocolo = CmbC9(0).ListIndex
    '
    '    If Option2(1).Value Then
    '        DatosConfig.SC10_RS232 = CmbSerie(1).Text & "," & CmbSerie(2).Text & "," & CmbSerie(3).Text & ",1," & (1 + CmbSerie(0).ListIndex)
    '    End If
    'End If
    
    'If Option1(3).Value Then
    '    DatosConfig.UsaEco1 = True
    '    DatosConfig.PuertoECO = CmbSerie(0).ListIndex + 1
    'End If
    
    'If Option1(4).Value Then
    '    DatosConfig.UsaEco2 = True
    '    DatosConfig.PuertoECO = CmbSerie(0).ListIndex + 1
    'End If
    
    'If Option1(5).Value Then
    '    DatosConfig.UsaEcoPrint = True
    '    DatosConfig.PuertoECO = CmbSerie(0).ListIndex + 1
    'End If
    
    'If Option2(3).Value Then
    '    DatosConfig.UsaModem = True
    '    DatosConfig.Modem_Cadena = Trim(TxtCadena.Text)
    '    DatosConfig.Modem_Puerto = CmbSerie(0).ListIndex + 1
    '    DatosConfig.Modem_Telefono = Trim(TxtTelefono.Text)
    '    DatosConfig.Modem_Velocidad = CmbSerie(1).ListIndex
    'End If
    '
    Escribe_Tienda MiTienda
    '
    EscribirParametrosInicio MiTienda.sPathS & "hydratouch.ini", DatosConfig
    '
    MsgBox CargaCadena(1144), vbInformation
    ActivaSingle MiTienda.Numero, MiTienda.Nombre, 61, MiTienda.sPathS, MiTienda.lScalePos
    
    Unload Me
End Sub



Private Sub TxtTCP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub TxtTCP_LostFocus()
    If IsNumeric(TxtTCP.Text) Then
        If Val(TxtTCP.Text) < 1024 Or Val(TxtTCP.Text) > 65535 Then
            MsgBox CargaCadena(958), vbExclamation
            Indice = 3
            Pon_Frame
            TxtTCP.SetFocus
        End If
    End If
End Sub




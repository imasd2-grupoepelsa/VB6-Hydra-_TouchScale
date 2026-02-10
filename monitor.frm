VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonitor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío de datos a Gama Alta por IP"
   ClientHeight    =   5460
   ClientLeft      =   3060
   ClientTop       =   2160
   ClientWidth     =   8235
   ControlBox      =   0   'False
   Icon            =   "monitor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1212
   End
   Begin VB.CommandButton cmdmonitorstop 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Detener"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1212
   End
   Begin VB.Timer TmrEnableCancel 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   840
   End
   Begin VB.Timer TmrConecta 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   5520
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   5520
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "monitor.frx":000C
            Key             =   """uno"""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "monitor.frx":0EE6
            Key             =   """dos"""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "monitor.frx":1DC0
            Key             =   """tres"""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar barra1 
      Height          =   468
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   3468
      _ExtentX        =   6112
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   1
      Max             =   140
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3540
      ItemData        =   "monitor.frx":2C9A
      Left            =   0
      List            =   "monitor.frx":2C9C
      TabIndex        =   0
      Top             =   0
      Width           =   8292
   End
   Begin VB.Label LblError 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "ERROR DE COMUNICACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   492
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   492
   End
   Begin VB.Shape LuzRecibe 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   132
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape LuzEnvio 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   9  'Not Mask Pen
      Height          =   132
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   372
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1092
   End
   Begin VB.Image Image1 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   8172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Menu mnuCopiar2 
      Caption         =   "Copiar"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
      End
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CIconos As Integer

Private Sub cmdmonitorstop_Click()
    Dim Arch As Integer
    Dim nC As Long

    If Not HayComandos Then
        frmMonitor.LuzEnvio.BackColor = vbBlack
        frmMonitor.LuzRecibe.BackColor = vbBlack
        frmMonitor.Label1.Caption = CargaCadena(273)    '"Deteniendo..."
        If DebugActivo Then CadenadeLog "Proceso de Comunicaciones Cancelado por Usuario"
        canCelar = 1
        CancelHumano = True
        HabilitarSalir
        frmMonitor.cmdmonitorstop.Enabled = False
    End If
End Sub


Private Sub Command1_Click()
    frmMonitor.List1.Clear
    canCelar = 0
    CadenadeLog "Pulsado Salir en FRMMONITOR...OrdenMulti:" & OrdenMulti & " FrmEpelsa.visible:" & CStr(frmEpelsa.Visible)
    If frmEpelsa.Visible = True And (OrdenMulti = "00") Then
        frmEpelsa.Enabled = True
        frmEpelsa.LblGA.Caption = ""
        frmEpelsa.LblSC10.Caption = ""
        frmEpelsa.SetFocus
    End If
    If HayMulti Then Multi_Salir = True

    frmMonitor.Hide
End Sub


Private Sub Form_Load()
    Me.Icon = Form2.Icon
    LblError.Caption = CargaCadena(544)
    Image1.Picture = Form2.Imagen.Picture
    left = (Screen.Width - Width) / 2
    top = (Screen.Height - Height) / 2
    Command1.Caption = CargaCadena(7)
    cmdmonitorstop.Caption = CargaCadena(908)
    frmMonitor.Caption = CargaCadena(4)
    LuzEnvio.BackColor = vbRed
    LuzRecibe.BackColor = vbBlue
    'DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
End Sub
Public Sub MostrarDato(LaCadena As String)
    If List1.ListCount > 10000 Then
        List1.Clear
    End If
    frmMonitor.List1.AddItem LaCadena
    If DebugActivo Then CadenadeLog LaCadena
    'If List1.ListCount > 13 Then
    'If List1.ListCount Mod 20 = 0 Then
    On Error Resume Next
    List1.TopIndex = List1.ListCount - 1
    On Error GoTo 0
    '   End If
    'End If
End Sub
Public Sub MostrarInfo(LaCadena As String)
    Dim izq As Integer
    Dim der As Integer
    Dim esp As Integer
    List1.AddItem "###########################################################################"
    List1.AddItem LaCadena
    List1.AddItem " "
    If DebugActivo Then CadenadeLog LaCadena
    'If List1.ListCount > 13 Then
    '    If List1.ListCount Mod 20 = 0 Then
    On Error Resume Next
    List1.TopIndex = List1.ListCount - 13
    On Error GoTo 0
    'End If
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
    canCelar = 1
End Sub

Private Sub TmrEnableCancel_Timer()
    cmdmonitorstop.Enabled = True
    frmMonitor.TmrEnableCancel.Enabled = False
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuCopiar2
    End If

End Sub

Private Sub mnuCopiar_Click()
    Dim Buf As String
    Dim bucle As Long
    For bucle = 0 To List1.ListCount - 1
        Buf = Buf & List1.List(bucle) & vbCrLf
    Next bucle
    Clipboard.SetText Buf
End Sub

Private Sub TmrConecta_Timer()
    Select Case CIconos
    Case 0
        Image2.Picture = ImageList1.ListImages(1).Picture
    Case 1
        Image2.Picture = ImageList1.ListImages(2).Picture
    Case 2
        Image2.Picture = ImageList1.ListImages(3).Picture
    End Select
    CIconos = CIconos + 1
    If CIconos = 3 Then CIconos = 0
End Sub

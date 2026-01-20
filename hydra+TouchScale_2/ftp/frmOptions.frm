VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "    Options"
   ClientHeight    =   4665
   ClientLeft      =   5685
   ClientTop       =   2160
   ClientWidth     =   3915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   480
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6060
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   0
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   9600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Frame fraModes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "Modes"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame fraConMod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Caption         =   "Connection Mode"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   2775
         Begin VB.OptionButton optConmode 
            BackColor       =   &H80000013&
            Caption         =   "Passive Connection"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optConmode 
            BackColor       =   &H80000013&
            Caption         =   "Active Connection"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame fraFileTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Caption         =   "File Transfer Mode"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   2775
         Begin VB.OptionButton optTranMode 
            BackColor       =   &H80000013&
            Caption         =   "ASCII File Transfer"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optTranMode 
            BackColor       =   &H80000013&
            Caption         =   "Binary File Transfer"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Line Line5 
         X1              =   3480
         X2              =   3480
         Y1              =   -120
         Y2              =   360
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   0
         X2              =   0
         Y1              =   240
         Y2              =   -120
      End
   End
   Begin VB.Frame fraConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "Connection/Proxy Server"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   3495
      Begin VB.CheckBox chkProxy 
         BackColor       =   &H80000013&
         Caption         =   "Use a proxy server for you're LAN (These settings will not apply to Dial-up or VPN Connections."
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtProxAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   0
         Y1              =   480
         Y2              =   -120
      End
      Begin VB.Line Line3 
         Visible         =   0   'False
         X1              =   3480
         X2              =   3480
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   -120
      End
      Begin VB.Label lblPort 
         BackColor       =   &H80000013&
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000013&
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3855
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3855
      Width           =   1095
   End
   Begin VB.Label cmdConnection 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connection"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label cmdModes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modes"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CreateEllipticRgn() As Long

'Private m_cN As cNeoCaption
'Private m_cSkin As cSkinConfiguration
Private m_iIndex As Long

Private Sub chkProxy_Click()
    If chkProxy = Checked Then
        txtProxAdd.Enabled = True
        txtProxAdd.BackColor = vbWhite
        txtPort.Enabled = True
        txtPort.BackColor = vbWhite
    Else
        txtProxAdd.Enabled = False
        txtProxAdd.BackColor = &H8000000B
        txtPort.Enabled = False
        txtPort.BackColor = &H8000000B
        
        txtProxAdd.Text = vbNullString
        txtPort.Text = vbNullString
    End If
End Sub

Private Sub cmdConnection_Click()
    fraConnection.Visible = True
    fraModes.Visible = False
    cmdConnection.BackColor = &H80000013
    cmdModes.BackColor = vbWhite
    Line1.Visible = False
    Line4.Visible = True
End Sub

Private Sub cmdModes_Click()
    fraModes.Visible = True
    fraConnection.Visible = False
    cmdConnection.BackColor = vbWhite
    cmdModes.BackColor = &H80000013
    Line1.Visible = False
    Line5.Visible = True
End Sub

Private Sub Form_Activate()
    'SetWindowRgn hwnd, CreateEllipticRgn(0, 0, 300, 200), True
    'center the form
   ' Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdApply_Click()
    cmdOK.Enabled = True
    Select Case True
        Case optTranMode(0)
            optTranMode(0) = Checked
        Case optTranMode(1)
            optTranMode(1) = Checked
    End Select
    SaveSetting App.EXEName, "Settings", "UserName", txtUser.Text
    SaveSetting App.EXEName, "Settings", "Password", txtPassword.Text
    SaveSetting App.EXEName, "Settings", "ProxyAddress", txtProxAdd.Text
    SaveSetting App.EXEName, "Settings", "Port", txtPort.Text
    SaveSetting App.EXEName, "Settings", "Binary", optTranMode(0).Value
    SaveSetting App.EXEName, "Settings", "ASCII", optTranMode(1).Value
    SaveSetting App.EXEName, "Settings", "Active", optConmode(0).Value
    SaveSetting App.EXEName, "Settings", "Passive", optConmode(1).Value
    SaveSetting App.EXEName, "Settings", "ProxyCheck", chkProxy.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub

Private Sub cmdOK_Click()
    Unload Me
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
Private Sub setColors()
  
      Me.BackColor = m_cN.MenuBackgroundColor
      frmOptions.BackColor = m_cN.MenuBackgroundColor
      frmOptions.ForeColor = m_cN.ActiveMenuColorOver
  
End Sub
Public Sub SelectNewSkin(ByVal iIndex As Long)
Dim sFont As New StdFont
   
   Set m_cSkin = New cSkinConfiguration
   With m_cSkin
         .Name = "Neo vbAccelerator"
         .ActiveCaptionColor = &HFFFFFF
         .InActiveCaptionColor = RGB(47, 122, 151)
         .ActiveMenuColor = &H0&
         .ActiveMenuColorOver = &HFFFFFF
         .InActiveMenuColor = RGB(47, 122, 151)
         .MenuBackgroundColor = &H80000013  'RGB(123, 189, 214)
         Set .CaptionFont = Me.Font
         Set .MenuFont = Me.Font
         Set .Caption = picResource(0).Picture
         Set .Borders = picResource(1).Picture
         .ButtonWidth = 19
         .ButtonHeight = 20
         .ActiveLeftEnd = 90
         .ActiveRightStart = 140
         .ActiveRightEnd = 240
         .InactiveOffset = 400
         
   End With
   
End Sub

Private Sub Form_Load()
Dim Index As Integer

optTranMode(0) = Checked
optConmode(0) = Checked
Line1.Visible = True

'Set m_cN = New cNeoCaption
    
   ' SelectNewSkin Index
      '  m_cN.Attach2 Me, m_cSkin
      '  setColors
      '  If Index = 5 Then
      '      frmMain.BackColor = RGB(140, 139, 139)
      '  End If

txtUser.Text = GetSetting(App.EXEName, "Settings", "UserName")
txtPassword.Text = GetSetting(App.EXEName, "Settings", "Password")
txtProxAdd.Text = GetSetting(App.EXEName, "Settings", "ProxyAddress")
txtPort.Text = GetSetting(App.EXEName, "Settings", "Port")
optTranMode(0).Value = GetSetting(App.EXEName, "Settings", "Binary", optTranMode(0).Value)
optTranMode(1).Value = GetSetting(App.EXEName, "Settings", "ASCII", optTranMode(1).Value)
optConmode(0).Value = GetSetting(App.EXEName, "Settings", "Active", optConmode(0).Value)
optConmode(1).Value = GetSetting(App.EXEName, "Settings", "Passive", optConmode(1).Value)
chkProxy.Value = GetSetting(App.EXEName, "Settings", "ProxyCheck", chkProxy.Value)


End Sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SYSTRAY.OCX"
Begin VB.Form FrmServidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de NTQ - Server"
   ClientHeight    =   4068
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5700
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4068
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   3360
      Top             =   3000
      _ExtentX        =   910
      _ExtentY        =   910
      InTray          =   0   'False
      TrayIcon        =   "FrmServidor.frx":0000
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   4080
      TabIndex        =   11
      Top             =   3600
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1332
   End
   Begin VB.Frame Frame2 
      Caption         =   "Puerto UDP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5412
      Begin VB.TextBox UDPPuerto 
         Height          =   372
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   852
      End
      Begin MSMask.MaskEdBox UDPHost 
         Height          =   372
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1692
         _ExtentX        =   2985
         _ExtentY        =   656
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Puerto : "
         Height          =   252
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Host : "
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puerto TCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5412
      Begin VB.TextBox TCPPuerto 
         Height          =   372
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   852
      End
      Begin MSMask.MaskEdBox TCPHost 
         Height          =   372
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   1692
         _ExtentX        =   2985
         _ExtentY        =   656
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "999.999.999.999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Puerto : "
         Height          =   252
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Host : "
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   972
      End
   End
   Begin MSWinsockLib.Winsock SockTCP 
      Left            =   4680
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockUDP 
      Left            =   5280
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "FrmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
    Dim Fichero As String
    If Index = 0 Then
        If Trim(TCPPuerto.Text) = "" Or Trim(UDPPuerto.Text) = "" Then
            MsgBox "Complete los campos de Puertos de Comunicaciones"
            Exit Sub
        End If
        Fichero = FreeFile()
        Open MiruTa & "\ntqserver.ini" For Output As Fichero
        Print #Fichero, TCPHost.Text
        Print #Fichero, TCPPuerto.Text
        Print #Fichero, UDPHost.Text
        Print #Fichero, UDPPuerto.Text
        Close #Fichero
        CI_TCP_host = TCPHost.Text
        CI_UDP_host = UDPHost.Text
        CI_TCP_port = TCPPuerto.Text
        CI_UDP_port = UDPPuerto.Text
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    If frmEpelsa.Visible = False Then
        frmEpelsa.Visible = True
    Else
        frmEpelsa.Visible = False
    End If
End Sub



Private Sub Form_Load()
    TCPPuerto.Text = CI_TCP_port
    UDPPuerto.Text = CI_UDP_port
    TCPHost.Text = CI_TCP_host
    UDPHost.Text = CI_UDP_host
    
End Sub

Private Sub TCPPuerto_KeyPress(KeyAscii As Integer)
    TCPPuerto.Locked = Checktexto(KeyAscii, 5, TCPPuerto.Text, True, False)
End Sub



Private Sub UDPPuerto_KeyPress(KeyAscii As Integer)
    UDPPuerto.Locked = Checktexto(KeyAscii, 5, UDPPuerto.Text, True, False)
End Sub

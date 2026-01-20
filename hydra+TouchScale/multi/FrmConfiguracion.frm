VERSION 5.00
Begin VB.Form FrmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración General"
   ClientHeight    =   4140
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtBackup 
      Height          =   288
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   3852
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Configurar formato de los datos de Exportación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   4572
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   120
      Picture         =   "FrmConfiguracion.frx":0000
      Top             =   1440
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      X1              =   120
      X2              =   7320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta donde depositar las copias de Seguridad :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   3252
   End
   Begin VB.Label Command2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   6000
      Picture         =   "FrmConfiguracion.frx":08CA
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   3372
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   3372
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   3372
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3372
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "FrmConfiguracion.frx":1194
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "FrmConfiguracion.frx":1A5E
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "FrmConfiguracion.frx":2328
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "FrmConfiguracion.frx":2BF2
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   9
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   8
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   7
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   6
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   5
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   4
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   3
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   2
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BufID As Integer
Private Sub Command1_Click(Index As Integer)
    Dim BufSec As Boolean
    Me.Enabled = False
    Select Case Index
        Case 0
            If HaySeguridad Then
                If ConsultarPermiso(15) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Me.Enabled = True
                    Exit Sub
                End If
            End If
            ActivaSingle 0, "", 62, "", False 'ojo D_TiendaActual(0).lScalePos
        Case 1
            If HaySeguridad Then
                If ConsultarPermiso(16) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Me.Enabled = True
                    Exit Sub
                End If
            End If
            ActivaSingle 0, "", 63, "", False 'ojo D_TiendaActual(0).lScalePos
        Case 2
            If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
                MsgBox CargaCadena(703), vbCritical
                Me.Enabled = True
                Exit Sub
            End If
            ActivaSingle 0, "", 64, "", False 'ojo D_TiendaActual(0).lScalePos
        Case 3
            If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
                MsgBox CargaCadena(703), vbCritical
                Me.Enabled = True
                Exit Sub
            End If
            BufSec = HaySeguridad
            ActivaSingle 0, "", 65, "", False 'ojo D_TiendaActual(0).lScalePos
            LeeMultiConfig
            If BufSec <> HaySeguridad Then ActualizaSeguridad
        Case 4
            If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
                MsgBox CargaCadena(703), vbCritical
                Me.Enabled = True
                Exit Sub
            End If
            FrmExportFormat.Show vbModal, Me
    End Select
    Me.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
    'ID = BufID
    Backup_Path = TxtBackup.Text
    If Len(Backup_Path) >= 1 Then
        If Right(Backup_Path, 1) = "\" Then Backup_Path = Left(Backup_Path, Len(Backup_Path) - 1)
    End If
    CargaIdiomas
    EscribeConfig
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Bucle As Integer
    TxtBackup.Text = Backup_Path
    CambiarIdioma
    BufID = ID
    Me.Icon = FrmPrincipal.Icon
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    If Dir(App.Path & "\langtouch", vbDirectory) <> "" Then
        For Bucle = 0 To 9
            If Dir(App.Path & "\langtouch\" & (Bucle + 1) & ".ico") <> "" Then
                On Error Resume Next
                Image1(Bucle).Picture = LoadPicture(App.Path & "\langtouch\" & (Bucle + 1) & ".ico")
                If Err.Number = 0 Then Image1(Bucle).Visible = True
                On Error GoTo 0
            Else
                Image1(Bucle).Visible = False
            End If
        Next Bucle
    End If
End Sub
Private Sub CambiarIdioma()
    Command1(0).Caption = CargaCadena(681)
    Command1(1).Caption = CargaCadena(680)
    Command1(2).Caption = CargaCadena(682)
    Command1(3).Caption = CargaCadena(684)
    Command2(0).Caption = CargaCadena(7)
    Label1.Caption = CargaCadena(1149)
    Command1(4).Caption = CargaCadena(1148)
    Me.Caption = CargaCadena(1081)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DescargaForm = True
End Sub

Private Sub Image1_Click(Index As Integer)
    ID = Index
    CargaIdiomas
    CambiarIdioma
End Sub

Private Sub Image2_Click(Index As Integer)
    Command1_Click Index
End Sub

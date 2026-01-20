VERSION 5.00
Begin VB.Form FrmExportFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formato de los datos de Exportación"
   ClientHeight    =   3705
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2400
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2160
      Width           =   372
   End
   Begin VB.ComboBox CmbAbsoluto 
      Height          =   288
      ItemData        =   "FrmExportFormat.frx":0000
      Left            =   2400
      List            =   "FrmExportFormat.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   1452
   End
   Begin VB.ComboBox CmbDiario 
      Height          =   288
      ItemData        =   "FrmExportFormat.frx":0004
      Left            =   2400
      List            =   "FrmExportFormat.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   1452
   End
   Begin VB.ComboBox CmbHora 
      Height          =   288
      ItemData        =   "FrmExportFormat.frx":0008
      Left            =   2400
      List            =   "FrmExportFormat.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   1452
   End
   Begin VB.ComboBox CmbFecha 
      Height          =   288
      ItemData        =   "FrmExportFormat.frx":0023
      Left            =   2400
      List            =   "FrmExportFormat.frx":003F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   6
      Left            =   4680
      TabIndex        =   13
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4200
      Picture         =   "FrmExportFormat.frx":0093
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Signo Negativo"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2052
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Signo Positivo"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2052
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valores con o sin signo"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos a Exportar"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Formato de Hora"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Formato de Fecha"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Formato de los datos de los Ficheros de Exportación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5652
   End
End
Attribute VB_Name = "FrmExportFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CambiarIdioma()
    Me.Caption = CargaCadena(1150)
    Label1.Caption = CargaCadena(1150)
    Label2(0).Caption = CargaCadena(946)
    Label2(1).Caption = CargaCadena(947)
    Label2(2).Caption = CargaCadena(1151)
    Label2(3).Caption = CargaCadena(1152)
    Label2(4).Caption = CargaCadena(948)
    Label2(5).Caption = CargaCadena(949)
    Label2(6).Caption = CargaCadena(7)
    CmbDiario.AddItem CargaCadena(1050)
    CmbDiario.AddItem CargaCadena(1049)
    
    CmbAbsoluto.AddItem CargaCadena(1153)
    CmbAbsoluto.AddItem CargaCadena(1154)
    
    
End Sub
Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = FrmPrincipal.Icon
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    CmbFecha.ListIndex = Exp_Fecha
    CmbHora.ListIndex = Exp_Hora
    If Exp_Diario Then
        CmbDiario.ListIndex = 1
    Else
        CmbDiario.ListIndex = 0
    End If
    If Exp_Absoluto Then
        CmbAbsoluto.ListIndex = 1
    Else
        CmbAbsoluto.ListIndex = 0
    End If
    Text1(0).Text = Exp_Positivo
    Text1(1).Text = Exp_Negativo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Guarda_Datos
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub
Private Sub Guarda_Datos()
    If Text1(0).Text = "" Then Text1(0).Text = "0"
    If Text1(1).Text = "" Then Text1(1).Text = "1"
    Exp_Positivo = Text1(0).Text
    Exp_Negativo = Text1(1).Text
    If CmbDiario.ListIndex = 1 Then
        Exp_Diario = True
    Else
        Exp_Diario = False
    End If
    If CmbAbsoluto.ListIndex = 1 Then
        Exp_Absoluto = True
    Else
        Exp_Absoluto = False
    End If
    Exp_Fecha = CmbFecha.ListIndex
    Exp_Hora = CmbHora.ListIndex
    EscribeConfig
    
End Sub


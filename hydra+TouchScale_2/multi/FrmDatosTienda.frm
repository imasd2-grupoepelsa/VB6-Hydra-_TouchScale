VERSION 5.00
Begin VB.Form FrmDatosTienda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Identificativos"
   ClientHeight    =   2655
   ClientLeft      =   2250
   ClientTop       =   360
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      Height          =   288
      Index           =   7
      Left            =   2160
      TabIndex        =   17
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      Height          =   288
      Index           =   6
      Left            =   3480
      TabIndex        =   14
      Text            =   "01"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   4092
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   1200
      Width           =   4092
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   4092
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   4092
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   4092
   End
   Begin VB.TextBox TxtTienda 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   4092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   16
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   7080
      Picture         =   "FrmDatosTienda.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "FrmDatosTienda.frx":08CA
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono de Contacto :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Postal :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Población"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   3372
   End
End
Attribute VB_Name = "FrmDatosTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CambiarIdioma()
    Label1(0).Caption = CargaCadena(552)
    Label1(1).Caption = CargaCadena(1127)
    Label1(2).Caption = CargaCadena(1128)
    Label1(3).Caption = CargaCadena(1129)
    Label1(4).Caption = CargaCadena(1130)
    Label1(5).Caption = CargaCadena(1131)
    Label2(0).Caption = CargaCadena(287)
    Label2(1).Caption = CargaCadena(288)
    Me.Caption = CargaCadena(1090)
End Sub
Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = FrmPrincipal.Icon
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    Pon_Los_Datos
End Sub
Private Sub Pon_Los_Datos()
    Dim MiTienda() As TipoTienda
    Dim Arch As Integer
    MiTienda = Consulta_Tiendas(FrmPrincipal.TiendaActual, 1)
    TxtTienda(0).Text = MiTienda(0).Nombre
    TxtTienda(1).Text = MiTienda(0).Direccion
    TxtTienda(2).Text = MiTienda(0).Poblacion
    TxtTienda(3).Text = MiTienda(0).CodPostal
    TxtTienda(4).Text = MiTienda(0).Provincia
    TxtTienda(5).Text = MiTienda(0).TlfContacto
    TxtTienda(6).Text = MiTienda(0).Grp
    TxtTienda(7).Text = MiTienda(0).sPathS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DescargaForm = True
End Sub

Private Sub Image1_Click(Index As Integer)
    Dim MiTienda As TipoTienda
    Dim miTmp As Boolean
    
    Select Case Index
        Case 0
            MiTienda.Numero = FrmPrincipal.TiendaActual
            MiTienda.Nombre = TxtTienda(0).Text
            MiTienda.Direccion = TxtTienda(1).Text
            MiTienda.Poblacion = TxtTienda(2).Text
            MiTienda.CodPostal = TxtTienda(3).Text
            MiTienda.Provincia = TxtTienda(4).Text
            MiTienda.TlfContacto = TxtTienda(5).Text
            MiTienda.Grp = TxtTienda(6).Text
            MiTienda.sPathS = TxtTienda(7).Text
            Escribe_Tienda MiTienda
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub


Private Sub TxtTienda_Change(Index As Integer)
   If Index = 6 Then
          If Val(TxtTienda(6).Text) > 99 Then
             TxtTienda(6).Text = ""
          End If
   End If

End Sub

Private Sub TxtTienda_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 6 Then
    If KeyAscii >= 48 And KeyAscii <= 57 Then
          If Val(TxtTienda(6).Text) > 99 Then
             TxtTienda(6).Text = ""
          End If
    Else
          KeyAscii = 0
    End If
   End If
   
End Sub

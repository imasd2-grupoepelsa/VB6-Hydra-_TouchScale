VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B58D002A-7750-11D6-A84F-00A0CCE04944}#1.0#0"; "SideBar.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00800000&
   Caption         =   "Saturno Multitienda"
   ClientHeight    =   8460
   ClientLeft      =   1875
   ClientTop       =   750
   ClientWidth     =   9540
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   9540
   Begin VB.CheckBox ChkDirect 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   7920
      Width           =   255
   End
   Begin VB.CheckBox CheckTransfer 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   6840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      Height          =   8055
      Left            =   2520
      ScaleHeight     =   7995
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.Image Image3 
         Height          =   495
         Left            =   0
         Picture         =   "FrmPrincipal.frx":1E72
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   612
         Left            =   1320
         Picture         =   "FrmPrincipal.frx":36A9
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   612
      End
   End
   Begin AspBigBar.SideBar Barra 
      Height          =   8055
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   14208
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      SmallImageList  =   "ImageList1"
      LargeImageList  =   "ImageList1"
      DisplayCaption  =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5520
      Top             =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar en la Tienda"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   3612
   End
   Begin RichTextLib.RichTextBox Info 
      Height          =   2895
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5106
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmPrincipal.frx":3F73
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3FF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":48CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":51A9
            Key             =   "hydra"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":5A83
            Key             =   "total"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":635D
            Key             =   "comu"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":6C37
            Key             =   "herram"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":7511
            Key             =   "listado"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Arbol 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   14208
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label LblDirect 
      Caption         =   "(Direct. Transfer)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Bal.dat/Shops"
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
      Height          =   495
      Index           =   5
      Left            =   6480
      TabIndex        =   13
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Image Command3 
      Height          =   495
      Index           =   5
      Left            =   5880
      Picture         =   "FrmPrincipal.frx":83EB
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   510
   End
   Begin VB.Label LabelTransfer 
      Caption         =   "Shop (Transfer/GRP.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   12
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Grp./Files"
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
      Height          =   495
      Index           =   4
      Left            =   6480
      TabIndex        =   10
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Image Command3 
      Height          =   495
      Index           =   4
      Left            =   5880
      Picture         =   "FrmPrincipal.frx":8CB5
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group."
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
      Index           =   3
      Left            =   6480
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Image Command3 
      Height          =   495
      Index           =   3
      Left            =   5880
      Picture         =   "FrmPrincipal.frx":957F
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   5880
      Picture         =   "FrmPrincipal.frx":9E49
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Copias de Seguridad"
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
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Image Command3 
      Height          =   495
      Index           =   2
      Left            =   5880
      Picture         =   "FrmPrincipal.frx":A713
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   1
      Left            =   7080
      TabIndex        =   5
      Top             =   6840
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   0
      Left            =   7080
      TabIndex        =   4
      Top             =   7320
      Width           =   1812
   End
   Begin VB.Image Command3 
      Height          =   372
      Index           =   1
      Left            =   9000
      Picture         =   "FrmPrincipal.frx":AFDD
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   372
   End
   Begin VB.Image Command3 
      Height          =   384
      Index           =   0
      Left            =   9000
      Picture         =   "FrmPrincipal.frx":B8A7
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   384
   End
   Begin VB.Menu mnuDatos 
      Caption         =   "Datos de la Tienda..."
      Visible         =   0   'False
      Begin VB.Menu mnuIntroducir 
         Caption         =   "Introducción de Datos"
         Begin VB.Menu mnuT 
            Caption         =   "Mostradores y Equipos"
            Index           =   0
         End
         Begin VB.Menu mnuT 
            Caption         =   "Códigos de Barras"
            Index           =   1
         End
         Begin VB.Menu mnuT 
            Caption         =   "Cabeceras y Leyendas"
            Index           =   2
         End
         Begin VB.Menu mnuT 
            Caption         =   "Publicidad y Ofertas"
            Index           =   3
         End
         Begin VB.Menu mnuT 
            Caption         =   "Textos Libres"
            Index           =   4
         End
         Begin VB.Menu mnuT 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuT 
            Caption         =   "Artículos"
            Index           =   6
         End
         Begin VB.Menu mnuT 
            Caption         =   "Cambios de Precio"
            Index           =   7
         End
         Begin VB.Menu mnuT 
            Caption         =   "Teclas"
            Index           =   8
         End
         Begin VB.Menu mnuT 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuT 
            Caption         =   "Tipos de IVA"
            Index           =   10
         End
         Begin VB.Menu mnuT 
            Caption         =   "Clientes"
            Index           =   11
         End
         Begin VB.Menu mnuT 
            Caption         =   "Vendedores"
            Index           =   12
         End
         Begin VB.Menu mnuT 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnuT 
            Caption         =   "Grupos de Conservación Euroscale"
            Index           =   14
         End
         Begin VB.Menu mnuT 
            Caption         =   "Grupos de Conservación SC10"
            Index           =   15
         End
      End
      Begin VB.Menu mnuNada 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuComu 
         Caption         =   "Comunicaciones"
         Begin VB.Menu mnuC 
            Caption         =   "Envío de Modificaciones"
            Index           =   0
         End
         Begin VB.Menu mnuC 
            Caption         =   "Datos Por Separado"
            Index           =   1
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Mostradores"
               Index           =   2
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Equipos"
               Index           =   3
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Códigos de Barras"
               Index           =   4
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Cabeceras y Leyendas"
               Index           =   5
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Publicidad"
               Index           =   6
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Teclas"
               Index           =   7
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Textos Libre"
               Index           =   8
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Artículos"
               Index           =   9
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Tipos IVA"
               Index           =   10
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Clientes"
               Index           =   11
            End
            Begin VB.Menu mnuC1 
               Caption         =   "Envío Vendedores"
               Index           =   12
            End
         End
         Begin VB.Menu mnuC 
            Caption         =   "-"
            Index           =   14
         End
         Begin VB.Menu mnuC 
            Caption         =   "Grandes Totales"
            Index           =   15
         End
         Begin VB.Menu mnuC 
            Caption         =   "Recibir Tickets"
            Index           =   16
         End
      End
   End
   Begin VB.Menu mnucrear 
      Caption         =   "Creación"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearTienda 
         Caption         =   "Crear Tienda"
      End
      Begin VB.Menu mnuListadoTIendas 
         Caption         =   "Listado de Tiendas"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TiendaActual As Integer
Public NoSalir As Boolean
'ScalePos 1.0.8, se hace pública
'Private D_TiendaActual() As TipoTienda
''''''''''''''
Public FolderQL As Boolean
Public FolderQL2 As Boolean
Private UsaGamaAlta As Boolean
Private UsaScalePos As Boolean
Private UsaSC10 As Boolean
Private UsaEco1 As Boolean
Private FamiliasFijas As Boolean
Private UsaEcoPrint As Boolean
Private TipoEcoPrint As Integer
Private TipoC9 As Boolean
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long
Private Sub Me_Init()
    Dim Bucle As Long
    
    Me_Width = Me.Width
    Me_Height = Me.Height
    ReDim ct_top(Me.Controls.Count)
    ReDim CT_Left(Me.Controls.Count)
    ReDim CT_Width(Me.Controls.Count)
    ReDim CT_Height(Me.Controls.Count)
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        ct_top(Bucle) = Me.Controls(Bucle).Top
        CT_Left(Bucle) = Me.Controls(Bucle).Left
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        On Error GoTo 0
    Next Bucle
End Sub
Private Sub Me_Resize()
    Dim RelW As Double
    Dim Bucle As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        RelW = Me.Width / Me_Width
        Me.Controls(Bucle).Width = RelW * CT_Width(Bucle)
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        Me.Controls(Bucle).Left = RelW * CT_Left(Bucle)
        CT_Left(Bucle) = Me.Controls(Bucle).Left
        
        '
        RelW = Me.Height / Me_Height
        Me.Controls(Bucle).Height = RelW * CT_Height(Bucle)
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        Me.Controls(Bucle).Top = RelW * ct_top(Bucle)
        ct_top(Bucle) = Me.Controls(Bucle).Top
        On Error GoTo 0
        '
    Next Bucle
    DrawGradient Picture1, 220, 220, 220, 0, 50, 0, Picture1.ScaleHeight, 1, True, 1, 1, 1
    If TiendaActual <> 0 Then
        
        Picture1.Height = 1
    End If
    Me_Width = Me.Width
    Me_Height = Me.Height
End Sub
Private Sub Arbol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        If Arbol.SelectedItem.Index > 1 Then
            
            PopupMenu mnuDatos
        Else
            
            PopupMenu mnucrear
        End If
    End If
End Sub
Private Sub MueveBarra(Expandir As Boolean)
    Dim Bucle As Integer
    
    Me.Enabled = False
    If Not Expandir Then
     
        For Bucle = 1 To (Me.Height - 850) Step 400
            Picture1.Height = Bucle
            Do_Events
        Next Bucle
        DrawGradient Picture1, 220, 220, 220, 0, 50, 0, Picture1.ScaleHeight, 1, True, 1, 1, 1
        Picture1.Height = (Me.Height - 850) '550
        Image1.Visible = True
        '
        Label2(0).Visible = True
        Label1(2).Visible = True
        Image2(0).Visible = True
        Command3(2).Visible = True
        
        Command3(3).Visible = True
        Label1(3).Visible = True
        Command3(4).Visible = True
        Label1(4).Visible = True

        '
    Else
        'If Picture.Height > 6000 Then

        Image1.Visible = False
        If Not (Picture1.Height > 300) Then
            For Bucle = 1 To (Me.Height - 850) Step 400
                Picture1.Height = Bucle
                Do_Events
            Next Bucle
        End If
        DrawGradient Picture1, 220, 220, 220, 0, 50, 0, Picture1.ScaleHeight, 1, True, 1, 1, 1
        For Bucle = (Me.Height - 850) To 1 Step -400
            Picture1.Height = Bucle
            Do_Events
        Next Bucle
        
        Picture1.Height = 1
        '
        Label2(0).Visible = False
        Label1(2).Visible = False
        Image2(0).Visible = False
        Command3(2).Visible = False
        Command3(3).Visible = False
        Label1(3).Visible = False
        Command3(4).Visible = False
        Label1(4).Visible = False
        
        '
    End If
    Me.Enabled = True
End Sub


Private Sub Arbol_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Resp As Boolean
    Dim Mover As Boolean
    
    If IsNumeric(Node.Text) Then
        Barra.Enabled = True
        Command1.Enabled = True
        If Val(Node.Text) <> TiendaActual Then Mover = True
        TiendaActual = Val(Node.Text)
        D_TiendaActual = Consulta_Tiendas(TiendaActual, 1)
        Refresca_Info_Tienda Val(Node.Text)
        Resp = LeeParamsTIenda
        If Not Resp Then
            MsgBox CargaCadena(1145), vbCritical
            End
        Else
            Pon_Barra
            Adapta_Barra
            If Mover Then MueveBarra True
        End If
    Else
        If TiendaActual <> 0 Then
            TiendaActual = 0
            MueveBarra False
        End If
        Barra.Enabled = False
        Command1.Enabled = False
        Refresca_Info_Tienda 0
    End If
End Sub





Private Sub Barra_FolderClick(Folder As AspBigBar.Folder)
    FolderQL2 = True
    
End Sub

Private Sub Barra_FolderItemClick(FolderItem As AspBigBar.FolderItem)
    '
    
    If FolderQL2 = True Then
        FolderQL2 = False
        Exit Sub
    End If
    If FolderQL = True Then
        Exit Sub
    Else
        FolderQL = True
    End If
    '
   
    Me.Enabled = False
    Me.MousePointer = ccHourglass
    Select Case Left(FolderItem.Key, 1)
        Case "D"
            mnuT_Click Val(Mid(FolderItem.Key, 2))
        Case "T"
            Select Case Val(Mid(FolderItem.Key, 2))
                Case 1
                    If TiendaActual < 100 Then
                        Aux_Ejecuta "gtarti", App.Path & "\t" & Format(TiendaActual, "00"), "", D_TiendaActual(0).lScalePos
                    Else
                        Aux_Ejecuta "gtarti", App.Path & "\t" & Format(TiendaActual, "000"), "", D_TiendaActual(0).lScalePos
                    End If
                Case 2
                    If TiendaActual < 100 Then
                        Aux_Ejecuta "gtsecs", App.Path & "\t" & Format(TiendaActual, "00"), "", D_TiendaActual(0).lScalePos
                    Else
                        Aux_Ejecuta "gtsecs", App.Path & "\t" & Format(TiendaActual, "000"), "", D_TiendaActual(0).lScalePos
                    End If
                Case 3
                    If TiendaActual < 100 Then
                        Aux_Ejecuta "gtvend", App.Path & "\t" & Format(TiendaActual, "00"), "", D_TiendaActual(0).lScalePos
                    Else
                        Aux_Ejecuta "gtvend", App.Path & "\t" & Format(TiendaActual, "000"), "", D_TiendaActual(0).lScalePos
                    End If
                Case 4
                    If Not UsaGamaAlta Then
                        MsgBox CargaCadena(1146), vbCritical
                        Me.Enabled = True
                        Me.MousePointer = ccDefault
                        FolderQL = False
                        Exit Sub
                    End If
                    ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, 66, "", D_TiendaActual(0).lScalePos
                Case 5
                    If UsaGamaAlta Or UsaSC10 Or (UsaEcoPrint) Then
                        ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, 67, "", D_TiendaActual(0).lScalePos
                    End If
                Case 6
                    If Not UsaGamaAlta Then
                        MsgBox CargaCadena(1146), vbCritical
                        Me.Enabled = True
                        Me.MousePointer = ccDefault
                        FolderQL = False
                        Exit Sub
                    End If
                    If TiendaActual < 100 Then
                    Aux_Ejecuta "pedido", App.Path & "\t" & Format(TiendaActual, "00"), _
                    App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "ped.sdf", D_TiendaActual(0).lScalePos
                    Else
                    Aux_Ejecuta "pedido", App.Path & "\t" & Format(TiendaActual, "000"), _
                    App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "ped.sdf", D_TiendaActual(0).lScalePos
                    End If
                Case 7
                    If Not UsaGamaAlta Then
                        MsgBox CargaCadena(1146), vbCritical
                        Me.Enabled = True
                        Me.MousePointer = ccDefault
                        FolderQL = False
                        Exit Sub
                    End If
                    If TiendaActual < 100 Then
                    Aux_Ejecuta "inventario", App.Path & "\t" & Format(TiendaActual, "00"), _
                    App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "inv.sdf", D_TiendaActual(0).lScalePos
                    Else
                    Aux_Ejecuta "inventario", App.Path & "\t" & Format(TiendaActual, "000"), _
                    App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "inv.sdf", D_TiendaActual(0).lScalePos
                    
                    End If
                Case 8
                    If Not (UsaGamaAlta Or (UsaEcoPrint And TipoEcoPrint = 2)) Then
                        MsgBox CargaCadena(1146), vbCritical
                        Me.Enabled = True
                        Me.MousePointer = ccDefault
                        FolderQL = False
                        Exit Sub
                    End If
                    If TiendaActual < 100 Then
                    Aux_Ejecuta "pedinv", App.Path & "\t" & Format(TiendaActual, "00"), _
                    App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "pedinv.sdf", D_TiendaActual(0).lScalePos
                    Else
                    Aux_Ejecuta "pedinv", App.Path & "\t" & Format(TiendaActual, "000"), _
                    App.Path & "\t" & Format(TiendaActual, "000") & "\" & Format(TiendaActual, "000") & "pedinv.sdf", D_TiendaActual(0).lScalePos
                    
                    End If
            End Select
        Case "C"
            mnuC_Click Val(Mid(FolderItem.Key, 2))
        Case "X"
            If HaySeguridad Then
                If ConsultarPermiso(17) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            Select Case Val(Mid(FolderItem.Key, 2))
                Case 1
                    ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, 61, "", D_TiendaActual(0).lScalePos
                Case 2
                    MostrarModal Me, FrmDatosTienda
                    Refresca_Info_Tienda TiendaActual
            End Select
    End Select
   
    FolderQL = False
    Me.MousePointer = 0
    Me.Enabled = True
    Barra.Enabled = True
    Me.SetFocus
   
End Sub




Private Sub CheckTransfer_Click()
    Dim nFich As Integer
    Dim sPass As String
    sPass = " EPFAB"
    'If Dir(App.Path & "\gigante.txt") <> "" And sPassG <> "" Then
    If sPassG <> "" Then
        If Dir(App.Path & "\passmh") <> "" Then
            nFich = FreeFile()
            Open App.Path & "\passmh" For Input As #nFich
            Line Input #nFich, sPass
            Close #nFich
            sPass = UCase(sPass)
        End If
        
        
        MostrarModal Me, FormP
        Unload FormP
        Set FormP = Nothing
        
        Me.SetFocus
        
        If UCase(sPassG) <> sPass Then
            If CheckTransfer.Value = 1 Then
                CheckTransfer.Value = 0
            Else
                CheckTransfer.Value = 1
            End If
            Exit Sub
        End If
    End If
    sPassG = "NONE"
    
    If CheckTransfer.Value = 0 Then
    '    CheckTransfer.Value = 0
        If Dir(App.Path & "\shoptras.txt") <> "" Then
            Kill App.Path & "\shoptras.txt"
        End If
        'Else
    Else
    '    CheckTransfer.Value = 1
        If Dir(App.Path & "\shoptras.txt") = "" Then
            nFich = FreeFile()
            Open App.Path & "\shoptras.txt" For Output As #nFich
            Close #nFich
        End If
    End If
    
End Sub

Private Sub ChkDirect_Click()
Dim nF As Integer
If ChkDirect.Value Then
    nF = FreeFile()
    Open App.Path & "\transfertable.cfg" For Output As #nF
    Print #nF, "direct"
    Close #nF
Else
    If Dir(App.Path & "\transfertable.cfg") <> "" Then Kill App.Path & "\transfertable.cfg"
End If
End Sub

Private Sub Command1_Click()
    Me.Visible = False
    If Dir(App.Path & "\ekobal_1.dat") <> "" And TiendaActual = 1 Then
        If Dir(App.Path & "\ekobal.dat") <> "" Then
            Kill App.Path & "\ekobal.dat"
        End If
        FileCopy App.Path & "\ekobal_1.dat", App.Path & "\ekobal.dat"
    End If
    If Dir(App.Path & "\baleko_1.dat") <> "" And TiendaActual = 1 Then
        If Dir(App.Path & "\ekobal.dat") <> "" Then
            Kill App.Path & "\ekobal.dat"
        End If
        FileCopy App.Path & "\baleko_1.dat", App.Path & "\ekobal.dat"
    End If
    If Dir(App.Path & "\ekobal_2.dat") <> "" And TiendaActual = 2 Then
        If Dir(App.Path & "\ekobal.dat") <> "" Then
            Kill App.Path & "\ekobal.dat"
        End If
        FileCopy App.Path & "\ekobal_2.dat", App.Path & "\ekobal.dat"
    End If
    If Dir(App.Path & "\baleko_2.dat") <> "" And TiendaActual = 2 Then
        If Dir(App.Path & "\ekobal.dat") <> "" Then
            Kill App.Path & "\ekobal.dat"
        End If
        FileCopy App.Path & "\baleko_2.dat", App.Path & "\ekobal.dat"
    End If
    
    ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, 0, "", D_TiendaActual(0).lScalePos
    Me.Visible = True
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim nMiFich As Integer
    Dim dB1 As Database
    Dim rStdB1 As Recordset
    Dim sMiFich As String
    Dim sMiPatron As String
    
    Select Case Index
        Case 1
            Unload Me
            End
        Case 0
            NoSalir = True
            Unload Me
            NoSalir = False
            FrmLogin.Show
        Case 2
            If HaySeguridad Then
                If ConsultarPermiso(17) = False Then
                    MsgBox CargaCadena(693), vbCritical
                    Exit Sub
                End If
            End If
            MostrarModal Me, FrmConfiguracion
            CambiarIdioma
        Case 3 'group
            MostrarModal Me, FrmGrupos
        Case 4 'captura ficheros por grupos
            If Dir(App.Path & "\tiendas.mdb") <> "" Then
                Set dB1 = OpenDatabase(App.Path & "\tiendas.mdb")
                Set rStdB1 = dB1.OpenRecordset("select * from grupos where grp<>0" & " order by tienda")
                If Not rStdB1.EOF Then
                    rStdB1.MoveFirst
                    Do While Not rStdB1.EOF
                        sMiFich = App.Path & "\" & Trim(rStdB1.Fields("fichero"))
                        sMiPatron = App.Path & "\etc\" & Trim(rStdB1.Fields("patron")) & ".FIA"
                        If Dir(sMiPatron) <> "" Then
                            sMiPatron = App.Path & "\" & Format(rStdB1.Fields("tienda"), "00") & Trim(rStdB1.Fields("patron"))
                            If Dir(sMiFich) <> "" Then
                                                        
                                frmWAIT.Show
                                frmWAIT.Label1.Caption = "GROUP: " & Format(rStdB1.Fields("grp"), "00")
                                frmWAIT.Label2.Caption = "SHOP:  " & Format(rStdB1.Fields("tienda"), "00")
                                frmWAIT.Refresh
                                
                                FileCopy sMiFich, sMiPatron
                                
                                
                                nMiFich = FreeFile()
                                Open App.Path & "\multiimp.ord" For Output As #nMiFich
                                Print #nMiFich, Format(rStdB1.Fields("tienda"), "00") & Trim(rStdB1.Fields("patron"))
                                Close #nMiFich
                                If Dir(App.Path & "\impenv.ord") <> "" Then
                                    ActivaSingle rStdB1.Fields("tienda"), "", 34, "", D_TiendaActual(0).lScalePos
                                Else
                                    ActivaSingle rStdB1.Fields("tienda"), "", 17, "", D_TiendaActual(0).lScalePos
                                End If
                            
                                frmWAIT.Hide
                                
                            End If
                        
                        End If
                        
                        rStdB1.MoveNext
                    Loop
                    rStdB1.Close
                    'dB1.Close
                End If
                dB1.Close
                Set dB1 = Nothing
            End If
    
        Case 5 'captura bal.dat de cada tienda
            If Dir(App.Path & "\tiendas.mdb") <> "" Then
                Set dB1 = OpenDatabase(App.Path & "\tiendas.mdb")
                Set rStdB1 = dB1.OpenRecordset("select * from grupos where tienda<>0" & " order by tienda")
                If Not rStdB1.EOF Then
                    rStdB1.MoveFirst
                    Do While Not rStdB1.EOF
                        sMiFich = App.Path & "\" & Format(rStdB1.Fields("tienda"), "00") & "bal.dat"
                        sMiPatron = App.Path & "\etc\" & "bal.dat.FIA"
                        If Dir(sMiPatron) <> "" Then

                            If Dir(sMiFich) <> "" Then
                                                        
                                frmWAIT.Show
                                frmWAIT.Label1.Caption = "Bal.Dat Import" '& Format(rStdB1.Fields("grp"), "00")
                                frmWAIT.Label2.Caption = "SHOP:  " & Format(rStdB1.Fields("tienda"), "00")
                                frmWAIT.Refresh
                                
                                FileCopy sMiFich, App.Path & "\bal.dat"
                                
                                
                                nMiFich = FreeFile()
                                Open App.Path & "\multiimp.ord" For Output As #nMiFich
                                Print #nMiFich, "\bal.dat"
                                Close #nMiFich
                                'If Dir(App.Path & "\impenv.ord") <> "" Then
                                '    ActivaSingle rStdB1.Fields("tienda"), "", 34
                                'Else
                                    ActivaSingle rStdB1.Fields("tienda"), "", 17, "", D_TiendaActual(0).lScalePos
                                'End If
                            
                                frmWAIT.Hide
                                
                            End If
                        
                        End If
                        
                        rStdB1.MoveNext
                    Loop
                    rStdB1.Close
                    'dB1.Close
                End If
                dB1.Close
                Set dB1 = Nothing
            End If
    
    End Select
End Sub

Public Sub Form_Load()
    Dim nF As Integer
    
    If Dir(App.Path & "\etc\ekobal.dat.fia") <> "" Then
        FrmPrincipal.WindowState = 2
        FrmPrincipal.BorderStyle = 0
    End If
    
    Me.Caption = "Saturno MultiStore " & MiVer
    sPassG = ""
    If Dir(App.Path & "\passmh") = "" Then
        nF = FreeFile()
        Open App.Path & "\passmh" For Output As #nF
        Print #nF, " EPFAB"
        Close #nF
    End If
    Picture1.Height = 8000 '7600
    Image1.Visible = True
    If Dir(App.Path & "\shoptras.txt") <> "" Then
        CheckTransfer.Value = 1
    End If
    If Dir(App.Path & "\transfertable.cfg") <> "" Then
        ChkDirect.Value = 1
    End If
    
    Refresca_tiendas
    Command1.Enabled = False
    Refresca_Info_Tienda 0
    If HaySeguridad Then
        Command3(0).Visible = True
        Label1(1).Visible = True
    Else
        Command3(0).Visible = False
        Label1(1).Visible = False
    End If
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    DrawGradient Picture1, 220, 220, 220, 0, 50, 0, Picture1.ScaleHeight, 1, True, 1, 1, 1
    Pon_Barra
    Barra.Enabled = False
    CambiarIdioma
    Me_Init
    Picture1.Height = (Me.Height - 800) '500
    Crea_Grp
End Sub
Public Sub CambiarIdioma()
   
    Label1(0).Caption = CargaCadena(7)
    Label1(1).Caption = CargaCadena(518)
    Label1(2).Caption = CargaCadena(1081)
    Label1(3).Caption = CargaCadena(1286)
    Label1(4).Caption = CargaCadena(1287)
    LabelTransfer.Caption = CargaCadena(1288)
    
    Command1.Caption = CargaCadena(1082)
    
    Label2(0).Caption = CargaCadena(1083)
    '
    mnuIntroducir.Caption = CargaCadena(1166)
    mnuComu.Caption = CargaCadena(4)
    mnuCrearTienda.Caption = CargaCadena(1165)
    mnuListadoTIendas.Caption = CargaCadena(1167)
    mnuT(0).Caption = CargaCadena(580)
    mnuT(1).Caption = CargaCadena(23)
    mnuT(2).Caption = CargaCadena(21)
    mnuT(3).Caption = CargaCadena(505)
    mnuT(4).Caption = CargaCadena(919)
    mnuT(6).Caption = CargaCadena(12)
    mnuT(7).Caption = CargaCadena(24)
    mnuT(8).Caption = CargaCadena(22)
    mnuT(10).Caption = CargaCadena(8)
    mnuT(11).Caption = CargaCadena(1056)
    mnuT(12).Caption = CargaCadena(25)
    mnuT(14).Caption = CargaCadena(406)
    mnuT(15).Caption = CargaCadena(410)
    mnuC(0).Caption = CargaCadena(151)
    mnuC(1).Caption = CargaCadena(763)
    mnuC(15).Caption = CargaCadena(161)
    mnuC(16).Caption = CargaCadena(1030)
    mnuC1(2).Caption = CargaCadena(738) & " " & CargaCadena(9)
    mnuC1(3).Caption = CargaCadena(738) & " " & CargaCadena(11)
    mnuC1(4).Caption = CargaCadena(738) & " " & CargaCadena(412)
    mnuC1(5).Caption = CargaCadena(738) & " " & CargaCadena(21)
    mnuC1(6).Caption = CargaCadena(738) & " " & CargaCadena(505)
    mnuC1(7).Caption = CargaCadena(738) & " " & CargaCadena(22)
    mnuC1(8).Caption = CargaCadena(738) & " " & CargaCadena(215)
    mnuC1(9).Caption = CargaCadena(738) & " " & CargaCadena(12)
    mnuC1(10).Caption = CargaCadena(738) & " " & CargaCadena(8)
    mnuC1(11).Caption = CargaCadena(738) & " " & CargaCadena(1056)
    mnuC1(12).Caption = CargaCadena(738) & " " & CargaCadena(13)
    '
End Sub
Private Sub Pon_Barra()
    Dim nCont As Integer
    
    Do_Events
    Limpia_Barra
    On Error Resume Next
    
    LeeMultiConfig
    
    Barra.AddFolder CargaCadena(1084), "P5" 'datos de artículos
    Barra.AddFolder CargaCadena(1085), "P4" 'estructura de tienda
    Barra.AddFolder CargaCadena(1086), "P3" 'otros datos de la tienda
    Barra.AddFolder CargaCadena(1087), "P2" 'comunicaciones con la tienda
    Barra.AddFolder CargaCadena(1088), "P1" 'totales de la tienda
    Barra.AddFolder CargaCadena(456) & " " & CargaCadena(108) & "/" & CargaCadena(109), "P7"
    'informe pedido inventario
    Barra.AddFolder CargaCadena(1089), "P6" 'configuración de la tienda
    
    'totales de la tienda
    Barra.Folders(5).AddItem CargaCadena(156), "T1", "total"
    Barra.Folders(5).AddItem CargaCadena(153), "T2", "total"
    Barra.Folders(5).AddItem CargaCadena(157), "T3", "total"
    Barra.Folders(5).AddItem CargaCadena(642), "T4", "total"
    Barra.Folders(5).AddItem CargaCadena(920), "T5", "total"
    ''''''''''''''''''''''
    'pedido e inventario
    Barra.Folders(6).AddItem CargaCadena(1091) & " (Euroscale)", "T6", "total"
    Barra.Folders(6).AddItem CargaCadena(1092) & " (Euroscale)", "T7", "total"
    Barra.Folders(6).AddItem CargaCadena(1196), "T8", "total"
    '''''''''''''''''''''
    'Barra.Folders(1).AddItem "", "N0", ""
    'datos de artículos...
    nCont = 0
    nMyPer = 65535
    nCont = mnuTTest(6, nCont)
    If nMyPer <> 65535 Then Barra.Folders(1).AddItem CargaCadena(12), "D6", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(7, nCont)
    If nMyPer <> 65535 Then Barra.Folders(1).AddItem CargaCadena(24), "D7", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(8, nCont)
    If nMyPer <> 65535 Then Barra.Folders(1).AddItem CargaCadena(321), "D8", "hydra"
    'If Dir(App.Path & "\gigante.txt") <> "" Then
        nMyPer = 65535
        nCont = mnuTTest(9, nCont)
        If nMyPer <> 65535 Then Barra.Folders(1).AddItem CargaCadena(151), "D9", "hydra"
    'End If
    If nCont = 0 Then 'de momento aparece pero sin opciones...
        'Barra.Folders("P5").Visible = False
        'Barra.Folders.Remove ("P5")
    End If
    nMyPer = 0
    ''''''''''''''''''''''
    ' estructura de la tienda...
    nMyPer = 65535
    nCont = 0
    nCont = mnuTTest(0, nCont)
    If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(580), "D0", "hydra"
    If FamiliasFijas Then
        nMyPer = 65535
        nCont = mnuTTest(18, nCont)
        If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(678), "D18", "hydra"
    End If
    nMyPer = 65535
    nCont = mnuTTest(1, nCont)
    If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(23), "D1", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(2, nCont)
    If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(21), "D2", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(3, nCont)
    If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(505), "D3", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(12, nCont)
    If nMyPer <> 65535 Then Barra.Folders(2).AddItem CargaCadena(13), "D12", "hydra"
    If nCont = 0 Then
        'Barra.Folders(2).Visible = False
        'Barra.Folders.Remove (2)
    End If
    nMyPer = 0
    ''''''''''''''''''''''''''''''
    'otros datos de la tienda...
    nMyPer = 65535
    nCont = 0
    nCont = mnuTTest(4, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(215), "D4", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(11, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(1056), "D11", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(10, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(8), "D10", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(14, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(406), "D14", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(15, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(410), "D15", "hydra"
    nMyPer = 65535
    nCont = mnuTTest(16, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(1114), "D16", "hydra" ' taras ecolabel
    nMyPer = 65535
    nCont = mnuTTest(17, nCont)
    If nMyPer <> 65535 Then Barra.Folders(3).AddItem CargaCadena(1115), "D17", "hydra" ' textos ecolabel
    If nCont = 0 Then
        'Barra.Folders(3).Visible = False
        'Barra.Folders.Remove (3)
    End If
    nMyPer = 0
    '''''''''''''''''''''''''''''
    '''''''''''''''''''''
    'Barra.Folders(4).AddItem CargaCadena(1132), "C17", "comu"
    'Barra.Folders(4).AddItem CargaCadena(151), "C0", "comu" 'envio de modificaciones
    'Barra.Folders(4).AddItem CargaCadena(1133), "C18", "comu"
    'Barra.Folders(4).AddItem CargaCadena(204), "C13", "comu"
    'Barra.Folders(4).AddItem CargaCadena(161), "C12", "comu"
    'Barra.Folders(4).AddItem CargaCadena(1030), "C16", "comu"
    '''''''''''''''''''''
    'comunicaciones con la tienda
    nMyPer = 65535
    nCont = 0
    nCont = mnuCTest(0, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(151), "C0", "comu" 'envio de modificaciones 0
    nMyPer = 65535
    nCont = mnuCTest(16, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(1030), "C16", "comu" 'recibir tiques 16
    nMyPer = 65535
    nCont = mnuCTest(12, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(161), "C12", "comu" 'grandes totales 12
    nMyPer = 65535
    nCont = mnuCTest(13, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(204), "C13", "comu" 'base maestra 13
    nMyPer = 65535
    nCont = mnuCTest(17, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(1132), "C17", "comu" 'rec totales + envio modificaciones 17
    nMyPer = 65535
    nCont = mnuCTest(18, nCont)
    If nMyPer <> 65535 Then Barra.Folders(4).AddItem CargaCadena(1133), "C18", "comu" 'recoger pedido + inventario 18
    If nCont = 0 Then
        'Barra.Folders(4).Visible = False
        'Barra.Folders.Remove (4)
    End If
    nMyPer = 0
    ''''''''''''''''''''
    'configuración de la tienda
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        If ConsultarPermiso(17) = True Then
            Barra.Folders(7).AddItem CargaCadena(1089), "X1", "herram"
            Barra.Folders(7).AddItem CargaCadena(1090), "X2", "herram"
        Else
            'Barra.Folders(7).Visible = False
            'Barra.Folders.Remove ("P6")
        End If
    Else
        Barra.Folders(7).AddItem CargaCadena(1089), "X1", "herram"
        Barra.Folders(7).AddItem CargaCadena(1090), "X2", "herram"
    End If
    

    On Error GoTo 0
End Sub
Private Sub Adapta_Barra()
    On Error Resume Next
    Barra.Folders("P7").Visible = True
    If Not UsaGamaAlta Then
        Barra.Folders(6).RemoveItem "T6"
        Barra.Folders(6).RemoveItem "T7"
        Barra.Folders(3).RemoveItem "D4"
        Barra.Folders(3).RemoveItem "D11"
        Barra.Folders(3).RemoveItem "D14"
        Barra.Folders(5).RemoveItem "T4"
        If Not UsaSC10 Then
            If Not (UsaEcoPrint And TipoEcoPrint = 2) Then
                Barra.Folders(5).RemoveItem "T5"
            End If
        End If
    End If
    If Not (UsaEcoPrint And TipoEcoPrint = 2) Then
        Barra.Folders(6).RemoveItem "T8"
        If Not UsaGamaAlta Then
            Barra.Folders("P7").Visible = False
            Barra.Folders(4).RemoveItem "C18"
        End If
    End If
    If Not ((UsaSC10 And (Not TipoC9)) Or UsaEco2) Then
        Barra.Folders(3).RemoveItem "D15"
    End If
    If (Not UsaGamaAlta) And (Not UsaSC10) Then
        Barra.Folders(1).RemoveItem "D8"
        Barra.Folders(2).RemoveItem "D3"
    End If
    If Not UsaEco2 Then
        Barra.Folders(3).RemoveItem "D16"
        Barra.Folders(3).RemoveItem "D17"
    End If
    If Not FamiliasFijas Then
        Barra.Folders(2).RemoveItem "D18"
    End If
    On Error GoTo 0
End Sub
Private Sub Limpia_Barra()
    On Error Resume Next
    
    Barra.Folders(5).RemoveItem "T1"
    
    Barra.Folders(5).RemoveItem "T2"
    Barra.Folders(5).RemoveItem "T3"
    Barra.Folders(5).RemoveItem "T4"
    Barra.Folders(5).RemoveItem "T5"
    Barra.Folders(6).RemoveItem "T6"
    Barra.Folders(6).RemoveItem "T7"
    Barra.Folders(6).RemoveItem "T8"
    '
    'Barra.Folders(1).RemoveItem "", "N0", ""
    Barra.Folders(1).RemoveItem "D6"
    Barra.Folders(1).RemoveItem "D7"
    Barra.Folders(1).RemoveItem "D8"
    Barra.Folders(2).RemoveItem "D0"
    Barra.Folders(2).RemoveItem "D1"
    Barra.Folders(2).RemoveItem "D2"
    Barra.Folders(2).RemoveItem "D3"
    Barra.Folders(2).RemoveItem "D18"
    Barra.Folders(2).RemoveItem "D12"
    Barra.Folders(3).RemoveItem "D4"
    Barra.Folders(3).RemoveItem "D11"
    Barra.Folders(3).RemoveItem "D10"
    Barra.Folders(3).RemoveItem "D14"
    Barra.Folders(3).RemoveItem "D15"
    '
    Barra.Folders(3).RemoveItem "D16"
    Barra.Folders(3).RemoveItem "D17"
    '
    Barra.Folders(4).RemoveItem "C0"
    Barra.Folders(4).RemoveItem "C17"
    Barra.Folders(4).RemoveItem "C18"
    Barra.Folders(4).RemoveItem "C13"
    Barra.Folders(4).RemoveItem "C12"
    Barra.Folders(4).RemoveItem "C16"
    Barra.Folders(7).RemoveItem "X1"
    Barra.Folders(7).RemoveItem "X2"
    On Error GoTo 0
    
End Sub
Private Sub Refresca_tiendas()
    Dim Num As Integer
    Dim Bucle As Integer
    Dim Tiendas() As TipoTienda
    Arbol.Nodes.Clear
    Num = Consulta_NTiendas
    If Num <> 0 Then
        Tiendas = Consulta_Tiendas(1, Num)
    End If
    Arbol.Nodes.Add , , "sys", CargaCadena(1126), 2
    If Num > 0 Then
        For Bucle = 0 To Num - 1
            If (Not HaySeguridad) Or (HaySeguridad And UsuarioActual.Permisos(7) = 0) Or (HaySeguridad And UsuarioActual.Permisos(7) <> 254 And (UsuarioActual.Permisos(7) = Tiendas(Bucle).Numero) Or (HaySeguridad And UsuarioActual.Permisos(7) = 255 And Tiendas(Bucle).Numero = 999)) Then
                Arbol.Nodes.Add "sys", tvwChild, , Tiendas(Bucle).Numero, 1
            End If
        Next Bucle
    End If
    Arbol.Nodes(1).Expanded = True
End Sub



Private Sub Form_Resize()
    
    If Dir(App.Path & "\etc\ekobal.dat.fia") <> "" Then
        Me.WindowState = vbMaximized
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Me_Resize
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Dir(App.Path & "\etc\ekobal.dat.fia") = "" Then
        If Not NoSalir Then End
    Else
        Cancel = 1
    End If
    

End Sub

Private Sub Image2_Click(Index As Integer)
    Select Case Index
        Case 0
            'Aux_Ejecuta "backupmulti", App.Path & "\t" & Format(TiendaActual, "00"), Backup_Path
            On Error Resume Next
            Aux_Ejecuta "backupmulti", App.Path, Backup_Path, D_TiendaActual(0).lScalePos
            If Err.Number <> 0 Then
                On Error GoTo 0
                Aux_Ejecuta "backupmulti", App.Path, Backup_Path, False
            End If
            'MostrarModal Me, FrmBackup
    End Select
            
End Sub
'MENÚ DE COMUNICACIONES
Private Sub mnuC_Click(Index As Integer)
    Dim Orden As Integer
    Dim MiTienda As TipoTienda
    Dim Permiso As Integer
    Dim Resp
    Orden = 0
    Select Case Index
        Case 0 'envio de modificaciones
            Orden = 2
            Permiso = 13
        Case 12 'grandes totales.
            Resp = MsgBox(CargaCadena(1013), vbExclamation + vbYesNo + vbDefaultButton2)
            If Resp = vbNo Then Exit Sub
            Orden = 15
            Permiso = 14
        Case 13 'base maestra.
            Orden = 1
            Permiso = 11
        Case 16 'recibir tiques.
            If Not ((UsaSC10 And Not TipoC9) Or (UsaGamaAlta)) Then
                If Not UsaScalePos Then
                    MsgBox CargaCadena(1146), vbCritical
                    Exit Sub
                End If
            End If
            Orden = 19
            Permiso = 14
        Case 17 'rec totales + env. modificaciones
            Orden = 68
            Permiso = 14
        Case 18 'recoger pedido e inventario.
            If Not (UsaGamaAlta Or (UsaEcoPrint And TipoEcoPrint = 2)) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 69
            Permiso = 14
    End Select
    If HaySeguridad Then
        If ConsultarPermiso(Permiso) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    If Orden <> 0 Then
        Me.Enabled = False
        ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, Orden, "", D_TiendaActual(0).lScalePos
        Me.Enabled = True
        Me.SetFocus
    End If

End Sub
Private Sub mnuC1_Click(Index As Integer)
    Dim Orden As Integer
    Dim Permiso As Integer
    Dim MiTienda As TipoTienda
    Orden = 0
    Select Case Index
        Case 9
            Permiso = 11
        Case Else
            Permiso = 12
    End Select
    Select Case Index
        Case 2
            Orden = 3
        Case 3
            Orden = 6
        Case 4
            Orden = 12
        Case 5
            Orden = 11
        Case 6
            Orden = 10
        Case 7
            Orden = 14
        Case 8
            Orden = 58
        Case 9
            Orden = 9
        Case 10
            Orden = 33
        Case 11
            Orden = 59
        Case 12
            Orden = 13
            
    End Select
    If HaySeguridad Then
        If ConsultarPermiso(Permiso) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Sub
        End If
    End If
    If Orden <> 0 Then
        Me.Enabled = False
        ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, Orden, "", D_TiendaActual(0).lScalePos
        Me.Enabled = True
        Me.SetFocus
    End If
End Sub

Private Sub mnuCrearTienda_Click()
    If HaySeguridad And UsuarioActual.Nombre <> "admin" Then
        MsgBox CargaCadena(703), vbCritical
        Me.Enabled = True
        Exit Sub
    End If
    MostrarModal Me, FrmAsistente
    Unload FrmAsistente
    Refresca_tiendas
    Refresca_Info_Tienda TiendaActual
End Sub

Private Sub mnuListadoTIendas_Click()
    Listar_Las_Tiendas
End Sub

'menú datos de artículos
'     datos de estructura
'     otros datos de la Tienda
Private Sub mnuT_Click(Index As Integer)
    Dim Orden As Integer
    Dim Permiso2 As Integer
    Dim Permiso As Integer
    Dim MiTienda As TipoTienda
    Dim dB As Database
    Dim Rst As Recordset
    
    Orden = 0
    Permiso = 0
    Permiso2 = 0
    Select Case Index
        Case 0 'mostradores y equipos.
            Orden = 20
            Permiso = 0
        Case 1 'códigos de Barras.
            Orden = 25
            Permiso = 7
        Case 2 'Cabeceras y leyendas.
            If UsaScalePos Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 24
            Permiso = 6
        Case 3 'Publicidad y ofertas.
            If (Not UsaSC10) And (Not UsaGamaAlta) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 23
            Permiso = 8
        Case 4
            If Not UsaGamaAlta Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 56
            Permiso = 6
        Case 6 'entrada en artículos...
            Orden = 21
            Permiso = 1
            Permiso2 = 24
        Case 7 'cambios de precio...
            Orden = 31
            Permiso = 2
        Case 8 'entrada en teclas...
            If Not UsaGamaAlta And Not UsaSC10 Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 26
            Permiso = 10
        
        Case 9 'envio de modificaciones...
            'If Dir(App.Path & "\gigante.txt") <> "" Then
                Orden = 2
                Permiso = 13
            'End If
            
        Case 10
            Orden = 32
            Permiso = 5
        Case 11
            If Not UsaGamaAlta Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 57
            Permiso = 9
        Case 12 'Vendedores
            If (Not UsaSC10) And (Not UsaGamaAlta) And (Not UsaScalePos) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 22
            Permiso = 9
        Case 14
            If Not UsaGamaAlta Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 27
            Permiso = 3
        Case 15
            If ((UsaSC10 And TipoC9) Or (Not UsaSC10)) And (Not UsaEco2) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 28
            Permiso = 3
        Case 16
            If (Not UsaEco2) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 70
        Case 17
            If (Not UsaEco2) Then
                MsgBox CargaCadena(1146), vbCritical
                Exit Sub
            End If
            Orden = 71
        Case 18 'Familias
            Orden = 72
            Permiso = 0
    End Select
    If HaySeguridad Then
        If Permiso2 = 0 Then
            If ConsultarPermiso(Permiso) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        Else
            If ConsultarPermiso(Permiso) = False And ConsultarPermiso(Permiso2) = False Then
                MsgBox CargaCadena(693), vbCritical
                Exit Sub
            End If
        End If
    End If
    If Orden <> 0 Then
        Me.Enabled = False
        If Orden = 2 And TiendaActual = 999 Then
            
            Set dB = OpenDatabase(App.Path & "\tiendas.mdb")
            Set Rst = dB.OpenRecordset("select * from grupos where tienda<>" & CStr(TiendaActual) & " order by tienda")
            If Not Rst.EOF Then
                Rst.MoveFirst
                Do While Not Rst.EOF
                    'MsgBox CStr(Rst.Fields("tienda")) & "Shop:" & CStr(Rst.Fields("tienda"))
                    ActivaSingle Rst.Fields("tienda"), "Shop:" & CStr(Rst.Fields("tienda")), Orden, Trim(Rst.Fields("path")), D_TiendaActual(0).lScalePos
                    Rst.MoveNext
                Loop
            
            End If
            Rst.Close
            dB.Close
            Set Rst = Nothing
            Set dB = Nothing
        Else
            On Error Resume Next
            ActivaSingle TiendaActual, D_TiendaActual(0).Nombre, Orden, "", D_TiendaActual(0).lScalePos
            On Error GoTo 0
        End If
        Me.Enabled = True
        Me.SetFocus
    End If
End Sub
Public Sub Refresca_Info_Tienda(NTienda As Integer)
    Dim MiTienda() As TipoTienda
    Dim MyChar As String
    If NTienda >= 1 Then
        MiTienda = Consulta_Tiendas(NTienda, 1)
        Info.Text = ""
        Info.Text = Info.Text & CargaCadena(46) & " : " & MiTienda(0).Numero & vbCrLf
        Info.Text = Info.Text & MiTienda(0).Nombre & vbCrLf
        Info.Text = Info.Text & MiTienda(0).Direccion & vbCrLf
        Info.Text = Info.Text & MiTienda(0).Poblacion & vbCrLf
        If MiTienda(0).CodPostal <> "" And MiTienda(0).Provincia <> "" Then
            MyChar = " - "
        Else
            MyChar = ""
        End If
        Info.Text = Info.Text & MiTienda(0).CodPostal & MyChar & MiTienda(0).Provincia & vbCrLf
        Info.Text = Info.Text & MiTienda(0).TlfContacto & vbCrLf
        Info.Text = Info.Text & "Group:" & Trim(MiTienda(0).Grp) & vbCrLf
        Info.Text = Info.Text & "Path :" & Trim(MiTienda(0).sPathS)
        sPathTien = Trim(MiTienda(0).sPathS)
        If Right(sPathTien, 1) = "\" Then
           sPathTien = Mid(sPathTien, 1, Len(sPathTien) - 1)
           
        End If
    Else
        Info.Text = ""
        Info.Text = CargaCadena(1125) '"Sistema de Tiendas"
    End If
    Info.SelStart = 0
    Info.SelLength = Len(Info.Text)
    Info.SelColor = vbWhite
    Info.SelLength = 0
End Sub

Public Function LeeParamsTIenda() As Boolean
    Dim Arch As Integer
    Dim bUF() As String
    Dim buf2 As String
    Dim NBuf As Integer
    Dim Resul As Boolean
    Dim sLocal As String
    
    Resul = True
    UsaGamaAlta = False
    UsaScalePos = False
    UsaSC10 = False
    TipoC9 = False
    UsaEco1 = False
    UsaEco2 = False
    UsaEcoPrint = False
    
    If TiendaActual < 100 Then
        sLocal = App.Path & "\t" & Format(TiendaActual, "00")
    Else
        sLocal = App.Path & "\t" & Format(TiendaActual, "000")
    End If
    
    If Dir(sLocal & "\hydratouch.ini") <> "" Then
        Arch = FreeFile()
        Open sLocal & "\hydratouch.ini" For Input As #Arch
        Do Until EOF(Arch)
            NBuf = NBuf + 1
            ReDim Preserve bUF(NBuf)
            Input #Arch, bUF(NBuf - 1), buf2
        Loop
        Close #Arch
        If NBuf >= 91 Then
            If bUF(3) = 1 Then UsaGamaAlta = True
            If D_TiendaActual(0).lScalePos Then
                UsaGamaAlta = False
                UsaScalePos = True
            End If
            If bUF(4) = 1 Then UsaSC10 = True
            If bUF(91) = "#TRUE#" Then TipoC9 = True
            If NBuf >= 125 Then
                If bUF(116) = "#TRUE#" Or bUF(116) = "1" Then UsaEco1 = True
                If bUF(118) = "#TRUE#" Or bUF(118) = "1" Then UsaEco2 = True
                If bUF(123) = "#TRUE#" Or bUF(123) = "1" Then UsaEcoPrint = True
                TipoEcoPrint = Val(bUF(124))
            End If
        Else
            Resul = False
        End If
        FamiliasFijas = False
        If NBuf >= 132 Then
            If bUF(132) = "1" Then FamiliasFijas = True
        End If
    Else
        Resul = False
    End If
    LeeParamsTIenda = Resul
End Function
Public Sub Listar_Las_Tiendas()
    Aux_Ejecuta "listado_tiendas", "", "", D_TiendaActual(0).lScalePos
End Sub


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Clientes"
   ClientHeight    =   5325
   ClientLeft      =   8160
   ClientTop       =   1335
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox MaskDescuento 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CheckBox ChkFactura 
      Caption         =   "Simplified"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Listado"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   2640
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   11
      Top             =   4800
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Index           =   1
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Width           =   1332
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      MaxLength       =   255
      TabIndex        =   8
      Top             =   3360
      Width           =   4572
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   255
      TabIndex        =   7
      Top             =   1440
      Width           =   4572
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   255
      TabIndex        =   6
      Top             =   3120
      Width           =   4572
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   255
      TabIndex        =   5
      Top             =   2880
      Width           =   4572
   End
   Begin VB.TextBox TxtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   4572
   End
   Begin VB.ComboBox CmbCodigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2040
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Discount:"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "NIF"
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
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   6495
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   492
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   2172
   End
   Begin VB.Label LblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   2400
      TabIndex        =   12
      Top             =   0
      Width           =   3972
   End
   Begin VB.Label Label2 
      Caption         =   "Datos Adicionales"
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
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1692
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyCliente() As DB_Cliente
Private NClientes As Integer
'DROP TABLE IF EXISTS `pcscale`.`customers`;
'CREATE TABLE  `pcscale`.`customers` (
'  `Id` int(11) NOT NULL AUTO_INCREMENT,
'  `TIN` varchar(16) NOT NULL DEFAULT '',
'  `Name` varchar(255) NOT NULL DEFAULT '',
'  `Address` varchar(255) DEFAULT '',
'  `City` varchar(255) DEFAULT '',
'  `Postalcode` varchar(80) DEFAULT '',
'  `Country` varchar(80) DEFAULT '',
'  `WEB` varchar(255) DEFAULT '',
'  `Phone` varchar(80) DEFAULT '',
'  `Fax` varchar(80) DEFAULT '',
'  `MobilePhone` varchar(80) DEFAULT '',
'  `eMail` varchar(80) DEFAULT '',
'  `Text` varchar(255) DEFAULT '',
'  `Discount` decimal(4,2) DEFAULT '0.00',
'  `SimplifiedInvoice` tinyint(1) DEFAULT '0',
'  `TicketOrder` tinyint(1) DEFAULT '0',
'  `Points` int(11) NOT NULL DEFAULT '0',
'  PRIMARY KEY (`Id`),
'  UNIQUE KEY `TIN` (`TIN`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB AUTO_INCREMENT=2223 DEFAULT CHARSET=utf8;

Private Sub CmbCodigo_Click()
    LblInfo.Caption = ""
    Pon_Cliente
End Sub

Private Sub CmbCodigo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case 48 To 59
            
            If Len(CmbCodigo.TexT) >= 5 Then KeyAscii = 0
        Case Else
            If KeyAscii = 13 Then TxtNombre.SetFocus
            KeyAscii = 0
    End Select
End Sub

Private Sub CmbCodigo_LostFocus()
    Pon_Cliente
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Resp As Integer
    Dim bucle As Integer
    Dim MiCli As DB_Cliente
    Select Case Index
        Case 0
            If IsNumeric(CmbCodigo.TexT) Then
                MiCli.codigo = CmbCodigo.TexT
                MiCli.Nombre = TxtNombre.TexT
                MiCli.Datos(0) = TxtData(0).TexT 'nif
                MiCli.Datos(1) = TxtData(1).TexT
                MiCli.Datos(2) = TxtData(2).TexT
                MiCli.Datos(3) = TxtData(3).TexT
                If MaskDescuento.TexT <> "" Then
                    MiCli.discount = MaskDescuento.TexT
                Else
                    MiCli.discount = 0
                End If
                If ChkFactura.Value = vbChecked Then
                    MiCli.factura = True
                Else
                    MiCli.factura = False
                End If
                Resp = Alta_Clientes(MiCli)
            End If
        Case 1
            If IsNumeric(CmbCodigo.TexT) Then
                Resp = Baja_Clientes(CmbCodigo.TexT)
            End If
        Case 2
            Unload Me
    End Select
    If Index = 0 Or Index = 1 Then
        Select Case Resp
            Case 0
                
                LblInfo.Caption = CargaCadena(527)
                'Rellena_Datos
            Case 1
                
                If Index = 0 Then
                    LblInfo.Caption = CargaCadena(439)
                Else
                    LblInfo.Caption = CargaCadena(529)
                    TxtNombre.TexT = ""
                    For bucle = 0 To 3
                        TxtData(bucle).TexT = ""
                    Next bucle
                    ChkFactura.Value = vbUnchecked
                    MaskDescuento.TexT = ""
                    'Rellena_Datos
                End If
        End Select
    End If
    Rellena_Datos
End Sub

Private Sub Command2_Click()
    Aux_Ejecuta "rclientes"
End Sub

Private Sub Form_Load()
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    CambiarIdioma
    Rellena_Datos
End Sub
Private Sub CambiarIdioma()
    Me.Caption = CargaCadena(1056)
    Command2.Caption = CargaCadena(490)
    LblInfo2.Caption = CargaCadena(524)
    Command1(0).Caption = CargaCadena(550)
    Command1(1).Caption = CargaCadena(38)
    Command1(2).Caption = CargaCadena(7)
    Label2(1).Caption = CargaCadena(552)
    Label2(0).Caption = CargaCadena(69)
    Label2(2).Caption = CargaCadena(1057)
    Label1.Caption = CargaCadena(1407)
End Sub
Private Sub Pon_Cliente()
    Dim bucle As Integer
    Dim B2 As Integer
    Dim eNCONTRADO As Boolean
    If IsNumeric(CmbCodigo.TexT) Then
        eNCONTRADO = False
        For bucle = 0 To NClientes - 1
            If MyCliente(bucle).codigo = Val(CmbCodigo.TexT) Then
                eNCONTRADO = True
                    TxtNombre.TexT = MyCliente(bucle).Nombre
                    For B2 = 0 To 3
                        TxtData(B2).TexT = MyCliente(bucle).Datos(B2)
                    Next B2
                    If MyCliente(bucle).factura Then
                        ChkFactura.Value = vbChecked
                    Else
                        ChkFactura.Value = vbUnchecked
                    End If
                    MaskDescuento.TexT = MyCliente(bucle).discount
                Exit For
            End If
        Next bucle
    Else
        eNCONTRADO = False
    End If
    If Not eNCONTRADO Then
        TxtData(2).Enabled = True
        LblInfo2.Caption = CargaCadena(524) '"Nuevo Dato"
        LblInfo2.BackColor = vbWhite
        TxtNombre.TexT = ""
        For bucle = 0 To 3
            TxtData(bucle).TexT = ""
        Next bucle
        ChkFactura.Value = vbUnchecked
        MaskDescuento.TexT = ""
        
    Else
        LblInfo2.Caption = CargaCadena(525) '"Modificación"
        LblInfo2.BackColor = vbYellow
        TxtData(2).Enabled = False
    End If
End Sub
Private Sub Rellena_Datos()
    
    Dim Registro As New RecordNet
    NClientes = 0
    CmbCodigo.Clear
    ReDim MyCliente(0)

    Registro.OpenRecordset ("select * from cliente where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                NClientes = NClientes + 1
                ReDim Preserve MyCliente(NClientes)
                CmbCodigo.AddItem .Fields("cod_cli")
                MyCliente(NClientes - 1).codigo = .Fields("cod_cli")
                MyCliente(NClientes - 1).Nombre = .Fields("nom_cli")
                MyCliente(NClientes - 1).Datos(0) = .Fields("Data1")
                MyCliente(NClientes - 1).Datos(1) = .Fields("Data2")
                MyCliente(NClientes - 1).Datos(2) = .Fields("Data3")
                MyCliente(NClientes - 1).Datos(3) = .Fields("Data4")
                MyCliente(NClientes - 1).factura = .Fields("factura")
                MyCliente(NClientes - 1).discount = .Fields("discount")
                .Movenext
            Loop
        End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub TxtData_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 3
            Select Case KeyAscii
                Case 13
                    KeyAscii = 0
                    Command1(0).SetFocus
                Case Else
            End Select
        Case Else
            Select Case KeyAscii
                Case 13
                    KeyAscii = 0
                    TxtData(Index + 1).SetFocus
                Case Else
            End Select
    End Select
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            TxtData(0).SetFocus
        Case Else
    End Select
End Sub

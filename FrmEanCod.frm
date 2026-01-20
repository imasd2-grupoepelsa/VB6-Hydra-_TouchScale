VERSION 5.00
Begin VB.Form FrmEanCod 
   ClientHeight    =   2595
   ClientLeft      =   6885
   ClientTop       =   1695
   ClientWidth     =   5805
   LinkTopic       =   "FrmEanCod"
   ScaleHeight     =   2595
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaxLength       =   30
      TabIndex        =   10
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   972
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   4212
      Begin VB.ComboBox CmbTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1212
      End
      Begin VB.ComboBox CmbNumero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Mostrador o Equipo"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2412
      End
      Begin VB.Label LblNumero 
         Caption         =   "Mostrador"
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1452
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4680
      TabIndex        =   0
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "FrmEanCod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DROP TABLE IF EXISTS `mbd`.`BarcodeCode128`;
'CREATE TABLE  `mbd`.`BarcodeCode128` (
'  `Id` int(11) DEFAULT NULL,
'  `Type` varchar(1) DEFAULT '0',
'  `Number` int(11) DEFAULT '0',
'  `Code128` varchar(30) DEFAULT '',
'  UNIQUE KEY `Id` (`Id`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;

Private Type dB_EanCod
    Equipo As Integer
    Mostrador As Integer
    Cod As String
End Type

Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
Private Sub CambiarIdioma()
    FrmEanCod.Caption = CargaCadena(725) & " (Code 128)"
    
    LblNumero.Caption = CargaCadena(302)
    LblInfo.Caption = CargaCadena(524)
    
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
   
    CmbTipo.AddItem CargaCadena(302)
    CmbTipo.AddItem CargaCadena(303)
    
End Sub
Private Function Desglosa_Codigos(LaCadena As String) As TipoCodigos
    Dim LosCodigos As TipoCodigos
    Dim Buffer As String
    Dim bucle As Long
    Dim cantidad As Integer
    cantidad = 0
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            cantidad = cantidad + 1
            ReDim Preserve LosCodigos.codigo(cantidad + 1)
            LosCodigos.codigo(cantidad) = Buffer
            LosCodigos.cantidad = cantidad
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Desglosa_Codigos = LosCodigos
End Function

Private Sub Refresca_Datos()
    
    Dim Registro As New RecordNet
    Dim bucle As Integer
    Dim posicion As Integer
    Dim eNCONTRADO As Boolean
    Dim ElCodigo As String
    
    Text1.TexT = ""

    LblInfo.BackColor = vbWhite
    LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
    CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
    If (CmbNumero.TexT = "") And (CmbTipo.ListIndex < 2) Then Exit Sub
    eNCONTRADO = False
    
    Select Case CmbTipo.ListIndex
    Case 0
        bucle = 1
        ElCodigo = ""
        Do Until Mid(CmbNumero.TexT, bucle, 1) = " " Or bucle > Len(CmbNumero.TexT)
            ElCodigo = ElCodigo & Mid(CmbNumero.TexT, bucle, 1)
            bucle = bucle + 1
        Loop
        Registro.OpenRecordset _
        ("select * from gen_sam where baja=false and tipo='cbCod' and destino=1 and numero=" & ElCodigo)
     
    Case 1
        Registro.OpenRecordset _
        ("select * from gen_sam where baja=false and tipo='cbCod' and destino=2 and numero=" & CmbNumero.TexT)
        
    End Select
    With Registro
        If Not .EOF Then
            eNCONTRADO = True
            Text1.TexT = .Fields("txt01")
        End If
    End With
    

    If eNCONTRADO Then
        LblInfo.BackColor = vbYellow
        LblInfo.Caption = CargaCadena(525)  '"Modificación"
        CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
    End If
End Sub

Private Sub CmbNumero_Click()
    Refresca_Datos
End Sub

Private Sub CmbTipo_Click()
    Dim bucle As Integer
    If CmbTipo.ListIndex = 0 Then
        Aux_Refresca_Mostradores CmbNumero
        LblNumero.Caption = CargaCadena(302)  '"Mostrador"
    Else
        Aux_Refresca_equipos CmbNumero
        LblNumero.Caption = CargaCadena(303)  '"Equipo"
    End If
    Refresca_Datos
End Sub

Private Sub Cmdaceptar_Click()
    
    Dim Registro As New RecordNet
    Dim bucle As Integer
    Dim ElCodigo As TipoCodigos
    Dim MiCodigo As dB_EanCod
    Dim nR As Integer
    
    '******************
    ' Condiciones de Error
    If CmbTipo.ListIndex < 2 Then
        If CmbNumero.TexT = "" Then
            If CmbTipo.TexT = CmbTipo.List(0) Then
                MsgBox CargaCadena(559)  '"Introduzca el número de mostrador"
            Else
                MsgBox CargaCadena(575)  '"Introduzca el número de equipo"
            End If
            Exit Sub
        End If
    End If
    '********************
    
    Select Case CmbTipo.ListIndex
        Case 0
            ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
            MiCodigo.Mostrador = ElCodigo.codigo(1)
            MiCodigo.Equipo = 0
        Case 1
            ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
            Registro.OpenRecordset ("select codi_ident from equipos where borrado=false and numero_eqp=" & _
            CmbNumero.TexT)
            MiCodigo.Mostrador = Registro.Fields("codi_ident")
            MiCodigo.Equipo = CmbNumero.TexT
    End Select
    MiCodigo.Cod = Trim(Text1.TexT)
    nR = Alta_BarrasCod(MiCodigo)
    Select Case nR
        Case 0
            LblInfo2.Caption = CargaCadena(527)  '"Datos Añadidos. Envíe Modificaciones"
        Case 1
            LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados. Envíe Modificaciones"
    End Select
End Sub

Private Sub CmdBorrar_Click()
    
    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim bucle As Integer
    Dim BucleGrande As Integer
    Dim ElCodigo As TipoCodigos
    Dim MiCodigo As dB_EanCod
    
    If CmbNumero.TexT = "" Then Exit Sub
    bucle = 1
    ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
    If CmbTipo.TexT = CmbTipo.List(0) Then
        MiCodigo.Equipo = 0
        MiCodigo.Mostrador = ElCodigo.codigo(1)
    Else
        
        Registro.OpenRecordset ("select codi_ident from equipos where numero_eqp=" & _
        CmbNumero.TexT & " and borrado=false")
        MiCodigo.Mostrador = Registro.Fields("codi_ident")
        MiCodigo.Equipo = ElCodigo.codigo(1)
        
    End If
    If Baja_BarrasCod(MiCodigo) = 0 Then
        
        LblInfo2.Caption = CargaCadena(556)  '"Datos Eliminados.Envíe Modificaciones"
        
    End If
    If CmbTipo.TexT = CmbTipo.List(0) Then
        Aux_Refresca_Mostradores CmbNumero
    Else
        Aux_Refresca_equipos CmbNumero
    End If
    Refresca_Datos
    
End Sub



Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = Form2.Icon
    'Image1.Picture = Form2.Imagen.Picture
    CmbTipo.TexT = CmbTipo.List(0)
    Aux_Refresca_Mostradores CmbNumero
    If CmbTipo.ListIndex = 0 Or CmbTipo.ListIndex = 1 Then
        If CmbNumero.ListCount > 0 Then CmbNumero.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LockBase = False
    frmEpelsa.Enabled = True
End Sub

Private Function Alta_BarrasCod(MiCabecera As dB_EanCod) As Integer
Dim myRs As DAO.Recordset
Dim myB As DAO.Database
Dim Resp As Integer
Dim Dest As Integer
Dim numero As Integer
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cbCod' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cbCod' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If myRs.EOF Then
        myRs.AddNew
        Resp = 0
    Else
        myRs.Edit
        Resp = 1
    End If
    myRs.Fields("tipo") = "cbCod"
    myRs.Fields("destino") = Dest
    myRs.Fields("numero") = numero
    myRs.Fields("redlocal") = False
    myRs.Fields("baja") = False
    myRs.Fields("txt01") = Trim(Text1.TexT)
    myRs.Fields("tran_txt") = " "
    myRs.Update
    myRs.Close
    Set myRs = Nothing
    CerrarBase myB
    Set myB = Nothing
    
    Alta_BarrasCod = Resp
End Function

'Baja_Cabeceras
Private Function Baja_BarrasCod(MiCabecera As dB_EanCod) As Integer
Dim myRs As DAO.Recordset
Dim myB As DAO.Database
Dim Resp As Integer
Dim Dest As Integer
Dim numero As Integer
Dim cm As String
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cbCod' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cbCod' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If Not myRs.EOF Then
        myRs.Edit
        myRs.Fields("baja") = True
        myRs.Fields("tran_txt") = " "
        myRs.Update
    End If
    
    Baja_BarrasCod = Resp
End Function




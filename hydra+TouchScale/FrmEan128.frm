VERSION 5.00
Begin VB.Form FrmEan128 
   ClientHeight    =   8040
   ClientLeft      =   6885
   ClientTop       =   1695
   ClientWidth     =   5805
   LinkTopic       =   "Form3"
   ScaleHeight     =   8040
   ScaleWidth      =   5805
   Begin VB.CheckBox CheckIA 
      Caption         =   "IA 01"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Códigos"
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
      Height          =   6135
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   5772
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 3922"
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
         Index           =   22
         Left            =   120
         TabIndex        =   33
         Top             =   5640
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 950"
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
         Index           =   21
         Left            =   120
         TabIndex        =   32
         Top             =   5400
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 8005"
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
         Index           =   20
         Left            =   120
         TabIndex        =   31
         Top             =   5160
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 7031"
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
         Index           =   19
         Left            =   120
         TabIndex        =   30
         Top             =   4920
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 7030"
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
         Index           =   18
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 426"
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
         Index           =   17
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 423"
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
         Index           =   16
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 422"
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
         Index           =   15
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 414"
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
         Index           =   14
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 400"
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
         Index           =   13
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 330"
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
         Index           =   12
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 310"
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
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 30"
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
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 251"
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
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 21"
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
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 17"
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
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 15"
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
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 13"
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
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 11"
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
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 10"
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
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 02"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   4095
      End
      Begin VB.CheckBox CheckIA 
         Caption         =   "IA 00"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   972
      Left            =   1560
      TabIndex        =   3
      Top             =   240
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
         ItemData        =   "FrmEan128.frx":0000
         Left            =   2880
         List            =   "FrmEan128.frx":0002
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
      Top             =   7440
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
      Top             =   7440
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
      Top             =   7440
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
      Top             =   360
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "FrmEan128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type dB_Ean128
    Equipo As Integer
    Mostrador As Integer
    Ia(23) As String
End Type

Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
Private Sub CambiarIdioma()
    Me.Caption = CargaCadena(725) & " (Ean128)"
    Frame2.Caption = " IA's "
    Label2.Caption = CargaCadena(720)
    
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
    For bucle = 0 To 22
        CheckIA(bucle).Value = vbUnchecked
    Next bucle
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
        ("select * from gen_sam where baja=false and tipo='cb128' and destino=1 and numero=" & ElCodigo)
     
    Case 1
        Registro.OpenRecordset _
        ("select * from gen_sam where baja=false and tipo='cb128' and destino=2 and numero=" & CmbNumero.TexT)
        
    End Select
    With Registro
        If Not .EOF Then
            eNCONTRADO = True
            CheckIA(0).Value = .Fields("txt01")
            CheckIA(1).Value = .Fields("txt02")
            CheckIA(2).Value = .Fields("txt03")
            CheckIA(3).Value = .Fields("txt04")
            CheckIA(4).Value = .Fields("txt05")
            CheckIA(5).Value = .Fields("txt06")
            CheckIA(6).Value = .Fields("txt07")
            CheckIA(7).Value = .Fields("txt08")
            CheckIA(8).Value = .Fields("txt09")
            CheckIA(9).Value = .Fields("txt10")
            CheckIA(10).Value = .Fields("txt11")
            CheckIA(11).Value = .Fields("txt12")
            CheckIA(12).Value = .Fields("txt13")
            CheckIA(13).Value = .Fields("txt14")
            CheckIA(14).Value = .Fields("txt15")
            CheckIA(15).Value = .Fields("txt16")
            CheckIA(16).Value = .Fields("txt17")
            CheckIA(17).Value = .Fields("txt18")
            CheckIA(18).Value = .Fields("txt19")
            CheckIA(19).Value = .Fields("txt20")
            CheckIA(20).Value = .Fields("txt21")
            CheckIA(21).Value = .Fields("txt22")
            CheckIA(22).Value = .Fields("txt23")
            
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
    Dim MiCodigo As dB_Ean128
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
    For bucle = 0 To 22
        MiCodigo.Ia(bucle) = CheckIA(bucle).Value
    Next bucle
    nR = Alta_Barras128(MiCodigo)
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
    Dim MiCodigo As dB_Ean128
    
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
    If Baja_Barras128(MiCodigo) = 0 Then
        
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

Private Function Alta_Barras128(MiCabecera As dB_Ean128) As Integer
Dim myRs As DAO.Recordset
Dim myB As DAO.Database
Dim Resp As Integer
Dim Dest As Integer
Dim numero As Integer
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb128' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb128' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
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
    myRs.Fields("tipo") = "cb128"
    myRs.Fields("destino") = Dest
    myRs.Fields("numero") = numero
    myRs.Fields("redlocal") = False
    myRs.Fields("baja") = False
    myRs.Fields("txt01") = CheckIA(0).Value
    myRs.Fields("txt02") = CheckIA(1).Value
    myRs.Fields("txt03") = CheckIA(2).Value
    myRs.Fields("txt04") = CheckIA(3).Value
    myRs.Fields("txt05") = CheckIA(4).Value
    myRs.Fields("txt06") = CheckIA(5).Value
    myRs.Fields("txt07") = CheckIA(6).Value
    myRs.Fields("txt08") = CheckIA(7).Value
    myRs.Fields("txt09") = CheckIA(8).Value
    myRs.Fields("txt10") = CheckIA(9).Value
    myRs.Fields("txt11") = CheckIA(10).Value
    myRs.Fields("txt12") = CheckIA(11).Value
    myRs.Fields("txt13") = CheckIA(12).Value
    myRs.Fields("txt14") = CheckIA(13).Value
    myRs.Fields("txt15") = CheckIA(14).Value
    myRs.Fields("txt16") = CheckIA(15).Value
    myRs.Fields("txt17") = CheckIA(16).Value
    myRs.Fields("txt18") = CheckIA(17).Value
    myRs.Fields("txt19") = CheckIA(18).Value
    myRs.Fields("txt20") = CheckIA(19).Value
    myRs.Fields("txt21") = CheckIA(20).Value
    myRs.Fields("txt22") = CheckIA(21).Value
    myRs.Fields("txt23") = CheckIA(22).Value
    myRs.Fields("tran_txt") = " "
    myRs.Update
    myRs.Close
    Set myRs = Nothing
    CerrarBase myB
    Set myB = Nothing
    
    Alta_Barras128 = Resp
End Function

'Baja_Cabeceras
Private Function Baja_Barras128(MiCabecera As dB_Ean128) As Integer
Dim myRs As DAO.Recordset
Dim myB As DAO.Database
Dim Resp As Integer
Dim Dest As Integer
Dim numero As Integer
Dim cm As String
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb128' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb128' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If Not myRs.EOF Then
        myRs.Edit
        myRs.Fields("baja") = True
        myRs.Fields("tran_txt") = " "
        myRs.Update
    End If
    
    Baja_Barras128 = Resp
End Function




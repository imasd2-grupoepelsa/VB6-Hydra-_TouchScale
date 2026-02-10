VERSION 5.00
Begin VB.Form FrmTextoLibre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Textos Libres"
   ClientHeight    =   4230
   ClientLeft      =   6900
   ClientTop       =   2115
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAdicionales 
      Caption         =   "Adicionales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   41
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton CmdPegar 
      BackColor       =   &H80000018&
      Caption         =   "&Pegar"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   240
      Width           =   1092
   End
   Begin VB.CommandButton CmdCopiar 
      BackColor       =   &H80000018&
      Caption         =   "&Copiar"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   240
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
      Height          =   372
      Left            =   5760
      TabIndex        =   25
      Top             =   3720
      Width           =   1212
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
      Height          =   372
      Left            =   5760
      TabIndex        =   24
      Top             =   2160
      Width           =   1212
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
      Height          =   372
      Left            =   5760
      TabIndex        =   23
      Top             =   1680
      Width           =   1212
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   2280
      TabIndex        =   22
      Top             =   3840
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   2280
      TabIndex        =   21
      Top             =   3600
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   2280
      TabIndex        =   20
      Top             =   3360
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   2280
      TabIndex        =   19
      Top             =   3120
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   2280
      TabIndex        =   18
      Top             =   2880
      Width           =   3252
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3840
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3600
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3360
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3120
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   732
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   2280
      TabIndex        =   17
      Top             =   2640
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Top             =   2400
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   2280
      TabIndex        =   15
      Top             =   2160
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Top             =   1920
      Width           =   3252
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   2280
      TabIndex        =   13
      Top             =   1680
      Width           =   3252
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2160
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   732
   End
   Begin VB.ComboBox CmbFuente 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   732
   End
   Begin VB.ComboBox CmbEquipo 
      Height          =   288
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.ComboBox CmbSeccion 
      Height          =   288
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1332
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Equipo"
      Height          =   252
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   1452
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mostrador"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H80000018&
      Height          =   252
      Left            =   3120
      TabIndex        =   40
      Top             =   720
      Width           =   3852
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3852
   End
   Begin VB.Label Label2 
      Caption         =   "5"
      Height          =   252
      Index           =   9
      Left            =   120
      TabIndex        =   37
      Top             =   3840
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "4"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   36
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "3"
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   35
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   33
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "5"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "4"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "3"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   2052
   End
End
Attribute VB_Name = "FrmTextoLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Portapapeles
    texto(10) As String
    Tipo(10) As String
End Type
Private Porta As Portapapeles
Private UsaSeccion As Boolean
Private Modificacion As Boolean
Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
Private QuantumLeap As Boolean
Private Sub CambiarIdioma()
    Dim bucle As Integer
    'FrmTextoLibre.Caption = CargaCadena( 716)
    'Label1.Caption = CargaCadena( 716)
    Me.Caption = CargaCadena(919)

    LblInfo.Caption = CargaCadena(524)
    CmdCopiar.Caption = CargaCadena(717)
    CmdPegar.Caption = CargaCadena(718)
    Option1.Caption = CargaCadena(302)
    Option2.Caption = CargaCadena(303)
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
    'CmdAdicionales.Caption = CargaCadena(1351)
    For bucle = 0 To 9
        Label2(bucle).Caption = CargaCadena(539) & " " & (bucle + 1)    'CargaCadena( 317) & " " & (Bucle + 1)
    Next bucle
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


Private Sub CmbEquipo_Click()
    Refresca_Datos
End Sub



'Private Sub CmbFuente_Click(Index As Integer)
'    Dim a As Integer
'    Dim Max As Integer
'    a = CmbFuente(Index).TexT
'    If a = 0 Or a = 1 Or a = 2 Or a = 3 Or a = 16 Or a = 17 Or a = 18 Or a = 19 Then
'        Max = 25
'    Else
'        If a = 8 Or a = 9 Or a = 14 Or a = 15 Or a = 24 Or a = 25 Or a = 30 Or a = 31 _
         '        Or a = 36 Or a = 37 Or a = 42 Or a = 43 Or a = 48 Or a = 49 Or a = 54 Or a = 55 _
         '        Or a = 60 Or a = 61 Or a = 66 Or a = 67 Or a = 72 Or a = 73 Then
'            Max = 12
'        Else
'            If a = 6 Or a = 7 Or a = 12 Or a = 13 Or a = 22 Or a = 23 Or a = 28 Or a = 29 _
             '            Or a = 34 Or a = 35 Or a = 40 Or a = 41 Or a = 46 Or a = 47 Or a = 52 Or a = 53 _
             '            Or a = 58 Or a = 59 Or a = 64 Or a = 65 Or a = 70 Or a = 71 Or a = 76 Or a = 77 Then
'                Max = 16
'            Else
'                Max = 24
'            End If
'        End If
'    End If
'    If Len(TxtTexto(Index).TexT) > Max Then
'        TxtTexto(Index).TexT = left(TxtTexto(Index).TexT, Max)
'    End If
'    TxtTexto(Index).MaxLength = Max
'End Sub

Private Sub CmbSeccion_Click()
    Refresca_Datos
End Sub

Private Sub Cmdaceptar_Click()

    Dim Registro As New RecordNet
    Dim ElNumero As TipoCodigos
    Dim bucle As Integer
    Dim MiCabecera As DB_Cabeceras
    If (CmbEquipo.TexT = "" And Not UsaSeccion) Or (CmbSeccion.TexT = "" And UsaSeccion) Then Exit Sub
    'If Modificacion Then
    If UsaSeccion Then
        ElNumero = Desglosa_Codigos(CmbSeccion.TexT)
        MiCabecera.Mostrador = ElNumero.codigo(1)
        MiCabecera.Equipo = 0
    Else

        Registro.OpenRecordset ("select * from equipos where numero_eqp=" _
                              & CmbEquipo.TexT & " and borrado=false")
        ElNumero = Desglosa_Codigos(CmbEquipo.TexT)
        MiCabecera.Mostrador = Registro.Fields("codi_ident")

        MiCabecera.Equipo = ElNumero.codigo(1)
    End If
    'For bucle = 0 To 4
    '    MiCabecera.Cabecera(bucle) = TxtTexto(bucle).TexT
    '    MiCabecera.TLCabecera(bucle) = CmbFuente(bucle).TexT
    'Next bucle
    'For bucle = 0 To 4
    '    MiCabecera.leyenda(bucle) = TxtTexto(bucle + 5).TexT
    '    MiCabecera.TLLeyenda(bucle) = CmbFuente(bucle + 5).TexT
    'Next bucle
    Select Case Alta_cabeceras(MiCabecera)
    Case 1
        LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados.Envíe Modificaciones."
    Case 0
        LblInfo2.Caption = CargaCadena(527)  '"Datos Agregados.Envíe Modificaciones."
    End Select
    Refresca_Datos
    If UsaSeccion Then
        CmbSeccion.SetFocus
    Else
        CmbEquipo.SetFocus
    End If
End Sub

Private Function Alta_cabeceras(MiCabecera As DB_Cabeceras) As Integer
    Dim myRs As dao.Recordset
    Dim myB As dao.Database
    Dim Resp As Integer
    Dim Dest As Integer
    Dim numero As Integer
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='free' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='free' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
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
    myRs.Fields("tipo") = "free"
    myRs.Fields("destino") = Dest
    myRs.Fields("numero") = numero
    myRs.Fields("redlocal") = False
    myRs.Fields("baja") = False
    myRs.Fields("txt01") = Trim(TxtTexto(0))
    myRs.Fields("txt02") = Trim(TxtTexto(1))
    myRs.Fields("txt03") = Trim(TxtTexto(2))
    myRs.Fields("txt04") = Trim(TxtTexto(3))
    myRs.Fields("txt05") = Trim(TxtTexto(4))
    myRs.Fields("txt06") = Trim(TxtTexto(5))
    myRs.Fields("txt07") = Trim(TxtTexto(6))
    myRs.Fields("txt08") = Trim(TxtTexto(7))
    myRs.Fields("txt09") = Trim(TxtTexto(8))
    myRs.Fields("txt10") = Trim(TxtTexto(9))
    myRs.Fields("tran_txt") = " "
    myRs.Update
    myRs.Close
    Set myRs = Nothing
    CerrarBase myB
    Set myB = Nothing

    Alta_cabeceras = Resp
End Function

'Baja_Cabeceras
Private Function Baja_cabeceras(MiCabecera As DB_Cabeceras) As Integer
    Dim myRs As dao.Recordset
    Dim myB As dao.Database
    Dim Resp As Integer
    Dim Dest As Integer
    Dim numero As Integer
    Dim cm As String
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='free' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='free' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If Not myRs.EOF Then
        myRs.Edit
        myRs.Fields("baja") = True
        myRs.Fields("tran_txt") = " "
        myRs.Update
    End If

    Baja_cabeceras = Resp
End Function

'Private Sub CmdAdicionales_Click()
'    FrmTextoLibre.Enabled = False
'    FrmAdicionales.Show
'
'End Sub

Private Sub CmdBorrar_Click()

    Dim Registro As New RecordNet
    Dim ElCodigo As TipoCodigos
    Dim MiCabecera As DB_Cabeceras
    If UsaSeccion = True Then
        If CmbSeccion.TexT = "" Then Exit Sub
    Else
        If CmbEquipo.TexT = "" Then Exit Sub
    End If
    If UsaSeccion Then
        ElCodigo = Desglosa_Codigos(CmbSeccion.TexT)
        MiCabecera.Mostrador = ElCodigo.codigo(1)
        MiCabecera.Equipo = 0
    Else

        Registro.OpenRecordset ("select * from equipos where borrado=false and numero_eqp=" & _
                                CmbEquipo.TexT)
        MiCabecera.Mostrador = Registro.Fields("codi_ident")

        MiCabecera.Equipo = CmbEquipo.TexT
    End If
    If Baja_cabeceras(MiCabecera) = 0 Then
        LblInfo2.Caption = CargaCadena(556)  '"Datos Eliminados.Envíe Modificaciones"
    End If
    RefrescaSeccion
    RefrescaEquipo
    CmbSeccion.TexT = CmbSeccion.List(0)
    CmbEquipo.TexT = CmbEquipo.List(0)
    Refresca_Datos

    If UsaSeccion Then
        CmbSeccion.SetFocus
    Else
        CmbEquipo.SetFocus
    End If
End Sub

Private Sub CmdCopiar_Click()
    Dim bucle As Integer
    For bucle = 0 To 9
        Porta.texto(bucle) = TxtTexto(bucle).TexT
        'Porta.Tipo(bucle) = CmbFuente(bucle).TexT
    Next bucle
End Sub

Private Sub cmdpegar_click()
    Dim bucle As Integer
    For bucle = 0 To 9
        TxtTexto(bucle).TexT = Porta.texto(bucle)
        'CmbFuente(bucle).TexT = Porta.Tipo(bucle)
    Next bucle
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Dim bucle As Integer
    Dim Bucle2 As Integer

    CambiarIdioma
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    QuantumLeap = False
    Option1.Value = True
    CmbSeccion.Enabled = True
    CmbEquipo.Enabled = False
    UsaSeccion = True
    Modificacion = False
    'If lNewArt = False Then
    CmdAdicionales.Enabled = False
    CmdAdicionales.Visible = False
    'End If
    For bucle = 0 To 9
        CmbFuente(bucle).Clear
        CmbFuente(bucle).Enabled = False
        CmbFuente(bucle).Visible = False
    Next bucle
    'For bucle = 0 To 9
    '    For Bucle2 = 0 To 77
    '        CmbFuente(bucle).AddItem Bucle2
    '    Next Bucle2
    'Next bucle
    'For bucle = 0 To 9
    '    CmbFuente(bucle).TexT = CmbFuente(bucle).List(0)
    'Next bucle
    RefrescaSeccion
    RefrescaEquipo
End Sub
Private Sub RefrescaSeccion()
    Aux_Refresca_Mostradores CmbSeccion
End Sub
Private Sub RefrescaEquipo()
    Aux_Refresca_equipos CmbEquipo


End Sub

Private Sub Form_Unload(Cancel As Integer)

    LockBase = False
    frmEpelsa.Enabled = True

End Sub
Private Sub deborefrescar()
    If UsaSeccion = True Then
        If CmbSeccion.TexT <> "" Then
            Refresca_Datos
        End If
    Else
        If CmbEquipo.TexT <> "" Then
            Refresca_Datos
        End If
    End If

End Sub
Private Sub Option1_Click()
    If Option1.Value = True Then
        UsaSeccion = True
        CmbEquipo.Enabled = False
        CmbSeccion.Enabled = True
    Else
        UsaSeccion = False
        CmbEquipo.Enabled = True
        CmbSeccion.Enabled = False
    End If
    Call deborefrescar
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        UsaSeccion = False
        CmbEquipo.Enabled = True
        CmbSeccion.Enabled = False
    Else
        UsaSeccion = True
        CmbEquipo.Enabled = False
        CmbSeccion.Enabled = True
    End If
    Call deborefrescar
End Sub
Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 4 And TxtTexto(5).Enabled = False Then
            TxtTexto(7).SetFocus
            Exit Sub
        Else
            If Index < 9 Then
                TxtTexto(Index + 1).SetFocus
                Exit Sub
            Else
                CmdAceptar.SetFocus
            End If
        End If
    End If
    TxtTexto(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub
    If Len(TxtTexto(Index).TexT) = 25 Then
        TxtTexto(Index).Locked = True
    End If
End Sub
'Private Sub Actualiza_Fuentes(Numero As Integer)
'    Dim Bucle1 As Long
'    Dim Bucle2 As Long
'    For Bucle1 = 0 To 9
'        CmbFuente(Bucle1).Clear
'        For Bucle2 = 0 To Numero
'            CmbFuente(Bucle1).AddItem Bucle2
'        Next Bucle2
'        CmbFuente(Bucle1).TexT = CmbFuente(Bucle1).List(0)
'    Next Bucle1
'End Sub
Private Sub Refresca_Datos()

    Dim Registro As New RecordNet
    Dim RegSec As New RecordNet
    Dim RegBal As New RecordNet
    Dim bucle As Integer
    Dim Bucle1 As Integer
    Dim Bucle2 As Integer
    Dim Limite As Boolean
    Dim ElMostrador As TipoCodigos
    Dim B2 As Integer

    'Actualiza_Fuentes 77
    ElMostrador = Desglosa_Codigos(CmbSeccion.TexT)
    If UsaSeccion Then
        Registro.OpenRecordset _
                ("select * from gen_sam where tipo='free' and destino=1 and numero=" & ElMostrador.codigo(1) & " and baja=FALSE")
    Else
        Registro.OpenRecordset _
                ("select * from gen_sam where tipo='free' and destino=2 and numero=" & CmbEquipo.TexT & " and baja=FALSE")
    End If
    With Registro
        For bucle = 0 To 9
            TxtTexto(bucle).TexT = ""
            'CmbFuente(bucle).TexT = CmbFuente(bucle).List(0)
        Next bucle
        If .EOF Then
            LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
            LblInfo.BackColor = vbWhite
            CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
            Modificacion = False
            '**************
            ' fuentes de letra
            '**************
            'Actualiza_Fuentes 77
            '**************
            ' /fuentes de letra
            '**************
        Else
            LblInfo.Caption = CargaCadena(525)  '"Modificación"
            LblInfo.BackColor = vbYellow
            CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
            Modificacion = True
            'Actualiza_Fuentes 77
            '**************
            ' /fuentes de letra
            '**************
            '*****************
            For B2 = 1 To 10
                If .Fields("txt" & Format(B2, "00")) <> "" Then
                    TxtTexto(B2 - 1).TexT = Trim(.Fields("txt" & Format(B2, "00")))
                End If
            Next B2
            'For B2 = 0 To 9
            '    If .Fields("tletra" & Format(B2, "00")) <> "" Then
            '        CmbFuente(B2).TexT = Convertir(.Fields("tletra" & Format(B2, "00")), False)
            '    End If
            'Next B2
            'For B2 = 10 To 49
            '    FrmAdicionales.TxtA(B2 - 10).TexT = Trim(.Fields("texto" & Format(B2, "00")))
            'Next B2

        End If
    End With
End Sub
Private Function convertir1(cadena As String) As String
    If Len(cadena) = 1 Then
        convertir1 = cadena
        Exit Function
    End If
    convertir1 = Chr(Val(cadena) + 48)
End Function
Private Function Convertir(cadena As String, Limite As Boolean) As String
    Dim MiValor As Integer
    MiValor = Val(Asc(cadena) - 48)
    If Limite And MiValor > 3 Then MiValor = 3
    Convertir = MiValor
End Function

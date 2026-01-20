VERSION 5.00
Begin VB.Form FrmFamilias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   6795
   ClientLeft      =   3420
   ClientTop       =   2355
   ClientWidth     =   5640
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5640
   Begin VB.TextBox TxtLabel2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   19
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox TxtLabel1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   17
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox ChkTraza 
      Caption         =   "Trazabilidad"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TxtOrden 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Asignación Familias a Mostradores"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Vacuno"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   3600
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3600
      Width           =   1452
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
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   2652
   End
   Begin VB.ComboBox CmbFamilias 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label9 
      Caption         =   "Tipo de IVA a programar en artículos de la Familia"
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
      Left            =   960
      TabIndex        =   21
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label8 
      Caption         =   "Etiquetas para artículos de la Familia"
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
      Left            =   960
      TabIndex        =   20
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Etq. 2"
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
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Etq. 1"
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
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "TVA:"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image picture1 
      Height          =   2655
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Orden:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "FrmFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const BALSAM = False

Private YaExiste As Boolean
Private nEvents As Long
Private nMinPos As Integer
Private WithEvents MyEvento As ClsEvento
Attribute MyEvento.VB_VarHelpID = -1

Private Sub Cambiar_idioma()
    Me.Caption = CargaCadena(678)
    Label1.Caption = CargaCadena(67)
    Label2.Caption = CargaCadena(539)
    Label3.Caption = "Etq.Fam:" 'CargaCadena(623)
    Label4.Caption = "Orden:"
    Command1.Caption = CargaCadena(287)
    Command4.Caption = "Asignación Familias a Mostradores"
    Command2.Caption = CargaCadena(7)
    Command3.Caption = CargaCadena(38)
    Command4.Caption = CargaCadena(1459)
    Check1.Caption = CargaCadena(676)
    If Dir(App.Path & "\test.tst") = "" Then
        Command4.Visible = False
    End If
#If BALSAM = False Then '''ver este
    ChkTraza.Visible = True
#End If

End Sub
Private Sub CmbFamilias_Click()
    nQueImagen = 0
    sQueNombreImagen = ""
    Refresca_Datos
End Sub
Private Sub Refresca_Datos()
    Dim Registro As New RecordNet
    Dim Token As Long
    Text1.TexT = ""
    Text2.TexT = ""
    TxtOrden.TexT = ""
    Check1.Value = vbUnchecked
    ChkTraza.Value = vbUnchecked
    
    Set picture1.Picture = Nothing
    Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
    'nQueImagen = 0
    'sQueNombreImagen = ""

    Registro.OpenRecordset "select max(index) as nmaxi from fam_code"
    If Not Registro.EOF Then
        If Not IsNull(Registro.Fields("nmaxi")) Then
            nMinPos = Registro.Fields("nmaxi")
        Else
            nMinPos = 0
        End If
    Else
        nMinPos = 0
    End If
    
    If IsNumeric(Val(CmbFamilias.TexT)) Then
        Registro.OpenRecordset "select * from fam_code where codi_fam=" & Val(CmbFamilias.TexT)
        If Not Registro.EOF Then
            If Not IsNull(Registro.Fields("txt_fam")) Then
                Text1.TexT = Registro.Fields("txt_fam")
            Else
                Text1.TexT = "Familia " & CStr(Val(CmbFamilias.TexT))
            End If
            If Not IsNull(Registro.Fields("etiqueta")) Then
                Text2.TexT = Trim(Registro.Fields("etiqueta"))
            Else
                Text2.TexT = ""
            End If
            If Not IsNull(Registro.Fields("index")) Then
                TxtOrden.TexT = CStr(Registro.Fields("index"))
            Else
                TxtOrden.TexT = CStr(nMinPos + 1)
            End If
            If Not IsNull(Registro.Fields("tax")) Then
                Text3.TexT = CStr(Registro.Fields("tax"))
            Else
                Text3.TexT = 0
            End If
            If Not IsNull(Registro.Fields("label1")) Then
                TxtLabel1.TexT = CStr(Registro.Fields("label1"))
            Else
                TxtLabel1.TexT = ""
            End If
            If Not IsNull(Registro.Fields("label2")) Then
                TxtLabel2.TexT = CStr(Registro.Fields("label2"))
            Else
                TxtLabel2.TexT = ""
            End If
            
            If Registro.Fields("posicion") <> 0 Then
                Check1.Value = vbChecked
            Else
                Check1.Value = vbUnchecked
            End If
            If Registro.Fields("istraza") Then
                ChkTraza.Value = vbChecked
            Else
                ChkTraza.Value = vbUnchecked
            End If
            
            On Error Resume Next
            If sQueNombreImagen = "" Then
                If Trim(Registro.Fields("imagen")) <> "" Then
                    If Dir(App.Path & "\images\" & Trim(Registro.Fields("imagen"))) <> "" Then
                        Token = InitGDIPlus
                        Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(Registro.Fields("imagen")))
                        FreeGDIPlus Token
                        sQueNombreImagen = Trim(Registro.Fields("imagen"))
                    Else
                        Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
                    End If
                Else
                    Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
                End If
            Else
                Token = InitGDIPlus
                Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\images\" & Trim(sQueNombreImagen))
                FreeGDIPlus Token
            End If
            On Error GoTo 0
            Command3.Enabled = True
        Else
            Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
            TxtOrden.TexT = CStr(nMinPos + 1)
        End If
    End If
End Sub

Private Sub Picture1_Click()
Dim Token As Long
    nQueImagen = 0
    sQueNombreImagen = ""
    If Val(CmbFamilias.TexT) <> 0 Then
        Load FrmImagen
        FrmImagen.Show
        Do While FrmImagen.Visible = True
            Do_Events
        Loop
        On Error Resume Next
        Token = InitGDIPlus
        Set picture1.Picture = LoadPictureGDIPlus(sQueNombreImagen)
        FreeGDIPlus Token '
        If Err.Number <> 0 Then
            On Error GoTo 0
        End If
        On Error GoTo 0
    Else
        Token = InitGDIPlus
        Set picture1.Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
        FreeGDIPlus Token '
    End If
End Sub

Private Sub CmbFamilias_GotFocus()
    nQueImagen = 0
    'sQueNombreImagen = ""
    Refresca_Datos
End Sub

Private Sub CmbFamilias_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyBack
        Case vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub CmbFamilias_LostFocus()
    nQueImagen = 0
    'sQueNombreImagen = ""
    
    Refresca_Datos
End Sub

Private Function Mensaje_Advertencia(ByVal nSecs As Long, secs() As Long) As Boolean
    Dim Registro As New RecordNet
    Dim Resp As Boolean
    Dim nResp As Long
    Resp = True
    Registro.OpenRecordset ("select count(codigo) from articulo where borrado=false and codi_fam=" & Val(CmbFamilias.TexT))
    If Registro.Fields(1) > 0 Then
        nResp = MsgBox(CargaCadena(1206), vbYesNo + vbExclamation + vbDefaultButton2)
        If nResp = vbNo Then Resp = False
    End If
    Mensaje_Advertencia = Resp
End Function
Private Sub Command1_Click()
    Dim bucle As Long
    Dim Resp As Long
    Dim lTraza As Boolean
    Dim mbase As DAO.Database
    
    Me.Enabled = False
    If Not IsNumeric(CmbFamilias.TexT) Then
        MsgBox CargaCadena(1205), vbCritical
        Me.Enabled = True
        Exit Sub
    End If
    If ChkTraza.Value = vbChecked Then
        lTraza = True
    Else
        lTraza = False
    End If
    Resp = Alta_Familias(Val(CmbFamilias.TexT), Text1.TexT, Check1.Value, Text2.TexT, False, MyEvento, Val(TxtOrden.TexT), separaFich(Trim(sQueNombreImagen)), lTraza, Val(Text3.TexT), Trim(TxtLabel1.TexT), Trim(TxtLabel2.TexT))
    
    If Len(Trim(TxtLabel1.TexT)) <> 0 Or Len(Trim(TxtLabel2.TexT)) <> 0 Then
        Set mbase = AbrirBase
        mbase.Execute "update articulo set tran_plu=' ' where codi_fam=" & CStr(Val(CmbFamilias.TexT))
        mbase.Close
        Set mbase = Nothing
    End If
    
    Select Case Resp
        Case 0
        '    MsgBox CargaCadena(421), vbInformation
            Refresca_Familias
            Refresca_Datos
        Case 1
            MsgBox CargaCadena(751), vbInformation
        Case 4
        Case 61
            MsgBox CargaCadena(1207), vbInformation
    End Select
    Me.Enabled = True
    If FrmFamMos.Visible Then
        FrmFamMos.Refresca_Familias
    End If
    If CmbFamilias.Enabled = True Then CmbFamilias.SetFocus
End Sub

Private Sub Command2_Click()
    If FrmFamMos.Visible Then
        Unload FrmFamMos
    End If
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim Registro As New RecordNet
    Dim Resp As Long
     
    If Not IsNumeric(CmbFamilias.TexT) Then Exit Sub
    Me.Enabled = False
    Registro.OpenRecordset ("select * from articulo where borrado=false and codi_fam=" & Val(CmbFamilias.TexT))
    If Not Registro.EOF Then
        Resp = MsgBox(CargaCadena(1208), vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbNo Then GoTo SalBaja
    End If
    Set Registro = Nothing
    Resp = Baja_Familias(Val(CmbFamilias.TexT), MyEvento)
    If Resp = 0 Then
        MsgBox "OK"
        Refresca_Familias
        Refresca_Datos
        If FrmFamMos.Visible Then
            FrmFamMos.Refresca_Familias
        End If
    End If
SalBaja:
    Me.Enabled = True
End Sub

Private Sub Command4_Click()
    MostrarModal FrmFamMos
End Sub

Private Sub Form_Load()
    Set MyEvento = New ClsEvento
    sQueNombreImagen = ""
    Me.Icon = Form2.Icon
    'picture1.left = CmdColor(1).left
    '
    Set picture1.Picture = LoadPicture(App.Path & "\vacio.jpg")
    Me.Icon = Form2.Icon
    Command3.Enabled = False
    Cambiar_idioma
    Refresca_Familias
    If CmbFamilias.ListCount > 0 Then
        CmbFamilias.ListIndex = 0
    End If
    Command4.Visible = False
    Me.left = frmEpelsa.left
    If Dir(App.Path & "\test.tst") <> "" Then
        MostrarModal FrmFamMos
    End If
    
End Sub
Private Sub Refresca_Familias()
    Dim Registro As New RecordNet
    CmbFamilias.Clear
    Registro.OpenRecordset ("select distinct(codi_fam) from fam_code where posicion<>1525 order by codi_fam")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                CmbFamilias.AddItem .Fields("codi_fam")
                .Movenext
            Loop
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FrmFamMos.Visible Then
        Unload FrmFamMos
    End If

    frmEpelsa.Enabled = True
End Sub

Private Sub MyEvento_Evento1(lData As Long)
    nEvents = nEvents + 1
    If nEvents >= 10 Then
        Do_Events
        nEvents = 0
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case 38
            If CmbFamilias.Enabled = True Then CmbFamilias.SetFocus
    End Select
End Sub


'Private Sub Text2_GotFocus()
'Refresca_Datos
'End Sub
'Private Sub check1_GotFocus()
'Refresca_Datos
'End Sub

Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

'Private Sub TxtOrden_LostFocus()
'Dim MyRs As DAO.Recordset
'Dim MyBa As DAO.Database
'    Set MyBa = AbrirBase()
'    Set MyRs = MyBa.OpenRecordset("select * from fam_code where index=" & Trim(TxtOrden.TexT) & " and codi_fam<>" & Trim(CmbFamilias.TexT))
'    If Not MyRs.EOF Then
'        MsgBox "Atención, Orden de familia ya asignado.", vbExclamation
'        'If Val(TxtOrden.TexT) < nMinPos + 1 Then
'        '    TxtOrden.TexT = CStr(nMinPos + 1)
'        'End If
'        Refresca_Datos
'    End If
'    MyRs.Close
'    Set MyRs = Nothing
'    MyBa.Close
'    Set MyBa = Nothing
'End Sub

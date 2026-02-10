VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6227E780-EE80-4A81-B37C-ED4D8D3587EE}#1.0#0"; "HyperLbl.ocx"
Begin VB.Form FrmPrintLabelGeneric 
   Caption         =   "Impresión Etiquetas"
   ClientHeight    =   5625
   ClientLeft      =   2070
   ClientTop       =   4245
   ClientWidth     =   15420
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin HyperLbl.HyperLabel RTB 
      Height          =   3375
      Left            =   7920
      TabIndex        =   14
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FrmPrintLabelGeneric.frx":0000
      Left            =   5640
      List            =   "FrmPrintLabelGeneric.frx":0010
      TabIndex        =   13
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FrmPrintLabelGeneric.frx":0047
      Left            =   7560
      List            =   "FrmPrintLabelGeneric.frx":0060
      TabIndex        =   12
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox TxtNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "8"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox tXTcODIGO 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Width           =   972
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   11880
      TabIndex        =   6
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   5400
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mostrar ordenados por PLU"
      Height          =   252
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   3012
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mostrar ordenados por Código"
      Height          =   252
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   3012
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PLU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descriptivo"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Etiquetas (1 a 8):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label LblcoDIGO 
      Caption         =   "Código"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1332
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   5
      Top             =   4680
      Width           =   5172
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "FrmPrintLabelGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Elementos As Integer
Private NPrecios100g As Long
Private Codigos100g() As Long
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long

Private sMyOrder As String
Private sLastOrder As String
Private sTEXTO As String
Dim Ini As Integer
Dim lInicio As Boolean
Dim l2 As Boolean

Private Sub GenerateForm()
    Dim nO As Integer
    Dim OldX As Long, OldY As Long, nDC As Long
    Dim ctl As Control
    Dim L As Long
    Dim nPos As Integer
    Dim nPos1 As Integer
    Dim nf As Integer
    Dim lU As Boolean
    Dim s As String
    Dim nArrI(100) As Integer
    Dim nArrF(100) As Integer
    Dim nArrT(100) As Integer
    Dim nC As Integer

    Me.RTB.Caption = Replace(Me.RTB.Caption, "^", "")
    Me.RTB.Caption = Replace(Me.RTB.Caption, "[u] [/u]", "")
    Me.RTB.Caption = Replace(Me.RTB.Caption, "[b]", "[u]")
    Me.RTB.Caption = Replace(Me.RTB.Caption, "[/b]", "[/u]")

    Form3.ScaleMode = 6
    Form3.BackColor = vbWhite
    Form3.BorderStyle = 0
    Form3.AutoRedraw = True

    Form3.Width = 297
    Form3.Height = 210
    Form3.top = 0
    Form3.left = 0

    Form3.RTB(0).top = 0    '0
    Form3.RTB(0).left = 0    '0
    Form3.RTB(0).Width = 68    '68 '70
    Form3.RTB(0).Height = 99

    Form3.RTB(1).top = 0    '0
    Form3.RTB(1).left = 73    '69 'Form3.RTB(0).Width + 6
    Form3.RTB(1).Width = 68
    Form3.RTB(1).Height = 99

    Form3.RTB(2).top = 0    '0
    Form3.RTB(2).left = 148    '143 'Form3.RTB(0).Width + Form3.RTB(1).Width + 7
    Form3.RTB(2).Width = 68
    Form3.RTB(2).Height = 99

    Form3.RTB(3).top = 0    '0
    Form3.RTB(3).left = 221    '217 'Form3.RTB(0).Width + Form3.RTB(1).Width + Form3.RTB(2).Width + 8
    Form3.RTB(3).Width = 68
    Form3.RTB(3).Height = 99

    '---------------------------
    Form3.RTB(4).top = Form3.RTB(0).Height + 5
    Form3.RTB(4).left = 0    '0
    Form3.RTB(4).Width = 68    '68 '70
    Form3.RTB(4).Height = 99

    Form3.RTB(5).top = Form3.RTB(0).Height + 5
    Form3.RTB(5).left = 73    '69 'Form3.RTB(0).Width + 6
    Form3.RTB(5).Width = 68
    Form3.RTB(5).Height = 99

    Form3.RTB(6).top = Form3.RTB(0).Height + 5
    Form3.RTB(6).left = 148    '143 'Form3.RTB(0).Width + Form3.RTB(1).Width + 7
    Form3.RTB(6).Width = 68
    Form3.RTB(6).Height = 99

    Form3.RTB(7).top = Form3.RTB(0).Height + 5
    Form3.RTB(7).left = 221    '217 'Form3.RTB(0).Width + Form3.RTB(1).Width + Form3.RTB(2).Width + 8
    Form3.RTB(7).Width = 68
    Form3.RTB(7).Height = 99

    For nO = 0 To 7
        Form3.RTB(nO).TextRTF = ""
        Form3.RTB(nO).TexT = ""
        Form3.RTB(nO).Font.name = Trim(Combo3.TexT)
        Form3.RTB(nO).Font.size = Combo2.TexT
        Form3.RTB(nO).BackColor = vbWhite
    Next nO

    nC = 0
    For nO = 0 To Val(TxtNum.TexT) - 1
        nC = 0
        Form3.RTB(nO).TexT = Me.RTB.Caption
        nPos1 = 0
        Do
            nPos = InStr(nPos1 + 2, Form3.RTB(nO).TexT, "[u]")
            If nPos Then

                Form3.RTB(nO).TexT = left(Form3.RTB(nO).TexT, nPos - 1) & Mid(Form3.RTB(nO).TexT, nPos + 3)

                nPos1 = InStr(nPos + 2, Form3.RTB(nO).TexT, "[/u]")
                If nPos1 Then
                    Form3.RTB(nO).TexT = left(Form3.RTB(nO).TexT, nPos1 - 1) & Mid(Form3.RTB(nO).TexT, nPos1 + 4)

                    nC = nC + 1
                    nArrI(nC) = nPos - 1
                    nArrF(nC) = nPos1 - nPos

                    nArrT(nC) = 2

                Else

                    nC = nC + 1
                    nArrI(nC) = nPos - 1
                    nArrF(nC) = Len(Form3.RTB(nO).TexT) - nPos

                    nArrT(nC) = 2

                    Exit Do
                End If

            End If
        Loop While nPos

        For nPos = 1 To nC
            Form3.RTB(nO).SelStart = nArrI(nPos)
            Form3.RTB(nO).SelLength = nArrF(nPos)
            Form3.RTB(nO).SelUnderline = True
        Next nPos

    Next nO

    'lU = False
    'For nO = 0 To Val(TxtNum.TexT) - 1
    '    nPos = Len(Me.RTB.Caption)
    '    For nPos1 = 1 To nPos
    '        s = Mid(Me.RTB.Caption, nPos1, 1)
    '        If s = "[" And lU = False Then
    '            Form3.RTB(nO).Font.Underline = True
    '            lU = True
    '            nPos1 = nPos1 + 3
    '        Else
    '            If s = "[" And lU = True Then
    '                Form3.RTB(nO).Font.Underline = False
    '                lU = False
    '                nPos1 = nPos1 + 4
    '            Else
    '                Form3.RTB(nO).TexT = Form3.RTB(nO).TexT & s
    '            End If
    '        End If
    '
    '    Next nPos1
    'Next nO

    ''Recuperar
    Form3.Dialogo.CancelError = True
    On Error Resume Next
    Form3.Dialogo.ShowPrinter
    If Err.Number = 32755 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo nosvamos

    ''recuperar
    Printer.Orientation = 2

    ''recuperar
    Form3.Width = Screen.Width * 1.25           ' Establecer el ancho del formulario.
    Form3.Height = Screen.Height * 1.25         ' Establecer el alto del formulario.
    Form3.left = (Screen.Width - Width) / 2     ' Centrar el formulario horizontalmente.
    Form3.top = (Screen.Height - Height) / 2    ' Centrar el formulario Verticalmente.

    Form3.Refresh
    Do_Events

    ''recuperar
    Form3.PrintForm
    Do_Events

    Unload Form3

nosvamos:
    'ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL)
    'delete our device context
    'DeleteDC nDC

End Sub

Private Sub Me_Init()
    Dim bucle As Long
    Me_Width = Me.Width
    Me_Height = Me.Height
    ReDim ct_top(Me.Controls.Count)
    ReDim CT_Left(Me.Controls.Count)
    ReDim CT_Width(Me.Controls.Count)
    ReDim CT_Height(Me.Controls.Count)
    Dim K, L As Integer
    For bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        ct_top(bucle) = Me.Controls(bucle).top
        CT_Left(bucle) = Me.Controls(bucle).left
        If LCase(Me.Controls(bucle).name) <> "lista" And LCase(Me.Controls(bucle).name) <> "rtb" Then
            ct_top(bucle) = Me.Controls(bucle).top
            CT_Left(bucle) = Me.Controls(bucle).left

            CT_Width(bucle) = Me.Controls(bucle).Width
            CT_Height(bucle) = Me.Controls(bucle).Height
        Else
            bucle = bucle
            If LCase(Me.Controls(bucle).name) = "lista" Then
                K = bucle
                ct_top(bucle) = Me.Controls(bucle).top
                CT_Left(bucle) = Me.Controls(bucle).left

                CT_Width(bucle) = Me.Width / (2.75)
                CT_Height(bucle) = Me.Controls(bucle).Height
            Else
                L = bucle
            End If
        End If

        On Error GoTo 0
    Next bucle
    ct_top(L) = ct_top(K)
    CT_Left(L) = CT_Left(K) + CT_Width(K) + 100
    CT_Width(L) = Me.Width - (CT_Left(K) + Me.Lista.Width) - 100
    CT_Height(L) = Me.Lista.Height
End Sub
Private Sub Me_Resize()
    Dim RelW As Double
    Dim bucle As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 5000 Then Me.Height = 5000
    For bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        RelW = Me.Width / Me_Width
        Me.Controls(bucle).Width = RelW * CT_Width(bucle)
        CT_Width(bucle) = Me.Controls(bucle).Width
        Me.Controls(bucle).left = RelW * CT_Left(bucle)
        CT_Left(bucle) = Me.Controls(bucle).left

        '
        RelW = Me.Height / Me_Height
        If TypeName(Me.Controls(bucle)) <> "TextBox" _
           And TypeName(Me.Controls(bucle)) <> "CommandButton" _
           And TypeName(Me.Controls(bucle)) <> "MaskEdBox" _
           And TypeName(Me.Controls(bucle)) <> "HyperLabel" _
           Then
            Me.Controls(bucle).Height = RelW * CT_Height(bucle)
            CT_Height(bucle) = Me.Controls(bucle).Height
        Else
            If TypeName(Me.Controls(bucle)) = "HyperLabel" Then
                Me.Controls(bucle).Height = Round(RelW, 5) * CT_Height(bucle)
                CT_Height(bucle) = Me.Controls(bucle).Height
            End If
        End If
        Me.Controls(bucle).top = RelW * ct_top(bucle)
        ct_top(bucle) = Me.Controls(bucle).top
        On Error GoTo 0
        '
    Next bucle
    Me_Width = Me.Width
    Me_Height = Me.Height
End Sub
Private Sub CambiarIdioma()
    Label1.Caption = CargaCadena(302)
    Option1(0).Caption = CargaCadena(656)
    Option1(1).Caption = CargaCadena(657)
    LblcoDIGO.Caption = CargaCadena(69)
    'Command1(0).Caption = CargaCadena(287)
    Command1(1).Caption = CargaCadena(7)
    FrmPrintLabelGeneric.Caption = "Impresión Etiquetas."
    Lista.ColumnHeaders(1).TexT = CargaCadena(69)
    Lista.ColumnHeaders(2).TexT = CargaCadena(68)
    Lista.ColumnHeaders(3).TexT = CargaCadena(70)
    Label2.Caption = CargaCadena(1372)
    Command1(0).Caption = CargaCadena(466)
End Sub
Private Sub Refresca_porCodigo()

    Dim Registro As New RecordNet
    Dim MiNumero As Long
    Dim nC As Integer

    If Not (IsNumeric(tXTcODIGO.TexT)) Then Exit Sub
    MiNumero = tXTcODIGO.TexT

    If Option1(0).Value Then
        Registro.OpenRecordset ("select codi_ident from articulo where borrado=false and codigo=" & Val(tXTcODIGO.TexT))
    Else
        Registro.OpenRecordset ("select codi_ident from articulo where borrado=false and plu=" & Val(tXTcODIGO.TexT) & " and codi_ident=" & Combo1.TexT)

    End If
    With Registro
        If .EOF Then
            If Option1(0).Value = True Then
                LblInfo.Caption = CargaCadena(218)
            Else
                LblInfo.Caption = CargaCadena(221)
            End If
            tXTcODIGO.SetFocus
        Else
            'Combo1.TexT = .Fields("codi_ident")
            For nC = 0 To Combo1.ListCount - 1
                Combo1.ListIndex = nC
                If Combo1.TexT = .Fields("codi_ident") Then
                    Exit For
                End If
            Next nC
            Refresca_Datos
            If Option1(0).Value Then
                On Error GoTo NoHay
                Do While Val(Lista.SelectedItem.TexT) <> MiNumero
                    Lista.ListItems(Lista.SelectedItem.Index + 1).Selected = True
                    Lista.SelectedItem.EnsureVisible
                Loop
NoHay:
                On Error GoTo 0
                Refresca_Item
            Else
                Do While Val(Lista.SelectedItem.SubItems(1)) <> MiNumero
                    Lista.ListItems(Lista.SelectedItem.Index + 1).Selected = True
                    Lista.SelectedItem.EnsureVisible
                Loop
                Refresca_Item
            End If
        End If
    End With


End Sub

Private Sub Combo1_Click()
    If Combo1.TexT <> "" Then
        Refresca_Datos
    End If
    tXTcODIGO.Enabled = True
End Sub

'Private Sub Combo2_Change()
'       RTB.Font.Name = Trim(Combo3.TexT)
'       RTB.Font.Size = Combo2.ItemData
'       tXTcODIGO_KeyPress 13
'End Sub

'Private Sub Combo3_Change()
'       RTB.Font.Name = Trim(Combo3.TexT)
'       RTB.Font.Size = Combo2.ItemData
'       tXTcODIGO_KeyPress 13
'End Sub

Private Sub Combo3_Click()
    RTB.Font.name = Trim(Combo3.TexT)
    RTB.Font.size = Combo2.TexT
    'tXTcODIGO_KeyPress 13
    'Refresca_Item
    RTB.Caption = sTEXTO
End Sub
Private Sub Combo2_Click()
    RTB.Font.name = Trim(Combo3.TexT)
    RTB.Font.size = Combo2.TexT
    'tXTcODIGO_KeyPress 13
    'Refresca_Item
    RTB.Caption = sTEXTO
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Miarticulo As DB_Articulo
    Dim i As Integer

    Command1(0).Enabled = False
    'Command1(1).Enabled = False

    sTEXTO = RTB.Caption

    If Trim(TxtNum.TexT) = "" Then TxtNum.TexT = "1"



    '    ' Variable para el recordset
    '    Dim El_Recordset As New ADODB.Recordset
    '
    '    ' Nuevo objeto Recordset desconectado
    '    'Set El_Recordset = New ADODB.Recordset
    Select Case Index
    Case 0

        Call GenerateForm

        '
        '            ' Crea cuatro campos, tres de tipo String y uno de tipo Date
        '            With El_Recordset.Fields
        '                .Append "texto", adVarChar, 3000
        '            End With
        '
        '            ' Abre el recorset para poder agregar datos
        '            El_Recordset.Open
        '
        '            ' Agrega valores al recordset, es decir a los campos
        '            '********************************************************************
        '            With El_Recordset
        '                .AddNew Array("texto"), Array(sTEXTO)
        '            End With
        '            '********************************************************************
        '
        '            ' Asigna el recordset como fuente de datos del DataReport
        '            Set DataReport1.DataSource = El_Recordset
        '
        '            DataReport1.Orientation = vbPRORPortrait
        '            DataReport1.LeftMargin = 0
        '            DataReport1.RightMargin = 0
        '            DataReport1.TopMargin = 0
        '            DataReport1.BottomMargin = 0
        '            With Printer
        '                .PaperSize = vbPRPSA4 'Using A4 size
        '                '.Orientation = 2 'Landscape
        '            End With
        '            Load DataReport1
        '            For i = 1 To Val(TxtNum.TexT)
        '                DataReport1.Sections("Etiquetas").Controls.Item("Texto" & CStr(i)).Caption = sTEXTO  '"Texto" & CStr(i)
        '            Next i
        '            MuestraInformeEtiquetas
        '
        '            El_Recordset.Close
        '            Set El_Recordset = Nothing
    Case 1
        Unload Me
    End Select

    Command1(0).Enabled = True
    'Command1(1).Enabled = True

End Sub

Private Sub Form_Activate()
    RTB.Width = 4000    '3975
    RTB.Height = 7500
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    lInicio = False
    tXTcODIGO.TexT = FrmArticulo2.CmbCodigo.TexT
    Call Refresca_porCodigo

End Sub

Private Sub Form_Load()
    Dim Retorno As Boolean
    Dim MyPref As String
    lInicio = True
    CambiarIdioma
    sMyOrder = ""
    sLastOrder = ""
    Me_Init
    Me.Height = Screen.Height * 0.89
    Me.Icon = Form2.Icon
    Retorno = Refresca_Combo
    Lista.HotTracking = True
    MyPref = RPreferencias("CHGPRECIO1", 1)
    If MyPref <> "" Then
        If left(MyPref, 1) <> "0" Then
            Option1(1).Value = True
        Else
            Option1(0).Value = True
        End If
    End If
    If Retorno Then Refresca_Datos
    If Trim(FrmArticulo2.CmbCodigo.TexT) <> "" Then
        tXTcODIGO.TexT = FrmArticulo2.CmbCodigo.TexT
        Call Refresca_porCodigo
    End If
End Sub
Private Function Refresca_Combo() As Boolean

    Dim Registro As New RecordNet
    Dim Retorno As Boolean
    Retorno = True
    Combo1.Clear

    If HaySeguridad And UsuarioActual.Asociado <> 0 Then
        Registro.OpenRecordset ("select codi_ident from seccion where borrado=false and codi_ident=" & UsuarioActual.Asociado & " order by codi_ident")
        Combo1.Enabled = False
    Else
        Registro.OpenRecordset ("select codi_ident from seccion where borrado=false order by codi_ident")
    End If
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Combo1.AddItem .Fields("codi_ident")
                .Movenext
            Loop
            Combo1.TexT = Combo1.List(0)
        Else
            LblInfo.Caption = CargaCadena(399)
            Combo1.Enabled = False
            Retorno = False
        End If
    End With

    Refresca_Combo = Retorno

End Function
Private Sub Refresca_Datos()

    Dim Registro As New RecordNet
    Dim Miorden As String
    If sMyOrder = "" Then
        If Option1(0).Value = True Then
            Miorden = "codigo"
        Else
            Miorden = "plu"
        End If
    Else
        Miorden = sMyOrder
    End If

    Lista.ListItems.Clear

    Registro.OpenRecordset ("select codigo,plu,des_plu1,codi_ident from articulo where borrado=false and codi_ident=" & Val(Combo1.TexT) & " order by " & Miorden)
    If sMyOrder <> "" Then
        sLastOrder = sMyOrder
    End If
    sMyOrder = ""
    Elementos = 1
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Lista.ListItems.Add Elementos, , .Fields("codigo")
                Lista.ListItems(Elementos).SubItems(1) = .Fields("plu")
                Lista.ListItems(Elementos).SubItems(2) = .Fields("des_plu1")
                Elementos = Elementos + 1
                .Movenext
            Loop
            Lista.Enabled = True
            'Lista.SetFocus
            Refresca_Item
            Command1(0).Enabled = True
            Command1(0).Visible = True
        Else
            'Lista.Enabled = False
            Command1(0).Enabled = False
        End If
    End With

    Set Registro = Nothing

End Sub

Private Function anegrita(ByVal sS As String) As String
    Dim nL As Integer
    'Dim SD As String
    Dim pos As Integer
    pos = 1
    'SD = ""
    nL = InStr(1, sS, "~")    '|
    If nL <> 0 Then
        'Ini = 0
        Do While nL <> 0

            If Ini = 0 Then
                sS = Mid(sS, pos, nL - 1) & "[u]" & Mid(sS, nL + 1)
                Ini = 1
            Else
                sS = Mid(sS, pos, nL - 1) & "[/u]" & Mid(sS, nL + 1)
                Ini = 0
            End If
            'sS = Mid(sS, nL + 1)
            nL = InStr(nL + 1, sS, "~")
            'pos = nL
        Loop
    End If

    'nL = InStr(1, sS, "^") '|
    'If nL <> 0 Then
    '    Do While nL <> 0
    '        sS = Mid(sS, pos, nL - 1) & Chr(13) & Chr(10) & Mid(sS, nL + 1)
    '        nL = InStr(nL + 1, sS, "^")
    '    Loop
    'End If

    ''''''''''
    'pos = 1
    ''SD = ""
    'nL = InStr(1, sS, ",")
    'If nL <> 0 Then
    '    Do While nL <> 0
    '        'If Ini = 0 Then
    '            sS = Mid(sS, pos, nL) & " " & Mid(sS, nL + 1)
    '            'Ini = 1
    '        'Else
    '        '    sS = Mid(sS, pos, nL - 1) & "[/u]" & Mid(sS, nL + 1)
    '        '    Ini = 0
    '        'End If
    '        'sS = Mid(sS, nL + 1)
    '        nL = InStr(nL + 1, sS, ",")
    '        'pos = nL
    '    Loop
    'End If
    '''''''''''
    'pos = 1
    ''SD = ""
    'nL = InStr(1, sS, ":")
    'If nL <> 0 Then
    '    Do While nL <> 0
    '        'If Ini = 0 Then
    '            sS = Mid(sS, pos, nL) & " " & Mid(sS, nL + 1)
    '            'Ini = 1
    '        'Else
    '        '    sS = Mid(sS, pos, nL - 1) & "[/u]" & Mid(sS, nL + 1)
    '        '    Ini = 0
    '        'End If
    '        'sS = Mid(sS, nL + 1)
    '        nL = InStr(nL + 1, sS, ":")
    '        'pos = nL
    '    Loop
    'End If
    '''''''''''
    'pos = 1
    ''SD = ""
    'nL = InStr(1, sS, ".")
    'If nL <> 0 Then
    '    Do While nL <> 0
    '        'If Ini = 0 Then
    '            sS = Mid(sS, pos, nL) & " " & Mid(sS, nL + 1)
    '            'Ini = 1
    '        'Else
    '        '    sS = Mid(sS, pos, nL - 1) & "[/u]" & Mid(sS, nL + 1)
    '        '    Ini = 0
    '        'End If
    '        'sS = Mid(sS, nL + 1)
    '        nL = InStr(nL + 1, sS, ".")
    '        'pos = nL
    '    Loop
    'End If
    '''''''''''

    anegrita = sS
End Function
Private Function aCR(ByVal sS As String) As String
    Dim nL As Integer
    Dim pos As Integer
    pos = 1
    nL = InStr(1, sS, "^")    '|
    If nL <> 0 Then
        Do While nL <> 0
            sS = Mid(sS, pos, nL - 1) & Chr(13) & Chr(10) & Mid(sS, nL + 1)
            nL = InStr(nL + 1, sS, "^")
        Loop
    End If

    aCR = sS
End Function
Private Sub Refresca_Datos_Texto()

    Dim Registro As New RecordNet
    Dim Miorden As String
    Dim i As Integer
    sTEXTO = ""
    'RTB.TexT = ""
    RTB.Caption = ""
    If Option1(0).Value = True Then
        Miorden = "codigo"
    Else
        Miorden = "plu"
    End If

    Ini = 0

    Registro.OpenRecordset ("select * from articulo where borrado=false and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))

    With Registro
        If Not .EOF Then

            'sTEXTO = "         " & Trim(.Fields("des_plu1")) & " " & vbCr & vbLf

            RTB.Caption = "           " & "[b]" & Trim(.Fields("des_plu1")) & "[/b]" & "[]"    'vbCr & vbLf

            For i = 2 To 21    '11 To 21
                On Error GoTo sigue
                If i > 9 Then
                    If i = 10 Then
                        'sTEXTO = sTEXTO & Trim(.Fields("des_plu0"))
                        'RTB.Caption = RTB.Caption & anegrita(Trim(.Fields("des_plu0")))
                        If Not IsNull(.Fields("des_plu0")) Then
                            RTB.Caption = RTB.Caption & Trim(.Fields("des_plu0"))
                        End If
                    Else
                        If i = 11 Then
                            'sTEXTO = sTEXTO & Trim(.Fields("des_pluX"))
                            'RTB.Caption = RTB.Caption & anegrita(Trim(.Fields("des_pluX")))
                            If Not IsNull(.Fields("des_pluX")) Then
                                RTB.Caption = RTB.Caption & Trim(.Fields("des_pluX"))
                            End If
                        Else
                            'sTEXTO = sTEXTO & Trim(.Fields("des_plu" & CStr(i - 1)))
                            'RTB.Caption = RTB.Caption & anegrita(Trim(.Fields("des_plu" & CStr(i - 1))))
                            If Not IsNull(.Fields("des_plu" & CStr(i - 1))) Then
                                RTB.Caption = RTB.Caption & Trim(.Fields("des_plu" & CStr(i - 1)))
                            End If
                        End If
                    End If
                Else
                    'sTEXTO = sTEXTO & Trim(.Fields("des_plu" & CStr(i)))
                    'RTB.Caption = RTB.Caption & anegrita(Trim(.Fields("des_plu" & CStr(i))))
                    If Not IsNull(.Fields("des_plu" & CStr(i))) Then
                        RTB.Caption = RTB.Caption & Trim(.Fields("des_plu" & CStr(i)))
                    End If
                End If
sigue:
                On Error GoTo 0
            Next i

        End If
    End With

    'RTB.Caption = anegrita(RTB.Caption)

    Set Registro = Nothing

    Registro.OpenRecordset ("select * from lintxt2040 where mostrador=" & CStr(Val(Combo1.TexT)) & " and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))

    With Registro
        If Not .EOF Then
            For i = 21 To 80
                On Error GoTo sigue1
                'sTEXTO = sTEXTO & Trim(.Fields("txt_" & CStr(i)))
                'RTB.Caption = RTB.Caption & anegrita(Trim(.Fields("txt_" & CStr(i))))
                If Not IsNull(.Fields("txt_" & CStr(i))) Then
                    RTB.Caption = RTB.Caption & Trim(.Fields("txt_" & CStr(i)))
                End If
sigue1:
                On Error GoTo 0
            Next i
        End If
    End With
    RTB.Caption = anegrita(RTB.Caption)
    RTB.Caption = Replace(RTB.Caption, Chr(94), vbCrLf)
    sTEXTO = RTB.Caption

End Sub

Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmArticulo2.Enabled = True
End Sub

Private Sub Lista_Click()
    Refresca_Item
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Lista.ColumnHeaders(1).TexT = CargaCadena(69)
    Lista.ColumnHeaders(2).TexT = CargaCadena(68)
    Lista.ColumnHeaders(3).TexT = CargaCadena(70)

    Select Case ColumnHeader.Index
    Case 1
        Option1(0).Value = True
        Lista.ColumnHeaders(1).TexT = "==> " & CargaCadena(69) & " <=="
    Case 2
        Option1(1).Value = True
        Lista.ColumnHeaders(2).TexT = "==> " & CargaCadena(68) & " <=="
    Case Else
        If ColumnHeader.Index = 3 Then
            sMyOrder = "des_plu1"
            Lista.ColumnHeaders(3).TexT = "==> " & CargaCadena(70) & " <=="
        Else
            If ColumnHeader.Index = 4 Then
                If UsaEuro Then
                    sMyOrder = "euros"
                Else
                    sMyOrder = "precio"
                End If
                Lista.ColumnHeaders(4).TexT = "==> " & CargaCadena(658) & " <=="
            Else
                If sLastOrder <> "" Then
                    sMyOrder = sLastOrder
                    If sMyOrder = "des_plu1" Then Lista.ColumnHeaders(3).TexT = "==> " & CargaCadena(70) & " <=="
                    If sMyOrder = "euros" Or sMyOrder = "precio" Then Lista.ColumnHeaders(4).TexT = "==> " & CargaCadena(658) & " <=="
                End If
            End If
        End If
    End Select
    If Combo1.TexT <> "" Then Refresca_Datos
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Refresca_Item
End Sub

Private Sub Option1_Click(Index As Integer)
    If Combo1.TexT <> "" Then Refresca_Datos
    If Option1(0).Value = True Then
        LblcoDIGO.Caption = CargaCadena(69)  '"Código"
    Else
        LblcoDIGO.Caption = CargaCadena(68)  '"PLU"
    End If
End Sub

'Private Sub MuestraInformeEtiquetas()'
'
'    DescargaReport = False
'    DataReport1.Show
'    Do Until DescargaReport
'        Sleep (100)
'        Do_Events
'    Loop
'
'End Sub

Private Sub Refresca_Item()

    On Error GoTo noHayItem
    If Option1(0).Value = True Then
        tXTcODIGO.TexT = Lista.SelectedItem.TexT
    Else
        tXTcODIGO.TexT = Lista.SelectedItem.SubItems(1)
    End If
    LblInfo.Caption = Lista.SelectedItem.SubItems(2)
    Call Refresca_Datos_Texto
noHayItem:
    If Err.Number <> 0 Then
        On Error GoTo 0
        'If Option1(0).Value = True Then
        tXTcODIGO.TexT = Trim(FrmArticulo2.CmbCodigo.TexT)  'Lista.SelectedItem.TexT
        'Else
        '    tXTcODIGO.TexT = Lista.SelectedItem.SubItems(1)
        'End If
        'LblInfo.Caption = Lista.SelectedItem.SubItems(2)
        Call Refresca_Datos_Texto

    End If
    On Error GoTo 0

End Sub


Private Sub tXTcODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lista.ListItems.Count > 0 Then
        Select Case KeyCode
        Case 40
            If Lista.SelectedItem.Index < Lista.ListItems.Count Then
                Lista.ListItems.Item(Lista.SelectedItem.Index + 1).Selected = True
                Lista.ListItems(Lista.SelectedItem.Index).EnsureVisible
                Refresca_Item
            End If
        Case 38
            If Lista.SelectedItem.Index > 1 Then
                Lista.ListItems.Item(Lista.SelectedItem.Index - 1).Selected = True
                Lista.ListItems(Lista.SelectedItem.Index).EnsureVisible
                Refresca_Item
            End If
        Case 33
            If Lista.SelectedItem.Index > 10 Then
                Lista.ListItems.Item(Lista.SelectedItem.Index - 10).Selected = True
                Lista.ListItems(Lista.SelectedItem.Index).EnsureVisible
                Refresca_Item
            End If
        Case 34
            If Lista.SelectedItem.Index < (Lista.ListItems.Count - 10) Then
                Lista.ListItems.Item(Lista.SelectedItem.Index + 10).Selected = True
                Lista.ListItems(Lista.SelectedItem.Index).EnsureVisible
                Refresca_Item
            End If

        End Select
    End If
End Sub

Private Sub tXTcODIGO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Refresca_porCodigo
        Exit Sub
    End If
    If Option1(0).Value = True Then
        tXTcODIGO.Locked = Checktexto(KeyAscii, 6, tXTcODIGO.TexT, True, False)
    Else
        tXTcODIGO.Locked = Checktexto(KeyAscii, 4, tXTcODIGO.TexT, True, False)
    End If

End Sub

'Private Sub tXTcODIGO_LostFocus()
'    lblinfo.Caption = ""
'    Refresca_porCodigo
'End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 49 Or KeyAscii > 56) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmArtText 
   Caption         =   "Textos de Artículos"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   4245
   ClientWidth     =   15345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   15345
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkT 
      Caption         =   "T5"
      Enabled         =   0   'False
      Height          =   195
      Index           =   5
      Left            =   10320
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox ChkT 
      Caption         =   "Text4"
      Height          =   195
      Index           =   4
      Left            =   9600
      TabIndex        =   34
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox ChkT 
      Caption         =   "Text3"
      Height          =   195
      Index           =   3
      Left            =   8880
      TabIndex        =   33
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox ChkT 
      Caption         =   "Text2"
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   32
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox ChkT 
      Caption         =   "Text1"
      Height          =   195
      Index           =   1
      Left            =   7440
      TabIndex        =   31
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox ChkT 
      Caption         =   "(Lin1-Lin40)"
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   30
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdNS 
      Caption         =   "NO_Sel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   29
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdNK 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   28
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton CmdK 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   27
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton CmdNB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   26
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   25
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton CmdNU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   24
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton CmdCursiva 
      Caption         =   "SEL_Italic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   23
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdBold 
      Caption         =   "SEL_Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   22
      Top             =   4800
      Width           =   1695
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PLU"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descriptivo"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "CR/LF ON"
      Height          =   255
      Left            =   11520
      TabIndex        =   21
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Accept."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   18
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdUnder 
      Caption         =   "SEL_Under-Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdCR 
      Caption         =   "CR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   15
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton CmdU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton CmdConvert 
      Caption         =   "Convert to Scale."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   3495
      Left            =   11160
      TabIndex        =   12
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmArtText.frx":0000
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3495
      Left            =   6720
      TabIndex        =   11
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmArtText.frx":00C9
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FrmArtText.frx":0192
      Left            =   6480
      List            =   "FrmArtText.frx":01A5
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FrmArtText.frx":01EB
      Left            =   8520
      List            =   "FrmArtText.frx":0201
      TabIndex        =   9
      Top             =   360
      Width           =   1695
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
      Left            =   11400
      TabIndex        =   6
      Top             =   5280
      Width           =   1332
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
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      Height          =   135
      Left            =   6720
      TabIndex        =   20
      Top             =   5880
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   5655
      Left            =   10920
      TabIndex        =   19
      Top             =   240
      Width           =   135
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
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   4680
      Width           =   3855
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
Attribute VB_Name = "FrmArtText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
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
Dim sCodigoG As String
Dim sPluG As String
Dim nIndexT As Integer
Dim lEnC As Boolean
Dim sSecG As String
#Const BALSAM = False

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

        CT_Width(bucle) = Me.Controls(bucle).Width
        CT_Height(bucle) = Me.Controls(bucle).Height
        On Error GoTo 0
    Next bucle
    ChkT(0).Value = vbChecked
    ChkT(1).Value = vbUnchecked
    ChkT(2).Value = vbUnchecked
    ChkT(3).Value = vbUnchecked
    ChkT(4).Value = vbUnchecked
    nIndexT = 0

    #If BALSAM = True Then
        ChkT(0).Visible = False
        ChkT(0).Enabled = False
        ChkT(1).Visible = False
        ChkT(1).Enabled = False
        ChkT(2).Visible = False
        ChkT(2).Enabled = False
        ChkT(3).Visible = False
        ChkT(3).Enabled = False
        ChkT(4).Visible = False
        ChkT(4).Enabled = False
        nIndexT = 1
    #Else
        If Dir(App.Path & "\5l.txt") <> "" Then
            For bucle = 0 To 4
                ChkT(bucle).Visible = False
                ChkT(bucle).Enabled = False

            Next bucle
        End If
    #End If

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
           And TypeName(Me.Controls(bucle)) <> "RichTextBox" _
           Then
            Me.Controls(bucle).Height = RelW * CT_Height(bucle)
            CT_Height(bucle) = Me.Controls(bucle).Height
        Else
            If TypeName(Me.Controls(bucle)) = "RichTextBox" Then
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
    Lista.ColumnHeaders(3).Width = Lista.Width - (Lista.ColumnHeaders(2).Width + Lista.ColumnHeaders(1).Width)
End Sub
Private Sub CambiarIdioma()
    Label1.Caption = CargaCadena(302)
    Option1(0).Caption = CargaCadena(656)
    Option1(1).Caption = CargaCadena(657)
    LblcoDIGO.Caption = CargaCadena(69)
    Command1(1).Caption = CargaCadena(7)
    CmdBorrar.Caption = CargaCadena(38)
    CmdGuardar.Caption = CargaCadena(287)
    CmdConvert.Caption = "Convert."    'CargaCadena(539)
    FrmArtText.Caption = "Textos Artículos."
    Lista.ColumnHeaders(1).TexT = CargaCadena(69)
    Lista.ColumnHeaders(2).TexT = CargaCadena(68)
    Lista.ColumnHeaders(3).TexT = CargaCadena(70)

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
            ''dar de alta artículo...

            If Option1(0).Value = True Then
                LblInfo.Caption = CargaCadena(218)
            Else
                LblInfo.Caption = CargaCadena(221)
            End If
            'tXTcODIGO.SetFocus

        Else
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

Private Sub Chk1_Click()
    Dim nf As Integer
    If Chk1.Value Then
        nf = FreeFile()
        Open App.Path & "\crlf.on" For Output As #nf
        Close #nf
    Else
        If Dir(App.Path & "\crlf.on") <> "" Then
            Kill App.Path & "\crlf.on"
        End If
    End If
End Sub

Private Sub ChkT_Click(Index As Integer)
    Dim n As Integer

    If Dir(App.Path & "\5l.txt") <> "" Then Exit Sub

    If lEnC = False Then
        lEnC = True
        For n = 0 To 5
            ChkT(n).Value = vbUnchecked
        Next n
        'For n = 0 To 5
        '    If n = Index Then
        '        'ChkT(n).Value = vbChecked
        '    Else
        '        ChkT(n).Value = vbUnchecked
        '    End If
        'Next n
        ChkT(Index).Value = vbChecked
        nIndexT = Index
        Refresca_Datos_Texto
        lEnC = False
    End If
End Sub

Private Sub CmdB_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 3
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "<b>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "<b>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub


Private Sub CmdBold_Click()
    If Len(RTB1.SelText) > 0 Then
        RTB1.SelBold = True
    End If
End Sub

Private Sub CmdCursiva_Click()
    If Len(RTB1.SelText) > 0 Then
        RTB1.SelItalic = True
    End If
End Sub

Private Sub CmdBorrar_Click()
    RTB1.TexT = ""
End Sub


''''''''''''''''''''''''''''''''
Private Sub Cmdconvert_Click()
    Dim old_start As Integer
    Dim old_length As Integer
    Dim nC, nX As Integer
    Dim st As String
    Dim lSelU As Boolean
    Dim lSelB As Boolean
    Dim lSelI As Boolean
    Dim i As Integer
    Dim fifo As Integer
    Dim lifo As Integer

    'Options:
    '+H              add an HTML header and footer
    '+G              add a generator Metatag
    '+T="MyTitle"    add a title (only works if +H is used)
    '+CR             add a carraige return after all <br>s
    '+I              keep html codes intact
    '+F=X            default font size (blanks out any changes to this size - saves on space)
    '-FF             ignore font faces

    ''''
    RTB.TexT = ""

    '''''''''''''''''''''''''''''''''''''''''''
    ''RTB1.TexT = Replace(RTB1.TexT, " ", " ~")
    'sT = rtf2html3(RTB1.TextRTF, "-FF")
    ''sT = Replace(sT, " ~", " ")
    'sT = Replace(sT, "<font size=2>", "")
    'sT = Replace(sT, "</font>", "")
    'If Chk1.Value = vbUnchecked Then
    '    sT = Replace(sT, "<br>", "")
    'Else
    '    sT = Replace(sT, "<br>", Chr(10))
    'End If
    '
    'RTB.TexT = sT
    'Exit Sub
    '''''''''''''''''''''''''''''''''''''''''''''

    If Len(RTB1.TexT) = 0 Then Exit Sub
    st = ""
    lSelU = False
    lSelB = False
    lSelI = False
    nX = Len(RTB1.TexT)
    fifo = 0
    lifo = 8
    For nC = 1 To Len(RTB1.TexT)
        RTB1.SelStart = nC
        RTB1.SelLength = 1
        If Mid(RTB1.TexT, nC, 1) <> Chr(13) Then    'And Mid(RTB1.TexT, nC, 1) <> " " Then 'Mid(RTB1.TexT, nC, 1) <> Chr(13) And Mid(RTB1.TexT, nC, 1) <> Chr(10) Then

            If ((RTB1.SelUnderline) And lSelU = False) Or ((RTB1.SelBold) And lSelB = False) Or ((RTB1.SelItalic) And lSelI = False) Then
                If (RTB1.SelUnderline) And lSelU = False Then
                    st = st & "<u>"
                    lSelU = True
                    fifo = fifo + 4
                End If
                If (RTB1.SelBold) And lSelB = False Then
                    st = st & "<b>"
                    lSelB = True
                    fifo = fifo + 2
                End If
                If (RTB1.SelItalic) And lSelI = False Then
                    st = st & "<i>"
                    lSelI = True
                    fifo = fifo + 1
                End If
                st = st & Mid(RTB1.TexT, nC, 1)
            Else
                'fifo
                '<u>  <b>  <i>
                ' 0    0    0                0
                ' 0    0    1   </i>         1
                ' 0    1    0   </b>         2
                ' 0    1    1   </i></b>     3
                ' 1    0    0   </u>         4
                ' 1    0    1   </i></u>     5
                ' 1    1    0   </b></u>     6
                ' 1    1    1   </i></b></u> 7

                st = st & Mid(RTB1.TexT, nC, 1)
                If fifo = 7 Then
                    'Do While lSelU = True Or lSelB = True Or lSelI = True
                    If (lSelI And RTB1.SelItalic = False) Or (lSelI And nC = nX) Then
                        st = st & "</i>"
                        lSelI = False
                        fifo = fifo - 1
                    End If
                    If (lSelB And RTB1.SelBold = False) Or (lSelB And nC = nX) Then
                        st = st & "</b>"
                        lSelB = False
                        fifo = fifo - 2
                    End If
                    If (lSelU And RTB1.SelUnderline = False) Or (lSelU And nC = nX) Then
                        st = st & "</u>"
                        lSelU = False
                        fifo = fifo - 4
                    End If
                    'Loop
                    'fifo = 0
                Else
                    If (lSelI And RTB1.SelItalic = False) Or (lSelI And nC = nX) Then
                        'sT = sT & "</i>"
                        lSelI = False
                    End If
                    If (lSelB And RTB1.SelBold = False) Or (lSelB And nC = nX) Then
                        'sT = sT & "</b>"
                        lSelB = False
                    End If
                    If (lSelU And RTB1.SelUnderline = False) Or (lSelU And nC = nX) Then
                        'sT = sT & "</u>"
                        lSelU = False
                    End If
                End If
                Select Case fifo
                Case 1
                    If lSelI = False Then
                        fifo = fifo - 1
                        st = st & "</i>"
                        GoTo finSelect
                    End If
                Case 2
                    If lSelB = False Then
                        fifo = fifo - 2
                        st = st & "</b>"
                        GoTo finSelect
                    End If
                    '<u>  <b>  <i>
                Case 3    ' 0    1    1   </i></b>     3
                    If lSelI = False And lSelB = False Then
                        fifo = fifo - 3
                        st = st & "</i></b>"
                        GoTo finSelect
                    Else
                        If lSelI = False Then
                            fifo = fifo - 1
                            st = st & "</i>"
                            GoTo finSelect
                        End If
                        If lSelB = False Then
                            fifo = fifo - 2
                            st = st & "</b>"
                            GoTo finSelect
                        End If
                    End If
                Case 4
                    If lSelU = False Then
                        fifo = fifo - 4
                        st = st & "</u>"
                        GoTo finSelect
                    End If
                    '<u>  <b>  <i>
                Case 5    ' 1    0    1   </i></u>     5
                    If lSelI = False And lSelU = False Then
                        fifo = fifo - 5
                        st = st & "</i></u>"
                        GoTo finSelect
                    Else
                        If lSelI = False Then
                            fifo = fifo - 1
                            st = st & "</i>"
                            GoTo finSelect
                        End If
                        If lSelU = False Then
                            fifo = fifo - 4
                            st = st & "</u>"
                            GoTo finSelect
                        End If
                    End If

                Case 6    ' 1    1    0   </b></u>     6
                    If lSelB = False And lSelU = False Then
                        fifo = fifo - 6
                        st = st & "</b></u>"
                        GoTo finSelect
                    Else
                        If lSelB = False Then
                            fifo = fifo - 4
                            st = st & "</b>"
                            GoTo finSelect
                        End If
                        If lSelU = False Then
                            fifo = fifo - 2
                            st = st & "</u>"
                            GoTo finSelect
                        End If
                    End If

                Case 7    ' 1    1    1   </i></b></u> 7
                    'If lSelI = False And lSelB = False And lSelU = False Then
                    '    fifo = fifo - 7
                    '    sT = sT & "</i></b></u>"
                    '    GoTo finSelect
                    'Else
                    '    If lSelB = False And lSelU = False Then
                    '        fifo = fifo - 6
                    '        sT = sT & "</b></u>"
                    '        GoTo finSelect
                    '    End If
                    '    If lSelI = False And lSelB = False Then
                    '        fifo = fifo - 3
                    '        sT = sT & "</i></b>"
                    '        GoTo finSelect
                    '    End If
                    '    If lSelI = False And lSelU = False Then
                    '        fifo = fifo - 5
                    '        sT = sT & "</i></u>"
                    '        GoTo finSelect
                    '    End If
                    '    If lSelB = False Then
                    '        fifo = fifo - 4
                    '        sT = sT & "</b>"
                    '        GoTo finSelect
                    '    End If
                    '    If lSelU = False Then
                    '        fifo = fifo - 2
                    '        sT = sT & "</u>"
                    '        GoTo finSelect
                    '    End If
                    '    If lSelI = False Then
                    '        fifo = fifo - 1
                    '        sT = sT & "</i>"
                    '        GoTo finSelect
                    '    End If
                    '
                    'End If

                End Select
finSelect:

            End If
        Else
            'If Mid(RTB1.TexT, nC, 1) = Chr(13) Then
            st = st & "<br>"
            'Else
            '    sT = sT & " "
            'End If
        End If
    Next nC

    st = Replace(st, "<u><u>", "")
    st = Replace(st, "<b><b>", "")
    st = Replace(st, "<i><i>", "")
    st = Replace(st, "</u></u>", "</u>")
    st = Replace(st, "</b></b>", "</b>")
    st = Replace(st, "</i></i>", "</i>")
    st = Replace(st, "<br><br>", "<br>")

    st = Replace(st, "<br>" & Chr(10) & "<u>", "<br><u>")
    st = Replace(st, "<br>" & Chr(10) & "<b>", "<br><b>")
    st = Replace(st, "<br>" & Chr(10) & "<i>", "<br><i>")

    st = Replace(st, "<br><u>" & Chr(10), "<br><u>")
    st = Replace(st, "<br><b>" & Chr(10), "<br><b>")
    st = Replace(st, "<br><i>" & Chr(10), "<br><i>")

    If Right(st, 4) = "<br>" Then
        st = Mid(st, 1, Len(st) - 4)
    End If

    If Len(st) > 1 Then
        On Error GoTo sigueAA
        If Mid(st, Len(st) - 4, 5) = "<br>" & Chr(10) Then
            st = Mid(st, 1, Len(st) - 5)
            'sT = Mid(sT, 1, Len(sT) - 5) & Chr(10)
        End If


        If Mid(st, Len(st) - 6, 7) = "<br><u>" Then
            st = Mid(st, 1, Len(st) - 6) & "<u>"
        End If
        If Mid(st, Len(st) - 6, 7) = "<br><b>" Then
            st = Mid(st, 1, Len(st) - 6) & "<b>"
        End If
        If Mid(st, Len(st) - 6, 7) = "<br><i>" Then
            st = Mid(st, 1, Len(st) - 6) & "<i>"
        End If
sigueAA:

    End If

    'sT = Replace(sT, Chr(10), "")
    'sT = Replace(sT, "<br>", vbLf)
    st = Replace(st, "<br>", "")
    If Chk1.Value = vbUnchecked Then
        st = Replace(st, Chr(10), "")
    End If
    st = Replace(st, Chr(13) & Chr(10), Chr(10))


    'RTB.TexT = Trim(LblInfo.Caption) & vbCrLf & sT
    RTB.TexT = st
End Sub

Private Sub CmdGuardar_Click()
    Call Guarda_Texto
End Sub

Private Sub CmdK_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 3
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "<i>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "<i>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub


Private Sub CmdNB_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 4
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "</b>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "</b>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub



Private Sub cmdNK_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 4
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "</i>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "</i>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub



Private Sub CmdNS_Click()
    RTB1.SelUnderline = False
    RTB1.SelBold = False
    RTB1.SelItalic = False

End Sub

Private Sub CmdU_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 3
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "<u>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "<u>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub

Private Sub CmdNU_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 4
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & "</u>" & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = "</u>" & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub

Private Sub CmdCR_Click()
    Dim ntmp As Integer
    ntmp = RTB.SelStart + 1
    If RTB.SelStart > 0 Then
        RTB.TexT = Mid(RTB.TexT, 1, RTB.SelStart) & Chr(10) & Mid(RTB.TexT, RTB.SelStart + 1)
    Else
        RTB.TexT = Chr(10) & RTB.TexT
    End If
    RTB.SelStart = ntmp
    RTB.SetFocus
End Sub

Private Sub cmdUnder_Click()
    If Len(RTB1.SelText) > 0 Then
        RTB1.SelUnderline = True
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.TexT <> "" Then
        Refresca_Datos
    End If
    tXTcODIGO.Enabled = True
End Sub

Private Sub Combo3_Click()
    RTB.Font.name = Trim(Combo3.TexT)
    RTB.Font.size = Combo2.TexT
End Sub
Private Sub Combo2_Click()
    RTB.Font.name = Trim(Combo3.TexT)
    RTB.Font.size = Combo2.TexT
End Sub

Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    lInicio = False
    'tXTcODIGO.TexT = FrmArticulo2.CmbCodigo.TexT
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
        'tXTcODIGO.TexT = FrmArticulo2.CmbCodigo.TexT
        Call Refresca_porCodigo
    End If
    If Dir(App.Path & "\crlf.on") <> "" Then
        Chk1.Value = vbChecked
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
                Lista.ListItems.Add Elementos, , .Fields("codigo")    '''
                Lista.ListItems(Elementos).SubItems(1) = .Fields("plu")
                Lista.ListItems(Elementos).SubItems(2) = .Fields("des_plu1")
                If Elementos = 1 Then
                    tXTcODIGO.TexT = .Fields("codigo")
                End If
                Elementos = Elementos + 1
                .Movenext
            Loop
            Lista.Enabled = True
            Refresca_Item
        End If
    End With

    Set Registro = Nothing

End Sub

Private Sub Refresca_Datos_Texto()
    Dim Registro As dao.Recordset
    Dim mBase As dao.Database
    Dim Miorden As String
    Dim i As Integer

    If lInicio Then Exit Sub
    If Not (IsNumeric(tXTcODIGO.TexT)) Then Exit Sub

    sTEXTO = ""
    RTB.TexT = ""
    If Option1(0).Value = True Then
        Miorden = "codigo"
    Else
        Miorden = "plu"
    End If

    Ini = 0

    Set mBase = AbrirBase

    If nIndexT = 0 Then
        'Set Registro = mBase.OpenRecordset("select * from articulo where borrado=false and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))
        Set Registro = mBase.OpenRecordset("select * from articulo where borrado=false and codigo=" & sCodigoG)  'Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))
        'where mostrador=" & CStr(Val(Combo1.TexT)) & " and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))
        With Registro
            If Not .EOF Then
                'RTB.TexT = Trim(.Fields("des_plu1")) & vbCrLf
                For i = 2 To 21
                    On Error GoTo sigue
                    If i > 9 Then
                        If i = 10 Then
                            If Not IsNull(.Fields("des_plu0")) Then
                                RTB.TexT = RTB.TexT & Trim(.Fields("des_plu0"))
                            End If
                        Else
                            If i = 11 Then
                                If Not IsNull(.Fields("des_pluX")) Then
                                    RTB.TexT = RTB.TexT & Trim(.Fields("des_pluX"))
                                End If
                            Else
                                If Not IsNull(.Fields("des_plu" & CStr(i - 1))) Then
                                    RTB.TexT = RTB.TexT & Trim(.Fields("des_plu" & CStr(i - 1)))
                                End If
                            End If
                        End If
                    Else
                        If Not IsNull(.Fields("des_plu" & CStr(i))) Then
                            RTB.TexT = RTB.TexT & Trim(.Fields("des_plu" & CStr(i)))
                        End If
                    End If
sigue:
                    On Error GoTo 0
                Next i

            End If
        End With

        Set Registro = Nothing

        Set Registro = mBase.OpenRecordset("select * from lintxt2040 where mostrador=" & CStr(Val(Combo1.TexT)) & " and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))

        With Registro
            If Not .EOF Then
                For i = 21 To 40
                    On Error GoTo sigue1
                    If Not IsNull(.Fields("txt_" & CStr(i))) Then
                        RTB.TexT = RTB.TexT & Trim(.Fields("txt_" & CStr(i)))
                    End If
sigue1:
                    On Error GoTo 0
                Next i
            End If
        End With

        Registro.Close
        Set Registro = Nothing

    Else

        'Registro.OpenRecordset ("select * from articulo where borrado=false and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))

        'With Registro
        '    If Not .EOF Then
        '        RTB.TexT = Trim(.Fields("des_plu1")) & vbCrLf
        '    End If
        'End With

        'Set Registro = Nothing

        Set Registro = mBase.OpenRecordset("select * from text15 where mostrador=" & CStr(Val(Combo1.TexT)) & " and " & Miorden & "=" & CStr(Val(tXTcODIGO.TexT)))

        With Registro
            If Not .EOF Then
                If Not IsNull(Trim(.Fields("text" & CStr(nIndexT)))) Then
                    'If Asc(Trim(.Fields("text" & CStr(nIndexT)))) <> 0 Then
                    RTB.TexT = RTB.TexT & Trim(.Fields("text" & CStr(nIndexT)))
                    'End If
                End If
            End If
        End With
        Registro.Close
        Set Registro = Nothing

    End If

    mBase.Close
    Set mBase = Nothing

    sTEXTO = RTB.TexT

End Sub

Private Sub Guarda_Texto()
    Dim Registro As dao.Recordset
    Dim Miorden As String
    Dim i As Integer
    Dim sOrg(40) As String
    Dim sRTB As String
    Dim Base As dao.Database
    Dim L2040 As Boolean

    On Error GoTo finGuarda

    sSecG = Trim(Combo1.TexT)


    sRTB = RTB.TexT


    Set Base = AbrirBase

    If nIndexT = 0 Then


        'If lNewArt Then
        '    For i = 0 To 19
        '        sOrg(i) = Mid(sRTB, (i * 25) + 1, 25)
        '    Next i
        '    For i = 20 To 39
        '        sOrg(i) = Mid(sRTB, (i * 34) + 1, 34)
        '    Next i
        'Else
        For i = 0 To 39
            sOrg(i) = Mid(sRTB, (i * 25) + 1, 25)
        Next i

        'End If
        L2040 = False
        For i = 20 To 39
            If Trim(sOrg(i)) <> "" Then
                L2040 = True
            End If
        Next i

        Set Registro = Base.OpenRecordset("select * from articulo where borrado=false and codigo=" & sCodigoG)
        With Registro
            If Not .EOF Then
                .Edit
                For i = 2 To 21
                    On Error GoTo sigue
                    If i > 9 Then
                        If i = 10 Then
                            .Fields("des_plu0") = sOrg(i - 2)
                        Else
                            If i = 11 Then
                                .Fields("des_pluX") = sOrg(i - 2)
                            Else
                                .Fields("des_plu" & CStr(i - 1)) = sOrg(i - 2)
                            End If
                        End If
                    Else
                        .Fields("des_plu" & CStr(i)) = sOrg(i - 2)
                    End If
sigue:
                    On Error GoTo 0
                Next i
                .Fields("tran_plu") = ""
                .Update
            End If
        End With

        If L2040 Then
            Set Registro = Nothing
            Set Registro = Base.OpenRecordset("select * from lintxt2040 where codigo=" & sCodigoG)
            With Registro
                If Not .EOF Then
                    .Edit
                Else
                    .AddNew
                    .Fields("codigo") = sCodigoG
                    .Fields("plu") = sPluG
                    .Fields("mostrador") = sSecG
                End If

                For i = 21 To 40
                    On Error GoTo sigue1
                    .Fields("txt_" & CStr(i)) = sOrg(i - 1)
sigue1:
                    On Error GoTo 0
                Next i
                .Fields("tran_txt2") = ""
                .Update
            End With
        Else
            Set Registro = Nothing
            Set Registro = Base.OpenRecordset("select * from lintxt2040 where codigo=" & sCodigoG)
            If Not Registro.EOF Then
                Registro.Delete
            End If
        End If

    Else

        Set Registro = Base.OpenRecordset("select * from articulo where borrado=false and codigo=" & sCodigoG)
        With Registro
            If Not .EOF Then
                .Edit
                .Fields("tran_plu") = ""
                .Update
                L2040 = True
            Else
                L2040 = False
            End If
        End With
        Registro.Close
        Set Registro = Nothing
        If L2040 Then
            Set Registro = Base.OpenRecordset("select * from text15 where codigo=" & sCodigoG)
            With Registro
                If Not .EOF Then
                    .Edit
                Else
                    .AddNew
                    .Fields("codigo") = sCodigoG
                    .Fields("plu") = sPluG
                    .Fields("mostrador") = sSecG
                End If
                If Trim(sRTB) <> "" Then
                    #If BALSAM = True Then
                        .Fields("text" & CStr(nIndexT)) = Mid(sRTB, 1, 1000)
                    #Else
                        .Fields("text" & CStr(nIndexT)) = Mid(sRTB, 1, 1500)
                    #End If
                Else
                    .Fields("text" & CStr(nIndexT)) = " "
                End If
                .Fields("tran_txt15") = ""
                .Update
            End With
        End If
    End If

    Base.Close
    Set Base = Nothing

    MsgBox CargaCadena(439), vbInformation

finGuarda:
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox CargaCadena(969), vbCritical
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Mybase As dao.Database
    Dim cm, cm1, cm2 As String
    Set Mybase = AbrirBase
    Mybase.Execute "delete from text15 where (trim(text1)='' or isnull(text1))and (trim(text2)='' or isnull(text2)) and (trim(text3)='' or isnull(text3)) and (trim(text4)='' or isnull(text4)) and (trim(text5)='' or isnull(text5))"
    cm1 = "delete from lintxt2040 where "
    cm2 = "update articulo set tran_plu='' where codigo in (select codigo from lintxt2040 where "
    cm = "(trim(txt_21)='' or isnull(txt_22)) and (trim(txt_23)='' or isnull(txt_23)) and "
    cm = cm & "(trim(txt_24)='' or isnull(txt_24)) and (trim(txt_25)='' or isnull(txt_25)) and (trim(txt_26)='' or isnull(txt_26)) and "
    cm = cm & "(trim(txt_27)='' or isnull(txt_27)) and (trim(txt_28)='' or isnull(txt_28)) and (trim(txt_29)='' or isnull(txt_29)) and (trim(txt_30)='' or isnull(txt_30)) and "
    cm = cm & "(trim(txt_31)='' or isnull(txt_31)) and (trim(txt_32)='' or isnull(txt_32)) and (trim(txt_33)='' or isnull(txt_33)) and (trim(txt_34)='' or isnull(txt_34)) and "
    cm = cm & "(trim(txt_35)='' or isnull(txt_35)) and (trim(txt_36)='' or isnull(txt_36)) and (trim(txt_37)='' or isnull(txt_37)) and (trim(txt_38)='' or isnull(txt_38)) and "
    cm = cm & "(trim(txt_39)='' or isnull(txt_39)) and (trim(txt_40)='' or isnull(txt_40)) and (trim(txt_41)='' or isnull(txt_41)) and (trim(txt_42)='' or isnull(txt_42)) and "
    cm = cm & "(trim(txt_43)='' or isnull(txt_43)) and (trim(txt_44)='' or isnull(txt_44)) and (trim(txt_45)='' or isnull(txt_45)) and (trim(txt_46)='' or isnull(txt_46)) and "
    cm = cm & "(trim(txt_47)='' or isnull(txt_47)) and (trim(txt_48)='' or isnull(txt_48)) and (trim(txt_49)='' or isnull(txt_49)) and (trim(txt_50)='' or isnull(txt_50)) and "
    cm = cm & "(trim(txt_51)='' or isnull(txt_51)) and (trim(txt_52)='' or isnull(txt_52)) and (trim(txt_53)='' or isnull(txt_53)) and (trim(txt_54)='' or isnull(txt_54)) and "
    cm = cm & "(trim(txt_55)='' or isnull(txt_55)) and (trim(txt_56)='' or isnull(txt_56)) and (trim(txt_57)='' or isnull(txt_57)) and (trim(txt_58)='' or isnull(txt_58)) and "
    cm = cm & "(trim(txt_59)='' or isnull(txt_59)) and (trim(txt_60)='' or isnull(txt_60)) and (trim(txt_61)='' or isnull(txt_61)) and (trim(txt_62)='' or isnull(txt_62)) and "
    cm = cm & "(trim(txt_63)='' or isnull(txt_63)) and (trim(txt_64)='' or isnull(txt_64)) and (trim(txt_65)='' or isnull(txt_65)) and (trim(txt_66)='' or isnull(txt_66)) and "
    cm = cm & "(trim(txt_67)='' or isnull(txt_67)) and (trim(txt_68)='' or isnull(txt_68)) and (trim(txt_69)='' or isnull(txt_69)) and (trim(txt_70)='' or isnull(txt_70)) and "
    cm = cm & "(trim(txt_71)='' or isnull(txt_71)) and (trim(txt_72)='' or isnull(txt_72)) and (trim(txt_73)='' or isnull(txt_73)) and (trim(txt_74)='' or isnull(txt_74)) and "
    cm = cm & "(trim(txt_75)='' or isnull(txt_75)) and (trim(txt_76)='' or isnull(txt_76)) and (trim(txt_77)='' or isnull(txt_77)) and (trim(txt_78)='' or isnull(txt_78)) and "
    cm = cm & "(trim(txt_79)='' or isnull(txt_79)) and (trim(txt_80)='' or isnull(txt_80))"
    cm2 = cm2 & cm & ")"
    Mybase.Execute cm2
    cm1 = cm1 & cm
    Mybase.Execute cm1
    Mybase.Close
    Set Mybase = Nothing
    frmEpelsa.Enabled = True
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

Private Sub Refresca_Item()

    On Error GoTo noHayItem
    sCodigoG = Lista.SelectedItem.TexT
    sPluG = Lista.SelectedItem.SubItems(1)
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
        'tXTcODIGO.TexT = Trim(FrmArticulo2.CmbCodigo.TexT)  'Lista.SelectedItem.TexT
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


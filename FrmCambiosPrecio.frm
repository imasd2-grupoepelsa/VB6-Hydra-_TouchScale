VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCambiosPrecio 
   Caption         =   "Cambios de Precio"
   ClientHeight    =   5565
   ClientLeft      =   1740
   ClientTop       =   1920
   ClientWidth     =   7905
   Icon            =   "FrmCambiosPrecio.frx":0000
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Paso automático al siguiente elemento de la lista"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   7572
   End
   Begin VB.TextBox tXTcODIGO 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   11
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
      Left            =   6480
      TabIndex        =   9
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox TxtNuevoPrecio 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   1560
      TabIndex        =   12
      Top             =   5160
      Width           =   972
   End
   Begin VB.TextBox TxtPrecioActual 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   1560
      TabIndex        =   6
      Top             =   4920
      Width           =   972
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
      NumItems        =   5
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
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Precio Actual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nuevo Precio"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Desc.:"
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
      TabIndex        =   17
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Lbl100g 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   372
      Left            =   2640
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   3492
   End
   Begin VB.Label LblcoDIGO 
      Caption         =   "Código"
      Height          =   252
      Left            =   120
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   4680
      Width           =   5172
   End
   Begin VB.Label Label3 
      Caption         =   "Nuevo Precio : "
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "Precio Actual : "
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   1332
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
Attribute VB_Name = "FrmCambiosPrecio"
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


Private Sub Me_Init()
    Dim bucle As Long
    Me_Width = Me.Width
    Me_Height = Me.Height
    ReDim ct_top(Me.Controls.Count)
    ReDim CT_Left(Me.Controls.Count)
    ReDim CT_Width(Me.Controls.Count)
    ReDim CT_Height(Me.Controls.Count)
    For bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        ct_top(bucle) = Me.Controls(bucle).top
        CT_Left(bucle) = Me.Controls(bucle).left
        CT_Width(bucle) = Me.Controls(bucle).Width
        CT_Height(bucle) = Me.Controls(bucle).Height
        On Error GoTo 0
    Next bucle
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
           Then
            Me.Controls(bucle).Height = RelW * CT_Height(bucle)
            CT_Height(bucle) = Me.Controls(bucle).Height
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
    Check1.Caption = CargaCadena(1198)
    Label1.Caption = CargaCadena(302)
    Option1(0).Caption = CargaCadena(656)
    Option1(1).Caption = CargaCadena(657)
    LblcoDIGO.Caption = CargaCadena(69)
    Label2.Caption = CargaCadena(658)
    Label3.Caption = CargaCadena(659)
    Command1(0).Caption = CargaCadena(287)
    Command1(1).Caption = CargaCadena(7)
    FrmCambiosPrecio.Caption = CargaCadena(24)
    Lista.ColumnHeaders(1).TexT = CargaCadena(69)
    Lista.ColumnHeaders(2).TexT = CargaCadena(68)
    Lista.ColumnHeaders(3).TexT = CargaCadena(70)
    Lista.ColumnHeaders(4).TexT = CargaCadena(658)
    Lista.ColumnHeaders(5).TexT = CargaCadena(659)
    Lbl100g.Caption = CargaCadena(1194)
End Sub
Private Sub Refresca_porCodigo()

    Dim Registro As New RecordNet
    Dim MiNumero As Long
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
            TxtPrecioActual.TexT = ""
            TxtNuevoPrecio.TexT = ""
            tXTcODIGO.SetFocus
        Else
            Combo1.TexT = .Fields("codi_ident")
            Refresca_Datos
            If Option1(0).Value Then
                Do While Val(Lista.SelectedItem.TexT) <> MiNumero
                    Lista.ListItems(Lista.SelectedItem.Index + 1).Selected = True
                    Lista.SelectedItem.EnsureVisible
                Loop
                Refresca_Item
            Else
                Do While Val(Lista.SelectedItem.SubItems(1)) <> MiNumero
                    Lista.ListItems(Lista.SelectedItem.Index + 1).Selected = True
                    Lista.SelectedItem.EnsureVisible
                Loop
                Refresca_Item
            End If
            TxtNuevoPrecio.SetFocus
        End If
    End With


End Sub

Private Sub Text1_Change()
    Refresca_Datos
End Sub
Private Sub Combo1_Click()
    Text1.TexT = ""

    If Combo1.TexT <> "" Then
        Refresca_Datos
        TxtNuevoPrecio.Enabled = True
    Else
        TxtNuevoPrecio.Enabled = False
    End If
    tXTcODIGO.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Miarticulo As DB_Articulo
    Select Case Index
    Case 0
        If Trim(TxtNuevoPrecio.TexT <> "") Then
            If RechazarPlu0 Then
                If Val(TxtNuevoPrecio.TexT) = 0 Then
                    MsgBox CargaCadena(368), vbCritical
                    Exit Sub
                End If
            End If
            Miarticulo.codigo = Lista.SelectedItem.TexT
            Miarticulo.Mostrador = Combo1.TexT
            Miarticulo.precio = ToDouble(TxtNuevoPrecio.TexT)
            Select Case Modificacion_Articulo_Precio(Miarticulo)
            Case 1
                LblInfo.Caption = CargaCadena(436)  '"Precio Modificado, envíe modificaciones"
                'If xlcampo Then
                '    If descAuto Then
                '        ModificacionesPendientes(0) = (ModificacionesPendientes(0) Or 128)
                '    Else
                '        lCogeTiquet = False
                '        AN_Articulos True
                '        lCogeTiquet = True
                '    End If
                'End If
            Case Else
                LblInfo.Caption = CargaCadena(660)  '"Sin Modificación"
            End Select
        Else
            LblInfo.Caption = CargaCadena(660)  '"Sin Modificación"
        End If
        If TxtNuevoPrecio.TexT <> "" Then
            Lista.SelectedItem.SubItems(4) = TxtNuevoPrecio.TexT
        Else
            Lista.SelectedItem.SubItems(4) = TxtPrecioActual.TexT
        End If
        If Check1.Value = vbChecked Then
            If Lista.SelectedItem.Index < Lista.ListItems.Count Then
                Lista.ListItems(Lista.SelectedItem.Index + 1).Selected = True
            Else
                Lista.ListItems(Lista.SelectedItem.Index).Selected = True
            End If
            Lista.ListItems(Lista.SelectedItem.Index).EnsureVisible
            Refresca_Item
            TxtNuevoPrecio.SetFocus
            'Refresca_Datos
        Else
            If Lista.SelectedItem.Index < Lista.ListItems.Count Then
                'Lista.ListItems.Item(Lista.SelectedItem.Index + 1).Selected = True
                'Refresca_Item
                tXTcODIGO.TexT = ""
                TxtNuevoPrecio.TexT = ""
                tXTcODIGO.SetFocus
            End If
        End If
    Case 1
        If Dir(App.Path & "\gigante.txt") <> "" Then
            Call generaGIAN
        End If
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
    Dim Retorno As Boolean
    Dim MyPref As String
    CambiarIdioma
    Text1.TexT = ""
    sMyOrder = ""
    sLastOrder = ""
    Me_Init
    Me.Height = Screen.Height * 0.89
    Me.Icon = Form2.Icon
    Retorno = Refresca_Combo
    Lista.HotTracking = True
    'tXTcODIGO.Enabled = False
    MyPref = RPreferencias("CHGPRECIO1", 1)
    If MyPref <> "" Then
        If left(MyPref, 1) <> "0" Then
            Option1(1).Value = True
        Else
            Option1(0).Value = True
        End If
    End If
    MyPref = RPreferencias("CHGPRECIO2", 1)
    If MyPref <> "" Then
        If left(MyPref, 1) <> "1" Then
            Check1.Value = vbUnchecked
        Else
            Check1.Value = vbChecked
        End If
    End If
    If Retorno Then Refresca_Datos
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
    Lbl100g.Visible = False
    NPrecios100g = 0
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
    If Text1.TexT = "" Then
        Registro.OpenRecordset _
                ("select codigo,plu,des_plu1,euros,precio,prc100g from articulo where borrado=false and codi_ident=" & Val(Combo1.TexT) & " order by " & Miorden)
        If sMyOrder <> "" Then
            sLastOrder = sMyOrder
        End If
    Else
        Registro.OpenRecordset _
                ("select codigo,plu,des_plu1,euros,precio,prc100g from articulo where borrado=false and codi_ident=" & Val(Combo1.TexT) & " and instr(1,des_plu1,'" & Trim(Text1.TexT) & "')<>0 order by " & Miorden)

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
                If UsaEuro Then
                    Lista.ListItems(Elementos).SubItems(3) = ToString(.Fields("Euros"))
                Else
                    Lista.ListItems(Elementos).SubItems(3) = ToString(.Fields("precio"))
                End If
                If .Fields("prc100g") Then
                    NPrecios100g = NPrecios100g + 1
                    ReDim Preserve Codigos100g(NPrecios100g)
                    If Option1(0).Value Then
                        Codigos100g(NPrecios100g - 1) = .Fields("codigo")
                    Else
                        Codigos100g(NPrecios100g - 1) = .Fields("plu")
                    End If
                End If
                Lista.ListItems(Elementos).SubItems(4) = ""
                Elementos = Elementos + 1

                .Movenext
            Loop
            Lista.Enabled = True
            'Lista.SetFocus
            Refresca_Item
            Command1(0).Enabled = True
        Else
            Lista.Enabled = False
            Command1(0).Enabled = False
        End If
    End With


End Sub



Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Option1(1).Value = True Then
        WPreferencias "CHGPRECIO1", 1, "1"
    Else
        WPreferencias "CHGPRECIO1", 1, "0"
    End If
    If Check1.Value = vbChecked Then
        WPreferencias "CHGPRECIO2", 1, "1"
    Else
        WPreferencias "CHGPRECIO2", 1, "0"
    End If
    frmEpelsa.Enabled = True
    Me.Enabled = False
End Sub

Private Sub Lista_Click()
    Refresca_Item
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Lista.ColumnHeaders(1).TexT = CargaCadena(69)
    Lista.ColumnHeaders(2).TexT = CargaCadena(68)
    Lista.ColumnHeaders(3).TexT = CargaCadena(70)
    Lista.ColumnHeaders(4).TexT = CargaCadena(658)

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
    Text1.TexT = ""
    If Combo1.TexT <> "" Then Refresca_Datos
    If Option1(0).Value = True Then
        LblcoDIGO.Caption = CargaCadena(69)  '"Código"
    Else
        LblcoDIGO.Caption = CargaCadena(68)  '"PLU"
    End If
End Sub
Private Sub Refresca_Item()
    If Option1(0).Value = True Then
        tXTcODIGO.TexT = Lista.SelectedItem.TexT
    Else
        tXTcODIGO.TexT = Lista.SelectedItem.SubItems(1)
    End If
    TxtPrecioActual.TexT = Lista.SelectedItem.SubItems(3)
    TxtNuevoPrecio.TexT = Lista.SelectedItem.SubItems(4)
    If IsNumeric(tXTcODIGO.TexT) Then CompruebaPrecio100 (Val(tXTcODIGO.TexT))

End Sub
Private Sub CompruebaPrecio100(MyCod As Long)
    Dim bucle As Long
    Lbl100g.Visible = False
    For bucle = 0 To NPrecios100g - 1
        If MyCod = Codigos100g(bucle) Then
            Lbl100g.Visible = True
            Exit For
        End If
    Next bucle
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
        TxtNuevoPrecio.SetFocus
        Exit Sub
    End If
    If Option1(0).Value = True Then
        tXTcODIGO.Locked = Checktexto(KeyAscii, 6, tXTcODIGO.TexT, True, False)
    Else
        tXTcODIGO.Locked = Checktexto(KeyAscii, 4, tXTcODIGO.TexT, True, False)
    End If

End Sub

Private Sub tXTcODIGO_LostFocus()
    LblInfo.Caption = ""
    Refresca_porCodigo
End Sub

Private Sub TxtNuevoPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TxtNuevoPrecio_KeyPress(KeyAscii As Integer)
    Dim Comas As Boolean
    Dim bucle As Integer
    If "1,1" = 1.1 Then
        If KeyAscii = 46 Then KeyAscii = 44
    Else
        If KeyAscii = 44 Then KeyAscii = 46
    End If
    If decimales > 0 Or UsaEuro Then
        TxtNuevoPrecio.Locked = Checktexto(KeyAscii, 7, TxtNuevoPrecio.TexT, True, True)
    Else
        TxtNuevoPrecio.Locked = Checktexto(KeyAscii, 7, TxtNuevoPrecio.TexT, True, False)
    End If
    If KeyAscii = 13 And Command1(0).Enabled = True Then Command1_Click (0)    'Command1(0).SetFocus
    If KeyAscii = 8 Then Exit Sub
    Comas = False
    For bucle = 1 To Len(TxtNuevoPrecio.TexT)
        If Mid(TxtNuevoPrecio.TexT, bucle, 1) = "," Or Mid(TxtNuevoPrecio.TexT, bucle, 1) = "." Then
            Comas = True
            Exit For
        End If
    Next bucle
    If Comas Then
        'TxtNuevoPrecio.Locked = Checktexto(KeyAscii, 10, TxtNuevoPrecio.Text, True, True)
        If (Not UsaEuro And Len(TxtNuevoPrecio.TexT) - bucle >= decimales) _
           Or (UsaEuro And Len(TxtNuevoPrecio.TexT) - bucle >= 2) Then
            TxtNuevoPrecio.Locked = True
        End If
    End If
End Sub

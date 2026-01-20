VERSION 5.00
Begin VB.Form FrmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Teclas"
   ClientHeight    =   4620
   ClientLeft      =   12210
   ClientTop       =   960
   ClientWidth     =   4215
   Icon            =   "FrmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4215
   Visible         =   0   'False
   Begin VB.TextBox TxtBy 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdKillData 
      Caption         =   "Eliminar Datos."
      Height          =   855
      Left            =   2760
      Picture         =   "FrmSelect.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton CmdEliQ 
      Caption         =   "Eliminar Calidad."
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton CmdNext 
      Height          =   3015
      Left            =   3600
      Picture         =   "FrmSelect.frx":22B4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CmdPrev 
      Height          =   3015
      Left            =   0
      Picture         =   "FrmSelect.frx":31F6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CmdPrevOk 
      Caption         =   "Guardar y Siguiente."
      Height          =   855
      Left            =   120
      Picture         =   "FrmSelect.frx":4138
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label LblBy 
      Caption         =   "Código:"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image ImgTec 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   0
      Left            =   480
      Picture         =   "FrmSelect.frx":457A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label LblTec1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NARANJA PREMIUM"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label LblTec2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
End
Attribute VB_Name = "FrmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLocateBy As Integer

Private Sub CmdEliQ_Click()
    'TeclaTmp.qLevel = 0
    'TeclaTmp.qLevelName = ""
    'Set FrmSelect.ImgQ(0).Picture = Nothing
    'FrmSelect.ImgQ(0).Visible = False
End Sub
            
Private Sub CmdKillData_Click()
Dim MyBaseDES As DAO.Database
Dim MyRsDes As DAO.Recordset

    TxtBy.Text = ""

    Set MyBaseDES = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsDes = MyBaseDES.OpenRecordset("select * from items where key=" & CStr(TeclaTmp.Key))
    If Not MyRsDes.EOF Then 'Debe existir
        MyRsDes.Delete
    End If
    MyRsDes.Close
    Set MyRsDes = Nothing
    MyBaseDES.Close
    Set MyBaseDES = Nothing
    
    Call FrmKeyTouch.cargaSubGrupoTeclas(nGrupoActual, nSubGrupoActual)
    
    Call CmdNext_Click
    Call CmdPrev_Click

End Sub

'TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey = MyRsORG.Fields("inikey")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey = MyRsORG.Fields("endkey")
'nGrupoActual, nSubGrupoActual
Private Sub CmdNext_Click()
TxtBy.Text = ""

    If (Val(LblTec2(0).Caption + 1)) <= TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).EndKey Then 'máxima tecla
        'Set FrmSelect.ImgQ(0).Picture = Nothing
        FrmKeyTouch.imgTecClick 1000, Val(LblTec2(0).Caption + 1)
    Else
        'Set FrmSelect.ImgQ(0).Picture = Nothing
        FrmKeyTouch.imgTecClick 1000, TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey
    End If
    
    'If TeclaTmp.qLevel <> 0 Then
    '    CmdEliQ.Enabled = True
    'Else
        CmdEliQ.Enabled = False
    'End If
    
End Sub
Private Sub CmdPrev_Click()
TxtBy.Text = ""
    If (Val(LblTec2(0).Caption - 1)) >= TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey Then  'mínima tecla
        'Set FrmSelect.ImgQ(0).Picture = Nothing
        FrmKeyTouch.imgTecClick 1000, Val(LblTec2(0).Caption - 1)
    Else
        'Set FrmSelect.ImgQ(0).Picture = Nothing
        FrmKeyTouch.imgTecClick 1000, TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).EndKey
    End If

    'If TeclaTmp.qLevel <> 0 Then
    '    CmdEliQ.Enabled = True
    'Else
        CmdEliQ.Enabled = False
    'End If

End Sub

Private Sub CmdPrevOk_Click()
Dim MyBaseORG As DAO.Database
Dim MyBaseDES As DAO.Database
Dim MyRsORG As DAO.Recordset
Dim MyRsDes As DAO.Recordset
Dim myRsTMP As DAO.Recordset
Dim sS As String
Dim lEsAlta As Boolean
Dim nPlu As Integer

    TxtBy.Text = ""
    
    Set MyBaseORG = DAO.OpenDatabase(App.Path & "\dbase.mdb")
    Set MyBaseDES = DAO.OpenDatabase(App.Path & "\touchscale.mdb")
    Set MyRsORG = MyBaseORG.OpenRecordset("select * from articulo where codigo=" & CStr(TeclaTmp.Code))
    If Not MyRsORG.EOF Then 'Debe existir
        Set MyRsDes = MyBaseDES.OpenRecordset("select * from items where code=" & CStr(TeclaTmp.Code))
        If MyRsDes.EOF Then 'alta
            If lAllItem Then
                MyRsDes.AddNew
                lEsAlta = True
            Else
                Exit Sub
            End If
        Else ' modificación
            MyRsDes.Edit
        End If
        ' entrada datos
        If lEsAlta Then
            Set myRsTMP = MyBaseDES.OpenRecordset("select max(plu) as mplu from items")
            If myRsTMP.EOF Then
                nPlu = 1
            Else
                nPlu = myRsTMP.Fields("mplu") + 1
            End If
            myRsTMP.Close
            Set myRsTMP = Nothing
            MyRsDes.Fields("plu") = nPlu
            MyRsDes.Fields("ean13") = MyRsORG.Fields("art_cb")
        End If
            
        MyRsDes.Fields("code") = TeclaTmp.Code
        MyRsDes.Fields("name") = Mid(TeclaTmp.Name, 1, 25)
        MyRsDes.Fields("shortname") = TeclaTmp.ShortName
        MyRsDes.Fields("family") = MyFam
        MyRsDes.Fields("codi_ident") = nSecSer
        MyRsDes.Fields("counter") = nSecSer
        MyRsDes.Fields("section") = nSecSer
        
        MyRsDes.Fields("icon_path") = App.Path & "\images"
        MyRsDes.Fields("grp") = MyFam
        MyRsDes.Fields("subgroup") = TeclaTmp.SubGroup
        MyRsDes.Fields("set") = 1
        MyRsDes.Fields("key") = TeclaTmp.Key

        'MyRsDes.Fields("text") = MyRsORG.Fields("des_plu2") & MyRsORG.Fields("des_plu3") & MyRsORG.Fields("des_plu4") & _
        '                          MyRsORG.Fields("des_plu5") & MyRsORG.Fields("des_plu6") & MyRsORG.Fields("des_plu7") & _
        '                          MyRsORG.Fields("des_plu8") & MyRsORG.Fields("des_plu9") & MyRsORG.Fields("des_plu0")
        
        MyRsDes.Fields("text") = ""
        MyRsDes.Fields("icon_name") = Trim(TeclaTmp.icon_name)
        
        MyRsDes.Fields("price") = MyRsORG.Fields("euros")
        If MyRsORG.Fields("codi_pes") = "W" Then
            MyRsDes.Fields("saleform") = 1
        Else
            MyRsDes.Fields("saleform") = 0
        End If
        MyRsDes.Fields("vat_type") = MyRsORG.Fields("tipo_iva")
        MyRsDes.Fields("vat_value") = 0
        'Set myRsTMP = MyBaseORG.OpenRecordset("select * from iva where tipo=" & CStr(MyRsORG.Fields("tipo_iva")))
        'If myRsTMP.EOF Then
        '   MyRsDes.Fields("vat_value") = 0
        'Else
        '    MyRsDes.Fields("vat_value") = myRsTMP.Fields("valor")
        'End If
        'myRsTMP.Close
        'Set myRsTMP = Nothing

        MyRsDes.Fields("expiring") = MyRsORG.Fields("caducidad")
        MyRsDes.Fields("bestbefore") = 0 'no tenemos dato
        MyRsDes.Fields("tare") = MyRsORG.Fields("tara")
        MyRsDes.Fields("label_value") = 0
        MyRsDes.Fields("label_text") = ""
        
        MyRsDes.Fields("qualitynr") = 0
        MyRsDes.Fields("qualityname") = "NO"
        
        'n = TeclaTmp.Code
        'n = TeclaTmp.Counter
        'n = TeclaTmp.Group
        'n = TeclaTmp.icon_name
        'n = TeclaTmp.Key
        'n = TeclaTmp.Name
        'n = TeclaTmp.PLU
        'n = TeclaTmp.qLevel
        'n = TeclaTmp.qLevelName
        'n = TeclaTmp.Section
        'n = TeclaTmp.ShortName
        'n = TeclaTmp.SubGroup
        ''''''''''
        MyRsDes.Update
        ''''''''''
    
    
    'Else
    '    If TeclaTmp.Code = 0 Then
    '        MsgBox "No ha seleccionado ARTÍCULO...", vbExclamation
    '    Else
    '        MsgBox "Inconsistencia en bases de datos...", vbCritical
    '        'cerrar lo abierto y aplicación
    '    End If
    End If

Salir:
    On Error Resume Next
    MyRsDes.Close
    Set MyRsDes = Nothing
    MyRsORG.Close
    Set MyRsDes = Nothing
    MyBaseORG.Close
    Set MyBaseORG = Nothing
    MyBaseDES.Close
    Set MyBaseDES = Nothing
    On Error GoTo 0
    
    Call FrmKeyTouch.cargaSubGrupoTeclas(nGrupoActual, nSubGrupoActual)
    
    Call CmdNext_Click

End Sub

Public Sub Form_Activate()
Dim SetWindowPosReturn As Long
If lAcT = True Then
    SetWindowPosReturn = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
        FrmKeyTouch.CmdEnvCambios.Visible = False
        FrmKeyTouch.CmdExit.Visible = False
        FrmKeyTouch.TxtVer.Visible = False
        lAcT = False
       
End If

End Sub

Private Sub Form_Load()
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).IniKey = MyRsORG.Fields("inikey")
'TTecGrupos.Indice(nnF).SubGrupo(nnnF).EndKey = MyRsORG.Fields("endkey")
'nGrupoActual, nSubGrupoActual
    Me.Caption = "TECLAS:" & Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).IniKey, "000") & " Hasta " & _
                             Format(TTecGrupos.Indice(nGrupoActual).SubGrupo(nSubGrupoActual).EndKey, "000")
    If TeclaTmp.qLevel <> 0 Then
        CmdEliQ.Enabled = True
    Else
        CmdEliQ.Enabled = False
    End If
    'If FrmSelect.Visible = True Then
    '    FrmKeyTouch.CmdEnvCambios.Visible = False
    '    FrmKeyTouch.CmdEnvTodo.Visible = False
    '    FrmKeyTouch.CmdExit.Visible = False
    'End If
    nLocateBy = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim n As Integer
    For n = 0 To 39 'veraquí
        FrmKeyTouch.LblTec1(n).Enabled = True
        FrmKeyTouch.LblTec2(n).Enabled = True
        FrmKeyTouch.ImgTec(n).Enabled = True
        'If FrmKeyTouch.ImgQ(n).Picture <> 0 Then
        '
        'End If
    Next n
    
    FrmKeyTouch.CmdEnvCambios.Visible = True
    FrmKeyTouch.CmdExit.Visible = True
    FrmKeyTouch.TxtVer.Visible = True
    
    FrmKeyTouch.Lista.Sorted = False

    FrmKeyTouch.Lista.Visible = False
    'FrmKeyTouch.ListaImg.Visible = False
    'FrmKeyTouch.ListaCalidad.Visible = False
    lAcT = True
    
End Sub

Private Sub ImgTec_Click(Index As Integer)
    If nLocateBy = 0 Then
        nLocateBy = 1
    Else
        nLocateBy = 0
    End If
    If nLocateBy = 0 Then
        LblBy.Caption = "Código:"
        TxtBy.MaxLength = 6
    Else
        LblBy.Caption = "PLU:"
        TxtBy.MaxLength = 4
    End If
End Sub

Private Sub ImgTec_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ImgTec(0).ToolTipText = "Código:" & CStr(TeclaTmp.Code) & " PLU:" & CStr(TeclaTmp.PLU) & " Imagen:" & TeclaTmp.icon_name
End Sub


Private Sub TxtBy_KeyPress(KeyAscii As Integer)
Dim LI As ListItem
Dim PrevIndex As Long
Dim SearchText As String

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        'If nLocateBy = 0 Then
        '    With FrmKeyTouch.Lista
        '        Set lI = FrmKeyTouch.Lista.FindItem(Format(Val(Trim(TxtBy.Text)), "000000"), lvwText, 1, lvwPartial)
        '        'If FindItem(txtToFind, lvwSubItem).ListSubItems(1) = Format(Val(Trim(TxtBy.Text)), "000000") Then
        '        If Not lI Is Nothing Then
        '            .ListItems(lI.Index).Selected = True
        '            .SetFocus
        '            .ListItems(lI.Index).EnsureVisible
        '            FrmKeyTouch.Lista_ItemClick lI
        '            .Refresh
        '        End If
        '    End With
        '    Set lI = Nothing
        'Else
        '    With FrmKeyTouch.Lista
        '        Set lI = FrmKeyTouch.Lista.FindItem(Format(Val(Trim(TxtBy.Text)), "0000"), lvwText, 2, lvwPartial)
        '        If Not lI Is Nothing Then
        '            .ListItems(lI.Index).Selected = True
        '            .SetFocus
        '            FrmKeyTouch.Lista_ItemClick lI
        '        End If
        '    End With
        '    Set lI = Nothing
       '
        'End If
        'Columnindex = 1
        PrevIndex = 1
        If nLocateBy = 0 Then
            With FrmKeyTouch.Lista
                Set LI = FrmKeyTouch.Lista.FindItem(Format(Val(Trim(TxtBy.Text)), "000000"), lvwText, 1, lvwPartial)
                If Not LI Is Nothing Then
                    .ListItems(LI.Index).Selected = True
                    .SetFocus
                    .ListItems(LI.Index).EnsureVisible
                    FrmKeyTouch.Lista_ItemClick LI
                    .Refresh
                Else
                    MsgBox "Artículo NO LOCALIZADO..."
                End If
            End With
            Set LI = Nothing
        Else
            Do
                SearchText = Format(Val(TxtBy.Text), "0000")
                Set LI = FrmKeyTouch.Lista.FindItem(SearchText, lvwSubItem, PrevIndex, False)
                If Not (LI Is Nothing) Then
                    'if the column we are searching for has the searchtext, then we are done
                    If LI.ListSubItems(1).Text = SearchText Then
                        FrmKeyTouch.Lista.ListItems(LI.Index).Selected = True
                        FrmKeyTouch.Lista.SetFocus
                        FrmKeyTouch.Lista.ListItems(LI.Index).EnsureVisible
                        FrmKeyTouch.Lista_ItemClick LI
                        FrmKeyTouch.Lista.Refresh

                        Exit Do
                    Else
                        'if it is the last item on the list and we still haven't found an item,
                        'it doesn't exist.
                        If FrmKeyTouch.Lista.ListItems.Count = LI.Index Then Exit Do
                        'or else just keep looping, adding previndex so we don't search that item
                        'again
                        PrevIndex = LI.Index + 1
                    End If
                Else
                    'None found, so just exit
                    MsgBox "Artículo NO LOCALIZADO..."
                    Exit Do
                End If
            Loop
        End If
        
        TxtBy.Text = ""
        
    End If
End Sub

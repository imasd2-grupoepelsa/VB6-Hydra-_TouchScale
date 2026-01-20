VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFamMos 
   BorderStyle     =   0  'None
   Caption         =   "Asociación de Familias a Mostradores"
   ClientHeight    =   6315
   ClientLeft      =   5505
   ClientTop       =   2370
   ClientWidth     =   10485
   Icon            =   "FrmFamMos.frx":0000
   LinkTopic       =   "Asociación de Familias a Mostradores"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtOrden 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   12
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000018&
      Caption         =   "Command4"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000018&
      Caption         =   "Command4"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1212
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
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   6240
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
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox CmbFamilias 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1932
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "FrmFamMos.frx":1E72
      Left            =   120
      List            =   "FrmFamMos.frx":1E79
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1440
      Width           =   6375
   End
   Begin MSComctlLib.TreeView Arbol 
      Height          =   6255
      Left            =   6720
      TabIndex        =   15
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11033
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "( press CR or tab to assign)"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Pos.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label LblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
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
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Label Label3 
      Caption         =   "Mostradores:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   5175
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
Attribute VB_Name = "FrmFamMos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private YaExiste As Boolean
Private nEvents As Long
Private WithEvents MyEvento As ClsEvento
Attribute MyEvento.VB_VarHelpID = -1
Dim MyX, MyY As Single
Dim lEnCarga As Boolean
Dim lRefresca As Boolean

Private Sub Cambiar_idioma()
    Me.Caption = "Asociación de Familias a Mostradores" 'CargaCadena(678)
    Label1.Caption = CargaCadena(67)
    Label2.Caption = CargaCadena(539)
    Label3.Caption = CargaCadena(593)
    Command1.Caption = CargaCadena(287)
    Command2.Caption = CargaCadena(7)
    'Command3.Caption = CargaCadena(38)
    Command4(0).Caption = CargaCadena(311)
    Command4(1).Caption = CargaCadena(1045)
    Label5.Caption = CargaCadena(1460) & "/ 0=No Pos."
    Label4.Caption = CargaCadena(1377)
End Sub
Private Sub CmbFamilias_Click()
    lEnCarga = True
    Refresca_Datos
    Refresca_Secciones
    lEnCarga = False
End Sub
Private Sub Refresca_Datos()
    Dim Registro As New RecordNet
    Dim Token As Long
    
    
    Text1.TexT = ""
    If IsNumeric(Val(CmbFamilias.TexT)) Then
        'sQueNombreImagen = Trim((CmbFamilias.TexT))
        Registro.OpenRecordset "select codi_fam,txt_fam,imagen from fam_code where codi_fam=" & Val(CmbFamilias.TexT)
        
        If Not Registro.EOF Then
            If Not IsNull(Registro.Fields("txt_fam")) Then
                Text1.TexT = Registro.Fields("txt_fam")
            Else
                Text1.TexT = "Familia " & CStr(Val(CmbFamilias.TexT))
            End If
        End If
    End If
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
'CUIDADO ERROR
'Private Sub CmbFamilias_LostFocus()
'    Refresca_Datos
'    Refresca_Secciones
'End Sub
''''''''''''''''

'Private Function Mensaje_Advertencia(ByVal nSecs As Long, secs() As Long) As Boolean
'    Dim Registro As New RecordNet
'    Dim scad As String
'    Dim Resp As Boolean
'    Dim nResp As Long
'    Dim Bucle As Long
'    Resp = True
'    scad = ""
'    If nSecs > 0 Then
'        For Bucle = 0 To nSecs - 1
'            scad = scad & "codi_ident<>" & secs(Bucle) & " and "
'        Next Bucle
'        scad = left(scad, Len(scad) - 5)
'        Registro.OpenRecordset ("select count(codigo) from articulo where borrado=false and codi_fam=" & Val(CmbFamilias.TexT) & " and " & scad)
'        If Registro.Fields(1) > 0 Then
'            nResp = MsgBox(CargaCadena(1206), vbYesNo + vbExclamation + vbDefaultButton2)
'            If nResp = vbNo Then Resp = False
'        End If
'    End If
'    Mensaje_Advertencia = Resp
'End Function


Private Sub Command1_Click()
    Dim bucle As Long
    Dim nSecs As Long
    Dim secs() As Long
    Dim Resp As Long
    
    If Not IsNumeric(CmbFamilias.TexT) Then
        MsgBox CargaCadena(1205), vbCritical
        Me.Enabled = True
        Exit Sub
    End If
    '
    '
    If lRefresca Then
        ReDim secs(1)
        If Not (List1.Selected(List1.ListIndex)) Then Exit Sub
        nSecs = 1
        secs(0) = Val(left(List1.List(List1.ListIndex), 4))
        Resp = Alta_Familias_Counter(Val(CmbFamilias.TexT), Text1.TexT, nSecs, secs, False, MyEvento, sQueNombreImagen, secs(0), Val(TxtOrden.TexT))
        If Resp = 0 Then
            Refresca_Secciones
        Else
            If Resp = 2 Then
                TxtOrden.TexT = "0"
                MsgBox "Posición ya existente en la Sección... Posición no Válida.", vbOKOnly
                
            End If
        End If
        
    Else
    '
    '    Me.Enabled = False
    '
    '    For bucle = 0 To List1.ListCount - 1
    '        If List1.Selected(bucle) Then
    '            nSecs = nSecs + 1
    '            ReDim Preserve secs(nSecs)
    '            secs(nSecs - 1) = Val(left(List1.List(bucle), 4))
    '        End If
    '    Next bucle
    '    If nSecs = 0 Then
    '        ReDim secs(1)
    '        secs(0) = 0
    '    End If
    '    LblInfo.Caption = ""
    '    LblInfo.Visible = True
    '    Resp = Alta_Familias_Counter(Val(CmbFamilias.TexT), Text1.TexT, nSecs, secs, False, MyEvento, sQueNombreImagen, secs(0), 32767)
    '    Select Case Resp
    '        Case 0
    '            MsgBox CargaCadena(421), vbInformation
    '            'Refresca_Familias
    '            'Refresca_Datos
    '            'Refresca_Secciones
    '        Case 1
    '            MsgBox CargaCadena(751), vbInformation
    '        Case 4
    '        Case 61
    '            MsgBox CargaCadena(1207), vbInformation
    '    End Select
    '    LblInfo.Visible = False
    '    Me.Enabled = True
    '    If CmbFamilias.Enabled = True Then CmbFamilias.SetFocus
    '
    End If
    Refresca_Arbol Val(left(List1.TexT, 3))
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

'Private Sub Command3_Click()
'    Dim Registro As New RecordNet
'    Dim Resp As Long
'
'    If Not IsNumeric(CmbFamilias.TexT) Then Exit Sub
'    Me.Enabled = False
'    Registro.OpenRecordset ("select * from articulo where borrado=false and codi_fam=" & Val(CmbFamilias.TexT))
'    If Not Registro.EOF Then
'        Resp = MsgBox(CargaCadena(1208), vbYesNo + vbDefaultButton2 + vbExclamation)
'        If Resp = vbNo Then Exit Sub
'        LblInfo.Caption = ""
'        LblInfo.Visible = True
'    End If
'    Set Registro = Nothing
'    Resp = Baja_Familias(Val(CmbFamilias.TexT), MyEvento)
'    If Resp = 0 Then
'        MsgBox "OK"
'        Refresca_Familias
'        Refresca_Datos
'    End If
'    LblInfo.Visible = False
'    Me.Enabled = True
'End Sub

Private Sub Command4_Click(Index As Integer)
    Dim bucle As Long
    Dim nSecs As Long
    Dim secs() As Long
    Dim Resp As Long
    
    If List1.ListCount > 0 Then
        For bucle = 0 To List1.ListCount - 1
            Select Case Index
                Case 0
                    List1.Selected(bucle) = True
                Case 1
                    List1.Selected(bucle) = False
            End Select
        Next bucle
    End If
End Sub

Private Sub Form_Activate()
    Me.top = FrmFamilias.top
    Me.left = FrmFamilias.left + FrmFamilias.Width
    Me.Height = FrmFamilias.Height
End Sub

Private Sub Form_Load()
    Set MyEvento = New ClsEvento
    Set Image1.Picture = Form2.Imagen.Picture
    Command2.Visible = False
    Command3.Enabled = False
    Command3.Visible = False
    Cambiar_idioma
    lEnCarga = True
    Refresca_Familias
    If CmbFamilias.ListCount > 0 Then
        CmbFamilias.ListIndex = 0
    End If
    lEnCarga = False
    
End Sub
Public Sub Refresca_Familias()
    Dim Registro As New RecordNet
    CmbFamilias.Clear
    Registro.OpenRecordset ("select distinct(codi_fam),posicion from fam_code where posicion<>1525 order by codi_fam")
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
Private Sub Refresca_Secciones()
    Dim Registro As New RecordNet
    Dim MyMos As String
    Dim bucle As Long
    Command3.Enabled = False
    Command3.Visible = False
    YaExiste = False
    List1.Clear
    
    If Not IsNumeric(CmbFamilias.TexT) Then Exit Sub
    Registro.OpenRecordset ("select distinct(codi_ident),descripcio,nombre,secc_maqui from seccion where borrado=false")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                MyMos = Str(.Fields("codi_ident"))
                FamMaqui(.Fields("codi_ident")) = .Fields("secc_maqui")
                If Len(MyMos) < 4 Then MyMos = MyMos & Space(4 - Len(MyMos))
                List1.AddItem MyMos & " - " & .Fields("nombre")
                
               
                .Movenext
            Loop
        End If
    End With
    Registro.OpenRecordset ("select distinct(codi_ident),index from familias where (codi_fam=" & Val(CmbFamilias.TexT) & " and borrado=false) order by codi_ident")
    With Registro
        If Not .EOF Then
            YaExiste = True
            'Command3.Enabled = True
            .MoveFirst
            Do Until .EOF
                MyMos = Str(.Fields("codi_ident"))
                For bucle = 0 To List1.ListCount - 1
                    'If Val(left(List1.List(bucle), 4)) = Val(MyMos) Then
                        List1.Selected(bucle) = False
                        List1.List(bucle) = List1.List(bucle) '& " Pos.:" & Format(.Fields("index"), "0000")
                        'If .Fields("index") <> 0 Then
                        '    TxtOrden.TexT = CStr(.Fields("index"))
                        'Else
                        '    TxtOrden.TexT = "0"
                        'End If
                        'Exit For
                    
                    'End If
                Next bucle
                .Movenext
            Loop
            
            .MoveFirst
            Do Until .EOF
                MyMos = Str(.Fields("codi_ident"))
                For bucle = 0 To List1.ListCount - 1
                    If Val(left(List1.List(bucle), 4)) = Val(MyMos) Then
                        'List1.Selected(bucle) = True
                        'List1.List(bucle) = List1.List(bucle) & " Pos.:" & Format(.Fields("index"), "0000")
                        If .Fields("index") <> 0 Then
                            List1.Selected(bucle) = True
                            TxtOrden.TexT = CStr(.Fields("index"))
                            List1.List(bucle) = List1.List(bucle) & " Pos.:" & Format(.Fields("index"), "0000")
                        Else
                            TxtOrden.TexT = "0"
                            List1.List(bucle) = List1.List(bucle) '& " Pos.:" & Format(.Fields("index"), "0000")
                        End If
                        Exit For
                    
                    End If
                Next bucle
                .Movenext
            Loop
        End If
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub




Private Sub List1_Click()
Dim nOrd As Integer
Dim nSecs As Long
Dim secs() As Long
Dim Resp As Long
Dim lBaja As Boolean
    If lRefresca Then Exit Sub
    
    lRefresca = True
    
    lBaja = False
    
    Refresca_Arbol Val(left(List1.TexT, 3))
    
    nOrd = Val(Right(List1.TexT, 4))
    If List1.Selected(List1.ListIndex) Then
        If nOrd <> 0 Then
            TxtOrden.TexT = CStr(nOrd)
        Else
            TxtOrden.TexT = "0"
        End If
    Else 'poner a cero y refrescar
        TxtOrden.TexT = "0"
    End If
    If lEnCarga = False Then
        ReDim secs(1)
        nSecs = 1
        secs(0) = Val(left(List1.List(List1.ListIndex), 4))
        If Not (List1.Selected(List1.ListIndex)) Then lBaja = True
        If TxtOrden.TexT <> "0" Then
            Resp = Alta_Familias_Counter(Val(CmbFamilias.TexT), Text1.TexT, nSecs, secs, False, MyEvento, sQueNombreImagen, secs(0), Val(TxtOrden.TexT), lBaja)
        End If
    End If
    
    lRefresca = False
End Sub



Private Sub MyEvento_Evento1(lData As Long)
    LblInfo.Caption = CargaCadena(69) & " : " & lData
    nEvents = nEvents + 1
    If nEvents >= 10 Then
        Do_Events
        nEvents = 0
    End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 
    Select Case KeyCode
        Case 40
            If List1.ListCount > 0 Then List1.SetFocus
        Case 38
            If CmbFamilias.Enabled = True Then CmbFamilias.SetFocus
    End Select
End Sub


Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'guardar y refrescar
        lRefresca = True
        If lEnCarga = False Then
            'If Val(TxtOrden.TexT) > 0 Then
            Command1_Click
        End If
        lRefresca = False
    Else
        If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

'Private Sub txtorden_LostFocus()
'    lRefresca = True
'    If lEnCarga = False Then
'        Command1_Click
'    End If
'    lRefresca = False
'End Sub

Public Sub Refresca_Arbol(nS As Integer)
    
    Dim Registro As New RecordNet
    Dim Regfamilia As New RecordNet
    Arbol.Nodes.Clear
    'Arbol.ImageList = Imagenes
    'Arbol.Nodes.Add , , "S", "Sección"
    Arbol.Nodes.Add , "S", "S" & CStr(nS), CargaCadena(47) & ":" & nS
    
    Registro.OpenRecordset ("select codi_fam,index from familias where index<>0 and secc_maqui=" & CStr(nS) & " order by index")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                ''
                '' meter familias
                ''
                Arbol.Nodes.Add "S" & CStr(nS), tvwChild, "F" & .Fields("codi_fam"), CargaCadena(67) & ":" & CStr(.Fields("codi_fam")) & "  Pos.:" & CStr(.Fields("index"))
                .Movenext
            
            Loop
        End If
    End With
    Arbol.Nodes(1).Expanded = True

End Sub



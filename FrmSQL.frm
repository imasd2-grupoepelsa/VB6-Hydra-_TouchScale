VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSQL 
   Caption         =   "Consultas SQL a la base"
   ClientHeight    =   5445
   ClientLeft      =   1155
   ClientTop       =   2475
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
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
      Index           =   1
      Left            =   9120
      TabIndex        =   6
      Top             =   4800
      Width           =   1572
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
      ForeColor       =   &H8000000D&
      Height          =   3345
      Left            =   8640
      TabIndex        =   5
      Top             =   480
      Width           =   2172
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar datos Seleccionados"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   3015
   End
   Begin MSComctlLib.ListView Lista 
      CausesValidation=   0   'False
      Height          =   3852
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   6800
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
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
      Index           =   0
      Left            =   7320
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   612
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   6372
   End
   Begin VB.Label LblRegistro 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   2172
   End
End
Attribute VB_Name = "FrmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long
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
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
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
Private Sub CmdBorrar_Click()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Integer
    Set Base = OpenDatabase(Base_General)
    Dim Resp
    Resp = MsgBox(CargaCadena(995), vbYesNo + vbExclamation + vbDefaultButton2)
    If Resp <> vbYes Then Exit Sub
    On Error GoTo fin
    Set Registro = Base.OpenRecordset(Text1.TexT)
    Registro.MoveFirst
    For bucle = 1 To Lista.ListItems.Count
        If Lista.ListItems(bucle).Selected = True Then
            Delete_Record Registro
        End If
        If Not Registro.EOF Then Registro.Movenext
    Next bucle
    CerrarBase Base
    Command1_Click (0)
    Exit Sub
fin:
    CerrarBase Base
    MsgBox CargaCadena(996), vbCritical
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Integer
    Dim Eventos As Integer
    Dim Adicional As String
    Dim MyCont As Long
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    FrmSQL.Enabled = False
    CmdBorrar.Enabled = False
    Lista.ListItems.Clear
    Lista.ColumnHeaders.Clear
    LblRegistro = "0"
    Set Base = AbrirBase(True)
    On Error GoTo final
    Set Registro = Base.OpenRecordset(Text1.TexT, , dbReadOnly + dbOpenForwardOnly)
    On Error GoTo 0
    With Registro

        If .Fields.Count > 0 Then


            For bucle = 0 To .Fields.Count - 1
                Select Case .Fields(bucle).type
                Case dbText
                    Adicional = "(" & CargaCadena(539) & ")"
                Case dbBoolean
                    Adicional = "(" & CargaCadena(1072) & ")"
                Case dbDate
                    Adicional = "(" & CargaCadena(1074) & ")"
                Case Else
                    Adicional = "(" & CargaCadena(1073) & ")"
                End Select
                Lista.ColumnHeaders.Add , , .Fields(bucle).name & " " & Adicional
                MyCont = 1
                LblRegistro.Caption = 0
                If Not .EOF Then
                    .MoveFirst
                    Do Until .EOF
                        LblRegistro.Caption = Val(LblRegistro.Caption) + 1
                        If bucle = 0 Then
                            If Not IsNull(.Fields(0)) Then
                                Lista.ListItems.Add , , .Fields(0)
                            Else
                                Lista.ListItems.Add , , ""
                            End If
                        Else
                            If Not IsNull(.Fields(bucle)) Then
                                Lista.ListItems(MyCont).SubItems(bucle) = .Fields(bucle)
                                MyCont = MyCont + 1
                            End If
                        End If
                        Eventos = Eventos + 1
                        If Eventos > 5000 Then
                            Do_Events
                            Eventos = 0
                        End If
                        .Movenext
                        If Lista.ListItems.Count > 20000 Then
                            MsgBox CargaCadena(997), vbCritical
                            CerrarBase Base
                            Exit Sub
                        End If
                    Loop
                End If
                If Not .BOF Then .MoveFirst
            Next bucle
        End If
    End With
    CerrarBase Base
    FrmSQL.Enabled = True
    Exit Sub
final:
    On Error GoTo 0
    LblRegistro = ""
    MsgBox CargaCadena(998), vbCritical
    FrmSQL.Enabled = True
End Sub

Private Sub Form_Load()
'Me.Width = (Screen.Width * 3 / 4) - 200
'Lista.Width = Me.Width - 100
'Text1.Width = Me.Width - 100
'Command1.Left = (Me.Width / 2) - (Command1.Width / 2)
'Me.Left = 100
    Me_Init
    CambiarIdioma
    Image1.Picture = Form2.Imagen.Picture
    Me.Icon = Form2.Icon
    Toma_Tablas
End Sub

Private Sub CambiarIdioma()
    Me.Caption = CargaCadena(1069)
    CmdBorrar.Caption = CargaCadena(1071)
    Command1(0).Caption = CargaCadena(742)
    Command1(1).Caption = CargaCadena(7)
End Sub

Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub LblRegistro_Click()
    Dim Mybase As Database
    Dim Myresp As Variant
    FrmNewLogin.ElUsuario = "FABRICA"
    FrmNewLogin.Form_Load
    FrmNewLogin.Show vbModal
    If LoginSucceeded Then
        Myresp = MsgBox("¿Realizar el borrado deseado?", vbYesNo)
        If Myresp = vbYes Then
            Set Mybase = AbrirBase
            Mybase.Execute "delete * from cabecera where typtic='12'"
            Mybase.Execute "delete * from tickets where typtic='12'"
            Mybase.Execute "delete * from gtsecs where modo=12"
            Mybase.Execute "delete * from gtarti where modo=12"
            Mybase.Execute "delete * from gtinfven where modo=12"
            Mybase.Execute "delete * from gtvend where modo=12"
            Mybase.Execute "delete * from gthora"
            'End If
            Mybase.Close
            Set Mybase = Nothing
            MsgBox "Borrado realizado."
            salir_programa
        Else
            salir_programa
        End If
    Else
        salir_programa
    End If

End Sub

Private Sub List1_DblClick()
    If List1.ListCount > 2 Then
        If List1.ListIndex >= 2 Then
            Text1.TexT = "select * from " & List1.TexT
            Command1_Click (0)
        End If
    End If
End Sub


Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CmdBorrar.Enabled = True
End Sub
Private Sub Toma_Tablas()
    Dim Base As dao.Database
    Dim bucle As Integer
    List1.Clear
    List1.AddItem CargaCadena(1070)    '"TABLAS DISPONIBLES"
    List1.AddItem ""
    Set Base = OpenDatabase(Base_General)
    For bucle = 0 To Base.TableDefs.Count - 1
        If Len(Base.TableDefs(bucle).name) >= 4 Then
            If UCase(left(Base.TableDefs(bucle).name, 4)) <> "MSYS" Then List1.AddItem Base.TableDefs(bucle).name
        Else
            List1.AddItem Base.TableDefs(bucle).name
        End If
    Next bucle
    CerrarBase Base
End Sub

Private Sub Text1_Change()
    CmdBorrar.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click (0)
End Sub

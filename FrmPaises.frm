VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaises 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Códigos de Países"
   ClientHeight    =   2580
   ClientLeft      =   6900
   ClientTop       =   2175
   ClientWidth     =   7320
   Icon            =   "FrmPaises.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView Lista 
      Height          =   2052
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   2652
      _ExtentX        =   4683
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cod"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox CmbNombre 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   2292
   End
   Begin VB.ComboBox CmbCodigo 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   1452
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
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
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
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ISO 3166"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2652
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H80000009&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2172
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del País"
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1932
   End
   Begin VB.Label Label2 
      Caption         =   "Código del País"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "FrmPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EsNuevo As Boolean
Public BloquearBase As Boolean
Private Sub CambiarIdioma()

    Label2(0).Caption = CargaCadena(415)
    Label2(1).Caption = CargaCadena(423)
    LblInfo.Caption = CargaCadena(524)
    CmdAceptar.Caption = CargaCadena(550)
    CmdSalir.Caption = CargaCadena(7)
    CmdBorrar.Caption = CargaCadena(38)
    FrmPaises.Caption = CargaCadena(412)
End Sub
Private Sub Refresca_Datos(EsCodigo As Boolean)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Set Base = OpenDatabase(Base_General)
    If EsCodigo Then
        Set Registro = Base.OpenRecordset _
                       ("select * from paises where borrado=false and tipo=1 and codigo=" & Val(CmbCodigo.TexT))
    Else
        Set Registro = Base.OpenRecordset _
                       ("select * from paises where tipo=1 and nombre=" & Chr(34) & Trim(CmbNombre.TexT) & Chr(34))
    End If
    With Registro
        If Not .EOF Then
            EsNuevo = False
            LblInfo.Caption = CargaCadena(525)  '"Modificación"
            LblInfo.BackColor = vbYellow
            CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"

            Select Case EsCodigo
            Case False
                CmbCodigo.TexT = .Fields("codigo")
            Case True
                CmbNombre.TexT = .Fields("Nombre")
            End Select
        Else
            EsNuevo = True
            CmbNombre.TexT = ""
            LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
            LblInfo.BackColor = vbWhite
            CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
        End If
    End With
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close

End Sub


Private Sub Refresca_Codigos()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    CmbCodigo.Clear
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from paises where borrado=false and tipo=1")
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            CmbCodigo.AddItem .Fields("codigo")
            .Movenext
        Loop
    End With
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close

End Sub

Private Sub CmbCodigo_Click()
    Refresca_Datos True
End Sub

Private Sub CmbCodigo_KeyPress(KeyAscii As Integer)
    CmbCodigo.Locked = False
    If KeyAscii = 13 Then
        CmbNombre.SetFocus
        Exit Sub
    End If
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        CmbCodigo.Locked = True
        Exit Sub
    End If
    If Len(CmbCodigo.TexT) = 3 Then
        CmbCodigo.Locked = True
        Exit Sub
    End If
End Sub



Private Sub CmbCodigo_LostFocus()
    Refresca_Datos True

End Sub

Private Sub CmbNombre_Click()

'If EsNuevo Then CmbCodigo.Text = ""
'Refresca_datos False

End Sub

Private Sub CmbNombre_KeyPress(KeyAscii As Integer)
    CmbNombre.Locked = False
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
        Exit Sub
    End If
    If KeyAscii = 8 Then Exit Sub
    If Len(CmbNombre.TexT) = 100 Then
        CmbNombre.Locked = True
        Exit Sub
    End If
End Sub

Private Sub CmbNombre_LostFocus()

'If EsNuevo Then CmbCodigo.Text = ""
'Refresca_datos False

End Sub

Private Sub Cmdaceptar_Click()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim CodigoPais As String
    CmbCodigo.Enabled = True
    CmbCodigo.SetFocus
    If CmbCodigo.TexT = "" Then
        MsgBox CargaCadena(438)  '"Se necesita un código de País"
        Exit Sub
    End If
    If CmbNombre.TexT = "" Then
        MsgBox CargaCadena(367)  '"Se necesita un Nombre de País"
        Exit Sub
    End If
    Set Base = OpenDatabase(Base_General)
    If EsNuevo Then
        Set Registro = Base.OpenRecordset("select * from paises where borrado=false and tipo=1")
        Registro.AddNew
        Registro!Modificado = False
        Registro!Tipo = 1
        FrmPaises.Caption = CargaCadena(527)  '"Dato Añadido.Envíe Modificaciones"
    Else
        Set Registro = Base.OpenRecordset _
                       ("select * from paises where borrado=false and tipo=1 and codigo=" & CmbCodigo.TexT)
        Edit_Record Registro
        Registro!Modificado = True
        FrmPaises.Caption = CargaCadena(439)  '"Dato Modificado.Envíe Modificaciones"
    End If
    Registro!codigo = CmbCodigo.TexT
    Registro!Nombre = Trim(CmbNombre.TexT)
    Registro!Valor = Trim(CmbNombre.TexT)
    Registro.Fields("abrv_2") = "  "
    Registro.Fields("abrv_3") = "   "

    Registro!borrado = False
    Registro!tran_pais = " "
    Registro.Update
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Set Base = Nothing
    'Workspaces(0).close

    CodigoPais = CmbCodigo.TexT
    Refresca_Codigos

    Call Carga_Codigos

    CmbCodigo.TexT = ""
    CmbNombre.TexT = ""
    If FrmPaises.Tag <> "" Then
        frmFichasVacuno.Combo1(FrmPaises.Tag) = CodigoPais
        Unload Me
    End If
End Sub

'Private Sub CmdBorrar_Click()
'    Dim Base As DAO.Database
'    Dim Registro As DAO.Recordset
'    Dim Bucle As Long
'    Dim Encontrado As Boolean
'    If CmbCodigo.TexT = "" Then Exit Sub
'    Encontrado = False
'    Set Base = OpenDatabase(Base_General)
'    '********************************************************************
'    ' 1º Comprueba que el país no está siendo utilizado en ninguna ficha de vacuno
'    '********************************************************************
'    Set Registro = Base.OpenRecordset _
     '    ("select codigo,codnacimiento,codcrianza,codsacrificio,coddespiece,codproduccion from fichavacuno where borrado=false")
'    With Registro
'        If Not .EOF Then
'            .MoveFirst
'            Do Until .EOF
'                For Bucle = 1 To 5
'                    If Val(CmbCodigo.TexT) = Val(.Fields(Bucle)) Then
'                        'MsgBox "El país no puede ser borrado porque está siendo Utilizado en la Ficha de Vacuno nº " & !codigo
'                        MsgBox CargaCadena(372) & " " & !codigo + 1
'                        Exit Sub
'                    End If
'                Next Bucle
'                .Movenext
'            Loop
'        End If
'    End With
'    Registro.Close
'    '*************************************
'    ' 2º Procede al borrado
'    '*************************************
'
'    Set Registro = Base.OpenRecordset _
     '    ("select * from paises where tipo=1 and borrado=false and codigo=" & CmbCodigo.TexT)
'    Edit_Record Registro
'    Registro!borrado = True
'    Registro!tran_pais = " "
'    Registro.Update
'    '1.7.2 CerrarBase Base
'    '1.7.2
'    Registro.Close
'    Set Registro = Nothing
'    Base.Close
'    Set Base = Nothing
'    'Workspaces(0).close
'
'    Refresca_Codigos
'
'    CmbCodigo.TexT = ""
'    CmbNombre.TexT = ""
'    lblinfo.Caption = CargaCadena(524)  '"Nuevo Dato"
'    lblinfo.BackColor = vbWhite
'    FrmPaises.Caption = CargaCadena(556)  '"Dato Borrado.Envíe Modificaciones"
'
'End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CambiarIdioma
    Carga_Codigos
    Me.Icon = Form2.Icon
    Refresca_Codigos
    Refresca_Datos True
    EsNuevo = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmEpelsa.Enabled = True
End Sub
Private Sub Carga_Codigos()
    Dim Mybase As dao.Database
    Dim myRec As dao.Recordset
    Set Mybase = OpenDatabase(Base_General)
    Set myRec = Mybase.OpenRecordset("select * from paises where tipo=1")
    If Not myRec.EOF Then
        Lista.ListItems.Clear
        myRec.MoveFirst
        Do While Not myRec.EOF
            Lista.ListItems.Add , , myRec.Fields("valor")
            Lista.ListItems(Lista.ListItems.Count).SubItems(1) = Format(myRec.Fields("codigo"), "000")

            myRec.Movenext
        Loop

    End If
    myRec.Close
    Set myRec = Nothing
    Mybase.Close
    Set Mybase = Nothing


End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CmbCodigo.TexT = Item.SubItems(1)
    Do_Events
    Refresca_Datos True
    If Trim(CmbNombre.TexT) = "" Then
        CmbNombre.TexT = left(Item.TexT, 25)
    End If

End Sub

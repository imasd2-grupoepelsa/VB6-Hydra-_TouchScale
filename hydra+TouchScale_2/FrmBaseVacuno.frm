VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBaseVacuno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Total Diario de Vacuno"
   ClientHeight    =   4980
   ClientLeft      =   2025
   ClientTop       =   2355
   ClientWidth     =   6720
   Icon            =   "FrmBaseVacuno.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   252
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   4560
      Width           =   1332
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox CmbModo 
      Height          =   315
      ItemData        =   "FrmBaseVacuno.frx":000C
      Left            =   2640
      List            =   "FrmBaseVacuno.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1092
   End
   Begin VB.ComboBox CmbFicha 
      Height          =   315
      ItemData        =   "FrmBaseVacuno.frx":0028
      Left            =   2640
      List            =   "FrmBaseVacuno.frx":002A
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "FrmBaseVacuno.frx":002C
      Left            =   120
      List            =   "FrmBaseVacuno.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1692
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Nº de Ficha :"
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
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label LblFecha 
      Caption         =   "Fecha :"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Totales de las Fichas de Vacuno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "FrmBaseVacuno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmbFicha_Click()
    Refresca_Datos
End Sub

Private Sub CmbFicha_KeyPress(KeyAscii As Integer)
    CmbFicha.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        CmbFicha.Locked = True
        Exit Sub
    End If
    If Len(CmbFicha.TexT) = 3 And (KeyAscii >= 48 And KeyAscii <= 57) Then
        CmbFicha.Locked = True
        Exit Sub
    End If
End Sub

Private Sub CmbFicha_LostFocus()
    Refresca_Datos
End Sub

Private Sub CmbModo_Click()
    Refresca_Datos
End Sub

Private Sub CmbModo_LostFocus()
    Refresca_Datos
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
If Combo1.TexT = Combo1.List(0) Then
     TxtFecha.Enabled = True
     TxtFecha.TexT = Format(Now(), "dd/mm/yyyy")
     LblFecha.Enabled = True
Else
     TxtFecha.Enabled = False
     TxtFecha.TexT = "__/__/____"
     LblFecha.Enabled = False
End If
Refresca_Datos
End Sub

Private Sub Combo1_LostFocus()
If Combo1.TexT = Combo1.List(0) Then
     TxtFecha.Enabled = True
     TxtFecha.TexT = Format(Now(), "dd/mm/yyyy")
     LblFecha.Enabled = True
Else
     TxtFecha.Enabled = False
     TxtFecha.TexT = "__/__/____"
     LblFecha.Enabled = False
End If
Refresca_Datos
End Sub

Private Sub CambiarIdioma()
    FrmBaseVacuno.Caption = CargaCadena(642)
    Label1.Caption = CargaCadena(642)
    LblFecha.Caption = CargaCadena(434)
    CmbModo.List(0) = CargaCadena(332) 'red
    CmbModo.List(1) = CargaCadena(333) 'local
    Combo1.List(0) = CargaCadena(118)
    Combo1.List(1) = CargaCadena(124)
    Label2.Caption = CargaCadena(666)
    CmdSalir.Caption = CargaCadena(7)
End Sub
Private Sub Form_Load()
    CambiarIdioma
    Image1.Picture = Form2.Imagen.Picture
    Me.Icon = Form2.Icon
    Combo1.TexT = Combo1.List(0)
    CmbModo.TexT = CmbModo.List(0)
    TxtFecha.TexT = Format(Now(), "dd/mm/yyyy")
    Refresca_fichas
    CmbFicha.TexT = CmbFicha.List(0)
    Refresca_Datos
End Sub
Private Sub Refresca_fichas()
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    CmbFicha.Clear
    CmbFicha.AddItem "(" & CargaCadena(311) & ")"
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select distinct codigo from fichavacuno")
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            CmbFicha.AddItem Format(!codigo + 1, "000")
            .Movenext
        Loop
    End With
    'CerrarBase Base
    '1.7.2
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Registro = Nothing
    Set Base = Nothing
    'Workspaces(0).close

    
End Sub
Private Sub Refresca_Datos()
Dim Base As DAO.Database
Dim Registro As DAO.Recordset
Dim Contador As Integer
Dim lEsAcumulado As Boolean

    lEsAcumulado = False
    
    Lista.ListItems.Clear
    Lista.ColumnHeaders.Clear
    Lista.ColumnHeaders.Add , , db.CargaCadena(1252), 650 '"Ficha"
    Lista.ColumnHeaders.Add , , db.CargaCadena(119) '"Nº Operaciones"
    Lista.ColumnHeaders.Add , , db.CargaCadena(131) '"Nº Unidades"
    Lista.ColumnHeaders.Add , , db.CargaCadena(132) '"Importe"
    Lista.ColumnHeaders.Add , , db.CargaCadena(130) '"Peso"
    Set Base = OpenDatabase(Base_General)
    '******************************************
    'Total o acumulado, general o en particular
    '******************************************
    If CmbFicha.TexT = CmbFicha.List(0) Then
        If Combo1.TexT = Combo1.List(0) Then
            Set Registro = Base.OpenRecordset("select * from gtvacuno where fecha=" & Chr(34) & TxtFecha.TexT & Chr(34))
        Else
            lEsAcumulado = True
            Set Registro = Base.OpenRecordset _
            ("select ficha,sum(lnumoperaciones),sum(lnumunidades),sum(limporte),sum(lpeso),sum(vnumoperaciones),sum(vnumunidades),sum(vimporte),sum(vpeso) from gtvacuno group by ficha")
        End If
    Else
        If Combo1.TexT = Combo1.List(0) Then
            Set Registro = Base.OpenRecordset _
            ("select * from gtvacuno where fecha=" & Chr(34) & TxtFecha.TexT & Chr(34) & "and ficha=" & Val(CmbFicha.TexT) - 1)
        Else
            lEsAcumulado = True
            Set Registro = Base.OpenRecordset _
            ("select ficha,sum(lnumoperaciones),sum(lnumunidades),sum(limporte),sum(lpeso),sum(vnumoperaciones),sum(vnumunidades),sum(vimporte),sum(vpeso) from gtvacuno where ficha=" _
            & Val(CmbFicha.TexT) - 1 & " group by ficha")
        End If
    End If
    
    ' ************
    ' Datos de la lista
    ' ************
    
    Contador = 1
    
      With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Lista.ListItems.Add , , Format(!ficha + 1, "000")
                Select Case CmbModo.TexT
                    Case CmbModo.List(0)
                        If lEsAcumulado Then
                            Lista.ListItems(Contador).SubItems(1) = .Fields(5)
                            Lista.ListItems(Contador).SubItems(2) = .Fields(6)
                            Lista.ListItems(Contador).SubItems(3) = .Fields(7)
                            Lista.ListItems(Contador).SubItems(4) = .Fields(8) ' 1.7.5 / 1000
                        
                        Else
                            Lista.ListItems(Contador).SubItems(1) = .Fields("vnumoperaciones")
                            Lista.ListItems(Contador).SubItems(2) = .Fields("vnumunidades")
                            Lista.ListItems(Contador).SubItems(3) = .Fields("vimporte")
                            Lista.ListItems(Contador).SubItems(4) = .Fields("vpeso") '1.7.5  / 1000
                        End If
                    Case CmbModo.List(1)
                        If lEsAcumulado Then
                            Lista.ListItems(Contador).SubItems(1) = .Fields(1)
                            Lista.ListItems(Contador).SubItems(2) = .Fields(2)
                            Lista.ListItems(Contador).SubItems(3) = .Fields(3)
                            Lista.ListItems(Contador).SubItems(4) = .Fields(4) '1.7.5 / 1000
                        Else
                            Lista.ListItems(Contador).SubItems(1) = .Fields("lnumoperaciones")
                            Lista.ListItems(Contador).SubItems(2) = .Fields("lnumunidades")
                            Lista.ListItems(Contador).SubItems(3) = .Fields("limporte")
                            Lista.ListItems(Contador).SubItems(4) = .Fields("lpeso") '1.7.5 / 1000
                        
                        End If
                End Select
                Contador = Contador + 1
                .Movenext
            Loop
        End If
    End With
    
    'CerrarBase Base
    '1.7.2
    Registro.Close
    Base.Close
    Set Registro = Nothing
    Set Base = Nothing
    'Workspaces(0).close
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

Private Sub TxtFecha_LostFocus()
    Refresca_Datos
End Sub

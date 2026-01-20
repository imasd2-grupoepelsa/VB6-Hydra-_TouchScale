VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFichasVacuno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fichas de Vacuno"
   ClientHeight    =   4785
   ClientLeft      =   1170
   ClientTop       =   1710
   ClientWidth     =   12030
   Icon            =   "FichasVacuno.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbIdentificacion 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   6120
      TabIndex        =   31
      Top             =   600
      Width           =   5535
      Begin MSMask.MaskEdBox MaskD 
         Height          =   252
         Index           =   0
         Left            =   2400
         TabIndex        =   18
         Top             =   2280
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   3000
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox CmbDef 
         Appearance      =   0  'Flat
         Height          =   288
         ItemData        =   "FichasVacuno.frx":000C
         Left            =   1920
         List            =   "FichasVacuno.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
      End
      Begin MSMask.MaskEdBox MaskD 
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskD 
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   20
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Sacrificio"
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
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Despiece"
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
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   2640
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Nacimiento"
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
         Left            =   240
         TabIndex        =   42
         Top             =   2280
         Width           =   1932
      End
      Begin VB.Label Label4 
         Caption         =   "Código país de nacimiento"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   2532
      End
      Begin VB.Label Label5 
         Caption         =   "Código país de crianza"
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
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   2652
      End
      Begin VB.Label Label6 
         Caption         =   "Código país de Sacrificio"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   2532
      End
      Begin VB.Label Label7 
         Caption         =   "Código país de despiece"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   2652
      End
      Begin VB.Label Label8 
         Caption         =   "Código país de producción"
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
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label13 
         Caption         =   "Definido Por :"
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
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   360
      TabIndex        =   25
      Top             =   600
      Width           =   5655
      Begin VB.ComboBox CmbSexo 
         Height          =   315
         Left            =   3720
         TabIndex        =   47
         Text            =   "Combo2"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtVerD 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox TxtVerD 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox CmbRaza 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FichasVacuno.frx":004B
         Left            =   3720
         List            =   "FichasVacuno.frx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.ComboBox CmbCategoria 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FichasVacuno.frx":004F
         Left            =   3360
         List            =   "FichasVacuno.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtVerD 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Index           =   2
         Left            =   3960
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Index           =   3
         Left            =   3960
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.TextBox TxtAnimal 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Texto Adicional"
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
         Index           =   5
         Left            =   120
         TabIndex        =   46
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Lote"
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
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Peso"
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
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Sexo"
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
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Edad"
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
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.Label Label2 
         Caption         =   "Identificación del Animal"
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
         TabIndex        =   30
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Categoría del Animal"
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
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.Label Label10 
         Caption         =   "Raza"
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
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.Label Label11 
         Caption         =   "RSI de sala de sacrificio"
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
         TabIndex        =   27
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "RSI de sala de especie"
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
         TabIndex        =   26
         Top             =   840
         Width           =   3015
      End
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
      Left            =   1560
      TabIndex        =   22
      Top             =   4200
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
      Left            =   9840
      TabIndex        =   23
      Top             =   4200
      Width           =   1092
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
      Left            =   240
      TabIndex        =   21
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "Número de Ficha"
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
      TabIndex        =   38
      Top             =   120
      Width           =   1812
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   6735
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6600
      TabIndex        =   24
      Top             =   120
      Width           =   2772
   End
End
Attribute VB_Name = "frmFichasVacuno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EsNuevo As Boolean
'Private Sub CargaRazaCategoria()
'Dim MyBase As DAO.Database
'Dim MyRec As DAO.Recordset
'Set MyBase = DAO.OpenDatabase(Base_General)
'Set MyRec = MyBase.OpenRecordset("select * from paises where tipo=2") ' Razas
'If Not MyRec.EOF Then
'    Me.CmbRaza.Clear
'    MyRec.MoveFirst
'    Do While Not MyRec.EOF
'        Me.CmbRaza.AddItem MyRec.Fields("valor")
'        MyRec.Movenext
'    Loop
'End If
'Set MyRec = MyBase.OpenRecordset("select * from paises where tipo=4") ' categorias
'If Not MyRec.EOF Then
'    'Me.CmbCategoria.Clear
'    MyRec.MoveFirst
'    Do While Not MyRec.EOF
'        Me.CmbCategoria.AddItem MyRec.Fields("valor")
'        MyRec.Movenext
'    Loop
'End If
'Set MyRec = MyBase.OpenRecordset("select * from paises where tipo=3") ' categorias
'If Not MyRec.EOF Then
'    Me.CmbSexo.Clear
'    MyRec.MoveFirst
'    Do While Not MyRec.EOF
'        Me.CmbSexo.AddItem MyRec.Fields("valor")
'        MyRec.Movenext
'    Loop
'End If'
'
'MyRec.Close
'Set MyRec = Nothing
'MyBase.Close
'Set MyBase = Nothing
'End Sub
Private Sub CambiarIdioma()
    With Me
        .Caption = CargaCadena(673)
        .Label3.Caption = CargaCadena(666)
        .Label2.Caption = CargaCadena(667)
        .Label11.Caption = CargaCadena(668)
        .Label12.Caption = CargaCadena(669)
        '.Label9.Caption = CargaCadena(670)
        '.Label10(0).Caption = CargaCadena(671)
        '.Label13.Caption = CargaCadena(672)
        '.Label8.Caption = CargaCadena(661)
        .Label4.Caption = CargaCadena(662)
        .Label5.Caption = CargaCadena(663)
        .Label6.Caption = CargaCadena(664)
        .Label7(0).Caption = CargaCadena(665)
        'If Not VacunoD Then
        '    .Caption = .Caption & " (Ver. C)"
        'Else
        '    .Caption = .Caption & " (Ver. D)"
        'End If
        .LblInfo.Caption = CargaCadena(524)
        .CmdAceptar.Caption = CargaCadena(550)
        .CmdBorrar.Caption = CargaCadena(38)
        .CmdSalir.Caption = CargaCadena(7)
        '.CmbDef.List(0) = CargaCadena(661)
        '.CmbDef.List(1) = CargaCadena(662)
        '.Label10(1) = CargaCadena(1038)
        '.Label10(2) = CargaCadena(1039)
        '.Label10(3) = CargaCadena(130)
        .Label10(4) = CargaCadena(1040)
        '.Label10(5) = CargaCadena(1041)
        .Label7(1) = CargaCadena(1042)
        '.Label7(2) = CargaCadena(1043)
        .Label7(3) = CargaCadena(1044)
        
        ''''''''''''''''''''''''''''''
        'Call CargaRazaCategoria
        'Hydra+Touch ... En carga de categorias tener en cuenta cargacadena
        '.CmbCategoria.AddItem db.CargaCadena(1241) '"Sin Categoria"
        '.CmbCategoria.AddItem db.CargaCadena(1242) '"Vaca Joven"
        '.CmbCategoria.AddItem db.CargaCadena(1243) '"Toro"
        '.CmbCategoria.AddItem db.CargaCadena(1244) '"Buey"
        '.CmbCategoria.AddItem db.CargaCadena(1245) '"Vaca Adulta"
        '.CmbCategoria.AddItem db.CargaCadena(1246) '"Ternera"
        '.CmbCategoria.AddItem db.CargaCadena(1247) '"Becerro"
        ''2.0.24
        '.CmbCategoria.AddItem "Ternera V"
        '.CmbCategoria.AddItem "Ternera Z"
        '.CmbCategoria.AddItem "Añojo"
        '.CmbCategoria.AddItem "Novillo"
        '.CmbCategoria.AddItem "Cebón"
        '.CmbCategoria.AddItem "Eral"
        '''''''''
        'Hydra+touch en carga Razas tener en cuenta cargacadena...
        '.CmbRaza.AddItem db.CargaCadena(1248) '"Sin Raza"
        '.CmbRaza.AddItem db.CargaCadena(1249) '"Lechera"
        '.CmbRaza.AddItem db.CargaCadena(1250) '"Mixta"
        '.CmbRaza.AddItem db.CargaCadena(1251) '"Carne"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    End With
End Sub
Private Sub Refresca_Identificacion()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select codigo from fichavacuno where borrado=false")
    CmbIdentificacion.Clear
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                CmbIdentificacion.AddItem !codigo '+ 1
                .Movenext
            Loop
        End If
    End With
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
End Sub
Private Sub Refresca_Datos()
Dim Identificador As String
Dim Base As dao.Database
Dim Registro As dao.Recordset
Dim sSQL As String
    
    If CmbIdentificacion.TexT = "" Then Exit Sub 'Or CmbCategoria.ListIndex = -1 Then Exit Sub
    
    Identificador = CmbIdentificacion.TexT
    If Identificador = "" Then Identificador = "999"
    Set Base = OpenDatabase(Base_General)
    sSQL = "SELECT * " & _
           "FROM fichavacuno " & _
           "WHERE borrado=false and codigo =" & Chr(34) & (Identificador) & Chr(34)
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        EsNuevo = False
        LblInfo.Caption = CargaCadena(525)  '"Modificación"
        LblInfo.BackColor = vbYellow
        CmdAceptar.Caption = CargaCadena(36)  '"Modificar"
        With Registro
            .MoveFirst
            If !codnacimiento <> 0 Then
                Combo1(0).TexT = !codnacimiento
            Else
                Combo1(0).TexT = ""
            End If
            If !codcrianza <> 0 Then
                Combo1(1).TexT = !codcrianza
            Else
                Combo1(1).TexT = ""
            End If
            Combo1(2).TexT = !codsacrificio
            Combo1(3).TexT = !coddespiece
            'If !codproduccion <> 0 Then
            '    Combo1(4).TexT = !codproduccion
            'Else
            '    Combo1(4).TexT = ""
            'End If
            TxtAnimal.TexT = !Identificador
            Text1(0).TexT = !rsisacrificio
            Text1(1).TexT = !rsidespiece
            On Error Resume Next
            'Text1(2).Text = !categoria
            'Text1(3).Text = !raza
            'CmbCategoria.ListIndex = !Categoria
            'CmbRaza.ListIndex = !Raza
            'CmbSexo.ListIndex = !Sexo
            On Error GoTo 0
            'If !codproduccion <> 0 Then
            '    Combo1(0).Enabled = False
            '    Combo1(1).Enabled = False
            '    Combo1(4).Enabled = True
            '    CmbDef.TexT = CmbDef.List(0)
            'Else
                Combo1(0).Enabled = True
                Combo1(1).Enabled = True
                'Combo1(4).Enabled = False
                'CmbDef.TexT = CmbDef.List(1)
            'End If
            '*****************************
            ' versión D
            '*****************************
            'TxtVerD(0).TexT = !Edad
            'TxtVerD(1).TexT = !Sexo
            'touchscale barna no tiene sexo...
            'CmbSexo.ListIndex = !Sexo
            '''''''''''''''''''
            TxtVerD(3).TexT = !lote
            'TxtVerD(4).TexT = !texto
          
            'TxtPeso.TexT = !Peso
            
            MaskD(0).TexT = AdaptaFecha(Format(!fecha_nac, "dd/mm/yy"))
            'touchscale barna no tiene fecha despiece
            'MaskD(1).TexT = AdaptaFecha(Format(!fecha_des, "dd/mm/yy"))
            '''''''''''''''''''''''''''''''''''''''''
            MaskD(2).TexT = AdaptaFecha(Format(!fecha_sac, "dd/mm/yy"))
        End With
    Else
        EsNuevo = True
        LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
        LblInfo.BackColor = vbWhite
        CmdAceptar.Caption = CargaCadena(550)  '"Añadir"
        Combo1(0).TexT = ""
        Combo1(1).TexT = ""
        Combo1(2).TexT = ""
        Combo1(3).TexT = ""
        'Combo1(4).TexT = ""
        TxtAnimal.TexT = ""
        Text1(0).TexT = ""
        Text1(1).TexT = ""
        'Text1(2).Text = ""
        'Touchscale barna no hay categoria, raza ni sexo...
        'CmbCategoria.ListIndex = 0
        'CmbRaza.ListIndex = 0
        'CmbSexo.ListIndex = 0
        
        'Text1(3).Text = ""
        'TxtVerD(0).TexT = ""

        TxtVerD(3).TexT = ""
        'TxtVerD(4).TexT = ""
        'TxtPeso.TexT = "0"
        MaskD(0).TexT = "01/01/00"
        'MaskD(1).TexT = "01/01/00"
        MaskD(2).TexT = "01/01/00"
    End If
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
End Sub
Private Sub Command1_Click()

End Sub


'Private Sub CmbCategoria_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        CmbRaza.SetFocus
'    End If
'End Sub

'Private Sub CmbDef_Click()
'    If CmbDef.TexT = CmbDef.List(0) Then
'        Combo1(0).TexT = ""
'        Combo1(1).TexT = ""
'        Combo1(0).Enabled = False
'        Combo1(1).Enabled = False
'        Combo1(4).Enabled = True
'    Else
'        Combo1(4).TexT = ""
'        Combo1(0).Enabled = True
'        Combo1(1).Enabled = True
'        Combo1(4).Enabled = False
'    End If
'End Sub

'Private Sub CmbDef_LostFocus()
'    If CmbDef.TexT = CmbDef.List(0) Then
'        Combo1(0).TexT = ""
'        Combo1(1).TexT = ""
'        Combo1(0).Enabled = False
'        Combo1(1).Enabled = False
'        Combo1(4).Enabled = True
'    Else
'        Combo1(4).TexT = ""
'        Combo1(0).Enabled = True
'        Combo1(1).Enabled = True
'        Combo1(4).Enabled = False
'    End If
'End Sub


Private Sub CmbIdentificacion_Click()
    Refresca_Datos
End Sub

Private Sub CmbIdentificacion_KeyPress(KeyAscii As Integer)
    CmbIdentificacion.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        TxtAnimal.SetFocus
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then
        CmbIdentificacion.Locked = True
        Exit Sub
    End If
    If Len(CmbIdentificacion.TexT) = 3 Then
        CmbIdentificacion.Locked = True
        Exit Sub
    End If
End Sub

Private Sub CmbIdentificacion_LostFocus()
    Refresca_Datos
    If CmbIdentificacion.TexT = "" Then Exit Sub
    If Val(CmbIdentificacion.TexT) > 999999 Or Val(CmbIdentificacion.TexT) < 1 Then
        MsgBox CargaCadena(400)  '"El número de ficha debe estar comprendido entre 1 y 256"
        CmbIdentificacion.SetFocus
    End If
End Sub



'Private Sub CmbRaza_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
 '       If VacunoD Then
 '           TxtVerD(0).SetFocus
'        Else
'            CmbDef.SetFocus
'        End If
'    End If
'End Sub

Private Sub Cmdaceptar_Click()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Long
    CmbIdentificacion.SetFocus
    '*********************************
    ' versión D
    '*********************************
    'If Trim(TxtVerD(0).TexT) <> "" Then
    '    If Val(TxtVerD(0).TexT) > 255 Then
    '        MsgBox CargaCadena(962), vbCritical '"Valor de Edad Incorrecto", vbCritical
    '        Exit Sub
    '    End If
    'End If
    'If TxtVerD(0).TexT = "" Then TxtVerD(0).TexT = "0"
    
    For bucle = 0 To 2
        If bucle <> 1 Then 'touchscale Barna no tiene fecha despiece...
            If (Not IsNumeric(left(MaskD(bucle).TexT, 2))) Or (Not IsNumeric(Right(MaskD(bucle).TexT, 2))) Or (Not IsNumeric(Mid(MaskD(bucle).TexT, 4, 2))) Then
                MsgBox CargaCadena(506), vbCritical
                MaskD(bucle).SetFocus
                Exit Sub
            End If
            If CheckFecha(MaskD(bucle).TexT) = False Then
                MsgBox CargaCadena(506), vbCritical
                MaskD(bucle).SetFocus
                Exit Sub
            End If
        End If
    Next bucle
    '*********************************
    
    If CmbIdentificacion.TexT = "" Then
        MsgBox CargaCadena(394)  '"Debe introducir un Número de ficha"
        Exit Sub
    End If
    
    Set Base = OpenDatabase(Base_General)
    '
    ' condiciones
    '
    If CmbIdentificacion.TexT = "" Then
        
        MsgBox CargaCadena(363)  '"Introduzca un identificador del Animal"
        Exit Sub
    End If
    ' A la espera de saber si son datos indispensables,
    ' si un campo es null se pone a "0"
    For bucle = 0 To 3
        If Text1(bucle).TexT = "" Then Text1(bucle).TexT = "0"
    Next bucle
    For bucle = 0 To 4
        If Combo1(bucle).TexT = "" Then Combo1(bucle).TexT = "0"
    Next bucle
    '
    '
    '
    If EsNuevo Then
        Set Registro = Base.OpenRecordset("select * from fichavacuno where borrado=false")
        frmFichasVacuno.Caption = CargaCadena(527)  '"Registro Añadido. Envíe Modificaciones."
        Registro.AddNew
        Registro!Modificado = False
        
        
    Else
        Set Registro = Base.OpenRecordset _
        ("select * from fichavacuno where borrado=false and codigo=" _
        & Chr(34) & Val(CmbIdentificacion.TexT) & Chr(34))
        frmFichasVacuno.Caption = CargaCadena(439)  '"Registro Modificado. Envíe Modificaciones."
        
        'On Error Resume Next
        
        Registro.Edit
        If Err.Number <> 0 Then
            On Error GoTo 0
            Registro.AddNew
            Registro!Modificado = False
        Else
            Registro!Modificado = True
        End If
        
        
        
    End If
    With Registro
        !borrado = False
        !Identificador = TxtAnimal.TexT
        !codigo = Val(CmbIdentificacion.TexT)
        !rsisacrificio = Text1(0).TexT
        !rsidespiece = Text1(1).TexT
        '!Categoria = CmbCategoria.ListIndex 'Text1(2).Text
        '!Raza = CmbRaza.ListIndex 'Text1(3).Text
        If InStr(1, Combo1(0).TexT, "-") <> 0 Then
            !codnacimiento = Mid(Combo1(0).TexT, 1, InStr(1, Combo1(0).TexT, "-") - 1)
        Else
            !codnacimiento = Val(Trim(Combo1(0).TexT))
        End If
        If InStr(1, Combo1(1).TexT, "-") <> 0 Then
            !codcrianza = Mid(Combo1(1).TexT, 1, InStr(1, Combo1(1).TexT, "-") - 1)
        Else
            !codcrianza = Val(Trim(Combo1(1).TexT))
        End If
        If InStr(1, Combo1(2).TexT, "-") <> 0 Then
            !codsacrificio = Mid(Combo1(2).TexT, 1, InStr(1, Combo1(2).TexT, "-") - 1)
        Else
            !codsacrificio = Val(Trim(Combo1(2).TexT))
        End If
        If InStr(1, Combo1(3).TexT, "-") <> 0 Then
            !coddespiece = Mid(Combo1(3).TexT, 1, InStr(1, Combo1(3).TexT, "-") - 1)
        Else
            !coddespiece = Val(Trim(Combo1(3).TexT))
        End If
        
        
        
        '!codcrianza = Mid(Combo1(1).TexT, 1, InStr(1, Combo1(1).TexT, "-") - 1)
        '!codsacrificio = Mid(Combo1(2).TexT, 1, InStr(1, Combo1(2).TexT, "-") - 1)
        '!coddespiece = Mid(Combo1(3).TexT, 1, InStr(1, Combo1(3).TexT, "-") - 1)
        
        'touchscale Barna no tiene pais de producción
        '!codproduccion = Mid(Combo1(4).TexT, 1, InStr(1, Combo1(4).TexT, "-") - 1)
        '''''''''''''''''''''''''''''''''''''''''''''
        
        '*************************************
        ' versión D
        '*************************************
        '!Edad = TxtVerD(0).TexT
        'touchscale barna no tiene sexo...
        '!Sexo = CmbSexo.ListIndex 'TxtVerD(1).TexT
        ''''''''''''''''''''
        '!Peso = ToDouble(TxtPeso.TexT)
        !lote = TxtVerD(3).TexT
        '!texto = TxtVerD(4).TexT
        !fecha_nac = DateSerial(Right(MaskD(0).TexT, 2), Mid(MaskD(0).TexT, 4, 2), left(MaskD(0).TexT, 2))
        'TouchScale Barna no tirne fecha despiece...
        '!fecha_des = DateSerial(Right(MaskD(1).TexT, 2), Mid(MaskD(1).TexT, 4, 2), left(MaskD(1).TexT, 2))
        ''''''''''''''''''''''''''''''''''''''''''
        !fecha_sac = DateSerial(Right(MaskD(2).TexT, 2), Mid(MaskD(2).TexT, 4, 2), left(MaskD(2).TexT, 2))
        '!valorcategoria = Trim(CmbCategoria.TexT)
        '!valorraza = Trim(CmbRaza.TexT)
        '!valorsexo = Trim(CmbSexo.TexT)
        !valorpnacimiento = Trim(Mid(Combo1(0).TexT, 5))
        !valorpcrianza = Trim(Mid(Combo1(1).TexT, 5))
        !valorpsacrificio = Trim(Mid(Combo1(2).TexT, 5))
        !valorpdespiece = Trim(Mid(Combo1(3).TexT, 5))
        '!valorpproduccion = Trim(Mid(Combo1(4).TexT, 5))
        
        !tranficha = " "
        .Update
    End With
    'CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    Refresca_Identificacion
    Refresca_Datos
    
End Sub


Private Sub CmdBorrar_Click()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    
    If CmbIdentificacion.TexT = "" Then Exit Sub
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset _
    ("select * from fichavacuno where codigo = " & Chr(34) & Val(CmbIdentificacion.TexT) _
    & Chr(34))
    If Registro.EOF Then Exit Sub
    With Registro
        Edit_Record Registro
        !borrado = True
        !tranficha = " "
        .Update
    End With
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    '1.7.2 CerrarBase Base
    Refresca_Identificacion
    Refresca_Datos
    frmFichasVacuno.Caption = CargaCadena(528)  '"Registro Borrado. Envíe Modificaciones."
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    Combo1(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        Select Case Index
            Case 4
                Combo1(2).SetFocus
            Case 0
                Combo1(1).SetFocus
            Case 1
                Combo1(2).SetFocus
            Case 2
                Combo1(3).SetFocus
            Case 3
                CmdAceptar.SetFocus
        End Select
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then
        Combo1(Index).Locked = True
        Exit Sub
    End If
    If Len(Combo1(Index).TexT) = 3 Then
        Combo1(Index).Locked = True
        Exit Sub
    End If
    
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Dim bucle As Long
    Dim eNCONTRADO As Boolean
    Dim Buffer As String
    eNCONTRADO = False
    
    If Combo1(Index).TexT = "" Then Exit Sub
    If Combo1(Index).ListCount > 0 Then
        For bucle = 0 To Combo1(Index).ListCount - 1
            If Val(Combo1(Index).List(bucle)) = Val(Combo1(Index).TexT) Then
                eNCONTRADO = True
                Exit For
            End If
        Next bucle
    End If
    If Not eNCONTRADO Then
        FrmPaises.CmbCodigo.TexT = Combo1(Index).TexT
        Combo1(Index).TexT = ""
        FrmPaises.Tag = Index
        FrmPaises.CmbCodigo.Enabled = False
        FrmPaises.BloquearBase = True
        frmFichasVacuno.Enabled = False
        If FrmPaises.Visible = False Then
            FrmPaises.Show
            Do Until FrmPaises.Visible = False
                Do_Events
            Loop
            frmEpelsa.Enabled = False
            Set FrmPaises = Nothing
        End If
        frmFichasVacuno.Enabled = True
        Buffer = Combo1(Index).TexT
        Refresca_codigoPais
        Combo1(Index).TexT = Buffer
        '
        Select Case Index
            Case 4
                Combo1(2).SetFocus
            Case 0
                Combo1(1).SetFocus
            Case 1
                Combo1(2).SetFocus
            Case 2
                If Combo1(3).Enabled = True Then
                    Combo1(3).SetFocus
                Else
                    Combo1(2).SetFocus
                End If
            Case 3
                CmdAceptar.SetFocus
        End Select
    End If
End Sub

Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = Form2.Icon
    If Not VacunoD Then
        'TxtVerD(0).Enabled = False
        'TxtVerD(1).Enabled = False
        'CmbSexo.Enabled = False
        TxtVerD(3).Enabled = False
        'TxtVerD(4).Enabled = False
        'TxtPeso.Enabled = False
        MaskD(0).Enabled = False
'        MaskD(1).Enabled = False
        MaskD(2).Enabled = False
        'Label10(1).Enabled = False
        'Label10(2).Enabled = False
        'Label10(3).Enabled = False
        Label10(4).Enabled = False
        'Label10(5).Enabled = False
        Label7(1).Enabled = False
        'Label7(2).Enabled = False
        Label7(3).Enabled = False
        TxtAnimal.MaxLength = 10
        Text1(0).MaxLength = 10
        Text1(1).MaxLength = 10
    End If
    Image1.Picture = Form2.Imagen.Picture
    'CmbDef.TexT = CmbDef.List(0)
    'Combo1(0).Enabled = False
    'Combo1(1).Enabled = False
    Refresca_Identificacion
    Refresca_codigoPais
End Sub
Private Sub Refresca_codigoPais()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Buffer(5) As String
    Dim bucle As Long
    For bucle = 0 To 4
        Buffer(bucle) = Combo1(bucle).TexT
        Combo1(bucle).Clear
    Next bucle
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select codigo,valor from paises where borrado=false and tipo=1")
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            For bucle = 0 To 4
                Combo1(bucle).AddItem CStr(!codigo) & "-" & Trim(!Valor) 'Format(!codigo, "000") & "-" & Trim(!Valor)
            Next bucle
             .Movenext
        Loop
    End With
    '1.7.2 CerrarBase Base
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    For bucle = 0 To 4
        Combo1(bucle).TexT = Buffer(bucle)
    Next bucle
End Sub



Private Sub Form_Unload(Cancel As Integer)
    LockBase = False
    frmEpelsa.Enabled = True
End Sub




Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Text1(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index < 1 Then
            Text1(Index + 1).SetFocus
        Else
            If Index = 1 Then
                'CmbCategoria.SetFocus
                TxtVerD(3).SetFocus
                Exit Sub
            End If
            'If Combo1(0).Enabled = True Then
                Combo1(0).SetFocus
            'Else
            '    Combo1(4).SetFocus
            'End If
        End If
        Exit Sub
    End If
    If Index = 2 Or Index = 3 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            Text1(Index).Locked = True
            Exit Sub
        End If
        If Len(Text1(Index).TexT) = 1 Then
            Text1(Index).Locked = True
            Exit Sub
        End If
    Else
        If Len(Text1(Index).TexT) = 20 Then
            Text1(Index).Locked = True
        End If
    End If
        
End Sub



Private Sub Text1_LostFocus(Index As Integer)
        'If Index = 2 Or Index = 3 Then
        '    If Val(Text1(Index).Text) > 65535 Then
        '        If Index = 2 Then MsgBox "La categoría del animal debe estar comprendida entre 0 y 65535"
        '        If Index = 3 Then MsgBox "La raza debe estar comprendida entre 0 y 65535"
        '        Text1(Index).SetFocus
        '    End If
        'End If
End Sub

Private Sub TxtAnimal_KeyPress(KeyAscii As Integer)
    TxtAnimal.Locked = False
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(0).SetFocus
        Exit Sub
    End If
    If Len(TxtAnimal.TexT) = 20 Then
        TxtAnimal.Locked = True
        Exit Sub
    End If
End Sub


'Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
'    If "1,1" = 1.1 Then
'        If KeyAscii = 46 Then KeyAscii = 44
'    Else
'        If KeyAscii = 44 Then KeyAscii = 46
'    End If
'    If (InStr(1, TxtPeso.TexT, ",") > 0) Or (InStr(1, TxtPeso.TexT, ".") > 0) Then
'        TxtPeso.MaxLength = 6
'    Else
'        TxtPeso.MaxLength = 5
'    End If
'    Select Case KeyAscii
'        Case vbKeyDelete
'        Case vbKeyBack
'        Case Else
'            If (KeyAscii = 44 Or KeyAscii = 46) Or (KeyAscii > 47 And KeyAscii < 60) Then
'
'                If IsNumeric(TxtPeso.TexT) Then
'                    If (InStr(1, TxtPeso.TexT, ",") <> 0) Or (InStr(1, TxtPeso.TexT, ".") <> 0) Then
'                            If KeyAscii = 44 Or KeyAscii = 46 Then
'                                KeyAscii = 0
'                            Else
'                                If Right(TxtPeso.TexT, 1) <> "," And Right(TxtPeso.TexT, 1) <> "." Then KeyAscii = 0
'                            End If
'                    End If
'                End If
'            Else
'                KeyAscii = 0
'            End If
'    End Select
'End Sub

Private Sub TxtVerD_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 48 To 59
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub
Private Function CheckFecha(myFecha As String) As Boolean
    Dim Anyo As Integer
    Dim Mes As Integer
    Dim Dia As Integer
    Dim Bisiesto As Boolean
    Dim Resp As Boolean
    Resp = True
    Anyo = Val(Right(myFecha, 2))
    If Anyo >= 80 Then
        Anyo = Anyo + 1900
    Else
        Anyo = Anyo + 2000
    End If
    If (Anyo Mod 4) = 0 Then Bisiesto = True
    Mes = Val(Mid(myFecha, 4, 2))
    Dia = Val(left(myFecha, 2))
    If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
        If Dia > 31 Then Resp = False
    Else
        If Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
            If Dia > 30 Then Resp = False
        Else
            If Mes = 2 Then
                If Bisiesto Then
                    If Dia > 29 Then Resp = False
                Else
                    If Dia > 28 Then Resp = False
                End If
            Else
                Resp = False
            End If
            
        End If
    End If
    CheckFecha = Resp
End Function

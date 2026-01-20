VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmListados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6120
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   1440
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Clientes"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Vacuno"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Vendedor"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Artículos"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Secciones"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5895
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
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
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   26
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Mostrador :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "FrmListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SimboloMonetario = "Pta."
Private Const UsaEuro = False
Private Type Tlistado
    Lcadena(15) As Integer
    Lentero(20) As Integer
    Ncadena(15) As Integer
    Nentero(20) As Integer
    Cadena(15) As String
    Entero(20) As String
End Type
Private Fecha As Integer
Private Modo As Integer
Private Listado As Tlistado
Private Ncadenas As Integer
Private Nenteros As String
Private NOrden As Integer

Private Sub Check1_Click()
    If Check1.Value = False Then
        TxtFecha(0).Enabled = False
        TxtFecha(1).Enabled = False
    Else
        TxtFecha(0).Enabled = True
        TxtFecha(1).Enabled = True
    End If
    CargaDatos
End Sub




Private Sub Combo1_Click(Index As Integer)
    CargaDatos
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim B As Integer
    Ncadenas = 0
    Nenteros = 0
    For B = 0 To 14
        Listado.Lcadena(B) = 800
    Next B
    For B = 0 To 19
        Listado.Lentero(B) = 800
    Next B
    Select Case Index
        Case 0
            Label3.Caption = "Totales por Mostrador"
            NOrden = 1
            Ncadenas = 0
            Nenteros = 8
            Fecha = 2
            Listado.Entero(0) = "Mostrador"
            Listado.Entero(1) = "SubSección"
            Listado.Entero(2) = "Familia"
            Listado.Entero(3) = "Tiquets"
            Listado.Entero(4) = "Operaciones"
            Listado.Entero(5) = "Unidades"
            Listado.Entero(6) = "Peso"
            If UsaEuro Then
                Listado.Entero(7) = "Euros"
            Else
                Listado.Entero(7) = SimboloMonetario
            End If
            Listado.Nentero(0) = 3
            Listado.Nentero(1) = 4
            Listado.Nentero(2) = 5
            Listado.Nentero(3) = 6
            Listado.Nentero(4) = 7
            Listado.Nentero(5) = 9
            Listado.Nentero(6) = 8
            Listado.Nentero(7) = 10
            Listado.Lentero(6) = 1100
            Listado.Lentero(7) = 1100
        Case 1
            Label3.Caption = "Totales por Artículos"
            NOrden = 2
            Ncadenas = 1
            Nenteros = 10
            Fecha = 2
            Listado.Entero(0) = "Código"
            Listado.Entero(1) = "Plu"
            Listado.Entero(2) = "Mostrador"
            Listado.Entero(3) = "Subsec"
            Listado.Entero(4) = "Familia"
            Listado.Entero(5) = "Operaciones"
            Listado.Entero(6) = "Unidades"
            Listado.Entero(7) = "Peso"
            Listado.Entero(8) = "Precio"
            If UsaEuro Then
                Listado.Entero(9) = "Euros"
            Else
                Listado.Entero(9) = SimboloMonetario
            End If
            Listado.Nentero(0) = 8
            Listado.Nentero(1) = 3
            Listado.Nentero(2) = 5
            Listado.Nentero(3) = 6
            Listado.Nentero(4) = 7
            Listado.Nentero(5) = 15
            Listado.Nentero(6) = 10
            Listado.Nentero(7) = 11
            Listado.Nentero(8) = 9
            Listado.Nentero(9) = 12
            Listado.Ncadena(0) = 4
            Listado.Cadena(0) = "Descriptivo"
        Case 2
            NOrden = 3
            Label3.Caption = "Totales por Vendedor"
            Ncadenas = 1
            Nenteros = 8
            Fecha = 2
            Listado.Entero(0) = "Mostrador"
            Listado.Entero(1) = "Cod. Vendedor"
            Listado.Entero(2) = "Tiquets"
            Listado.Entero(3) = "Operaciones"
            Listado.Entero(4) = "Unidades"
            Listado.Entero(5) = "Peso"
            Listado.Entero(6) = "Cancel."
            If UsaEuro Then
                Listado.Entero(7) = "Euros"
            Else
                Listado.Entero(7) = SimboloMonetario
            End If
            Listado.Nentero(0) = 3
            Listado.Nentero(1) = 4
            Listado.Nentero(2) = 9
            Listado.Nentero(3) = 10
            Listado.Nentero(4) = 14
            Listado.Nentero(5) = 11
            Listado.Nentero(6) = 7
            Listado.Nentero(7) = 8
            Listado.Cadena(0) = "Nombre"
            Listado.Ncadena(0) = 5
        Case 3
    End Select
    PreparaLista
    ActivaCombos
    CargaDatos
    If Me.Width = 2940 Then
        For B = 2940 To Screen.Width - 400 Step 200
            Me.Width = B '11040
            Me.Left = (Screen.Width / 2) - (B / 2)
            Lista.Width = Me.Width - Lista.Left
            Label3.Width = Me.Width - Label3.Left
        Next B
        'Me.Left = 0 'me.Width
    End If
    If Me.Height = 2290 Then
        For B = 2290 To 6990 Step 100
            Me.Height = B  '11040
            Me.Top = (Screen.Height / 2) - (B / 2)
            
        Next B
        
        'Me.Left = 0 'me.Width
    End If
    
End Sub

Private Sub PreparaLista()
    Dim B As Integer
    Lista.ListItems.Clear
    Lista.ColumnHeaders.Clear
    Lista.ColumnHeaders.Add , , "Fecha", 900
     
    For B = 0 To Ncadenas - 1
        Lista.ColumnHeaders.Add , , Listado.Cadena(B), Listado.Lcadena(B)
    Next B
    For B = 0 To Nenteros - 1
        Lista.ColumnHeaders.Add , , Listado.Entero(B), Listado.Lentero(B)
    Next B
End Sub
Private Sub CargaDatos()
    Dim Base As Database
    Dim B2 As Database
    Dim Registro As Recordset
    Dim R2 As Recordset
    Dim Bucle As Integer
    Dim MiFiltro As String
    MiFiltro = CreaFiltro()
    Lista.ListItems.Clear
    Set Base = AbrirBase()
    Set B2 = OpenDatabase(MiruTa & "\dirtmp\tmp.mdb")
    Set R2 = B2.OpenRecordset("select * from totales")
    With R2
        If Not .EOF Then .MoveFirst
        If Not .EOF Then
            Do Until .EOF
                If Not .EOF Then .Delete
                If Not .EOF Then .MoveNext
            Loop
        End If
    End With
    Select Case NOrden
        Case 1
            Set Registro = Base.OpenRecordset("select * from gtsecs " & MiFiltro)
        Case 2
            Set Registro = Base.OpenRecordset("select * from gtarti " & MiFiltro)
        Case 3
            Set Registro = Base.OpenRecordset("select * from gtvend " & MiFiltro)
    End Select
    With Registro
        If Not .EOF Then
            Do Until .EOF
                R2.AddNew
                For Bucle = 0 To Nenteros - 1
                    R2.Fields(4 + Bucle) = .Fields(Listado.Nentero(Bucle) - 1)
                Next Bucle
                For Bucle = 0 To Ncadenas - 1
                    R2.Fields(24 + Bucle) = .Fields(Listado.Ncadena(Bucle) - 1)
                Next Bucle
                R2.Update
                .MoveNext
            Loop
        End If
    End With
    Base.Close
    R2.Requery
    If Not R2.EOF Then
        R2.MoveFirst
        Do Until R2.EOF
            Lista.ListItems.Add , , "jaarl"
            For Bucle = 0 To Ncadenas - 1
                Lista.ListItems(Lista.ListItems.Count).SubItems(Bucle + 1) = R2.Fields(24 + Bucle)
            Next Bucle
            For Bucle = 0 To Nenteros - 1
                Lista.ListItems(Lista.ListItems.Count).SubItems(Bucle + Ncadenas + 1) = R2.Fields(4 + Bucle)
            Next Bucle
            
            R2.MoveNext
        Loop
    End If
    B2.Close
End Sub

Private Sub Command2_Click()
    Select Case NOrden
        Case 1
        Case 2
        Case 3
    End Select
End Sub

Private Sub Form_Load()
    Dim B As Integer
    '*******************
    ' hasta incluir nuevos listados
    Command1(3).Visible = False
    Command1(4).Visible = False
    Command1(5).Visible = False
    Command1(6).Visible = False
    '********************
    Label3.Caption = "Seleccione Total"
    Me.Width = 2940
    Me.Height = 2290
    For B = 0 To 5
        Combo1(B).Visible = False
        Label1(B).Visible = False
    Next B
    TxtFecha(0).Enabled = False
    TxtFecha(1).Enabled = False
    TxtFecha(0).Text = Format(Now(), "dd/mm/yy")
    TxtFecha(1).Text = Format(Now(), "dd/mm/yy")
    Check1.Value = vbUnchecked
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Dim B As Integer
'    If Me.Width > 2940 Then
'        For B = Screen.Width - 400 To 3140 Step -50
'            Me.Width = B
'            Me.Left = (Screen.Width / 2) - (B / 2)
'            DoEvents
'        Next B
        'Me.Left = 0 'me.Width
'    End If
'    If Me.Height > 2290 Then
'        For B = 6990 To 1 Step -10
'            Me.Height = B
'            Me.Top = (Screen.Height / 2) - (B / 2)
'            DoEvents
'        Next B
'    End If
'End Sub
Private Sub ActivaCombos()
    Dim B As Integer
    For B = 0 To 5
        Combo1(B).Visible = False
        Label1(B).Visible = False
    Next B
    ' (articulos, secciones, vendedores)
    If NOrden < 4 Then
        Combo1(0).Visible = True
        Label1(0).Visible = True
    End If
    ' (articulos)
    If NOrden = 2 Then
        Combo1(1).Visible = True
        Label1(1).Caption = "Código"
        Label1(1).Visible = True
        Combo1(2).Visible = True
        Label1(2).Caption = "PLU"
        Label1(2).Visible = True
    End If
    ' (vendedores)
    If NOrden = 3 Then
        Combo1(1).Visible = True
        Label1(1).Caption = "Vendedor"
        Label1(1).Visible = True
    End If
    '
    RellenaCombos
End Sub
Private Sub RellenaCombos()
    Dim Base As Database
    Dim Registro As Recordset
    Set Base = AbrirBase()
    '(articulos,seciones,vendedores)
    If NOrden < 4 Then
        Combo1(0).Clear
        Combo1(0).AddItem "Todos"
        Set Registro = Base.OpenRecordset("select codi_ident from seccion where borrado=false")
        RellenaUnCombo Registro, 0
    End If
    ' (vendedores)
    If NOrden = 3 Then
        Combo1(1).Clear
        Combo1(1).AddItem "Todos"
        Set Registro = Base.OpenRecordset("select ident_vend from vendedor where borrado=false")
        RellenaUnCombo Registro, 1
    End If
    ' (articulos)
    If NOrden = 2 Then
        Combo1(2).Clear
        Combo1(2).AddItem "Todos"
        Set Registro = Base.OpenRecordset("select codigo from articulo where borrado=false")
        RellenaUnCombo Registro, 1
        Combo1(3).Clear
        Combo1(3).AddItem "Todos"
        Set Registro = Base.OpenRecordset("select plu from articulo where borrado=false")
        RellenaUnCombo Registro, 2
    End If
    Base.Close
End Sub
Private Sub RellenaUnCombo(Registro As Recordset, Index As Integer)
    Combo1(Index).Clear
    Combo1(Index).AddItem "Todos"
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Combo1(Index).AddItem .Fields(0)
                .MoveNext
            Loop
        End If
    End With
    Combo1(Index).Text = Combo1(Index).List(0)
End Sub

Private Function CreaFiltro() As String
    Dim MiFiltro As String
    Dim mibuf As String
    Dim F1 As String
    Dim F2 As String
    MiFiltro = ""
    mibuf = ""
    ' fecha (articulos, sección, vendedor)
    If Check1.Value = vbChecked Then
        F1 = Mid(TxtFecha(0).Text, 4, 2) & "/" & Left(TxtFecha(0).Text, 2) & "/" & Right(TxtFecha(0).Text, 2)
        F2 = Mid(TxtFecha(1).Text, 4, 2) & "/" & Left(TxtFecha(1).Text, 2) & "/" & Right(TxtFecha(1).Text, 2)
        If NOrden < 4 Then
            MiFiltro = "(fecha >= #" & F1 & "# and fecha <= #" & F2 & "#)"
        End If
    End If
    
    If MiFiltro <> "" Then mibuf = " and "
    
    ' mostrador (articulos, sección, vendedor)
    If NOrden < 4 Then
        If IsNumeric(Combo1(0).Text) Then
            MiFiltro = MiFiltro & mibuf & " ( codi_ident =" & Val(Combo1(0).Text) & ")"
            mibuf = " and "
        End If
    End If
    
    If MiFiltro <> "" Then mibuf = " and "
    
    If MiFiltro <> "" Then MiFiltro = "where " & MiFiltro
    CreaFiltro = MiFiltro
End Function


Private Sub TxtFecha_LostFocus(Index As Integer)
    CargaDatos
End Sub

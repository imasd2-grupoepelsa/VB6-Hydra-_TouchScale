VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConsultaTiquets 
   Caption         =   "Consulta de Tiquets"
   ClientHeight    =   6690
   ClientLeft      =   225
   ClientTop       =   1395
   ClientWidth     =   11400
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   6480
      Top             =   960
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listar por Impresora"
      Height          =   372
      Left            =   3720
      TabIndex        =   13
      Top             =   960
      Width           =   2412
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   3840
      Top             =   480
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmConsultaTiquets.frx":0000
      Left            =   9000
      List            =   "FrmConsultaTiquets.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ListView Detalle 
      Height          =   1575
      Left            =   0
      TabIndex        =   9
      Top             =   5040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3132
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   11412
      _ExtentX        =   20135
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   3
      ItemData        =   "FrmConsultaTiquets.frx":002A
      Left            =   1800
      List            =   "FrmConsultaTiquets.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   2
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   1
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   0
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Detalle del tiquet seleccionado :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Máquinas"
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Balanza"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Sección"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmConsultaTiquets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyFilter As String
Private MyTabla As String
Private QL As Boolean
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long
Private nIndex As Integer
Private nFichHys As Integer

'Formas de Pago....
'1   Efectivo
'2   Tarjeta Crédito
'3   Tarjeta Débito
'4   EBT Cash
'5   Gift Card
'6   Cheque
'7   Pendiente
'14  Ticket Restaurant
'15  Chèque -Vacances
'16  Cheque Gourmet
'17  Restaurant Pass
'19  Cupón
'20  Others
'21  EBT FoodStamp
'22  Visa
'23  MasterCard
'24  American Express
'25  Traveller 's Check
'28  Multibanco
'30  Visa Electron
'31  Diners Club
'32  Big Pass
'33  Sodexho
'34  Pin
'35  Chipknip
'36  Cestaticket
'37  EC -Maestro
'38  Quick
'39  BPI -Express
'40  Gift Check
'41  Ripley
'42  Saga
'43  Otras tarjetas
'44  Cheque Comercial
'45  Cheque de Gobiernos
'46  CLAVE

Private Sub Me_Init()
    Dim Bucle As Long
    Me_Width = Me.Width
    Me_Height = Me.Height
    ReDim ct_top(Me.Controls.Count)
    ReDim CT_Left(Me.Controls.Count)
    ReDim CT_Width(Me.Controls.Count)
    ReDim CT_Height(Me.Controls.Count)
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        ct_top(Bucle) = Me.Controls(Bucle).top
        CT_Left(Bucle) = Me.Controls(Bucle).left
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        On Error GoTo 0
    Next Bucle
End Sub
Private Sub Me_Resize()
    Dim RelW As Double
    Dim Bucle As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        RelW = Me.Width / Me_Width
        Me.Controls(Bucle).Width = RelW * CT_Width(Bucle)
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        Me.Controls(Bucle).left = RelW * CT_Left(Bucle)
        CT_Left(Bucle) = Me.Controls(Bucle).left
        
        '
        RelW = Me.Height / Me_Height
        If TypeName(Me.Controls(Bucle)) <> "TextBox" _
        And TypeName(Me.Controls(Bucle)) <> "CommandButton" _
        And TypeName(Me.Controls(Bucle)) <> "MaskEdBox" _
        Then
            Me.Controls(Bucle).Height = RelW * CT_Height(Bucle)
            CT_Height(Bucle) = Me.Controls(Bucle).Height
        End If
        Me.Controls(Bucle).top = RelW * ct_top(Bucle)
        ct_top(Bucle) = Me.Controls(Bucle).top
        On Error GoTo 0
        '
    Next Bucle
    Me_Width = Me.Width
    Me_Height = Me.Height
End Sub
Private Sub CmbFiltro_Click(Index As Integer)
    If Not QL Then
        Crea_Filtro
        Refresca_Datos
    End If
End Sub


Private Sub Combo1_Click()
    If QL = False Then
        Select Case Combo1.ListIndex
            Case 0
                MyTabla = ""
            Case 1
                MyTabla = "SC10"
            Case 2
                MyTabla = "TQ"
        End Select
        If Dir(App.Path & "\ahorramas.cfg") <> "" Then
            MyTabla = "SC10"
        End If
        QL = True
        RellenaFiltros
        Refresca_Datos
        QL = False
    Else
        If Dir(App.Path & "\ahorramas.cfg") <> "" Then
            MyTabla = "SC10"
        'Else c2f 1.7.24
        '    MyTabla = ""
        End If
        'QL = True
        'RellenaFiltros
        'Refresca_Datos
        'QL = False
    
    End If
End Sub


Private Sub Command1_Click()
    Dialogo.CancelError = True
    On Error Resume Next
    Dialogo.ShowPrinter
    If Err.Number = 0 Then
        Do_Events
        ImprimeLosTiquets
    End If
End Sub


Private Sub Form_Load()
    QL = True 'c2f 1.7.24
    
    Me_Init
    Me.Label1(0).Caption = CargaCadena(47)
    Me.Label1(1).Caption = CargaCadena(147)
    Me.Label1(2).Caption = CargaCadena(55)
    Me.Label1(3).Caption = CargaCadena(434)
    Me.Label2.Caption = CargaCadena(1228)
    Me.Command1.Caption = CargaCadena(1229)
    Me.Label3.Caption = CargaCadena(921)
    Me.Caption = CargaCadena(112)
    Me.Icon = Form2.Icon
    Combo1.TexT = Combo1.List(0)
    'MyTabla = "" 1.7.23
    'If UsaGamaAlta And (Not usaSC10) And (Not UsaEcoPrint) Then
        Combo1.Enabled = False
    'End If
    'If usaSC10 And (Not UsaGamaAlta) Then
    '    MyTabla = "SC10"
    '    If (Not UsaEcoPrint) Or (TipoEcoPrint <> 2) Or (TipoEcoPrint <> 3) Then Combo1.Enabled = False
    '    Combo1.TexT = Combo1.List(1)
    'End If
    'If UsaEcoPrint And (TipoEcoPrint = 2 Or TipoEcoPrint = 3) And Not UsaGamaAlta And Not usaSC10 Then
    '    MyTabla = "TQ"
    '    'Combo1.Enabled = False c2f 1.7.24
    '    Combo1.TexT = Combo1.List(2)
    'End If
    'Detalle.ColumnHeaders.Add , , "Linea", 800
    Detalle.ColumnHeaders.Add , , CargaCadena(148), 800
    
    'Detalle.ColumnHeaders.Add , , "Código", 1000
    Detalle.ColumnHeaders.Add , , CargaCadena(69), 1000
    
    'Detalle.ColumnHeaders.Add , , "PLU", 1000
    Detalle.ColumnHeaders.Add , , CargaCadena(68), 1000
    
    'Detalle.ColumnHeaders.Add , , "Texto", 2000
    Detalle.ColumnHeaders.Add , , CargaCadena(539), 2000
    
    'Detalle.ColumnHeaders.Add , , "Precio", 1000
    Detalle.ColumnHeaders.Add , , CargaCadena(71), 1000
    
    'Detalle.ColumnHeaders.Add , , "Peso", 1000
    Detalle.ColumnHeaders.Add , , CargaCadena(130), 1000
    
    'Detalle.ColumnHeaders.Add , , "Unidades", 1000
    Detalle.ColumnHeaders.Add , , CargaCadena(131), 1000
    
    'Detalle.ColumnHeaders.Add , , "Total"
    Detalle.ColumnHeaders.Add , , CargaCadena(488)
    
    'Detalle.ColumnHeaders.Add , , "Cancelación"
    'Detalle.ColumnHeaders.Add , , CargaCadena(290)
    Detalle.ColumnHeaders.Add , , CargaCadena(1224)
    
    'Detalle.ColumnHeaders.Add , , "Tipo IVA", 500
    Detalle.ColumnHeaders.Add , , CargaCadena(553), 500
    
    MyFilter = ""
    RellenaFiltros
    Me.Enabled = False
    Do_Events
    Crea_Filtro
    Refresca_Datos
    Me.Enabled = True
    QL = False
End Sub
Private Sub RellenaFiltros()
    Dim Base  As DAO.Database
    Dim Registro As DAO.Recordset
    Dim Bucle As Integer
    Dim stCampo As String
    Dim Elegido As Integer
    Dim sbas As String
    
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sbas
       Close #nFichHys
       Set Base = OpenDatabase(sbas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If
    '*****************
    ' datos de filtros
    '*****************
    For Bucle = 0 To 2
        CmbFiltro(Bucle).Clear
    Next Bucle
    For Bucle = 0 To 2
        CmbFiltro(Bucle).Clear
        'CmbFiltro(Bucle).AddItem "(Todo)"
        CmbFiltro(Bucle).AddItem CargaCadena(311)
        Select Case Bucle
            Case 0
                stCampo = "seccion"
            Case 1
                stCampo = "balanza"
            Case 2
                stCampo = "fecha"
        End Select
        Set Registro = Base.OpenRecordset("select distinct " & stCampo & " from cabecera" & MyTabla)
        With Registro
            If Not .EOF Then
                Do Until .EOF
                    CmbFiltro(Bucle).AddItem .Fields(stCampo)
                    .Movenext
                Loop
            End If
        End With
    Next Bucle
    Me.CmbFiltro(3).AddItem CargaCadena(311)    '"Todo"
    Me.CmbFiltro(3).AddItem CargaCadena(336)    '"Venta"
    'Me.CmbFiltro(3).AddItem CargaCadena(741)    '"Super"
    Me.CmbFiltro(3).AddItem "Fact."
    Me.CmbFiltro(3).AddItem CargaCadena(195)    '"Envasado"
    Me.CmbFiltro(3).AddItem CargaCadena(196)    '"Autoservicio"
    'formas de pago...
    'Me.CmbFiltro(3).AddItem CargaCadena(601)    '"Moneda"
    'Me.CmbFiltro(3).AddItem CargaCadena(1214)   '"Cheque"
    'Me.CmbFiltro(3).AddItem CargaCadena(1215)   '"Tarjeta"
    'Me.CmbFiltro(3).AddItem "Tick.Rest."        '"Tick.Rest."
    'Me.CmbFiltro(3).AddItem CargaCadena(1216)   '"Genérico"
    'Me.CmbFiltro(3).AddItem CargaCadena(1217)   '"Anulado"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Me.CmbFiltro(3).AddItem "Efectivo"
    Me.CmbFiltro(3).AddItem "Tarjeta Crédito"
    Me.CmbFiltro(3).AddItem "Tarjeta Débito"
    Me.CmbFiltro(3).AddItem "EBT Cash"
    Me.CmbFiltro(3).AddItem "Gift Card"
    Me.CmbFiltro(3).AddItem "Cheque"
    Me.CmbFiltro(3).AddItem "Pendiente"
    'del 8 pasa al 13
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    '''''''''''''''''
    Me.CmbFiltro(3).AddItem "Ticket Restaurant"
    Me.CmbFiltro(3).AddItem "Chèque -Vacances"
    Me.CmbFiltro(3).AddItem "Cheque Gourmet"
    Me.CmbFiltro(3).AddItem "Restaurant Pass"
    ''''''''''''''''' 18
    Me.CmbFiltro(3).AddItem "--"
    '''''''''''''''''
    Me.CmbFiltro(3).AddItem "Cupón"
    Me.CmbFiltro(3).AddItem "Others"
    Me.CmbFiltro(3).AddItem "EBT FoodStamp"
    Me.CmbFiltro(3).AddItem "Visa"
    Me.CmbFiltro(3).AddItem "MasterCard"
    Me.CmbFiltro(3).AddItem "American Express"
    Me.CmbFiltro(3).AddItem "Traveller 's Check"
    '26 y 27
    Me.CmbFiltro(3).AddItem "--"
    Me.CmbFiltro(3).AddItem "--"
    ''''''''
    Me.CmbFiltro(3).AddItem "Multibanco"
    '29
    Me.CmbFiltro(3).AddItem "--"
    '''
    Me.CmbFiltro(3).AddItem "Visa Electron"
    Me.CmbFiltro(3).AddItem "Diners Club"
    Me.CmbFiltro(3).AddItem "Big Pass"
    Me.CmbFiltro(3).AddItem "Sodexho"
    Me.CmbFiltro(3).AddItem "Pin"
    Me.CmbFiltro(3).AddItem "Chipknip"
    Me.CmbFiltro(3).AddItem "Cestaticket"
    Me.CmbFiltro(3).AddItem "EC -Maestro"
    Me.CmbFiltro(3).AddItem "Quick"
    Me.CmbFiltro(3).AddItem "BPI -Express"
    Me.CmbFiltro(3).AddItem "Gift Check"
    Me.CmbFiltro(3).AddItem "Ripley"
    Me.CmbFiltro(3).AddItem "Saga"
    Me.CmbFiltro(3).AddItem "Otras tarjetas"
    Me.CmbFiltro(3).AddItem "Cheque Comercial"
    Me.CmbFiltro(3).AddItem "Cheque de Gobiernos"
    Me.CmbFiltro(3).AddItem "CLAVE"
    ''''''''''''''''''''''''''''''''''
    'Me.CmbFiltro(3).AddItem "Cobrado"
    'Me.CmbFiltro(3).AddItem "No Cobrado"
    'Me.CmbFiltro(3).AddItem "Trainning"
    'Me.CmbFiltro(3).AddItem "Cliente"
    For Bucle = 0 To 3
        CmbFiltro(Bucle).TexT = CmbFiltro(Bucle).List(0)
    Next Bucle
    If CmbFiltro(2).ListCount > 1 Then
        CmbFiltro(2).TexT = CmbFiltro(2).List(CmbFiltro(2).ListCount - 1)
    End If
End Sub
Private Sub Refresca_Datos()
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Dim MiRefresco As Integer
    Dim sOrder As String
    Dim sbas As String
        
    sOrder = " order by nume"
    
    If nIndex = 0 Then nIndex = 3
    Select Case nIndex
    Case 1
        sOrder = " order by d_fecha"
    Case 2
        sOrder = " order by d_hora"
    Case 3
        sOrder = " order by nume"
    Case 4
        sOrder = " order by seccion"
    Case 5
        sOrder = " order by balanza"
    Case 6
        sOrder = " order by vended"
    Case Else
        sOrder = " order by nume"
    End Select
    
    Detalle.ListItems.Clear
    Lista.ColumnHeaders.Clear
    Lista.ColumnHeaders.Add , , CargaCadena(434), 1000 '"Fecha"
    Lista.ColumnHeaders.Add , , CargaCadena(435), 1000 '"Hora"
    Lista.ColumnHeaders.Add , , CargaCadena(1218), 1000 '"Numero"
    Lista.ColumnHeaders.Add , , CargaCadena(47), 700 '"Sección"
    Lista.ColumnHeaders.Add , , CargaCadena(147), 700 '"Balanza"
    Lista.ColumnHeaders.Add , , CargaCadena(13), 1000 '"Vendedor"
    Lista.ColumnHeaders.Add , , CargaCadena(1219), 700 '"Líneas"
    Lista.ColumnHeaders.Add , , CargaCadena(792) '"Líneas Canceladas"
    Lista.ColumnHeaders.Add , , CargaCadena(132) '"Importe"
    Lista.ColumnHeaders.Add , , CargaCadena(790) '"Cliente"
    Lista.ColumnHeaders.Add , , "EAN13"
    Lista.ColumnHeaders.Add , , CargaCadena(55) '"Tipo"
    Lista.ColumnHeaders.Add , , CargaCadena(1221) '"Estado"
    Lista.ColumnHeaders.Add , , CargaCadena(1220), 1300 '"Forma de pago"
    Lista.ColumnHeaders.Add , , CargaCadena(1222), 1500 '"Importe entregado"
    Lista.ColumnHeaders.Add , , CargaCadena(1223), 700 '"Cambio"
    Lista.ColumnHeaders.Add , , CargaCadena(1276)
    Lista.ListItems.Clear
    
    Crea_Filtro
    
    
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sbas
       Close #nFichHys
       Set Base = OpenDatabase(sbas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If
    
    'Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cabecera" & MyTabla & " " & MyFilter & sOrder)
    With Registro
        If Not .EOF Then
            Do Until .EOF
                MiRefresco = MiRefresco + 1
                If MiRefresco > 30 Then
                    MiRefresco = 0
                    Do_Events
                End If
                Lista.ListItems.Add , , !Fecha
                Lista.ListItems(Lista.ListItems.Count).SubItems(1) = !hora
                Lista.ListItems(Lista.ListItems.Count).SubItems(2) = Val(!nume)
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = !seccion
                Lista.ListItems(Lista.ListItems.Count).SubItems(4) = !Balanza
                Lista.ListItems(Lista.ListItems.Count).SubItems(5) = !vended
                Lista.ListItems(Lista.ListItems.Count).SubItems(6) = !numlin
                Lista.ListItems(Lista.ListItems.Count).SubItems(7) = !lincan
                Lista.ListItems(Lista.ListItems.Count).SubItems(8) = !Importe
                Lista.ListItems(Lista.ListItems.Count).SubItems(9) = !codcli
                Lista.ListItems(Lista.ListItems.Count).SubItems(10) = !codbar
                '1.0.0 (0=Factura pasa a 7 red, 8 local
                '       1=tique pasa a 1 red, 4 local
                '       2=Envasado pasa a 2 red, 5 local
                '       3=AutoServicio pasa a 3 red, 6 local
                
                Select Case !TYPTIC
                    Case 1
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
                    Case 2
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
                    Case 3
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
                    Case 4
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
                    Case 5
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
                    Case 6
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
                    Case 7
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
                    Case 8
                        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
                End Select
                If (Val(!STPETIC) And 64) <> 0 Or (Val(!STPETIC) And 128) <> 0 Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(12) = CargaCadena(1224) '"Cancelado"
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(12) = CargaCadena(1225) '"Normal"
                End If
                'formas de pago........
                'If !tipoPago = CLng(0) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(601) '"Moneda"
                'ElseIf !tipoPago = CLng(1) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1214) '"Cheque"
                'ElseIf !tipoPago = CLng(2) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1215) '"Tarjeta"
                'ElseIf !tipoPago = CLng(3) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Tick.Rest."      '"Tick.Rest."
                'ElseIf !tipoPago = CLng(4) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1216) '"Genérico"
                'ElseIf !tipoPago = CLng(255) Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1217) '"Anulado"
                'ElseIf !tipoPago = 12 Then
                '    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Trainning"
                'End If
                '''''''''''''''''''''''''''

                If !tipoPago = CLng(1) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Efectivo" 'CargaCadena(601) '"Moneda"
                ElseIf !tipoPago = CLng(2) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Tarjeta Crédito" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(3) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Tarjeta Débito" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(4) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "EBT Cash" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(5) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Gift Card" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(6) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cheque" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(7) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Pendiente" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(8) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(9) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(10) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(11) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(12) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(13) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(14) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Ticket Restaurant" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(15) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Chèque -Vacances" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(16) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cheque Gourmet" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(17) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Restaurant Pass" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(18) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(19) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cupón" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(20) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Others" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(21) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "EBT FoodStamp" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(22) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Visa" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(23) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "MasterCard" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(24) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "American Express" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(25) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Traveller 's Check" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(26) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(27) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(28) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Multibanco" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(29) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "--" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(30) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Visa Electron" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(31) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Diners Club" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(32) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Big Pass" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(33) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Sodexho" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(34) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Pin" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(35) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Chipknip" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(36) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cestaticket" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(37) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "EC -Maestro" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(38) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Quick" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(39) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "BPI -Express" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(40) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Gift Check" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(41) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Ripley" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(42) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Saga" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(43) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Otras tarjetas" 'cargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(44) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cheque Comercial" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(45) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Cheque de Gobiernos" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(46) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "CLAVE" 'CargaCadena(1214) '"Cheque"
                End If
                '''''''''''''''''''''''''''
                If IsNull(!cantidadEntre) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(14) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(14) = quitaNulo(!cantidadEntre)
                End If
                If IsNull(!Cambio) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(15) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(15) = quitaNulo(!Cambio)
                End If
                If IsNull(!descuento) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(16) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(16) = quitaNulo(!descuento)
                End If
                
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
    If Lista.ListItems.Count = 0 Then
        Lista.Enabled = False
    Else
        Lista.Enabled = True
    End If
End Sub
Private Function quitaNulo(ByVal nD As Double) As Double
    If IsNull(nD) Then nD = 0
    quitaNulo = nD
End Function
Private Sub Crea_Filtro()
    Dim Bucle As Integer
    Dim Filtro(4) As String
    For Bucle = 0 To 3
        Select Case CmbFiltro(Bucle).ListIndex
            Case 0
                Filtro(Bucle) = ""
            Case Else
                Select Case Bucle
                    Case 0
                        Filtro(Bucle) = " (seccion=" & Chr(34) & CmbFiltro(Bucle).TexT & Chr(34) & ") "
                    Case 1
                        Filtro(Bucle) = " (balanza=" & CmbFiltro(Bucle).TexT & ") "
                    Case 2
                        Filtro(Bucle) = " (fecha=" & Chr(34) & CmbFiltro(Bucle).TexT & Chr(34) & ") "
                    Case 3
                        '1.0.0 (0=Factura pasa a 7 red, 8 local
                        '       1=tique pasa a 1 red, 4 local
                        '       2=Envasado pasa a 2 red, 5 local
                        '       3=AutoServicio pasa a 3 red, 6 local
                        Select Case CmbFiltro(Bucle).ListIndex
                            Case 1
                                Filtro(Bucle) = " ( (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "4" & Chr(34) & "))"
                            Case 2
                                Filtro(Bucle) = " ( (typtic=" & Chr(34) & "7" & Chr(34) & " or typtic=" & Chr(34) & "8" & Chr(34) & "))"
                            Case 3
                                Filtro(Bucle) = " ( (typtic=" & Chr(34) & "2" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & "))"
                            Case 4
                                Filtro(Bucle) = " ( (typtic=" & Chr(34) & "3" & Chr(34) & " or typtic=" & Chr(34) & "6" & Chr(34) & "))"
                            Case 5
                                Filtro(Bucle) = "(tipoPago=" & CStr(1) & ")"
                            Case 6
                                Filtro(Bucle) = "(tipoPago=" & CStr(2) & ")"
                            Case 7
                                Filtro(Bucle) = "(tipoPago=" & CStr(3) & ")"
                            Case 8
                                Filtro(Bucle) = "(tipoPago=" & CStr(4) & ")"
                            Case 9
                                Filtro(Bucle) = "(tipoPago=" & CStr(5) & ")"
                            Case 10
                                Filtro(Bucle) = "(tipoPago=" & CStr(6) & ")"
                            Case 11
                                Filtro(Bucle) = "(tipoPago=" & CStr(7) & ")"
                            Case 12
                                Filtro(Bucle) = "(tipoPago=" & CStr(8) & ")"
                            Case 13
                                Filtro(Bucle) = "(tipoPago=" & CStr(9) & ")"
                            Case 14
                                Filtro(Bucle) = "(tipoPago=" & CStr(10) & ")"
                            Case 15
                                Filtro(Bucle) = "(tipoPago=" & CStr(11) & ")"
                            Case 16
                                Filtro(Bucle) = "(tipoPago=" & CStr(12) & ")"
                            Case 17
                                Filtro(Bucle) = "(tipoPago=" & CStr(13) & ")"
                            Case 18
                                Filtro(Bucle) = "(tipoPago=" & CStr(14) & ")"
                            Case 19
                                Filtro(Bucle) = "(tipoPago=" & CStr(15) & ")"
                            Case 20
                                Filtro(Bucle) = "(tipoPago=" & CStr(16) & ")"
                            Case 21
                                Filtro(Bucle) = "(tipoPago=" & CStr(17) & ")"
                            Case 22
                                Filtro(Bucle) = "(tipoPago=" & CStr(18) & ")"
                            Case 23
                                Filtro(Bucle) = "(tipoPago=" & CStr(19) & ")"
                            Case 24
                                Filtro(Bucle) = "(tipoPago=" & CStr(20) & ")"
                            Case 25
                                Filtro(Bucle) = "(tipoPago=" & CStr(21) & ")"
                            Case 26
                                Filtro(Bucle) = "(tipoPago=" & CStr(22) & ")"
                            Case 27
                                Filtro(Bucle) = "(tipoPago=" & CStr(23) & ")"
                            Case 28
                                Filtro(Bucle) = "(tipoPago=" & CStr(24) & ")"
                            Case 29
                                Filtro(Bucle) = "(tipoPago=" & CStr(25) & ")"
                            Case 30
                                Filtro(Bucle) = "(tipoPago=" & CStr(26) & ")"
                            Case 31
                                Filtro(Bucle) = "(tipoPago=" & CStr(27) & ")"
                            Case 32
                                Filtro(Bucle) = "(tipoPago=" & CStr(28) & ")"
                            Case 33
                                Filtro(Bucle) = "(tipoPago=" & CStr(29) & ")"
                            Case 34
                                Filtro(Bucle) = "(tipoPago=" & CStr(30) & ")"
                            Case 35
                                Filtro(Bucle) = "(tipoPago=" & CStr(31) & ")"
                            Case 36
                                Filtro(Bucle) = "(tipoPago=" & CStr(32) & ")"
                            Case 37
                                Filtro(Bucle) = "(tipoPago=" & CStr(33) & ")"
                            Case 38
                                Filtro(Bucle) = "(tipoPago=" & CStr(34) & ")"
                            Case 39
                                Filtro(Bucle) = "(tipoPago=" & CStr(35) & ")"
                            Case 40
                                Filtro(Bucle) = "(tipoPago=" & CStr(36) & ")"
                            Case 41
                                Filtro(Bucle) = "(tipoPago=" & CStr(37) & ")"
                            Case 42
                                Filtro(Bucle) = "(tipoPago=" & CStr(38) & ")"
                            Case 43
                                Filtro(Bucle) = "(tipoPago=" & CStr(39) & ")"
                            Case 44
                                Filtro(Bucle) = "(tipoPago=" & CStr(40) & ")"
                            Case 45
                                Filtro(Bucle) = "(tipoPago=" & CStr(41) & ")"
                            Case 46
                                Filtro(Bucle) = "(tipoPago=" & CStr(42) & ")"
                            '//cas.v117
                        End Select
                End Select
        End Select
    Next Bucle
    MyFilter = ""
    For Bucle = 0 To 3
        If Filtro(Bucle) <> "" Then
            MyFilter = MyFilter & " and " & Filtro(Bucle)
        End If
    Next Bucle
    If Len(MyFilter) >= 4 Then
        If left(MyFilter, 4) = " and" Then
            MyFilter = "where " & Mid(MyFilter, 5)
        End If
    End If
End Sub


Private Sub Form_Resize()
    Me_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Dir(App.Path & "\ht.txt") <> "" Then
        Kill App.Path & "\ht.txt"
    Else
        frmEpelsa.Enabled = True
    End If
End Sub

Private Sub Lista_Click()
    Dim eNCONTRADO As Boolean
    Dim Base As DAO.Database
    Dim Registro As DAO.Recordset
    Dim CadenaEnlace As String
    Dim TYPTIC As Integer
    Dim Num As Long
    Dim sbas As String
    
'Case 1
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
'Case 2
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
'Case 3
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
'Case 4
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
'Case 5
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
'Case 6
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
'Case 7
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
'Case 8
'    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
    
    Detalle.ListItems.Clear
    Select Case Lista.SelectedItem.SubItems(11)
        Case CargaCadena(336)
            TYPTIC = 1
        Case CargaCadena(195)
            TYPTIC = 2
        Case CargaCadena(196)
            TYPTIC = 3
        Case CargaCadena(1213)
            TYPTIC = 4
        Case CargaCadena(199)
            TYPTIC = 5
        Case CargaCadena(200)
            TYPTIC = 6
        Case "Fact. Net" 'CargaCadena(199)
            TYPTIC = 7
        Case "Fact. Local" 'CargaCadena(200)
            TYPTIC = 8
    End Select
    CadenaEnlace = "date=" & Chr(34) & Lista.SelectedItem.TexT & Chr(34) & _
    " and val(nume)=" & Val(Lista.SelectedItem.SubItems(2)) & " and numbal=" & Val(Lista.SelectedItem.SubItems(4)) _
    & " and val(secc)=" & Val(Lista.SelectedItem.SubItems(3)) & " and val(vendor)=" _
    & Val(Lista.SelectedItem.SubItems(5)) & " and val(typtic)=" & Val(TYPTIC) & " and hour=" & _
    Chr(34) & Lista.SelectedItem.SubItems(1) & Chr(34)
    
    If Dir(App.Path & "\ht.txt") <> "" Then
       nFichHys = FreeFile()
       Open App.Path & "\ht.txt" For Input As #nFichHys
       Line Input #nFichHys, sbas
       Close #nFichHys
       Set Base = OpenDatabase(sbas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If
    
    'Set Base = OpenDatabase(Base_General)
    eNCONTRADO = False
    Set Registro = Base.OpenRecordset("select * from tickets" & MyTabla & " where " & CadenaEnlace & " order by val(numlin)")
    With Registro
        If Not .EOF Then
            Do Until .EOF
                eNCONTRADO = False
                Select Case Lista.SelectedItem.SubItems(12)
                    'Case "Normal"
                    Case CargaCadena(1225)
                        If (Val(!STPETIC) And 64) = 0 And (Val(!STPETIC) And 128) = 0 Then
                            eNCONTRADO = True
                        End If
                    'Case "Cancelado"
                    Case CargaCadena(1224)
                        If (Val(!STPETIC) And 64) <> 0 Or (Val(!STPETIC) And 128) <> 0 Then
                            eNCONTRADO = True
                        End If
                End Select
                If eNCONTRADO Then
                    Detalle.ListItems.Add , , !numlin
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(1) = !Code
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(2) = !Plu
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(3) = !Item
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(4) = !price
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(5) = !Weight
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(6) = !units
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(7) = !amount
                    If Val(!lincan) = 0 Then
                        'Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = "NO"
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = CargaCadena(286)
                    Else
                        'Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = "SÍ"
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = CargaCadena(285)
                    End If
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(9) = !IVA
                End If
                .Movenext
            Loop
        End If
    End With
    CerrarBase Base
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    nIndex = ColumnHeader.Index
    If nIndex < 7 Then
        Call Refresca_Datos
    End If
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Lista_Click
End Sub

Private Sub Lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Lista_Click
End Sub
Private Sub ImprimeLosTiquets()
Dim Bucle As Integer
Dim Bucle2 As Integer
Dim CuentaLineas As Long
    Printer.FontName = "Courier New"
    Printer.FontSize = 14
    Printer.FontUnderline = True
    'Printer.Print "Tickets Sistema  " & Combo1.Text & " Listado el " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Printer.Print CargaCadena(1226) & " " & Combo1.TexT & ". " & CargaCadena(1227) & "  " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Printer.FontUnderline = False
    For Bucle = 1 To Lista.ListItems.Count
        Printer.FontSize = 8
        Printer.Print ""
        Printer.FontBold = True
        Printer.Print "-----------------------------------------------------------------------------------------------------------"
        Printer.Print CargaCadena(1218) & " : " & Lista.ListItems(Bucle).SubItems(2) & "#" & _
                      CargaCadena(434) & " : " & Lista.ListItems(Bucle).TexT & " " & Lista.ListItems(Bucle).SubItems(1) & "#" & _
                      CargaCadena(55) & " : " & Lista.ListItems(Bucle).SubItems(11) & "#" & _
                      CargaCadena(1221) & " : " & Lista.ListItems(Bucle).SubItems(12) & "#" & _
                      "EAN13: " & Lista.ListItems(Bucle).SubItems(10) & "#" & _
                      CargaCadena(132) & " : " & Lista.ListItems(Bucle).SubItems(8)
        Printer.FontBold = False
        Printer.Print CargaCadena(302) & " : " & Lista.ListItems(Bucle).SubItems(3) & "#" & _
                      CargaCadena(147) & " : " & Lista.ListItems(Bucle).SubItems(4) & "#" & _
                      CargaCadena(13) & " : " & Lista.ListItems(Bucle).SubItems(5) & "#" & _
                      CargaCadena(790) & " : " & Lista.ListItems(Bucle).SubItems(9)
        Printer.Print CargaCadena(1219) & " : " & Lista.ListItems(Bucle).SubItems(6) & "#" & _
                      CargaCadena(290) & " : " & Lista.ListItems(Bucle).SubItems(7) & "#" & _
                      CargaCadena(1220) & " : " & Lista.ListItems(Bucle).SubItems(13) & "#" & _
                      CargaCadena(1222) & " : " & Lista.ListItems(Bucle).SubItems(14) & "#" & _
                      CargaCadena(1223) & " : " & Lista.ListItems(Bucle).SubItems(15)
        Printer.Print CargaCadena(1276) & " : " & Lista.ListItems(Bucle).SubItems(16)
        'Printer.Print ""
        Lista.ListItems(Bucle).Selected = True
        Lista_Click
        Printer.FontBold = True
        'Printer.Print "Nº  Cancel. Cód.   PLU  Texto                     Precio  Peso     Unid.    Total    T.IVA"
        Printer.Print "Nº" & "  " & _
                      CargaCadena(1224) & " " & _
                      CargaCadena(69) & "   " & _
                      CargaCadena(68) & "  " & _
                      CargaCadena(539) & "                     " & _
                      CargaCadena(71) & "  " & _
                      CargaCadena(130) & "     " & _
                      CargaCadena(131) & "    " & _
                      CargaCadena(488) & "    " & _
                      CargaCadena(553)
        Printer.FontBold = False
        For Bucle2 = 1 To Detalle.ListItems.Count
            Printer.FontItalic = True
            Printer.Print Detalle.ListItems(Bucle2).TexT & Space(4 - Len(Detalle.ListItems(Bucle2).TexT)) & _
                            Detalle.ListItems(Bucle2).SubItems(8) & Space(8) & _
                            Detalle.ListItems(Bucle2).SubItems(1) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(1))) & _
                            Detalle.ListItems(Bucle2).SubItems(2) & Space(5 - Len(Detalle.ListItems(Bucle2).SubItems(2))) & _
                            Detalle.ListItems(Bucle2).SubItems(3) & Space(26 - Len(Detalle.ListItems(Bucle2).SubItems(3))) & _
                            Detalle.ListItems(Bucle2).SubItems(4) & Space(8 - Len(Detalle.ListItems(Bucle2).SubItems(4))) & _
                            Detalle.ListItems(Bucle2).SubItems(5) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(5))) & _
                            Detalle.ListItems(Bucle2).SubItems(6) & Space(12 - Len(Detalle.ListItems(Bucle2).SubItems(6))) & _
                            Detalle.ListItems(Bucle2).SubItems(7) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(7))) & _
                            Detalle.ListItems(Bucle2).SubItems(9) & Space(8 - Len(Detalle.ListItems(Bucle2).SubItems(9)))
        Next Bucle2
        Printer.FontItalic = False
    Next Bucle
    Printer.EndDoc
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConsultaTiquets 
   Caption         =   "Consulta de Tiquets"
   ClientHeight    =   8325
   ClientLeft      =   2325
   ClientTop       =   2430
   ClientWidth     =   11415
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Resumen impuestos.Datos disponibles Balanza"
      Height          =   615
      Left            =   3720
      TabIndex        =   21
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "Exportar Ficheros."
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resumen Impuestos"
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   5
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   4
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CmdTxtFich 
      Caption         =   "Generar fichero de Texto. (TTTDDMMYYHHMMSS.TXT)"
      Height          =   615
      Left            =   8040
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
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
      Top             =   120
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
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   6120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3836
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
      Height          =   3615
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6376
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
      ItemData        =   "FrmConsultaTiquets.frx":002E
      Left            =   1800
      List            =   "FrmConsultaTiquets.frx":0035
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
   Begin VB.Label Label1 
      Caption         =   "Client."
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Detalle del tiquet seleccionado :"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   5880
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
Private nPor0 As Double
Private nBas0 As Double
Private nImp0 As Double
Private nPor1 As Double
Private nBas1 As Double
Private nImp1 As Double
Private nPor2 As Double
Private nBas2 As Double
Private nImp2 As Double
Private nPor3 As Double
Private nBas3 As Double
Private nImp3 As Double
Private nPor4 As Double
Private nBas4 As Double
Private nImp4 As Double
Private nSuperBase As Double
Private nSuperImp As Double
Private NmYpAG As Long
Private nTiquesRealizados As Long

Private MyDateMin As Date
Private MyDateMax As Date
#Const BALSAM = False

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
Private Sub CmbFiltro_Click(Index As Integer)

    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim sBas As String



    If Not QL Then
        If Index = 2 Then
            CmbFiltro(4).Clear
            If Dir(App.Path & "\ht.txt") <> "" Then
                nFichHys = FreeFile()
                Open App.Path & "\ht.txt" For Input As #nFichHys
                Line Input #nFichHys, sBas
                Close #nFichHys
                Set Base = OpenDatabase(sBas)
            Else
                Set Base = OpenDatabase(Base_General)
            End If

            'Set Base = OpenDatabase(Base_General)
            '*****************
            ' datos de filtros
            '*****************
            If CmbFiltro(2).ListIndex = 0 Then
                Set Registro = Base.OpenRecordset("select distinct fecha from cabecera")
                CmbFiltro(4).AddItem CargaCadena(311)
            Else
                Set Registro = Base.OpenRecordset("select distinct fecha" & " from cabecera where cdate(fecha)>=#" & Format(CDate(CmbFiltro(2).TexT), "mm/dd/yy") & "#")
            End If
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        CmbFiltro(4).AddItem .Fields("fecha")
                        .Movenext
                    Loop
                End If
            End With
            CmbFiltro(4).TexT = CmbFiltro(4).List(CmbFiltro(4).ListCount - 1)
        End If


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

Private Sub Command2_Click()
    Dialogo.CancelError = True
    On Error Resume Next
    Dialogo.ShowPrinter
    If Err.Number = 0 Then
        Do_Events
        ImprimeImpuestos
    End If
End Sub

Private Sub CmdTxtFich_Click()
    FrmConsultaTiquets.Enabled = False
    GeneraFichero
    MsgBox "OK..."
    FrmConsultaTiquets.Enabled = True
End Sub



Private Sub Command3_Click()
    Dim cm As String
    Dim limpreso As Boolean
    Dim lsal As Boolean
    Dim d1 As Date
    Dim d2 As Date

    'Dim dIni As Date
    'Dim dFin As Date

    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        Me.Enabled = False
        Load FrmPeriodo
        FrmPeriodo.Show
        lsal = False
        Sleep (10)
        Do_Events
        salir = 0
        Do While salir = 0
            'lsal = FrmPeriodo.LSALIR()
            Do_Events
        Loop
        Me.Enabled = True
        If salir = 69 Then

            d1 = myPER1    'CDate(FrmPeriodo.TxtFecha(0).TexT)
            d2 = myPER2    'CDate(FrmPeriodo.TxtFecha(1).TexT)

            Dialogo.CancelError = True
            On Error Resume Next
            Dialogo.ShowPrinter
            If Err.Number = 0 Then

                Do_Events
                Printer.FontName = "Courier New"
                NmYpAG = 1
                Call ImpCab

                Printer.FontSize = 12
                limpreso = False
                Printer.Print
                'Printer.Print " " & CargaCadena(116) & ":" & Format(MyDateMin, "dd/mm/yyyy") & " " & CargaCadena(117) & ":" & Format(MyDateMax, "dd/mm/yyyy")
                'Printer.Print
                'Printer.Print CargaCadena(1420)
                'Printer.Print

                If MyRecSet.State <> 0 Then
                    MyRecSet.Close
                    Set MyRecSet = Nothing
                End If
                MyRecSet.CursorLocation = adUseClient
                'MyRecSet.Open "SELECT sum(amount) as amount,sum(base) as base,sum(tax) as tax,Percentage from pcscale.totalvat where (date>='" & Format(MyDateMin, "yyyy-mm-dd") & "' and date<='" & Format(MyDateMax, "yyyy-mm-dd") & "') group by percentage", MyConnObj, adOpenKeyset, adLockOptimistic
                MyRecSet.Open "SELECT sum(amount) as amount,sum(base) as base,sum(tax) as tax,Percentage from pcscale.totalvat where (date>='" & Format(d1, "yyyy-mm-dd") & "' and date<='" & Format(d2, "yyyy-mm-dd") & "') group by percentage", MyConnObj, adOpenKeyset, adLockOptimistic
                nSuperBase = 0
                nSuperImp = 0
                If Not MyRecSet.EOF Then
                    'Printer.Print " " & CargaCadena(116) & ":" & Format(MyDateMin, "dd/mm/yyyy") & " " & CargaCadena(117) & ":" & Format(MyDateMax, "dd/mm/yyyy")
                    Printer.Print " " & CargaCadena(116) & ":" & Format(d1, "dd/mm/yyyy") & " " & CargaCadena(117) & ":" & Format(d2, "dd/mm/yyyy")
                    Printer.Print
                    Printer.Print CargaCadena(1420)
                    Printer.Print

                    MyRecSet.MoveFirst
                    Do While Not MyRecSet.EOF
                        If limpreso = False Then
                            Call MyPrintCabIva
                            limpreso = True
                        End If
                        Printer.Print "                     " & FormatMy(MyRecSet.Fields("percentage"), "00.00") & FormatMy(MyRecSet.Fields("base"), "        0.00") & FormatMy(MyRecSet.Fields("tax"), "                0.00")
                        nSuperBase = nSuperBase + MyRecSet.Fields("base")
                        nSuperImp = nSuperImp + MyRecSet.Fields("tax")
                        MyRecSet.Movenext
                    Loop
                End If
                If MyRecSet.State <> 0 Then
                    MyRecSet.Close
                    Set MyRecSet = Nothing
                End If

                Printer.Print "                     _____________________________________"
                If limpreso = True Then
                    Printer.FontSize = 14
                    Printer.Print CargaCadena(1429) & FormatMy(nSuperBase, "       " & "0.00") & FormatMy(nSuperImp, "             0.00")
                    Printer.Print CargaCadena(1430) & FormatMy(nSuperBase + nSuperImp, "       0.00")
                Else
                    Printer.FontSize = 14
                    Printer.Print
                    Printer.Print
                End If
                Printer.FontSize = 12
                Printer.Print

                'Printer.Print " (Nº Tiques Listados:" & CStr(Lista.ListItems.Count) & ")"
                'If lIva = False Then
                '    Printer.Print CargaCadena(1431) & CargaCadena(1432) & CStr(Lista.ListItems.Count) & ")"  ' & "Fecha Inicial:" & Format(MyDateMin, "dd/mm/yyyy") & " Fecha Final:" & Format(MyDateMax, "dd/mm/yyyy")
                'End If

                Printer.EndDoc

            End If
        Else
            Exit Sub
        End If

    Else
        MsgBox "No existe conexión con Balaza..."
    End If
End Sub

Private Sub Form_Load()
    QL = True    'c2f 1.7.24

    Me_Init
    Me.Label1(0).Caption = CargaCadena(47)
    Me.Label1(1).Caption = CargaCadena(147)
    Me.Label1(2).Caption = CargaCadena(55)
    Me.Label1(3).Caption = CargaCadena(87)    'CargaCadena(434) & " Ini."
    Me.Label1(4).Caption = CargaCadena(88)    'CargaCadena(434) & " Fin."
    Me.Label2.Caption = CargaCadena(1228)
    Me.Command1.Caption = CargaCadena(1229)
    Me.Label3.Caption = CargaCadena(921)
    Me.Caption = CargaCadena(112)
    Me.Icon = Form2.Icon
    Command2.Caption = CargaCadena(1420)
    CmdTxtFich.Caption = CargaCadena(1421)

    #If BALSAM = False Then
        Command3.Visible = True
    #End If


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

    Detalle.ColumnHeaders.Add , , CargaCadena(1422), 600

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
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Integer
    Dim stCampo As String
    Dim Elegido As Integer
    Dim sBas As String

    If Dir(App.Path & "\ht.txt") <> "" Then
        nFichHys = FreeFile()
        Open App.Path & "\ht.txt" For Input As #nFichHys
        Line Input #nFichHys, sBas
        Close #nFichHys
        Set Base = OpenDatabase(sBas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If
    '*****************
    ' datos de filtros
    '*****************
    For bucle = 0 To 5
        CmbFiltro(bucle).Clear
    Next bucle
    For bucle = 0 To 5
        If bucle <> 3 Then
            CmbFiltro(bucle).Clear
            'CmbFiltro(Bucle).AddItem "(Todo)"
            CmbFiltro(bucle).AddItem CargaCadena(311)
            Select Case bucle
            Case 0
                stCampo = "seccion"
            Case 1
                stCampo = "balanza"
            Case 2
                stCampo = "fecha"
            Case 4
                stCampo = "fecha"
            Case 5
                stCampo = "codcli"
            End Select
            Set Registro = Base.OpenRecordset("select distinct " & stCampo & " from cabecera" & MyTabla)
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        CmbFiltro(bucle).AddItem .Fields(stCampo)
                        .Movenext
                    Loop
                End If
            End With
        End If
    Next bucle
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
    Me.CmbFiltro(3).AddItem CargaCadena(1435)
    '--------------
    Me.CmbFiltro(3).AddItem CargaCadena(1436)
    Me.CmbFiltro(3).AddItem CargaCadena(1437)
    Me.CmbFiltro(3).AddItem CargaCadena(1438)
    Me.CmbFiltro(3).AddItem "Recar."
    Me.CmbFiltro(3).AddItem "Cancel."
    Me.CmbFiltro(3).AddItem "Order"
    ''''''''''''''''''''''''''''''''''
    'Me.CmbFiltro(3).AddItem "Cobrado"
    'Me.CmbFiltro(3).AddItem "No Cobrado"
    'Me.CmbFiltro(3).AddItem "Trainning"
    'Me.CmbFiltro(3).AddItem "Cliente"
    For bucle = 0 To 5
        CmbFiltro(bucle).TexT = CmbFiltro(bucle).List(0)
    Next bucle
    If CmbFiltro(2).ListCount > 1 Then
        CmbFiltro(2).TexT = CmbFiltro(2).List(CmbFiltro(2).ListCount - 1)
        QL = False
        CmbFiltro_Click (2)
        QL = True
    End If

    'If CmbFiltro(4).ListCount > 1 Then
    '    CmbFiltro(4).TexT = CmbFiltro(4).List(CmbFiltro(4).ListCount - 1)
    'End If

End Sub
Private Sub Refresca_Datos()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegCli As dao.Recordset

    Dim MiRefresco As Integer
    Dim sOrder As String
    Dim sBas As String

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
    Lista.ColumnHeaders.Add , , CargaCadena(434), 1000    '"Fecha"
    Lista.ColumnHeaders.Add , , CargaCadena(435), 1000    '"Hora"
    Lista.ColumnHeaders.Add , , CargaCadena(1218), 1000    '"Numero"
    Lista.ColumnHeaders.Add , , CargaCadena(47), 700    '"Sección"
    Lista.ColumnHeaders.Add , , CargaCadena(147), 700    '"Balanza"
    Lista.ColumnHeaders.Add , , CargaCadena(13), 1000    '"Vendedor"
    Lista.ColumnHeaders.Add , , CargaCadena(1219), 700    '"Líneas"
    Lista.ColumnHeaders.Add , , CargaCadena(792)    '"Líneas Canceladas"
    Lista.ColumnHeaders.Add , , CargaCadena(132)    '"Importe"
    Lista.ColumnHeaders.Add , , CargaCadena(790)    '"Cliente"
    Lista.ColumnHeaders.Add , , "EAN13"
    Lista.ColumnHeaders.Add , , CargaCadena(55)    '"Tipo"
    Lista.ColumnHeaders.Add , , CargaCadena(1221)    '"Estado"
    Lista.ColumnHeaders.Add , , CargaCadena(1220), 1300    '"Forma de pago"
    Lista.ColumnHeaders.Add , , CargaCadena(1222), 1500    '"Importe entregado"
    Lista.ColumnHeaders.Add , , CargaCadena(1223), 700    '"Cambio"
    Lista.ColumnHeaders.Add , , CargaCadena(1276), 770    'Descuentos
    Lista.ColumnHeaders.Add , , CargaCadena(1427), 2000

    Lista.ColumnHeaders.Add , , CargaCadena(1422), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1423), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1424), 700

    Lista.ColumnHeaders.Add , , CargaCadena(1422), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1423), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1424), 700

    Lista.ColumnHeaders.Add , , CargaCadena(1422), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1423), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1424), 700

    Lista.ColumnHeaders.Add , , CargaCadena(1422), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1423), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1424), 700

    Lista.ColumnHeaders.Add , , CargaCadena(1422), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1423), 700
    Lista.ColumnHeaders.Add , , CargaCadena(1424), 700

    Lista.ColumnHeaders.Add , , CargaCadena(1425), 2000
    Lista.ColumnHeaders.Add , , CargaCadena(1426), 2000

    Lista.ListItems.Clear

    Crea_Filtro


    If Dir(App.Path & "\ht.txt") <> "" Then
        nFichHys = FreeFile()
        Open App.Path & "\ht.txt" For Input As #nFichHys
        Line Input #nFichHys, sBas
        Close #nFichHys
        Set Base = OpenDatabase(sBas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If

    nPor0 = 0
    nBas0 = 0
    nImp0 = 0
    nPor1 = 0
    nBas1 = 0
    nImp1 = 0
    nPor2 = 0
    nBas2 = 0
    nImp2 = 0
    nPor3 = 0
    nBas3 = 0
    nImp3 = 0
    nPor4 = 0
    nBas4 = 0
    nImp4 = 0

    nTiquesRealizados = 0

    MyDateMin = Format("01/01/2099", "dd/mm/yyyy")
    MyDateMax = Format("01/01/1961", "dd/mm/yyyy")

    'Set Base = OpenDatabase(Base_General)
    If CmbFiltro(3).TexT <> CargaCadena(1435) Then
        If MyFilter = "" Then
            Set Registro = Base.OpenRecordset("select * from cabecera" & MyTabla & " where TYPTIC <> '12' " & sOrder)
        Else
            Set Registro = Base.OpenRecordset("select * from cabecera" & MyTabla & " " & MyFilter & " and TYPTIC <> '12' " & sOrder)
        End If
    Else
        Set Registro = Base.OpenRecordset("select * from cabecera" & MyTabla & " " & MyFilter & sOrder)
    End If
    With Registro
        If Not .EOF Then

            Me.Enabled = False

            Do Until .EOF

                nTiquesRealizados = nTiquesRealizados + 1

                MiRefresco = MiRefresco + 1
                If MiRefresco > 200 Then
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
                'Set RegCli = Base.OpenRecordset("select * from cliente where cod_cli=" & CStr(Val(!codcli)))
                Set RegCli = Base.OpenRecordset("select * from cliente where trim(ucase(data3))='" & Trim(UCase(!codcli)) & "'")
                If RegCli.EOF Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(33) = "Consumidor Final"
                    Lista.ListItems(Lista.ListItems.Count).SubItems(34) = "999999990"
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(33) = Trim(RegCli.Fields("nom_cli"))
                    'Lista.ListItems(Lista.ListItems.Count).SubItems(34) = Trim(RegCli.Fields("data1"))
                    Lista.ListItems(Lista.ListItems.Count).SubItems(34) = Trim(!codcli)    'Trim(RegCli.Fields("data3"))
                End If
                RegCli.Close
                Set RegCli = Nothing

                Select Case !TYPTIC
                Case 1
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336)    '"Venta"
                Case 2
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195)    '"Envasado" 'CargaCadena(741) '"Super"
                Case 3
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196)    '"Autoservicio" 'CargaCadena(195) '"Envasado"
                Case 4
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213)    '"Venta Local" 'CargaCadena(196) '"Autoservicio"
                Case 5
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199)    '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
                Case 6
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200)    '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
                Case 7
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net"    '"Envasado Local"
                Case 8
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local"    '"Autoservicio Local"
                Case 12
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Descart."
                Case 13
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Recar."
                Case 14
                    Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Cancel."

                End Select
                If (Val(!STPETIC) And 64) <> 0 Or (Val(!STPETIC) And 128) <> 0 Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(12) = CargaCadena(1224)    '"Cancelado"
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(12) = CargaCadena(1225)    '"Normal"
                End If

                If !tipoPago = CLng(0) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1436)    '"Efectivo" 'CargaCadena(601) '"Moneda"
                ElseIf !tipoPago = CLng(1) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1437)    '"Tarjeta Crédito" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(2) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = CargaCadena(1438)    '"Cheque" 'CargaCadena(1214) '"Cheque"
                ElseIf !tipoPago = CLng(12) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(13) = "Order"
                End If
                '''''''''''''''''''''''''''
                Lista.ListItems(Lista.ListItems.Count).SubItems(14) = !cantidadEntre
                Lista.ListItems(Lista.ListItems.Count).SubItems(15) = !Cambio
                Lista.ListItems(Lista.ListItems.Count).SubItems(16) = !descuento
                If IsNull(!factura) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(17) = Space(15)
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(17) = !factura
                End If
                Lista.ListItems(Lista.ListItems.Count).SubItems(18) = Format(!por0, "0.00")
                Lista.ListItems(Lista.ListItems.Count).SubItems(19) = !base0
                'Lista.ListItems(Lista.ListItems.Count).SubItems(20) = !imp0
                If IsNull(!imp0) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(20) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(20) = !imp0
                End If


                Lista.ListItems(Lista.ListItems.Count).SubItems(21) = Format(!por1, "0.00")
                Lista.ListItems(Lista.ListItems.Count).SubItems(22) = !base1
                'Lista.ListItems(Lista.ListItems.Count).SubItems(23) = !imp1
                If IsNull(!imp1) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(23) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(23) = !imp1
                End If


                Lista.ListItems(Lista.ListItems.Count).SubItems(24) = Format(!por2, "0.00")
                Lista.ListItems(Lista.ListItems.Count).SubItems(25) = !base2
                'Lista.ListItems(Lista.ListItems.Count).SubItems(26) = !imp2
                If IsNull(!imp2) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(26) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(26) = !imp2
                End If


                Lista.ListItems(Lista.ListItems.Count).SubItems(27) = Format(!por3, "0.00")
                Lista.ListItems(Lista.ListItems.Count).SubItems(28) = !base3
                'Lista.ListItems(Lista.ListItems.Count).SubItems(29) = !imp3
                If IsNull(!imp3) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(29) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(29) = !imp3
                End If

                Lista.ListItems(Lista.ListItems.Count).SubItems(30) = Format(!por4, "0.00")
                Lista.ListItems(Lista.ListItems.Count).SubItems(31) = !base4
                If IsNull(!imp4) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(32) = 0
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(32) = !imp4
                End If

                '2.0.5
                If Lista.ListItems(Lista.ListItems.Count).SubItems(12) <> CargaCadena(1224) And Lista.ListItems(Lista.ListItems.Count).SubItems(11) <> "Recar." And Lista.ListItems(Lista.ListItems.Count).SubItems(11) <> "Descart." Then
                    'CadenadeLog "Tique:" & !nume & CStr(!base0) & " " & CStr(!base1) & " " & CStr(!base2) & " " & CStr(!base3) & " " & CStr(!base4)
                    If nPor0 = 0 And !por0 <> 0 Then nPor0 = !por0
                    nBas0 = nBas0 + !base0
                    nImp0 = nImp0 + !imp0

                    If nPor1 = 0 And !por1 <> 0 Then nPor1 = !por1
                    nBas1 = nBas1 + !base1
                    nImp1 = nImp1 + !imp1

                    If nPor2 = 0 And !por2 <> 0 Then nPor2 = !por2
                    nBas2 = nBas2 + !base2
                    nImp2 = nImp2 + !imp2

                    If nPor3 = 0 And !por3 <> 0 Then nPor3 = !por3
                    nBas3 = nBas3 + !base3
                    nImp3 = nImp3 + !imp3

                    If nPor4 = 0 And !por4 <> 0 Then nPor4 = !por4
                    If Not IsNull(!base4) And Not IsNull(!imp4) Then
                        nBas4 = nBas4 + !base4
                        nImp4 = nImp4 + !imp4
                    Else
                        .Edit
                        !base4 = 0
                        !imp4 = 0
                        .Update
                        nBas4 = nBas4 + 0
                        nImp4 = nImp4 + 0

                    End If
                End If

                '''''''''''''
                If MyDateMin > !D_FECHA Then MyDateMin = !D_FECHA
                If MyDateMax < !D_FECHA Then MyDateMax = !D_FECHA

                .Movenext
            Loop

            Me.Enabled = True

        End If
    End With
    CerrarBase Base

    nTiquesRealizados = Lista.ListItems.Count

    If Lista.ListItems.Count = 0 Then
        Lista.Enabled = False
        nTiquesRealizados = 0
    Else
        Lista.Enabled = True
    End If
End Sub
Private Sub Crea_Filtro()
    Dim bucle As Integer
    Dim Filtro(5) As String
    For bucle = 0 To 5
        Select Case CmbFiltro(bucle).ListIndex
        Case 0
            If bucle <> 4 Or (LCase(CmbFiltro(bucle).TexT) = LCase(CargaCadena(311))) Then
                Filtro(bucle) = ""
            Else
                Filtro(bucle) = " (cdate(fecha)<=" & "#" & Format(CmbFiltro(bucle).TexT, "mm/dd/yy") & "#) "
            End If
        Case Else
            Select Case bucle
            Case 0
                Filtro(bucle) = " (seccion=" & Chr(34) & CmbFiltro(bucle).TexT & Chr(34) & ") "
            Case 1
                Filtro(bucle) = " (balanza=" & CmbFiltro(bucle).TexT & ") "
            Case 2
                Filtro(bucle) = " (cdate(fecha)>=#" & Format(CDate(CmbFiltro(bucle).TexT), "mm/dd/yy") & "#) "
            Case 4
                Filtro(bucle) = " (cdate(fecha)<=#" & Format(CDate(CmbFiltro(bucle).TexT), "mm/dd/yy") & "#) "
            Case 5
                'Filtro(Bucle) = " (val(codcli)=" & CmbFiltro(Bucle).TexT & ") "
                Filtro(bucle) = " (codcli='" & CmbFiltro(bucle).TexT & "') "
            Case 3
                '1.0.0 (0=Factura pasa a 7 red, 8 local
                '       1=tique pasa a 1 red, 4 local
                '       2=Envasado pasa a 2 red, 5 local
                '       3=AutoServicio pasa a 3 red, 6 local
                Select Case CmbFiltro(bucle).ListIndex
                Case 1
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "4" & Chr(34) & "))"
                Case 2
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "7" & Chr(34) & " or typtic=" & Chr(34) & "8" & Chr(34) & "))"
                Case 3
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "2" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & "))"
                Case 4
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "3" & Chr(34) & " or typtic=" & Chr(34) & "6" & Chr(34) & "))"
                Case 5
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "12" & Chr(34) & "))"
                Case 6
                    Filtro(bucle) = "(tipoPago=" & CStr(0) & ")"
                Case 7
                    Filtro(bucle) = "(tipoPago=" & CStr(1) & ")"
                Case 8
                    Filtro(bucle) = "(tipoPago=" & CStr(2) & ")"
                Case 9
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "13" & Chr(34) & "))"
                Case 10
                    Filtro(bucle) = " ( (typtic=" & Chr(34) & "14" & Chr(34) & "))"
                Case 11
                    Filtro(bucle) = "(tipoPago=" & CStr(12) & ")"
                    'Case 9
                    '    Filtro(bucle) = "(tipoPago=" & CStr(4) & ")"
                    'Case 10
                    '    Filtro(bucle) = "(tipoPago=" & CStr(5) & ")"
                    'Case 11
                    '    Filtro(bucle) = "(tipoPago=" & CStr(6) & ")"
                    'Case 12
                    '    Filtro(bucle) = "(tipoPago=" & CStr(7) & ")"
                    'Case 13
                    '    Filtro(bucle) = "(tipoPago=" & CStr(8) & ")"
                    'Case 14
                    '    Filtro(bucle) = "(tipoPago=" & CStr(9) & ")"
                    'Case 15
                    '    Filtro(bucle) = "(tipoPago=" & CStr(10) & ")"
                    'Case 16
                    '    Filtro(bucle) = "(tipoPago=" & CStr(11) & ")"
                    'Case 17
                    '    Filtro(bucle) = "(tipoPago=" & CStr(12) & ")"
                    'Case 18
                    '    Filtro(bucle) = "(tipoPago=" & CStr(13) & ")"
                    'Case 19
                    '    Filtro(bucle) = "(tipoPago=" & CStr(14) & ")"
                    'Case 20
                    '    Filtro(bucle) = "(tipoPago=" & CStr(15) & ")"
                    'Case 21
                    '    Filtro(bucle) = "(tipoPago=" & CStr(16) & ")"
                    'Case 22
                    '    Filtro(bucle) = "(tipoPago=" & CStr(17) & ")"
                    'Case 23
                    '    Filtro(bucle) = "(tipoPago=" & CStr(18) & ")"
                    'Case 24
                    '    Filtro(bucle) = "(tipoPago=" & CStr(19) & ")"
                    'Case 25
                    '    Filtro(bucle) = "(tipoPago=" & CStr(20) & ")"
                    'Case 26
                    '    Filtro(bucle) = "(tipoPago=" & CStr(21) & ")"
                    'Case 27
                    '    Filtro(bucle) = "(tipoPago=" & CStr(22) & ")"
                    'Case 28
                    '    Filtro(bucle) = "(tipoPago=" & CStr(23) & ")"
                    'Case 29
                    '    Filtro(bucle) = "(tipoPago=" & CStr(24) & ")"
                    'Case 30
                    '    Filtro(bucle) = "(tipoPago=" & CStr(25) & ")"
                    'Case 31
                    '    Filtro(bucle) = "(tipoPago=" & CStr(26) & ")"
                    'Case 32
                    '    Filtro(bucle) = "(tipoPago=" & CStr(27) & ")"
                    'Case 33
                    '    Filtro(bucle) = "(tipoPago=" & CStr(28) & ")"
                    'Case 34
                    '    Filtro(bucle) = "(tipoPago=" & CStr(29) & ")"
                    'Case 35
                    '    Filtro(bucle) = "(tipoPago=" & CStr(30) & ")"
                    'Case 36
                    '    Filtro(bucle) = "(tipoPago=" & CStr(31) & ")"
                    'Case 37
                    '    Filtro(bucle) = "(tipoPago=" & CStr(32) & ")"
                    'Case 38
                    '    Filtro(bucle) = "(tipoPago=" & CStr(33) & ")"
                    'Case 39
                    '    Filtro(bucle) = "(tipoPago=" & CStr(34) & ")"
                    'Case 40
                    '    Filtro(bucle) = "(tipoPago=" & CStr(35) & ")"
                    'Case 41
                    '    Filtro(bucle) = "(tipoPago=" & CStr(36) & ")"
                    'Case 42
                    '    Filtro(bucle) = "(tipoPago=" & CStr(37) & ")"
                    'Case 43
                    '    Filtro(bucle) = "(tipoPago=" & CStr(38) & ")"
                    'Case 44
                    '    Filtro(bucle) = "(tipoPago=" & CStr(39) & ")"
                    'Case 45
                    '    Filtro(bucle) = "(tipoPago=" & CStr(40) & ")"
                    'Case 46
                    '    Filtro(bucle) = "(tipoPago=" & CStr(41) & ")"
                    'Case 47
                    '    Filtro(bucle) = "(tipoPago=" & CStr(42) & ")"
                    ''//cas.v117
                End Select
            End Select
        End Select
    Next bucle
    MyFilter = ""
    For bucle = 0 To 5
        If Filtro(bucle) <> "" Then
            MyFilter = MyFilter & " and " & Filtro(bucle)
        End If
    Next bucle
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
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim CadenaEnlace As String
    Dim TYPTIC As Integer
    Dim Num As Long
    Dim sBas As String
    Dim nD As Integer
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
    'Cabecera.tipoTiquet
    Detalle.ListItems.Clear
    Select Case Lista.SelectedItem.SubItems(11)
    Case CargaCadena(336)    'venta
        TYPTIC = 1
    Case CargaCadena(195)    'envasado
        TYPTIC = 2
    Case CargaCadena(196)    'autoservicio
        TYPTIC = 3
    Case CargaCadena(1213)    'Venta Local
        TYPTIC = 4
    Case CargaCadena(199)    'Envasado Local
        TYPTIC = 5
    Case CargaCadena(200)    'Autoservicio local
        TYPTIC = 6
    Case "Fact. Net"    'CargaCadena(199)
        TYPTIC = 7
    Case "Fact. Local"    'CargaCadena(200)
        TYPTIC = 8
    Case "Descart."    '"Training"
        TYPTIC = 12
    Case "Recar."    'recargados
        TYPTIC = 13
    Case "Cancel."    'Cancelados
        TYPTIC = 14
    End Select

    nD = 0

    CadenaEnlace = "date=" & Chr(34) & Lista.SelectedItem.TexT & Chr(34) & _
                 " and val(nume)=" & Val(Lista.SelectedItem.SubItems(2)) & " and numbal=" & Val(Lista.SelectedItem.SubItems(4)) _
                 & " and val(secc)=" & Val(Lista.SelectedItem.SubItems(3)) & " and val(vendor)=" _
                 & Val(Lista.SelectedItem.SubItems(5)) & " and val(typtic)=" & Val(TYPTIC) & " and hour=" & _
                   Chr(34) & Lista.SelectedItem.SubItems(1) & Chr(34)

    CadenadeLog CadenaEnlace
    On Error GoTo sEEnlace

    If Dir(App.Path & "\ht.txt") <> "" Then
        nFichHys = FreeFile()
        Open App.Path & "\ht.txt" For Input As #nFichHys
        Line Input #nFichHys, sBas
        Close #nFichHys
        Set Base = OpenDatabase(sBas)
    Else
        Set Base = OpenDatabase(Base_General)
    End If

    'Set Base = OpenDatabase(Base_General)
    eNCONTRADO = False
    Set Registro = Base.OpenRecordset("select * from tickets" & MyTabla & " where " & CadenaEnlace & " order by val(numlin)")
    With Registro
        If Not .EOF Then
            Do Until .EOF
                nD = 1
                eNCONTRADO = False
                Select Case Lista.SelectedItem.SubItems(12)
                    'nD = 2
                    'Case "Normal"
                Case CargaCadena(1225)
                    nD = 3
                    If (Val(!STPETIC) And 64) = 0 And (Val(!STPETIC) And 128) = 0 Then
                        eNCONTRADO = True
                    End If
                    'Case "Cancelado"
                Case CargaCadena(1224)
                    nD = 4
                    If (Val(!STPETIC) And 64) <> 0 Or (Val(!STPETIC) And 128) <> 0 Then
                        eNCONTRADO = True
                    End If
                End Select
                If eNCONTRADO Then
                    nD = 5
                    Detalle.ListItems.Add , , !numlin
                    nD = 6
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(1) = !Code
                    nD = 7
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(2) = !Plu
                    nD = 8
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(3) = !Item
                    nD = 9
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(4) = !Price
                    nD = 10
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(5) = !Weight
                    nD = 11
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(6) = !units
                    nD = 12
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(7) = !amount
                    nD = 13
                    If Val(!lincan) = 0 Then
                        'Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = "NO"
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = CargaCadena(286)
                    Else
                        'Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = "SÍ"
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(8) = CargaCadena(285)
                    End If
                    nD = 14
                    Detalle.ListItems(Detalle.ListItems.Count).SubItems(9) = !IVA
                    nD = 15
                    If IsNull(!porcentaje) Then
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(10) = 0
                    Else
                        Detalle.ListItems(Detalle.ListItems.Count).SubItems(10) = !porcentaje
                    End If
                End If
                .Movenext
            Loop
        End If
    End With
sEEnlace:
    If Err.Number <> 0 Then
        MsgBox "Err. Det." & nD
    End If
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
Private Function FormatMy(ByVal s As String, ByVal sF As String) As String
    Dim nL As Integer
    s = Format(s, Trim(sF))
    If Len(s) < Len(sF) Then
        s = Space(Len(sF) - Len(s)) & s
    Else
        s = Mid(s, 1, Len(sF))
    End If
    FormatMy = s
End Function
Private Sub MyPrintCabIva()
    Printer.Print CargaCadena(1428)
    Printer.Print "                     _____________________________________"
End Sub
Private Sub ImpCab()
    Printer.FontSize = 14
    Printer.FontUnderline = True
    'Printer.Print " " & CargaCadena(1226) & " " & Combo1.TexT & ". " & CargaCadena(1227) & "  " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Printer.FontBold = True
    Printer.Print " " & Mid(Empresa, 1, 80) & " "
    Printer.FontUnderline = False
    Printer.Print " (pag. " & CStr(NmYpAG) & " )"
    Printer.FontBold = False
End Sub
'Private Sub ImpResu(Optional ByVal lIva As Boolean)
'Dim limpreso As Boolean
'    Printer.FontSize = 12
'    limpreso = False
'    Printer.Print
'    Printer.Print " " & CargaCadena(116) & ":" & Format(MyDateMin, "dd/mm/yyyy") & " " & CargaCadena(117) & ":" & Format(MyDateMax, "dd/mm/yyyy")
'    Printer.Print
'    If lIva Then
'        Printer.Print CargaCadena(1420)
'    Else
'        Printer.Print CargaCadena(1439)
'    End If
'    'Printer.Print " -----------------------------------------------------------------------------"
'    Printer.Print
'        If nPor0 <> 0 Or nBas0 <> 0 Then
'            If limpreso = False Then Call MyPrintCabIva
'            limpreso = True
'            Printer.Print "                     " & Format(nPor0, "00.00") & FormatMy(nBas0, "        0.00") & FormatMy(nImp0, "                0.00")
'        End If
'        If nPor1 <> 0 Or nBas1 <> 0 Then
'            If limpreso = False Then Call MyPrintCabIva
'            limpreso = True
'            Printer.Print "                     " & Format(nPor1, "00.00") & FormatMy(nBas1, "        0.00") & FormatMy(nImp1, "                0.00")
'        End If
'        If nPor2 <> 0 Or nBas2 <> 0 Then
'            If limpreso = False Then Call MyPrintCabIva
'            limpreso = True
'            Printer.Print "                     " & Format(nPor2, "00.00") & FormatMy(nBas2, "        0.00") & FormatMy(nImp2, "                0.00")
'        End If
'        If nPor3 <> 0 Or nBas3 <> 0 Then
'            If limpreso = False Then Call MyPrintCabIva
'            limpreso = True
'            Printer.Print "                     " & Format(nPor3, "00.00") & FormatMy(nBas3, "        0.00") & FormatMy(nImp3, "                0.00")
'        End If
'        If nPor4 <> 0 Or nBas4 <> 0 Then
'            If limpreso = False Then Call MyPrintCabIva
'            limpreso = True
'            Printer.Print "                     " & Format(nPor4, "00.00") & FormatMy(nBas4, "        0.00") & FormatMy(nImp4, "                0.00")
'        End If'
'
'    Printer.Print "                     _____________________________________"
'    If limpreso = True Then
'        Printer.FontSize = 14
'        Printer.Print CargaCadena(1429) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4, "       " & "0.00") & FormatMy(nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "             0.00")
'        Printer.Print CargaCadena(1430) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4 + nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "       0.00")
'    Else
'        Printer.FontSize = 14
'        Printer.Print
'        Printer.Print
'    End If
'    Printer.FontSize = 12
'    Printer.Print
'
'    'Printer.Print " (Nº Tiques Listados:" & CStr(Lista.ListItems.Count) & ")"
'    If lIva = False Then
'        Printer.Print CargaCadena(1431) & CargaCadena(1432) & CStr(Lista.ListItems.Count) & ")"  ' & "Fecha Inicial:" & Format(MyDateMin, "dd/mm/yyyy") & " Fecha Final:" & Format(MyDateMax, "dd/mm/yyyy")
'    End If
'End Sub

Private Sub ImpResu(Optional ByVal lIva As Boolean)
    Dim limpreso As Boolean
    Dim nImp0R, nImp1R, nImp2R, nImp3R, nImp4R As Double
    nImp0R = Round((nBas0 * nPor0) / 100, 2)
    nImp1R = Round((nBas1 * nPor1) / 100, 2)
    nImp2R = Round((nBas2 * nPor2) / 100, 2)
    nImp3R = Round((nBas3 * nPor3) / 100, 2)
    nImp4R = Round((nBas4 * nPor4) / 100, 2)
    Printer.FontSize = 12
    limpreso = False
    Printer.Print
    Printer.Print " " & CargaCadena(87) & ":" & Format(MyDateMin, "dd/mm/yyyy") & " " & CargaCadena(88) & ":" & Format(MyDateMax, "dd/mm/yyyy")
    Printer.Print
    'If lIva Then
    '    Printer.Print CargaCadena(1361) '" Resumen Impuestos."
    'Else
    '    Printer.Print CargaCadena(1361) '" Resumen Impuestos del Listado."
    'End If
    If lIva Then
        Printer.Print CargaCadena(1420)
    Else
        Printer.Print CargaCadena(1439)
    End If

    Printer.Print
    Printer.Print CargaCadena(134) & ": " & CStr(nTiquesRealizados)  '" Número de Tiques realizados:" & CStr(nTiquesRealizados)
    'Printer.Print " -----------------------------------------------------------------------------"
    Printer.Print
    If Dir(App.Path & "\sobrebase") = "" Then
        If nPor0 <> 0 Or nBas0 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor0, "00.00") & FormatMy(nBas0, "        0.00") & FormatMy(nImp0, "                0.00")
        End If
        If nPor1 <> 0 Or nBas1 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor1, "00.00") & FormatMy(nBas1, "        0.00") & FormatMy(nImp1, "                0.00")
        End If
        If nPor2 <> 0 Or nBas2 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor2, "00.00") & FormatMy(nBas2, "        0.00") & FormatMy(nImp2, "                0.00")
        End If
        If nPor3 <> 0 Or nBas3 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor3, "00.00") & FormatMy(nBas3, "        0.00") & FormatMy(nImp3, "                0.00")
        End If
        If nPor4 <> 0 Or nBas4 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor4, "00.00") & FormatMy(nBas4, "        0.00") & FormatMy(nImp4, "                0.00")
        End If
    Else
        Printer.Print "Impuesto calculado respecto a total de cada base."
        If nPor0 <> 0 Or nBas0 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor0, "00.00") & FormatMy(nBas0, "        0.00") & FormatMy(nImp0R, "                0.00")
        End If
        If nPor1 <> 0 Or nBas1 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor1, "00.00") & FormatMy(nBas1, "        0.00") & FormatMy(nImp1R, "                0.00")
        End If
        If nPor2 <> 0 Or nBas2 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor2, "00.00") & FormatMy(nBas2, "        0.00") & FormatMy(nImp2R, "                0.00")
        End If
        If nPor3 <> 0 Or nBas3 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor3, "00.00") & FormatMy(nBas3, "        0.00") & FormatMy(nImp3R, "                0.00")
        End If
        If nPor4 <> 0 Or nBas4 Then
            If limpreso = False Then Call MyPrintCabIva
            limpreso = True
            Printer.Print "                     " & Format(nPor4, "00.00") & FormatMy(nBas4, "        0.00") & FormatMy(nImp4R, "                0.00")
        End If

    End If
    Printer.Print "                     _____________________________________"

    If limpreso = True Then
        Printer.FontSize = 14
        If Dir(App.Path & "\sobrebase") = "" Then
            Printer.Print CargaCadena(1429) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4, "       " & "0.00") & FormatMy(nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "             0.00")
            Printer.Print CargaCadena(1430) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4 + nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "       0.00")
        Else
            Printer.Print CargaCadena(1429) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4, "       " & "0.00") & FormatMy(nImp0R + nImp1R + nImp2R + nImp3R + nImp4R, "             0.00")
            Printer.Print CargaCadena(1430) & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4 + nImp0R + nImp1R + nImp2R + nImp3R + nImp4R, "       0.00")
        End If
    Else
        Printer.FontSize = 14
        Printer.Print
        Printer.Print
    End If
    Printer.FontSize = 12
    Printer.Print

    'Printer.FontSize = 14
    'Printer.Print "              Totales:" & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4, "       " & "0.00") & FormatMy(nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "             0.00")
    'Printer.Print " TOTAL Base+Impuestos:" & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4 + nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "       0.00")
    'Printer.FontSize = 12
    'Printer.Print

    'Printer.Print " (Nº Tiques Listados:" & CStr(Lista.ListItems.Count) & ")"
    If lIva Then

    Else
        Printer.Print CargaCadena(1431) & CargaCadena(1432) & CStr(Lista.ListItems.Count) & ")"    ' & "Fecha Inicial:" & Format(MyDateMin, "dd/mm/yyyy") & " Fecha Final:" & Format(MyDateMax, "dd/mm/yyyy")
    End If
End Sub

Private Sub MyFichCabIva(ByRef nf As Integer)
    Print #nf, CargaCadena(1428)
    Print #nf, "                     _____________________________________"
End Sub

Private Sub ImprimeImpuestos()

    Printer.FontName = "Courier New"

    NmYpAG = 1

    Call ImpCab

    ImpResu True

    Printer.EndDoc

End Sub

Private Sub ImprimeLosTiquets()
    Dim bucle As Integer
    Dim Bucle2 As Integer
    Dim CuentaLineas As Long
    Dim limpreso As Boolean
    Dim nPag As Long
    Dim nLineas As Integer
    Dim nPar As Integer
    Dim nItem As Integer
    Dim sParte As String

    Printer.FontName = "Courier New"

    NmYpAG = 1
    nItem = 1

    Call ImpCab

    Call ImpResu

    nLineas = 18

    limpreso = False

    For bucle = 1 To Lista.ListItems.Count

        Lista.ListItems(bucle).Selected = True
        Lista_Click
        nPar = 0

        '2.0.5
        If Lista.ListItems(bucle).SubItems(12) <> CargaCadena(1224) Then
            nLineas = nLineas + 8 + Detalle.ListItems.Count
        Else
            nLineas = nLineas + 8
        End If
        ''''''''
        'nLineas = nLineas + 8 + Detalle.ListItems.Count

        limpreso = False
        If Val(Lista.ListItems(bucle).SubItems(18)) <> 0 Then
            limpreso = True
            nLineas = nLineas + 1
            nPar = nPar + 1
        End If
        If Val(Lista.ListItems(bucle).SubItems(21)) <> 0 Then
            limpreso = True
            nLineas = nLineas + 1
            nPar = nPar + 1
        End If
        If Val(Lista.ListItems(bucle).SubItems(24)) <> 0 Then
            limpreso = True
            nLineas = nLineas + 1
            nPar = nPar + 1
        End If
        If Val(Lista.ListItems(bucle).SubItems(27)) <> 0 Then
            limpreso = True
            nLineas = nLineas + 1
            nPar = nPar + 1
        End If
        If Val(Lista.ListItems(bucle).SubItems(30)) <> 0 Then
            limpreso = True
            nLineas = nLineas + 1
            nPar = nPar + 1
        End If
        If limpreso Then nLineas = nLineas + 1

        If nLineas > 85 Then
            Printer.NewPage
            NmYpAG = NmYpAG + 1
            Call ImpCab
            nLineas = 2 + 8 + Detalle.ListItems.Count + nPar
            If limpreso Then nLineas = nLineas + 1
        End If

        limpreso = False

        Printer.FontSize = 8
        'Printer.Print ""
        Printer.FontBold = True
        '2.0.5
        If Lista.ListItems(bucle).SubItems(12) <> CargaCadena(1224) Then
            Printer.Print " ---------------------------------------------------------------------------------------" & " N.Item:" & CStr(nItem)
            sParte = CargaCadena(132) & " : " & Lista.ListItems(bucle).SubItems(8)
        Else
            Printer.Print " ----------------------------" & CargaCadena(1224) & "  --------------------------------" & " N.Item:" & CStr(nItem)
            sParte = ""
        End If
        ''''''''''''''
        'Printer.Print " ---------------------------------------------------------------------------------------" & " N.Item:" & CStr(nItem)

        Printer.Print " " & CargaCadena(1427) & " : " & Lista.ListItems(bucle).SubItems(17)
        Printer.Print " " & CargaCadena(790) & ": " & Format(Val(Lista.ListItems(bucle).SubItems(9)), "000000") & _
                      CargaCadena(1425) & Mid(Lista.ListItems(bucle).SubItems(33), 1, 30) & CargaCadena(1426) & _
                      Mid(Lista.ListItems(bucle).SubItems(34), 1, 20)
        Printer.Print " " & CargaCadena(868) & " : " & Lista.ListItems(bucle).SubItems(2) & "#" & _
                      CargaCadena(434) & " : " & Lista.ListItems(bucle).TexT & " " & Lista.ListItems(bucle).SubItems(1) & "#" & _
                      CargaCadena(55) & " : " & Lista.ListItems(bucle).SubItems(11) & "#" & _
                      CargaCadena(1221) & " : " & Lista.ListItems(bucle).SubItems(12) & "#" & _
                      "EAN13: " & Lista.ListItems(bucle).SubItems(10) & "#" & _
                      sParte
        '2.0.5 CargaCadena(132) & " : " & Lista.ListItems(Bucle).SubItems(8)
        Printer.FontBold = False
        Printer.Print " " & CargaCadena(302) & " : " & Lista.ListItems(bucle).SubItems(3) & "#" & _
                      CargaCadena(147) & " : " & Lista.ListItems(bucle).SubItems(4) & "#" & _
                      CargaCadena(136) & " : " & Lista.ListItems(bucle).SubItems(5) & "#"    '& _
                                                                                             'CargaCadena(790) & " : " & Lista.ListItems(Bucle).SubItems(9)
        '2.0.5
        If Lista.ListItems(bucle).SubItems(12) <> CargaCadena(1224) Then
            Printer.Print " " & CargaCadena(1219) & " : " & Lista.ListItems(bucle).SubItems(6) & "#" & _
                          CargaCadena(290) & " : " & Lista.ListItems(bucle).SubItems(7) & "#" & _
                          CargaCadena(1220) & " : " & Lista.ListItems(bucle).SubItems(13) & "#" & _
                          CargaCadena(1222) & " : " & Lista.ListItems(bucle).SubItems(14) & "#" & _
                          CargaCadena(1223) & " : " & Lista.ListItems(bucle).SubItems(15)
        Else
            Printer.Print " "
        End If
        ''''''''
        Printer.Print " " & CargaCadena(1276) & " : " & Lista.ListItems(bucle).SubItems(16)

        '2.0.5
        If Lista.ListItems(bucle).SubItems(12) <> CargaCadena(1224) Then
            Printer.FontBold = True
            'Printer.Print "Nº  Cancel. Cód.   PLU  Texto                     Precio  Peso     Unid.    Total    T.IVA"
            Printer.Print " Nº" & "  " & _
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
                On Error Resume Next
                Printer.Print " " & Detalle.ListItems(Bucle2).TexT & Space(4 - Len(Detalle.ListItems(Bucle2).TexT)) & _
                              Detalle.ListItems(Bucle2).SubItems(8) & Space(8) & _
                              Detalle.ListItems(Bucle2).SubItems(1) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(1))) & _
                              Detalle.ListItems(Bucle2).SubItems(2) & Space(5 - Len(Detalle.ListItems(Bucle2).SubItems(2))) & _
                              Mid(Detalle.ListItems(Bucle2).SubItems(3), 1, 25) & Space(26 - Len(Mid(Detalle.ListItems(Bucle2).SubItems(3), 1, 25))) & _
                              Detalle.ListItems(Bucle2).SubItems(4) & Space(8 - Len(Detalle.ListItems(Bucle2).SubItems(4))) & _
                              Detalle.ListItems(Bucle2).SubItems(5) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(5))) & _
                              Detalle.ListItems(Bucle2).SubItems(6) & Space(12 - Len(Detalle.ListItems(Bucle2).SubItems(6))) & _
                              Detalle.ListItems(Bucle2).SubItems(7) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(7))) & _
                              Detalle.ListItems(Bucle2).SubItems(9) & _
                            " (" & Format(Detalle.ListItems(Bucle2).SubItems(10), "0.00") & " %)"
                If Err.Number <> 0 Then
                    limpreso = limpreso
                    On Error GoTo 0
                End If
            Next Bucle2
            Printer.FontItalic = False
            Printer.FontBold = True
            If Val(Lista.ListItems(bucle).SubItems(18)) <> 0 Then
                If limpreso = False Then Printer.Print CargaCadena(1433)
                limpreso = True
                Printer.Print " " & Format(Lista.ListItems(bucle).SubItems(18), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(19), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(20), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(21)) <> 0 Then
                If limpreso = False Then Printer.Print CargaCadena(1433)
                limpreso = True
                Printer.Print " " & Format(Lista.ListItems(bucle).SubItems(21), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(22), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(23), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(24)) <> 0 Then
                If limpreso = False Then Printer.Print CargaCadena(1433)
                limpreso = True
                Printer.Print " " & Format(Lista.ListItems(bucle).SubItems(24), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(25), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(26), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(27)) <> 0 Then
                If limpreso = False Then Printer.Print CargaCadena(1433)
                limpreso = True
                Printer.Print " " & Format(Lista.ListItems(bucle).SubItems(27), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(28), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(29), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(30)) <> 0 Then
                If limpreso = False Then Printer.Print CargaCadena(1433)
                limpreso = True
                Printer.Print " " & Format(Lista.ListItems(bucle).SubItems(30), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(31), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(32), "                0.00")
            End If
        End If
        ''''''''''''
        nItem = nItem + 1

        Printer.FontBold = False

    Next bucle
    If nLineas > 82 Then
        Printer.NewPage
        NmYpAG = NmYpAG + 1
        Call ImpCab
    End If
    Printer.Print " -----------------------------------------------------------------------------------------------------------"
    Printer.Print CargaCadena(1434) & CStr(NmYpAG)
    Printer.Print " -----------------------------------------------------------------------------------------------------------"

    Printer.EndDoc

End Sub


Private Sub GeneraFichero()
    Dim bucle As Integer
    Dim Bucle2 As Integer
    Dim CuentaLineas As Long
    Dim limpreso As Boolean
    Dim nf As Integer
    nf = FreeFile()
    Open App.Path & "\" & Format(TiendaActual, "000") & Format(Now(), "ddmmyyhhmmss") & ".txt" For Output As #nf
    Print #nf, " " & Empresa
    limpreso = False
    Print #nf, " "
    Print #nf, " " & CargaCadena(116) & ":" & Format(MyDateMin, "dd/mm/yyyy") & " " & CargaCadena(116) & ":" & Format(MyDateMax, "dd/mm/yyyy")
    Print #nf, " "
    Print #nf, CargaCadena(1439)
    Print #nf, " "
    If nBas0 <> 0 Then
        If limpreso = False Then Call MyFichCabIva(nf)
        limpreso = True
        Print #nf, "                     " & Format(nPor0, "00.00") & FormatMy(nBas0, "        0.00") & FormatMy(nImp0, "                0.00")
    End If
    If nBas1 <> 0 Then
        If limpreso = False Then Call MyFichCabIva(nf)
        limpreso = True
        Print #nf, "                     " & Format(nPor1, "00.00") & FormatMy(nBas1, "        0.00") & FormatMy(nImp1, "                0.00")
    End If
    If nBas2 <> 0 Then
        If limpreso = False Then Call MyFichCabIva(nf)
        limpreso = True
        Print #nf, "                     " & Format(nPor2, "00.00") & FormatMy(nBas2, "        0.00") & FormatMy(nImp2, "                0.00")
    End If
    If nBas3 <> 0 Then
        If limpreso = False Then Call MyFichCabIva(nf)
        limpreso = True
        Print #nf, "                     " & Format(nPor3, "00.00") & FormatMy(nBas3, "        0.00") & FormatMy(nImp3, "                0.00")
    End If
    If nBas4 <> 0 Then
        If limpreso = False Then Call MyFichCabIva(nf)
        limpreso = True
        Print #nf, "                     " & Format(nPor4, "00.00") & FormatMy(nBas4, "        0.00") & FormatMy(nImp4, "                0.00")
    End If
    Print #nf, "                     _____________________________________"

    Print #nf, "              Totales:" & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4, "            " & "0.00") & FormatMy(nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "               0.00")
    Print #nf, " TOTAL Base+Impuestos:" & FormatMy(nBas0 + nBas1 + nBas2 + nBas3 + nBas4 + nImp0 + nImp1 + nImp2 + nImp3 + nImp4, "            0.00")
    Print #nf, " "

    Print #nf, CargaCadena(1440)    ' & "Fecha Inicial:" & Format(MyDateMin, "dd/mm/yyyy") & " Fecha Final:" & Format(MyDateMax, "dd/mm/yyyy")
    limpreso = False
    For bucle = 1 To Lista.ListItems.Count
        limpreso = False
        Print #nf, " -----------------------------------------------------------------------------------------------------------"
        Print #nf, " " & CargaCadena(1427) & " : " & Lista.ListItems(bucle).SubItems(17)
        Print #nf, " " & CargaCadena(790) & ": " & Format(Val(Lista.ListItems(bucle).SubItems(9)), "000000") & _
                   CargaCadena(1425) & Mid(Lista.ListItems(bucle).SubItems(33), 1, 30) & CargaCadena(1426) & _
                   Mid(Lista.ListItems(bucle).SubItems(34), 1, 20)
        Print #nf, " " & CargaCadena(868) & " : " & Lista.ListItems(bucle).SubItems(2) & "#" & _
                   CargaCadena(434) & " : " & Lista.ListItems(bucle).TexT & " " & Lista.ListItems(bucle).SubItems(1) & "#" & _
                   CargaCadena(55) & " : " & Lista.ListItems(bucle).SubItems(11) & "#" & _
                   CargaCadena(1221) & " : " & Lista.ListItems(bucle).SubItems(12) & "#" & _
                   "EAN13: " & Lista.ListItems(bucle).SubItems(10) & "#" & _
                   CargaCadena(132) & " : " & Lista.ListItems(bucle).SubItems(8)
        Print #nf, " " & CargaCadena(302) & " : " & Lista.ListItems(bucle).SubItems(3) & "#" & _
                   CargaCadena(147) & " : " & Lista.ListItems(bucle).SubItems(4) & "#" & _
                   CargaCadena(136) & " : " & Lista.ListItems(bucle).SubItems(5) & "#"    '& _
                                                                                          'CargaCadena(790) & " : " & Lista.ListItems(Bucle).SubItems(9)
        Print #nf, " " & CargaCadena(1219) & " : " & Lista.ListItems(bucle).SubItems(6) & "#" & _
                   CargaCadena(290) & " : " & Lista.ListItems(bucle).SubItems(7) & "#" & _
                   CargaCadena(1220) & " : " & Lista.ListItems(bucle).SubItems(13) & "#" & _
                   CargaCadena(1222) & " : " & Lista.ListItems(bucle).SubItems(14) & "#" & _
                   CargaCadena(1223) & " : " & Lista.ListItems(bucle).SubItems(15)
        Print #nf, " " & CargaCadena(1276) & " : " & Lista.ListItems(bucle).SubItems(16)

        '2.0.5
        If Lista.ListItems(bucle).SubItems(12) <> CargaCadena(1224) Then

            Lista.ListItems(bucle).Selected = True
            Lista_Click
            Print #nf, " Nº" & "  " & _
                       CargaCadena(1224) & " " & _
                       CargaCadena(69) & "   " & _
                       CargaCadena(68) & "  " & _
                       CargaCadena(539) & "                     " & _
                       CargaCadena(71) & "  " & _
                       CargaCadena(130) & "     " & _
                       CargaCadena(131) & "    " & _
                       CargaCadena(488) & "    " & _
                       CargaCadena(553)
            For Bucle2 = 1 To Detalle.ListItems.Count
                Print #nf, " " & Detalle.ListItems(Bucle2).TexT & Space(4 - Len(Detalle.ListItems(Bucle2).TexT)) & _
                           Detalle.ListItems(Bucle2).SubItems(8) & Space(8) & _
                           Detalle.ListItems(Bucle2).SubItems(1) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(1))) & _
                           Detalle.ListItems(Bucle2).SubItems(2) & Space(5 - Len(Detalle.ListItems(Bucle2).SubItems(2))) & _
                           Mid(Detalle.ListItems(Bucle2).SubItems(3), 1, 25) & Space(26 - Len(Mid(Detalle.ListItems(Bucle2).SubItems(3), 1, 25))) & _
                           Detalle.ListItems(Bucle2).SubItems(4) & Space(8 - Len(Detalle.ListItems(Bucle2).SubItems(4))) & _
                           Detalle.ListItems(Bucle2).SubItems(5) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(5))) & _
                           Detalle.ListItems(Bucle2).SubItems(6) & Space(12 - Len(Detalle.ListItems(Bucle2).SubItems(6))) & _
                           Detalle.ListItems(Bucle2).SubItems(7) & Space(9 - Len(Detalle.ListItems(Bucle2).SubItems(7))) & _
                           Detalle.ListItems(Bucle2).SubItems(9) & _
                         " (" & Format(Detalle.ListItems(Bucle2).SubItems(10), "0.00") & " %)"
            Next Bucle2
            If Val(Lista.ListItems(bucle).SubItems(18)) <> 0 Then
                If limpreso = False Then Print #nf, CargaCadena(1433)    '" % IVA        BASE            Impuesto"
                limpreso = True
                Print #nf, " " & Format(Lista.ListItems(bucle).SubItems(18), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(19), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(20), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(21)) <> 0 Then
                If limpreso = False Then Print #nf, CargaCadena(1433)
                limpreso = True
                Print #nf, " " & Format(Lista.ListItems(bucle).SubItems(21), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(22), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(23), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(24)) <> 0 Then
                If limpreso = False Then Print #nf, CargaCadena(1433)
                limpreso = True
                Print #nf, " " & Format(Lista.ListItems(bucle).SubItems(24), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(25), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(26), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(27)) <> 0 Then
                If limpreso = False Then Print #nf, CargaCadena(1433)
                limpreso = True
                Print #nf, " " & Format(Lista.ListItems(bucle).SubItems(27), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(28), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(29), "                0.00")
            End If
            If Val(Lista.ListItems(bucle).SubItems(30)) <> 0 Then
                If limpreso = False Then Print #nf, CargaCadena(1433)
                limpreso = True
                Print #nf, " " & Format(Lista.ListItems(bucle).SubItems(30), "00.00") & FormatMy(Lista.ListItems(bucle).SubItems(31), "        0.00") & FormatMy(Lista.ListItems(bucle).SubItems(32), "                0.00")
            End If

        End If
        '''''''''

    Next bucle

    Close #nf
End Sub

Private Sub CmdExport_Click()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Dim sZip As String
    Dim nResp As Integer
    Dim siP As String
    Dim susR As String
    Dim spsS As String
    Dim Ret, nFich
    Dim lBF As Boolean
    Dim nf As Integer

    On Error GoTo siE1


    FrmConsultaTiquets.Enabled = False
    Screen.MousePointer = vbHourglass

    Contador = 0
    If Dir(App.Path & "\exp_hist.ini") <> "" And IsDate(CmbFiltro(2).TexT) And IsDate(CmbFiltro(4).TexT) Then
        Archivo = FreeFile()
        Open Miruta & "\exp_hist.ini" For Input As Archivo
        Do Until EOF(Archivo)
            Contador = Contador + 1
            ReDim Preserve Buffer(Contador + 1)
            Line Input #Archivo, Buffer(Contador)
        Loop
        Close #Archivo
        Load FrmExportar
        'FrmExportar.Hide
        If Contador > 0 Then
            If Dir(App.Path & "\ht.txt") = "" Then
                lBF = True
                nf = FreeFile()
                Open App.Path & "\ht.txt" For Output As #nf
                Print #nf, Base_General
                Close #nf
            End If
        End If
        For bucle = 1 To Contador
            Buffer(bucle) = Trim(Buffer(bucle))
            If Buffer(bucle) <> "" Then
                FrmExportar.MostrarMensajes = False
                FrmExportar.ExportarFicheroHIS Buffer(bucle), CDate(CmbFiltro(2).TexT), CDate(CmbFiltro(4).TexT)
                '2.0.2 frmMonitor.MostrarDato CargaCadena(984) & " " & Buffer(Bucle)
                If Dir(App.Path & "\getputfileftp.dat") <> "" And Dir(Exp_Path & "\" & Buffer(bucle)) <> "" Then
                    nFich = FreeFile()
                    Open App.Path & "\getputfileftp.dat" For Input As #nFich
                    'Do While Not EOF(nFich)
                    Line Input #nFich, siP
                    Line Input #nFich, susR
                    Line Input #nFich, spsS
                    'Loop
                    Close #nFich

                    nResp = ElPing(siP)

                    If nResp = 1 Or Dir(App.Path & "\noping.txt") <> "" Then

                        If nRetConnFTP = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0)    '0=no ha podido conectar, 1=se ha conectado
                            If Ret = 1 Then nRetConnFTP = 1
                        Else
                            Ret = 1
                        End If
                        If Ret = 1 Then    'ha sido posible conectar
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else

                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If

                                Ret = Zipear(Miruta, Buffer(bucle))


                            End If
                        End If
                        If Ret = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0)    '0=no ha podido conectar, 1=se ha conectado
                            'Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If

                                Ret = Zipear(Miruta, Buffer(bucle))

                            End If

                        End If

                    End If

                    If Ret = 0 Or nResp = 0 Then
                        nRetConnFTP = 0
                    End If

                End If

            End If
        Next bucle

        If Dir(App.Path & "\ht.txt") <> "" And lBF Then
            Kill App.Path & "\ht.txt"
        End If

        Set FrmExportar = Nothing

        MsgBox "Ficheros exportados", vbInformation

    Else
        MsgBox "No existe definición ficheros a exportar o Fechas no válidas...", vbExclamation
    End If
siE1:
    On Error GoTo 0
    FrmConsultaTiquets.Enabled = True
    Screen.MousePointer = vbDefault

End Sub


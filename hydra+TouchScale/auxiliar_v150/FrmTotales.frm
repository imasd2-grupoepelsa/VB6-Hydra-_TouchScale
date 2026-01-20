VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTotales 
   Caption         =   "Consulta de Totales"
   ClientHeight    =   7650
   ClientLeft      =   4050
   ClientTop       =   1185
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   9345
   Begin VB.CheckBox ChkLocal 
      Caption         =   "Al Imprimir Totalizar Red y Local Todos los Tipos de Venta."
      Height          =   255
      Left            =   1920
      TabIndex        =   44
      Top             =   7320
      Width           =   3735
   End
   Begin VB.CheckBox ChkTotPer 
      Caption         =   "Period. Tot."
      Height          =   195
      Left            =   240
      TabIndex        =   43
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "N.G.T."
      Height          =   255
      Left            =   5640
      TabIndex        =   42
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton OptFiltro 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   40
      Top             =   6960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.OptionButton OptFiltro 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   39
      Top             =   6600
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.OptionButton OptFiltro 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   6360
   End
   Begin VB.PictureBox Cortina 
      BackColor       =   &H00800000&
      Height          =   5775
      Left            =   1920
      ScaleHeight     =   5715
      ScaleWidth      =   7275
      TabIndex        =   25
      Top             =   360
      Width           =   7332
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Por Favor, espere..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   492
         Left            =   2040
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   3132
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   24
      Top             =   6960
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1920
      TabIndex        =   23
      Top             =   6960
      Width           =   1572
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   288
      ItemData        =   "FrmTotales.frx":0000
      Left            =   3600
      List            =   "FrmTotales.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   0
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   288
      Index           =   9
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5760
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   288
      Index           =   8
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5160
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   288
      Index           =   7
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4560
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   288
      Index           =   6
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3960
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   288
      Index           =   5
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3360
      Width           =   1692
   End
   Begin VB.ComboBox CmbFiltro 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   3
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmTotales.frx":0004
      Left            =   120
      List            =   "FrmTotales.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5772
      Left            =   1920
      TabIndex        =   22
      Top             =   360
      Width           =   7332
      _ExtentX        =   12938
      _ExtentY        =   10186
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "(Incluye totales no consolidados en la base)"
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
      Left            =   3720
      TabIndex        =   41
      Top             =   6960
      Visible         =   0   'False
      Width           =   3852
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2040
      TabIndex        =   37
      Top             =   6240
      Width           =   1212
   End
   Begin VB.Label LblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   35
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   2
      Left            =   6000
      TabIndex        =   34
      Top             =   6240
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label LblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblTot 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblTot 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblTot 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblTot 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblTot 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Totales por :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1920
      TabIndex        =   20
      Top             =   0
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label LblFiltro 
      Caption         =   "Label1"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   732
      Left            =   1920
      Top             =   6120
      Width           =   7332
   End
End
Attribute VB_Name = "FrmTotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TFiltro
    Tipo As Long
    Campo As String
    Indice As Integer
End Type
Private NFiltro0 As Integer

Private NFiltros As Integer
Private Filtros(10) As TFiltro
Private Ordenacion As Integer
Private Iniciando As Boolean
Private Me_Width As Long
Private Me_Height As Long
Private ct_top() As Long
Private CT_Left() As Long
Private CT_Width() As Long
Private CT_Height() As Long
Private Indices() As Integer
Private NIndices As Integer
Private QLTOT As Boolean
Private Function Redondea(ByVal dato As Double) As Double
    '************************************
    ' redondea la diferencia que
    ' se produce al tratar datos
    ' con decimales, con Access y VB
    '************************************
    If UsaEuro Then
        dato = Round(dato, 2)
    Else
        dato = Round(dato, decimales)
    End If
    Redondea = dato
    
End Function
Public Function importar_totVen(ByVal argRutaf As String) As Long
    '*****************************************
    ' GAMA BAJA : debido a la operativa de las
    ' máquinas y los usuarios, los totales
    ' recibidos se dejan en un fichero temporal
    ' y sólo se importan a la base de datos
    ' cuando se hace el fin de día (borrado de
    ' máquinas)
    '******************************************
    '**************
    Dim strLinea As String
    '*********************
   
    Dim conxf As Integer
    Dim rstGTVEN As New ADODB.Recordset
    Dim Registro As New ADODB.Recordset
    '************************
    Dim myFecha As Date
    Dim MyMostrador As Integer
    Dim MySeccMaqui As Integer
    Dim MyNombre As String
    Dim MyModo As Integer
    Dim MyVendedor As Integer
    Dim MyImporNeg As Double
    Dim MyImporte As Double
    Dim MyCancelado As Double
    Dim MyOp As Double
    Dim MyPes As Double
    Dim MyTiq As Double
    Dim MyOpCan As Double
    '************************
    On Error GoTo ret
    If Dir(argRutaf) = "" Then Exit Function

    
    conxf = FreeFile()
    Open argRutaf For Input Access Read Lock Write As #conxf
    Do While Not EOF(conxf)
        Line Input #conxf, strLinea
        '**************************************
        ' Obtiene información de cuadre
        '**************************************
        myFecha = DateSerial(Val(Mid(strLinea, 5, 2)), Val(Mid(strLinea, 3, 2)), Val(Mid(strLinea, 1, 2)))
        If Val(Mid(strLinea, 61, 1)) = 0 Then   '************ 1 -> venta
            MyModo = 1         '************ 4 -> envasado
        ElseIf Val(Mid(strLinea, 61, 1)) = 1 Then
            MyModo = 3
        End If
        MySeccMaqui = Val(Mid(strLinea, 7, 2))
        
        Set Registro = AbreRecordset _
        ("select codi_ident,secc_maqui from equipos where numero_eqp=" & MySeccMaqui & " and borrado=false")
        With Registro
            If Not .EOF Then
                MySeccMaqui = .Fields("secc_maqui")
                MyMostrador = .Fields("codi_ident")
            End If
        End With
        MyVendedor = Val(Mid(strLinea, 9, 2))
        '**********************************
        ' más datos
        If Val(Mid(strLinea, 11, 1)) = 1 Then
            MyImporte = Val(Mid(strLinea, 12, 7))
        Else
            MyImporte = (-1) * Val(Mid(strLinea, 12, 7))
        End If
        MyImporNeg = Abs(Val(Mid(strLinea, 19, 7)))
        MyCancelado = Val(Mid(strLinea, 26, 7))
        MyOp = Val(Mid(strLinea, 40, 7))
        MyOpCan = Val(Mid(strLinea, 47, 7))
        MyTiq = Val(Mid(strLinea, 54, 7))
        If UsaEuro Then
            MyImporNeg = Redondea(MyImporNeg / 100)
            MyImporte = Redondea(MyImporte / 100)
            MyCancelado = Redondea(MyImporte / 100)
        Else
            MyImporNeg = Redondea(MyImporNeg / (10 ^ decimales))
            MyImporte = Redondea(MyImporte / (10 ^ decimales))
            MyCancelado = Redondea(MyImporte / (10 ^ decimales))
        End If
        '**********************************
        ' Busca el registro adecuado
        '**********************************
        If MyOp <> 0 Or MyOpCan <> 0 Then
            Set rstGTVEN = New ADODB.Recordset
            rstGTVEN.Open _
            "select * from totales where ident_vend=" & MyVendedor & _
            " and codi_ident=" & MyMostrador & " and modo=" & MyModo & _
            " and fecha=#" & Format(myFecha, "mm/dd/yy") & "#", _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\dirtmptouch\tmp.mdb ;Persist Security Info=False", _
            adOpenDynamic, adLockOptimistic
            If rstGTVEN.EOF Then
                rstGTVEN.AddNew
                rstGTVEN.Fields("modo") = MyModo
                rstGTVEN.Fields("fecha") = myFecha '
                rstGTVEN.Fields("ident_vend") = MyVendedor '
                'rstGTVEN.Fields("secc_maqui") = MySeccMaqui '
                rstGTVEN.Fields("codi_ident") = MyMostrador
                Set Registro = AbreRecordset _
                ("select nombre from vendedor where borrado=false and ident_vend=" & MyVendedor)
                If Not Registro.EOF Then
                    rstGTVEN.Fields("Nombre") = Registro.Fields("Nombre")
                Else
                    rstGTVEN.Fields("Nombre") = " "
                End If
                rstGTVEN.Fields("peso") = 0
                rstGTVEN.Fields("pesetas") = 0
                'rstGTVEN.Fields("totcance") = 0
                'rstGTVEN.Fields("total") = 0
                'rstGTVEN.Fields("supertot") = 0
                rstGTVEN.Fields("tiquets") = 0
                rstGTVEN.Fields("operacion") = 0
                rstGTVEN.Fields("cancela") = 0
                rstGTVEN.Fields("unidad") = 0
                rstGTVEN.Fields("impor_neg") = 0
                rstGTVEN.Update
                
                'Set rstGTVEN = Base.OpenRecordset _
                '("select * from gtvend where ident_vend=" & MyVendedor & _
                '" and codi_ident=" & MyMostrador & " and modo=" & MyModo & _
                '" and fecha=#" & Format(myFecha, "mm/dd/yy") & "#")
            End If
                '
            On Error GoTo 0
            rstGTVEN.Fields("pesetas") = Redondea(rstGTVEN.Fields("pesetas") + MyImporte)
            'rstGTVEN.Fields("totcance") = (rstGTVEN.Fields("totcance") + MyCancelado)
            'rstGTVEN.Fields("total") = Redondea(rstGTVEN.Fields("pesetas") + MyImporte)
            'rstGTVEN.Fields("supertot") = Redondea(rstGTVEN.Fields("pesetas") + Abs(MyImporte) + MyCancelado)
            rstGTVEN.Fields("tiquets") = rstGTVEN.Fields("tiquets") + MyTiq
            rstGTVEN.Fields("operacion") = Round(rstGTVEN.Fields("operacion") + MyOp, 0)
            If Not IsNull(rstGTVEN.Fields("impor_neg")) Then
                rstGTVEN.Fields("impor_neg") = Redondea(rstGTVEN.Fields("impor_neg") + MyImporNeg)
            Else
                rstGTVEN.Fields("impor_neg") = MyImporNeg
            End If
            rstGTVEN.Fields("cancela") = rstGTVEN.Fields("cancela") + MyOpCan
            rstGTVEN.Update
            rstGTVEN.Close
            Set rstGTVEN = Nothing
            Label4.Visible = True
        End If
    Loop
    Close #conxf
    
    
    
    
    importar_totVen = 0
    On Error GoTo 0
    Exit Function
    '************
ret:
    importar_totVen = 1
    On Error GoTo 0
End Function
'***********
Public Function Importar_TotItem(ByVal argRutaf As String) As Long
    '*****************************************
    ' GAMA BAJA : debido a la operativa de las
    ' máquinas y los usuarios, los totales
    ' recibidos se dejan en un fichero temporal
    ' y sólo se importan a la base de datos
    ' cuando se hace el fin de día (borrado de
    ' máquinas)
    '******************************************
    Dim res As Long
    Dim strSQL As String
    Dim strLinea As String
    '*********************
    Dim conxf As Integer
    Dim rstARTI As ADODB.Recordset
    Dim rstGTARTI As New ADODB.Recordset
    Dim RegBal As ADODB.Recordset
    '**************************
    Dim MyMostrador As Integer
    Dim MyCodigo As Long
    Dim MyBalanza As Integer
    Dim MyDescriptivo As String
    Dim MySubSeccion As Integer
    Dim MyFamilia As Integer
    Dim MyPrecio As Double
    Dim MyTara As Long
    Dim MyCaducidad As Integer
    Dim MyPlu As Integer
    Dim myFecha As Date
    Dim MyTipoIVA As Integer
    Dim MyPositivo As Integer
    Dim MyModo As Integer
        On Error GoTo ret
        If Dir(argRutaf) <> "" Then
            conxf = FreeFile
            Open argRutaf For Input Access Read Lock Write As #conxf
            Do While Not EOF(conxf)
                Line Input #conxf, strLinea
                MyBalanza = Val(Mid(strLinea, 7, 2))
                myFecha = DateSerial(Val(Mid(strLinea, 5, 2)), Val(Mid(strLinea, 3, 2)), Val(Left(strLinea, 2)))
                MyPlu = Val(Mid(strLinea, 9, 4))
                MyModo = Mid(strLinea, 44, 1)
                MyPositivo = Val(Mid(strLinea, 22, 1))
                If MyPositivo <> 1 Then
                    MyPositivo = -1
                End If
                If MyModo = 0 Then
                    MyModo = 1
                Else
                    MyModo = 3
                End If
                MyMostrador = 0
                MyCodigo = 0
                MyDescriptivo = ""
                MySubSeccion = 0
                MyFamilia = 0
                MyCodigo = 0
                MyPrecio = 0
                MyTara = 0
                MyCaducidad = 0
                MyTipoIVA = 0
                '///////////////////////////////
                ' obtiene datos del mostrador
                Set RegBal = AbreRecordset("select codi_ident,secc_maqui from equipos where borrado=false and numero_eqp=" & MyBalanza)
                With RegBal
                    If Not .EOF Then
                        MyBalanza = .Fields("secc_maqui")
                        MyMostrador = .Fields("codi_ident")
                    End If
                End With
                RegBal.Close
                Set RegBal = Nothing
                '////////////////////////////////
                ' obtiene datos del artículos
                strSQL = "SELECT * FROM Articulo WHERE codi_ident=" & MyMostrador & " AND PLU=" & MyPlu & " AND BORRADO=FALSE"
                Set rstARTI = AbreRecordset(strSQL)
                    If Not rstARTI.BOF Then
                        MyDescriptivo = rstARTI.Fields("DES_PLU1")
                        MySubSeccion = rstARTI.Fields("CODI_SUB")
                        MyFamilia = rstARTI.Fields("CODI_FAM")
                        MyCodigo = rstARTI.Fields("CODIGO")
                        MyPrecio = rstARTI.Fields("PRECIO")
                        MyTara = rstARTI.Fields("TARA")
                        MyCaducidad = rstARTI.Fields("CADUCIDAD")
                        MyTipoIVA = rstARTI.Fields("TIPO_IVA")
                    End If
                rstARTI.Close
                Set rstARTI = Nothing
                '*********************************************
                strSQL = "SELECT * FROM totales where codi_ident=" & MyMostrador & " and codigo=" & MyCodigo & _
                " and modo=" & MyModo & " and fecha=#" & Format(myFecha, "mm/DD/yy") & "#"
                Set rstGTARTI = New ADODB.Recordset
                rstGTARTI.Open strSQL, _
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                App.Path & "\dirtmptouch\tmp.mdb ;Persist Security Info=False", _
                adOpenDynamic, adLockOptimistic
                If rstGTARTI.EOF Then
                    rstGTARTI.AddNew
                    rstGTARTI.Fields("codigo") = MyCodigo
                    rstGTARTI.Fields("plu") = MyPlu
                    rstGTARTI.Fields("codi_ident") = MyMostrador
                    'rstGTARTI.Fields("secc_maqui") = MyBalanza
                    rstGTARTI.Fields("des_plu1") = MyDescriptivo
                    rstGTARTI.Fields("codi_sub") = MySubSeccion
                    rstGTARTI.Fields("codi_fam") = MyFamilia
                    'rstGTARTI.Fields("precio") = MyPrecio
                    'rstGTARTI.Fields("tara") = MyTara
                    'rstGTARTI.Fields("caducidad") = MyCaducidad
                    'rstGTARTI.Fields("tipo_iva") = MyTipoIVA
                    rstGTARTI.Fields("fecha") = myFecha
                    'rstGTARTI.Fields("modificado") = False
                    rstGTARTI.Fields("operacion") = 0
                    rstGTARTI.Fields("pesetas") = 0
                    rstGTARTI.Fields("peso") = 0
                    rstGTARTI.Fields("unidades") = 0
                    rstGTARTI.Fields("modo") = MyModo
                    rstGTARTI.Update
                End If
                On Error GoTo 0
                rstGTARTI.Fields("peso") = Round(rstGTARTI.Fields("peso") + (Val(Mid(strLinea, 13, 9)) / 1000), 3)
                rstGTARTI.Fields("unidades") = rstGTARTI.Fields("unidades") + Val(Mid(strLinea, 37, 7))
                rstGTARTI.Fields("operacion") = rstGTARTI.Fields("operacion") + Val(Mid(strLinea, 30, 7))
                If UsaEuro Then
                    rstGTARTI.Fields("pesetas") = Redondea(rstGTARTI.Fields("pesetas") + (MyPositivo * Val(Mid(strLinea, 23, 7)) / 100))
                Else
                    rstGTARTI.Fields("pesetas") = Redondea(rstGTARTI.Fields("pesetas") + (MyPositivo * Val(Mid(strLinea, 23, 7)) / (10 ^ decimales)))
                End If
                rstGTARTI.Update
                rstGTARTI.Close
                Set rstGTARTI = Nothing
                Label4.Visible = True
            Loop
            Close #conxf
        End If
        Importar_TotItem = 0
        On Error GoTo 0
    Exit Function
ret:
    Importar_TotItem = 1
    On Error GoTo 0
End Function
Public Function Importar_TotMaquina(ByVal Ruta As String)
    Dim Archivo As Integer
    Dim Buf As String
    Dim RegSec As ADODB.Recordset
    Dim Registro As New ADODB.Recordset
    Dim myFecha As Date
    Dim MySeccMaqui As Integer
    Dim MyMostrador As Integer
    Dim MyImporte As Double
    Dim MyOp As Double
    Dim MyUnd As Double
    Dim MyNTiquets As Double
    Dim MyTipo As Integer
    If Dir(Ruta) <> "" Then
        Archivo = FreeFile()
        Open Ruta For Input As #Archivo
        Do While Not EOF(Archivo)
            Line Input #Archivo, Buf
            myFecha = DateSerial(Val(Mid(Buf, 5, 2)), Val(Mid(Buf, 3, 2)), Val(Mid(Buf, 1, 2)))
            MySeccMaqui = Val(Mid(Buf, 7, 2))
            Set RegSec = AbreRecordset _
            ("select codi_ident,secc_maqui from equipos where borrado=false and numero_eqp=" & MySeccMaqui)
            With RegSec
                If Not .EOF Then
                    MySeccMaqui = .Fields("secc_maqui")
                    MyMostrador = .Fields("codi_ident")
                Else
                    MySeccMaqui = 0
                    MyMostrador = 0
                End If
            End With
            MyImporte = Val(Mid(Buf, 10, 7))
            If UsaEuro Then
                MyImporte = MyImporte / 100
            Else
                MyImporte = MyImporte / (10 ^ decimales)
            End If
            If Mid(Buf, 9, 1) = "0" Then
                MyImporte = (-1) * MyImporte
            End If
            MyOp = Val(Mid(Buf, 17, 7))
            MyUnd = Val(Mid(Buf, 24, 7))
            MyNTiquets = Val(Mid(Buf, 31, 7))
            MyTipo = Val(Mid(Buf, 38, 1))
            If MyTipo = 0 Then
                MyTipo = 1
            Else
                MyTipo = 3
            End If
            If MyOp <> 0 And MyUnd <> 0 Then
                Set Registro = New ADODB.Recordset
                Registro.Open _
                "select * from totales where fecha=#" & Format(myFecha, "mm/dd/yy") & _
                "# and codi_ident=" & MyMostrador & " and codi_sub=0 and codi_fam=0 and modo=" & _
                MyTipo, _
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                App.Path & "\dirtmptouch\tmp.mdb ;Persist Security Info=False", _
                adOpenDynamic, adLockOptimistic
                With Registro
                    If .EOF Then
                        .AddNew
                        .Fields("fecha") = myFecha
                        .Fields("codi_ident") = MyMostrador
                        .Fields("codi_sub") = 0
                        .Fields("codi_fam") = 0
                        If FamiliasFijas Then .Fields("txt_fam") = ""
                        If SubseccionesFijas Then .Fields("txt_sub") = ""
                        '.Fields("secc_maqui") = MySeccMaqui
                        .Fields("modo") = MyTipo
                        .Fields("tiquets") = 0
                        .Fields("operacion") = 0
                        .Fields("peso") = 0
                        .Fields("unidades") = 0
                        .Fields("pesetas") = 0
                        .Update
                    End If
                    .Fields("tiquets") = .Fields("tiquets") + MyNTiquets
                    .Fields("operacion") = .Fields("operacion") + MyOp
                    .Fields("unidades") = .Fields("unidades") + MyUnd
                    .Fields("pesetas") = Redondea(.Fields("pesetas") + MyImporte)
                    Label4.Visible = True
                    .Update
                    .Close
                    Set Registro = Nothing
                End With
            End If
        Loop
        Close #Archivo
    End If
End Function
Private Sub CambiarIdioma()
    If id <> 0 Then Me.Caption = CargaCadena(19)
    Command1.Caption = CargaCadena(466)
    Command2.Caption = CargaCadena(7)
    Label1.Caption = CargaCadena(162)
    CmbTipo.AddItem CargaCadena(9)
    CmbTipo.AddItem CargaCadena(12)
    CmbTipo.AddItem CargaCadena(13)
    CmbTipo.AddItem CargaCadena(9) & "-N.G.T."
    Label2.Caption = CargaCadena(1199)
    ChkLocal.Caption = CargaCadena(1404)
End Sub
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
        ct_top(Bucle) = Me.Controls(Bucle).Top
        CT_Left(Bucle) = Me.Controls(Bucle).Left
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        CT_Height(Bucle) = Me.Controls(Bucle).Height
        On Error GoTo 0
    Next Bucle
End Sub
Private Sub Me_Resize()
    Dim RelW As Double
    Dim Bucle As Integer

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
    For Bucle = 0 To Me.Controls.Count - 1
        On Error Resume Next
        RelW = Me.Width / Me_Width
        Me.Controls(Bucle).Width = RelW * CT_Width(Bucle)
        CT_Width(Bucle) = Me.Controls(Bucle).Width
        Me.Controls(Bucle).Left = RelW * CT_Left(Bucle)
        CT_Left(Bucle) = Me.Controls(Bucle).Left
        
        '
        RelW = Me.Height / Me_Height
        If TypeName(Me.Controls(Bucle)) <> "TextBox" _
        And TypeName(Me.Controls(Bucle)) <> "CommandButton" _
        And TypeName(Me.Controls(Bucle)) <> "MaskEdBox" _
        Then
            Me.Controls(Bucle).Height = RelW * CT_Height(Bucle)
            CT_Height(Bucle) = Me.Controls(Bucle).Height
        End If
        Me.Controls(Bucle).Top = RelW * ct_top(Bucle)
        ct_top(Bucle) = Me.Controls(Bucle).Top
        On Error GoTo 0
        '
    Next Bucle
    Me_Width = Me.Width
    Me_Height = Me.Height
    If CmbTipo.Text <> "" Then IzarCortina
End Sub

Private Sub Init_Filter()
Dim Bucle As Integer
    NFiltros = 0
    Iniciando = True
    For Bucle = 0 To 9
        CmbFiltro(Bucle).Visible = False
        CmbFiltro(Bucle).Clear
        If Bucle <> 0 Then
            CmbFiltro(Bucle).AddItem "(" & CargaCadena(311) & ")" ' (TODO)
            CmbFiltro(Bucle).Text = CmbFiltro(Bucle).List(0)
        End If
        LblFiltro(Bucle).Visible = False
    Next Bucle
    Iniciando = False
End Sub

Private Sub Add_Filter(sTexto As String, sCampo As String, vbTipo As Long)
    Dim MyReg As New ADODB.Recordset
    Dim Bucle As Long
    Dim BufInicio As Boolean
    Dim Encontrado As Boolean
    Dim WhereSec As String
    If HaySeguridad And UsuarioActual.Asociado <> 0 Then WhereSec = "where codi_ident = " & UsuarioActual.Asociado
    If Not ((Not HaySeguridad) Or (HaySeguridad And UsuarioActual.Asociado = 0) Or (NFiltros <> 3)) Then
        QLTOT = True
        CmbFiltro(3).Clear
        CmbFiltro(3).AddItem UsuarioActual.Asociado
        CmbFiltro(3).ListIndex = 0
    End If
    'If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
    'MyReg.Open "select *" & " from totales " & WhereSec & " order by " & sCampo, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
    'Else
    MyReg.Open "select distinct " & sCampo & " from totales " & WhereSec & " order by " & sCampo, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
    'End If
    NFiltro0 = 0
    Filtros(NFiltros).Campo = sCampo
    Filtros(NFiltros).Tipo = vbTipo
    LblFiltro(NFiltros).Caption = sTexto
    With MyReg
        If (Not (.EOF And .BOF)) Then
            .MoveFirst
            Do Until .EOF
                If NFiltros = 0 Then
                    '1.0.0 (0=Factura pasa a 7 red, 8 local
                    '       1=tique pasa a 1 red, 4 local
                    '       2=Envasado pasa a 2 red, 5 local
                    '       3=AutoServicio pasa a 3 red, 6 local
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
                    
                    Select Case .Fields(sCampo)
                        'Case 0
                        '    CmbFiltro(0).AddItem CargaCadena(311) 'todos
                        Case 1
                            CmbFiltro(0).AddItem CargaCadena(336) '"Venta"
                        Case 2
                            CmbFiltro(0).AddItem CargaCadena(195) '"Envasado"
                        Case 3
                            CmbFiltro(0).AddItem CargaCadena(196) '"Autoservicio"
                        Case 4
                            CmbFiltro(0).AddItem CargaCadena(1213) '"Venta Local"
                        Case 5 'Venta local
                            CmbFiltro(0).AddItem CargaCadena(199) '"Envasado Local"
                        Case 6
                            CmbFiltro(0).AddItem CargaCadena(200) '"Autoservicio Local"
                        Case 7
                            CmbFiltro(0).AddItem "Fact. Net"
                        Case 8
                            CmbFiltro(0).AddItem "Fact. Local"
                        Case 12
                            CmbFiltro(0).AddItem "Descart."
                    End Select
                    Filtro0(NFiltro0) = .Fields(sCampo)
                    NFiltro0 = NFiltro0 + 1
                Else
                    Select Case vbTipo
                        Case 1 To 2
                            Encontrado = False
                            'For Bucle = 0 To CmbFiltro(NFiltros).ListCount - 1
                            '    If CmbFiltro(NFiltros).List(Bucle) = Format(.Fields(sCampo), "dd/mm/yy") Then Encontrado = True
                            'Next Bucle
                            If Not Encontrado Then CmbFiltro(NFiltros).AddItem Format(.Fields(sCampo), "dd/mm/yy")
                        Case 3
                            Encontrado = False
                            'For Bucle = 0 To CmbFiltro(NFiltros).ListCount - 1
                            '    If CmbFiltro(NFiltros).List(Bucle) = .Fields(sCampo) Then Encontrado = True
                            'Next Bucle
                            If Not Encontrado Then CmbFiltro(NFiltros).AddItem .Fields(sCampo)
                        Case Else
                            Encontrado = False
                            'For Bucle = 0 To CmbFiltro(NFiltros).ListCount - 1
                            '    If Val(CmbFiltro(NFiltros).List(Bucle)) = .Fields(sCampo) Then Encontrado = True
                            'Next Bucle
                            If (Not HaySeguridad) Or (HaySeguridad And UsuarioActual.Asociado = 0) Or (NFiltros <> 3) Then
                                If Not Encontrado Then CmbFiltro(NFiltros).AddItem .Fields(sCampo)
                            End If
                    End Select
                End If
                .MoveNext
            Loop
        End If
    End With
    LblFiltro(NFiltros).Visible = True
    CmbFiltro(NFiltros).Visible = True
    CmbFiltro(NFiltros).Enabled = True
    If NFiltros = 0 Then
        If CmbFiltro(0).ListCount > 0 Then
            BufInicio = Iniciando
            Iniciando = True
            CmbFiltro(0).ListIndex = 0
            Iniciando = BufInicio
        End If
    End If
    NFiltros = NFiltros + 1
End Sub

Private Sub Pon_Cabeceras(sData() As String, lTam() As Integer, ndata As Integer)
    Dim Bucle As Integer
    Lista_Clear
    
    If ndata > 0 Then
        For Bucle = 0 To ndata - 1
            Lista.ColumnHeaders.Add , "S" & Format(Bucle, "000"), sData(Bucle), lTam(Bucle)
        Next Bucle
    End If
End Sub

Private Sub CheckFechas()
    Dim Bucle As Integer
    CmbFiltro(2).Clear
    CmbFiltro(2).AddItem "(" & CargaCadena(311) & ")" '"(TODOS)"
    For Bucle = 1 To CmbFiltro(1).ListCount - 1
        If Bucle >= CmbFiltro(1).ListIndex Then
            CmbFiltro(2).AddItem CmbFiltro(1).List(Bucle)
        End If
    Next Bucle
    CmbFiltro(2).ListIndex = 0
End Sub


'Private Sub ChkLocal_Click()
'Dim a
'a = a
'End Sub

Private Sub ChkTotPer_Click()
'llamar a consulta datos
 Call Consulta_Datos
End Sub

Private Sub CmbFiltro_Click(Index As Integer)
    Dim MyReg As ADODB.Recordset
    If QLTOT Then
        QLTOT = False
        Exit Sub
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ' Fechas, inicial siempre menor que la final
    '///////////////////////////////////////////
    If Index = 1 Then
        QLTOT = True
        CheckFechas
    End If
    If CmbTipo.ListIndex = 1 Then
        '\\\\\\\\\\\\\\\\\\\\\\\\\\
        ' para totales por artículo
        '//////////////////////////
        If Index = 3 Then
            QLTOT = True
            'CmbFiltro(5).ListIndex = 0
            If CmbFiltro(3).ListIndex = 0 Then
                'CmbFiltro(5).Enabled = False
            Else
                'CmbFiltro(6).Enabled = True
                CmbFiltro(3).ListIndex = 1
                
                CmbFiltro(4).Clear
                CmbFiltro(4).AddItem "(" & CargaCadena(311) & ")"
                Set MyReg = New ADODB.Recordset
                MyReg.Open "select distinct codi_fam from totales where codi_ident=" & Val(CmbFiltro(3).Text) & " order by codi_fam", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
                With MyReg
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            CmbFiltro(4).AddItem .Fields("codi_fam")
                            .MoveNext
                        Loop
                    End If
                End With
                QLTOT = True
                CmbFiltro(4).ListIndex = 0
                
                
                CmbFiltro(6).Clear
                CmbFiltro(6).AddItem "(" & CargaCadena(311) & ")" '"(Todos)"
                Set MyReg = New ADODB.Recordset
                MyReg.Open "select distinct plu from totales where codi_ident=" & Val(CmbFiltro(3).Text) & " order by plu", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
                With MyReg
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            CmbFiltro(6).AddItem .Fields("plu")
                            .MoveNext
                        Loop
                    End If
                End With
                QLTOT = True
                CmbFiltro(6).ListIndex = 0
            End If
        End If
    End If
    If Not Iniciando Then Consulta_Datos
End Sub


Private Sub CmbTipo_Click()
Dim Registro As New ADODB.Recordset
Dim Informacion As New ADODB.Recordset
Dim MyBase As DAO.Database
Dim nre As DAO.Recordset
Dim sData() As String
Dim nPlus As Long
Dim lTam() As Integer
Dim cadSub As String
Dim cadFam As String
Dim sSQL As String
Dim sCon As String
'Dim lTngt As Boolean
    lngt = False
    Command3.Enabled = False
    Command3.Visible = False
    
    nPlus = 0
    OptFiltro(0).Visible = False
    OptFiltro(1).Visible = False
    OptFiltro(2).Visible = False
    ChkTotPer.Visible = False
    
    Timer1.Enabled = False
    CorrerCortina
    Init_Filter
    Ordenacion = 1
    nPlus = 0
    Select Case CmbTipo.ListIndex
        Case 0
            cadSub = ""
            nPlus = 0
            'If SubseccionesFijas Then
            '    cadSub = ",txt_sub"
            '    nPlus = nPlus + 1
            'End If
            If FamiliasFijas Then
                cadFam = ",txt_fam"
                nPlus = nPlus + 1
            End If
            sSQL = "SELECT modo,fecha,codi_ident,codi_fam" & cadFam & ",tiquets,operacion,peso,unidades,pesetas " & _
                   "FROM gtsecs"
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & Base_General & ";" & _
                   "Persist Security Info=False"
            Informacion.Open sSQL, sCon
            CreaTablaPropia "totales", Informacion
            Importar_TotMaquina NomFichero(0)
            sSQL = "SELECT * " & _
                   "FROM totales"
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";" & _
                   "Persist Security Info=False"
            Registro.Open sSQL, sCon
            Add_Filter CargaCadena(1190), "modo", 0
            Add_Filter CargaCadena(116), "fecha", 1
            Add_Filter CargaCadena(117), "fecha", 2
            Add_Filter CargaCadena(329), "codi_ident", 0
            'Add_Filter CargaCadena(10), "codi_sub", 0      '///cas.v120
            Add_Filter CargaCadena(67), "codi_fam", 0    '///cas.v120
            ReDim sData(8 + nPlus)
            ReDim lTam(8 + nPlus)
            sData(0) = CargaCadena(434)
            sData(1) = CargaCadena(329)
            '---
            'sData(2) = CargaCadena(10)
            'If SubseccionesFijas Then
            '    sData(3) = CargaCadena(539) & " " & CargaCadena(10)
            '    sData(4) = CargaCadena(67)
            '    If FamiliasFijas Then sData(5) = CargaCadena(539) & " " & CargaCadena(67)
            'Else
                sData(2) = CargaCadena(67)
                If FamiliasFijas Then sData(3) = CargaCadena(539) & " " & CargaCadena(67)
            'End If
            sData(3 + nPlus) = CargaCadena(134)
            LblDesc(0) = CargaCadena(134)
            sData(4 + nPlus) = CargaCadena(133)
            LblDesc(1) = CargaCadena(133)
            sData(6 + nPlus) = CargaCadena(131)
            LblDesc(3) = CargaCadena(131)
            sData(5 + nPlus) = CargaCadena(130)
            LblDesc(2) = CargaCadena(130)
            If UsaEuro Then
                sData(7 + nPlus) = "Euros"
                LblDesc(4) = "Euros"
            Else
                sData(7 + nPlus) = SimboloMonetario
                LblDesc(4) = SimboloMonetario
            End If
            lTam(0) = 999
            lTam(1) = 800 - (10 * nPlus)
            '---
            'lTam(2) = 999
            'If SubseccionesFijas Then
            '    lTam(3) = 1500
            '    lTam(4) = 800
            '    If FamiliasFijas Then lTam(5) = 1200
            'Else
                lTam(2) = 800
                If FamiliasFijas Then lTam(3) = 900
            'End If
            lTam(3 + nPlus) = 1000 - (100 * nPlus)
            lTam(4 + nPlus) = 1200 - (50 * nPlus)
            lTam(5 + nPlus) = 700
            lTam(6 + nPlus) = 999 - (50 * nPlus)
            lTam(7 + nPlus) = 700
            NIndices = 5
            ReDim Indices(5)
            Indices(0) = 3 + nPlus
            Indices(1) = 4 + nPlus
            Indices(2) = 5 + nPlus
            Indices(3) = 6 + nPlus
            Indices(4) = 7 + nPlus
            'Indices(5) = 8 + nPlus
            Pon_Cabeceras sData, lTam, 8 + nPlus
        Case 1
            cadSub = ""
            nPlus = 0
            'If SubseccionesFijas Then
            '    cadSub = ",txt_sub"
            '    nPlus = nPlus + 1
            'End If
            If FamiliasFijas Then
                cadFam = ",txt_fam"
                nPlus = nPlus + 1
            End If
            ''''''''''
            
            'If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then 'seguimos fecha
            
            'sSQL = "SELECT modo,fecha,codi_ident,codi_sub" & cadSub & ",codi_fam" & cadFam & ",codigo,plu,des_plu1,sum(operacion) as operacion,sum(unidades) as unidades ,sum(peso) as peso ,sum(pesetas) as pesetas FROM gtarti"
            'sSQL = "select codigo,plu,des_plu1,sum(operacion) as operacion,sum(unidades) as unidades ,sum(peso) as peso ,sum(pesetas) as pesetas from gtarti where fecha group by codigo,plu,des_plu1"
            'hacer group by
            
            'Else
            'sSQL = "SELECT modo,fecha,codi_ident,codi_sub" & cadSub & ",codi_fam" & cadFam & ",codigo,plu,des_plu1,operacion,unidades,peso,pesetas "
            sSQL = "SELECT modo,fecha,codi_ident,codi_fam" & cadFam & ",codigo,plu,des_plu1,operacion,unidades,peso,pesetas " & _
                   "FROM gtarti"
            'hacer group by
            'End If
            '''''''''''
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & Base_General & ";" & _
                   "Persist Security Info=False"
            Informacion.Open sSQL, sCon
            
            CreaTablaPropia "totales", Informacion
            Importar_TotItem NomFichero(1)
            Registro.Open "select * from totales", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
            If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
                Add_Filter CargaCadena(1190), "modo", 0
                Add_Filter CargaCadena(116), "fecha", 1
                Add_Filter CargaCadena(117), "fecha", 2
                Add_Filter CargaCadena(329), "codi_ident", 0
                'Add_Filter CargaCadena(10), "codi_sub", 0      '///cas.v120
                Add_Filter CargaCadena(67), "codi_fam", 0    '///cas.v120
                Add_Filter CargaCadena(69), "codigo", 0
                Add_Filter "PLU", "plu", 0
            Else
                Add_Filter CargaCadena(1190), "modo", 0
                Add_Filter CargaCadena(116), "fecha", 1
                Add_Filter CargaCadena(117), "fecha", 2
                Add_Filter CargaCadena(329), "codi_ident", 0
                'Add_Filter CargaCadena(10), "codi_sub", 0      '///cas.v120
                Add_Filter CargaCadena(67), "codi_fam", 0    '///cas.v120
                Add_Filter CargaCadena(69), "codigo", 0
                Add_Filter "PLU", "plu", 0
            End If
            '---
            'OptFiltro(0).Visible = True
            'OptFiltro(1).Visible = True
            'OptFiltro(2).Visible = True
            '----
            ChkTotPer.Visible = True
            '---
            'OptFiltro(0).Value = True
            'OptFiltro(0).Caption = CargaCadena(471)
            'OptFiltro(1).Caption = CargaCadena(329) & " + PLU"
            'OptFiltro(2).Caption = CargaCadena(69)
            ''''
            Iniciando = True
            'CmbFiltro(3).Enabled = False
            CmbFiltro(3).Enabled = True
            'CmbFiltro(4).Enabled = False
            CmbFiltro(4).Enabled = True
            'CmbFiltro(5).Enabled = False
            CmbFiltro(5).Enabled = True
            CmbFiltro(6).Enabled = True
            '---
            CmbFiltro(1).ListIndex = 0
            CmbFiltro(2).ListIndex = CmbFiltro(2).ListCount - 1
            Iniciando = False
            '---
            ReDim sData(10 + nPlus)
            ReDim lTam(10 + nPlus)
            sData(0) = CargaCadena(434)
            sData(1) = CargaCadena(329) 'seccion
            'sData(2) = CargaCadena(10)  'subSeccion
            'If SubseccionesFijas Then
            '    'sData(3) = CargaCadena(539) & " " & CargaCadena(10)    'txtSubSeccion
            '    sData(3) = "Txt. " & CargaCadena(10)     'txtSubSeccion
            '    sData(4) = CargaCadena(67)
            '    'If FamiliasFijas Then sData(5) = CargaCadena(539) & " " & CargaCadena(67)
            '    If FamiliasFijas Then sData(5) = "Txt. " & CargaCadena(67)
            'Else
                sData(2) = CargaCadena(67)
                'If FamiliasFijas Then sData(4) = CargaCadena(539) & " " & CargaCadena(67)
                If FamiliasFijas Then sData(3) = "Txt. " & CargaCadena(67)
            'End If
            sData(3 + nPlus) = CargaCadena(69)
            sData(4 + nPlus) = "PLU"
            sData(5 + nPlus) = CargaCadena(70)
            sData(6 + nPlus) = CargaCadena(133)
            sData(7 + nPlus) = CargaCadena(131)
            sData(8 + nPlus) = CargaCadena(130)
            'sData(9 + nPlus) = "Euros"
            
            If UsaEuro Then
                sData(9 + nPlus) = "Euros"
            Else
                sData(9 + nPlus) = SimboloMonetario
            End If
            
            LblDesc(0) = CargaCadena(133)
            LblDesc(1) = CargaCadena(131)
            LblDesc(2) = CargaCadena(130)
            If UsaEuro Then
                LblDesc(3) = "Euros"
            Else
                LblDesc(3) = SimboloMonetario
            End If
            lTam(0) = 999
            lTam(1) = 750
            'lTam(2) = 700
            lTam(2) = 1100
            lTam(3) = 1070
            lTam(4) = 850
            lTam(5) = 700
            lTam(6) = 800
            lTam(7) = 800
            lTam(8) = 800
            lTam(9) = 1100
            lTam(9 + nPlus) = 800
            lTam(10 + nPlus) = 1000
            NIndices = 4
            ReDim Indices(4)
            Indices(0) = 6 + nPlus
            Indices(1) = 7 + nPlus
            Indices(2) = 8 + nPlus
            Indices(3) = 9 + nPlus
            Pon_Cabeceras sData, lTam, 10 + nPlus
        Case 2
            'Informacion.Open "select modo,fecha,codi_ident,ident_vend,nombre,tiquets,cancela,impor_neg,operacion,unidad,peso,pesetas from gtvend", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
            'CreaTablaPropia "totales", Informacion
            'importar_totVen NomFichero(2)
            'Registro.Open "select * from totales", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
            'Add_Filter CargaCadena(1190), "modo", 0
            'Add_Filter CargaCadena(116), "fecha", 1
            'Add_Filter CargaCadena(117), "fecha", 2
            'Add_Filter CargaCadena(329), "codi_ident", 0
            'Add_Filter CargaCadena(470), "ident_vend", 0
            'ReDim sData(11)
            'ReDim lTam(11)
            'sData(0) = CargaCadena(434)
            'sData(1) = CargaCadena(329)
            'sData(2) = CargaCadena(470)
            'sData(3) = CargaCadena(13)
            'sData(4) = CargaCadena(134)
            'sData(5) = "Cancel."
            'sData(6) = "Imp.Neg."
            'sData(7) = CargaCadena(133)
            'sData(8) = CargaCadena(131)
            'sData(9) = CargaCadena(130)
            'If UsaEuro Then
            '    sData(10) = "Euros"
            'Else
            '    sData(10) = SimboloMonetario
            'End If
            'lTam(0) = 1000
            'lTam(1) = 800
            'lTam(2) = 800
            'lTam(3) = 1300
            'lTam(4) = 900
            'lTam(5) = 800
            'lTam(6) = 1000
            'lTam(7) = 1100
            'lTam(8) = 1000
            'lTam(9) = 800
            'lTam(10) = 800
            'LblDesc(0) = CargaCadena(133)
            'LblDesc(1) = CargaCadena(131)
            'LblDesc(2) = CargaCadena(130)
            'If UsaEuro Then
            '    LblDesc(3) = "Euros"
            'Else
            '    LblDesc(3) = SimboloMonetario
            'End If
            'NIndices = 4
            'ReDim Indices(4)
            'Indices(0) = 7
            'Indices(1) = 8
            'Indices(2) = 9
            'Indices(3) = 10
            'Pon_Cabeceras sData, lTam, 11
            Informacion.Open "select modo,fecha,codi_ident,ident_vend,nombre,cancela,impor_neg,tiquets,operacion,unidad,peso,pesetas from gtvend", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Base_General & ";Persist Security Info=False"
            CreaTablaPropia "totales", Informacion
            importar_totVen NomFichero(2)
            Registro.Open "select * from totales", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
            Add_Filter CargaCadena(1190), "modo", 0
            Add_Filter CargaCadena(116), "fecha", 1
            Add_Filter CargaCadena(117), "fecha", 2
            Add_Filter CargaCadena(329), "codi_ident", 0
            Add_Filter CargaCadena(470), "ident_vend", 0
            ReDim sData(11)
            ReDim lTam(11)
            sData(0) = CargaCadena(434)
            sData(1) = CargaCadena(329)
            sData(2) = CargaCadena(470)
            sData(3) = CargaCadena(13)
            'cambio tiques
            sData(4) = "Cancel." 'CargaCadena(134)
            sData(5) = "Imp.Neg." '"Cancel."
            sData(6) = CargaCadena(134) '"Imp.Neg."
            sData(7) = CargaCadena(133)
            sData(8) = CargaCadena(131)
            sData(9) = CargaCadena(130)
            If UsaEuro Then
                sData(10) = "Euros"
            Else
                sData(10) = SimboloMonetario
            End If
            lTam(0) = 1000
            lTam(1) = 800
            lTam(2) = 800
            lTam(3) = 1300
            lTam(4) = 900
            lTam(5) = 800
            lTam(6) = 1000
            lTam(7) = 1100
            lTam(8) = 1000
            lTam(9) = 800
            lTam(10) = 800
            LblDesc(0) = CargaCadena(133)
            LblDesc(1) = CargaCadena(131)
            LblDesc(2) = CargaCadena(130)
            If UsaEuro Then
                LblDesc(3) = "Euros"
            Else
                LblDesc(3) = SimboloMonetario
            End If
            LblDesc(4) = CargaCadena(134)
            NIndices = 5
            ReDim Indices(5)
            Indices(0) = 7
            Indices(1) = 8
            Indices(2) = 9
            Indices(3) = 10
            Indices(4) = 6 'cambio tiques 4
            Pon_Cabeceras sData, lTam, 11
            
        Case 3
            cadSub = ""
            nPlus = 0
            NFiltros = 1
            Command3.Enabled = True
            Command3.Visible = True
            
            'If SubseccionesFijas Then
            '    cadSub = ",txt_sub"
            '    nPlus = nPlus + 1
            'End If
            'If FamiliasFijas Then
            '    cadFam = ",txt_fam"
            '    nPlus = nPlus + 1
            'End If
            sSQL = "SELECT fecha,codi_ident,sum(pesetas) as pesetas " & _
                   "FROM gtsecs group by fecha,codi_ident"
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & Base_General & ";" & _
                   "Persist Security Info=False"
            Informacion.Open sSQL, sCon
            CreaTablaPropia "totales", Informacion
            Importar_TotMaquina NomFichero(0)
            sSQL = "SELECT fecha,codi_ident,pesetas " & _
                   "FROM totales"
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";" & _
                   "Persist Security Info=False"
            Registro.Open sSQL, sCon
            'Add_Filter CargaCadena(1190), "modo", 0
            Add_Filter CargaCadena(116), "fecha", 1
            Add_Filter CargaCadena(117), "fecha", 2
            Add_Filter CargaCadena(329), "codi_ident", 0
            'Add_Filter CargaCadena(10), "codi_sub", 0      '///cas.v120
            'Add_Filter CargaCadena(67), "codi_fam", 0    '///cas.v120
            CmbFiltro(0).Enabled = False
            CmbFiltro(4).Enabled = False
            CmbFiltro(5).Enabled = False
            CmbFiltro(1).Enabled = True
            CmbFiltro(2).Enabled = True
            CmbFiltro(3).Enabled = True
            ReDim sData(4 + nPlus)
            ReDim lTam(4 + nPlus)
            sData(0) = CargaCadena(434)
            sData(1) = CargaCadena(329)
            sData(2) = CargaCadena(132)
            sData(3) = "N.G.T"
            'If SubseccionesFijas Then
            '    sData(3) = CargaCadena(539) & " " & CargaCadena(10)
            '    sData(4) = CargaCadena(67)
            '    If FamiliasFijas Then sData(5) = CargaCadena(539) & " " & CargaCadena(67)
            'Else
            '    sData(3) = CargaCadena(67)
            '    If FamiliasFijas Then sData(4) = CargaCadena(539) & " " & CargaCadena(67)
            'End If
            'sData(4 + nPlus) = CargaCadena(134)
            'LblDesc(0) = CargaCadena(134)
            'sData(5 + nPlus) = CargaCadena(133)
            'LblDesc(1) = CargaCadena(133)
            'sData(7 + nPlus) = CargaCadena(131)
            'LblDesc(3) = CargaCadena(131)
            'sData(6 + nPlus) = CargaCadena(130)
            'LblDesc(2) = CargaCadena(130)
            If UsaEuro Then
                sData(2 + nPlus) = "Euros"
                LblDesc(4) = "Euros"
            Else
                sData(2 + nPlus) = SimboloMonetario
                LblDesc(4) = SimboloMonetario
            End If
            lTam(0) = 999
            lTam(1) = 800 - (10 * nPlus)
            lTam(2) = 999
            lTam(3) = 999
            'If SubseccionesFijas Then
            '    lTam(3) = 1500
            '    lTam(4) = 800
            '    If FamiliasFijas Then lTam(5) = 1200
            'Else
            '    lTam(3) = 800
            '    If FamiliasFijas Then lTam(4) = 900
            'End If
            'lTam(4 + nPlus) = 1000 - (100 * nPlus)
            'lTam(5 + nPlus) = 1200 - (50 * nPlus)
            'lTam(6 + nPlus) = 700
            'lTam(7 + nPlus) = 999 - (50 * nPlus)
            'lTam(8 + nPlus) = 700
            NIndices = 2
            ReDim Indices(2)
            Indices(0) = 0 + nPlus
            Indices(1) = 1 + nPlus
            Indices(2) = 2 + nPlus
            'Indices(3) = 7 + nPlus
            lngt = True
            
            Pon_Cabeceras sData, lTam, 4 + nPlus

    End Select
    Timer1.Enabled = True
    If Not lngt Then
        Muestra_Indices
    Else
        'Muestra_IndicesNGT
    
    End If
    Consulta_Datos
End Sub
Private Sub Muestra_Indices()
    Dim Bucle As Integer
    Dim B2 As Integer
    Dim MyVal As Double
    For Bucle = 0 To 4
        LblTot(Bucle).Visible = False
        LblDesc(Bucle).Visible = False
    Next Bucle
    For Bucle = 0 To NIndices - 1
        MyVal = 0
        LblTot(Bucle).Visible = True
        LblDesc(Bucle).Visible = True
        For B2 = 1 To Lista.ListItems.Count
            On Error Resume Next
            MyVal = MyVal + ToDouble((Lista.ListItems(B2).SubItems(Indices(Bucle))))
            On Error GoTo 0
        Next B2
        LblTot(Bucle).Caption = MyVal
    Next Bucle
End Sub
Private Sub Muestra_IndicesNGT()
    Dim Bucle As Integer
    Dim B2 As Integer
    Dim MyVal As Double
    For Bucle = 0 To 4
        LblTot(Bucle).Visible = False
        LblDesc(Bucle).Visible = False
    Next Bucle
    'For Bucle = 1 To NIndices - 1
        MyVal = 0
        LblTot(4).Visible = True
        LblDesc(4).Visible = True
        For B2 = 1 To Lista.ListItems.Count
            MyVal = MyVal + ToDouble((Lista.ListItems(B2).SubItems(2)))
        Next B2
        LblTot(4).Caption = MyVal
    'Next Bucle
End Sub

Private Function BuscaOrden() As String
    Dim MyReg As New ADODB.Recordset
    Dim Resp As String
    'If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
    '    Resp = " group by codigo"
    'Else
        MyReg.Open "select * from totales", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
        
        If ((lngt) And Ordenacion > 2) Then Ordenacion = 0
        
        Resp = MyReg.Fields(Ordenacion).Name
    
    'End If
    BuscaOrden = Resp
End Function
'modificar para totalizar por periodo artículos
Private Sub Consulta_Datos() 'Optional ByVal lngt As Boolean)
    Dim Bucle As Integer
    Dim CadFiltro, sCadTmp As String
    Dim Contador As Long
    Dim Registro As New ADODB.Recordset
    Dim nIni As Integer
    'Dim Arch As Integer
    Dim nLastGtTmp As Long
    Dim nI As Variant
    
    nLastGtTmp = nlastGT
    
    If IsNull(lngt) Then lngt = False
    If lngt Then
        nIni = 1
    Else
        nIni = 0
    End If
    
    CorrerCortina
    Lista.ListItems.Clear
    If CmbFiltro(0).ListCount = 0 And nIni = 0 Then
        Command1.Enabled = False
        IzarCortina
        Exit Sub
    End If
    For Bucle = nIni To NFiltros
        If (Bucle = 0) Or (CmbFiltro(Bucle).ListIndex > 0) Or ((Bucle = 3 And IsNumeric(CmbFiltro(3).Text)) And (HaySeguridad And (UsuarioActual.Asociado <> 0))) Then
            If CadFiltro = "" Then
                CadFiltro = " where "
            Else
                CadFiltro = CadFiltro & " and "
            End If
            CadFiltro = CadFiltro & "("
            If Bucle = 0 Then
                'If CmbTipo.ListIndex = 0 Then
                    CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & CStr(Filtro0(CmbFiltro(0).ListIndex)) & ")"
                'Else
                'If Filtro0(CmbFiltro(0).ListIndex) = 1 Then
                '    CadFiltro = CadFiltro & Filtros(Bucle).Campo & ">=" & CStr(Filtro0(CmbFiltro(0).ListIndex)) & " and " & Filtros(Bucle).Campo & "<=8" & ")"
                'Else
                '    CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & CStr(Filtro0(CmbFiltro(0).ListIndex)) & ")"
                'End If
                'End If
            Else
                Select Case Filtros(Bucle).Tipo
                    Case 1
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & ">=#" & Mid(CmbFiltro(Bucle).Text, 4, 2) & "/" & Left(CmbFiltro(Bucle).Text, 2) & "/" & Right(CmbFiltro(Bucle).Text, 2) & "#"
                    Case 2
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "<=#" & Mid(CmbFiltro(Bucle).Text, 4, 2) & "/" & Left(CmbFiltro(Bucle).Text, 2) & "/" & Right(CmbFiltro(Bucle).Text, 2) & "#"
                    Case 3
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & Chr(34) & CmbFiltro(Bucle).Text & Chr(34)
                    Case Else
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & CmbFiltro(Bucle).Text
                End Select
                CadFiltro = CadFiltro & ")"
            End If
        End If
    Next Bucle
    If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
        If InStr(1, UCase(CadFiltro), "FECHA") <> 0 Then
            nI = InStr(1, UCase(CadFiltro), "FECHA")
            CadFiltro = " where " & Mid(CadFiltro, nI - 1)
        Else
            CadFiltro = ""
        End If
        Registro.Open "select codi_fam, txt_fam,codigo,plu,des_plu1,sum(operacion) as operacion,sum(unidades) as unidades,sum(peso) as peso,sum(pesetas) as pesetas from totales" & CadFiltro & " group by codi_fam,txt_fam,codigo,plu,des_plu1", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
        If FamiliasFijas Then '2.0.12
            nI = 2 '4 '-- 5
        Else
            nI = 1 '3 '-- 4
        End If
        
    Else
        CadFiltro = CadFiltro & " order by " & BuscaOrden
        Registro.Open "select * from totales" & CadFiltro, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False"
        nI = 0
    End If
    
    
    With Registro
        If Not (.EOF And .BOF) Then
            
            For Bucle = 0 To Lista.ColumnHeaders.Count - 1  ' en lugar de 4 0
                .MoveFirst
                'If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
                '    Contador = 4
                'Else
                    Contador = 1
                'End If
                Do Until .EOF
                    If Bucle = 0 Then
                        If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
                            If Bucle < nI Then
                                Lista.ListItems.Add , , "---"
                            Else
                                Lista.ListItems.Add , , .Fields(Bucle)
                            End If
                        Else
                            Lista.ListItems.Add , , .Fields(Bucle + (1 - nIni))
                        End If
                    Else
                        
                        'c2f aquí incluir autonumeración gran total columna NGT
                        If lngt And (Bucle = 3) Then
                            Lista.ListItems(Contador).SubItems(Bucle) = nLastGtTmp
                            
                            nLastGtTmp = nLastGtTmp + 1
                        Else
                            'On Error Resume Next
                            If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
                                '---
                                'If Bucle < 4 Then
                                'If Bucle < 3 Then
                                If Bucle < nI Then
                                        Lista.ListItems(Contador).SubItems(Bucle) = "---"
                                Else

                                    If Not IsNull(.Fields(Bucle - nI)) Then '
                                        Lista.ListItems(Contador).SubItems(Bucle) = .Fields(Bucle - nI)
                                    Else
                                        Lista.ListItems(Contador).SubItems(Bucle - nI) = "---"
                                    End If
                                End If
                            Else
                                If Not IsNull(.Fields(Bucle + (1 - nIni))) Then
                                    Lista.ListItems(Contador).SubItems(Bucle) = .Fields(Bucle + (1 - nIni))
                                Else
                                    Lista.ListItems(Contador).SubItems(Bucle) = 0
                                End If
                            End If
                            'If Err.Number <> 0 Then
                            '    Contador = Contador
                            'End If
                            'On Error GoTo 0
                        End If
                        Contador = Contador + 1
                    End If
                    .MoveNext
                Loop
            Next Bucle
            'c2f aquí se guarda en fichero el último número de GT.
            'If lngt Then
            '    Arch = FreeFile()
            '    Open sPathNGT & "ngtord" For Output As #Arch
            '    Print #Arch, Format(nlastGT, "000000")
            '    Close #Arch
            'End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End With
    If CmbFiltro(0).ListCount = 0 Then
        For Bucle = 0 To 9
            CmbFiltro(Bucle).Enabled = False
        Next Bucle
        'For Bucle = 0 To 2
        '    OptFiltro(Bucle).Enabled = False
        'Next Bucle
        Command1.Enabled = False
    Else
        Command1.Enabled = True
        'For Bucle = 0 To 2
        '    OptFiltro(Bucle).Enabled = True
        'Next Bucle
    End If
    If Not lngt Then
        Muestra_Indices
    Else
        Muestra_IndicesNGT
        CmbFiltro(1).Enabled = True
        CmbFiltro(2).Enabled = True
        CmbFiltro(3).Enabled = True
        Command1.Enabled = True
        
    End If
    IzarCortina
End Sub
Private Sub ListaInforme()
Dim Bucle As Integer
Dim Campos() As String
Dim MyCont As Integer
Dim Longitudes() As Integer
Dim nTots As Integer
Dim Tots() As String
Dim txt_Sub As String
Dim CadFiltro As String
Dim RestaSub As Long
Dim sSQL As String
Dim a, a1, a2 As String

    a1 = ""
    a2 = ""
    a = ""
    
    If Not lngt Then
    
        If SubseccionesFijas Then
            txt_Sub = "txt_sub ,"
            RestaSub = 1
        End If
        
        If ChkLocal.Value = vbUnchecked Then
            CadFiltro = " where modo=" & Filtro0(CmbFiltro(0).ListIndex)
        Else
            CadFiltro = " where (modo>=1 and modo<=8)"
        End If
        For Bucle = 1 To NFiltros
           If HaySeguridad And UsuarioActual.Asociado <> 0 Then
                CadFiltro = " where codi_ident=" & UsuarioActual.Asociado
           End If
           If CmbFiltro(Bucle).ListIndex > 0 Then
                If CadFiltro = "" Then
                    CadFiltro = " where "
                Else
                    CadFiltro = CadFiltro & " and "
                End If
                CadFiltro = CadFiltro & "("
                Select Case Filtros(Bucle).Tipo
                    Case 1
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & ">=#" & Mid(CmbFiltro(Bucle).Text, 4, 2) & "/" & Left(CmbFiltro(Bucle).Text, 2) & "/" & Right(CmbFiltro(Bucle).Text, 2) & "#"
                        a1 = CmbFiltro(Bucle).Text
                    Case 2
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "<=#" & Mid(CmbFiltro(Bucle).Text, 4, 2) & "/" & Left(CmbFiltro(Bucle).Text, 2) & "/" & Right(CmbFiltro(Bucle).Text, 2) & "#"
                        a2 = CmbFiltro(Bucle).Text
                    Case 3
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & Chr(34) & CmbFiltro(Bucle).Text & Chr(34)
                    Case Else
                        CadFiltro = CadFiltro & Filtros(Bucle).Campo & "=" & CmbFiltro(Bucle).Text
                End Select
                CadFiltro = CadFiltro & ")"
            Else
                Select Case Filtros(Bucle).Tipo
                    Case 1
                        CmbFiltro(1).ListIndex = 1
                        a1 = CmbFiltro(1).Text
                    Case 2
                        CmbFiltro(2).ListIndex = CmbFiltro(2).ListCount - 1
                        a2 = CmbFiltro(2).Text
                End Select
            
            End If
        Next Bucle
    
    End If
    
    If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
    
    Else
        CadFiltro = CadFiltro & " order by " & BuscaOrden
    End If
    
    MyCont = 1
    ReDim Campos(Lista.ColumnHeaders.Count)
    ReDim Longitudes(Lista.ColumnHeaders.Count)
    For Bucle = 1 To Lista.ColumnHeaders.Count
        If Not lngt Then
            If Bucle <> 2 Then
                Campos(MyCont) = Lista.ColumnHeaders(Bucle).Text
                Longitudes(MyCont) = 7
                MyCont = MyCont + 1
            End If
        Else
            Campos(MyCont) = Lista.ColumnHeaders(Bucle).Text
            Longitudes(MyCont) = 7
            MyCont = MyCont + 1
        
        End If
    Next Bucle
    'aquí 1
    Select Case CmbTipo.ListIndex
        Case 0
            Longitudes(0) = 7 - (RestaSub * 2)
            Longitudes(1) = 6 - (RestaSub * 2)
            Longitudes(2) = 7 - (RestaSub * 2)
            Longitudes(3) = 7 - RestaSub
            Longitudes(4) = 7 - RestaSub
            Longitudes(5) = 7 - RestaSub
            nTots = 5
            ReDim Tots(5)
            Tots(1) = "tiquets"
            Tots(2) = "operacion"
            'c2f
            'Tots(3) = "unidades"
            'Tots(4) = "peso"
            Tots(4) = "unidades"
            Tots(3) = "peso"
            
            Tots(5) = "pesetas"
            'If FamiliasFijas Then
                Informe_Totales CargaCadena(153), "select codi_ident as agrupador,fecha,codi_fam,txt_fam,tiquets,operacion,peso,unidades,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            'Else
            '    Informe_Totales CargaCadena(153), "select codi_ident as agrupador,fecha,codi_sub, " & txt_Sub & " codi_fam,tiquets,operacion,peso,unidades,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            'End If
        
        Case 1
            Longitudes(2) = 5 - RestaSub
            Longitudes(3) = 5 - RestaSub
            Longitudes(4) = 5 - RestaSub
            Longitudes(5) = 5 - RestaSub
            Longitudes(6) = 9 - RestaSub
            Longitudes(7) = 5 - RestaSub
            Longitudes(8) = 5 - RestaSub
            If FamiliasFijas Then
                Longitudes(2) = 4 - RestaSub
                Longitudes(3) = 4 - RestaSub
                Longitudes(4) = 4 - RestaSub
                Longitudes(6) = 7 - RestaSub
            End If
            nTots = 4
            ReDim Tots(4)
            Tots(1) = "operacion"
            Tots(2) = "unidades"
            Tots(3) = "peso"
            Tots(4) = "pesetas"
            'If FamiliasFijas Then
                sSQL = "SELECT codi_ident AS agrupador,fecha, codi_fam,txt_fam,codigo,plu,des_plu1,operacion,unidades,peso,pesetas " & _
                       "FROM totales" & CadFiltro
            'Else
            '    sSQL = "SELECT codi_ident AS agrupador,fecha,codi_sub, " & txt_Sub & " codi_fam,codigo,plu,des_plu1,operacion,unidades,peso,pesetas " & _
            '           "FROM totales" & CadFiltro
            'End If
            'introducir group by
            a = a
            If CmbTipo.ListIndex = 1 And ChkTotPer.Value Then
                a = ""
                If InStr(1, UCase(CadFiltro), "FECHA>") <> 0 Then
                    a = " where " & Mid(CadFiltro, InStr(1, UCase(CadFiltro), "FECHA") - 1)
                End If
                
                '---sSQL = "select codi_ident as agrupador,'Periodo' as fecha,'0' as codisub, '0' as codi_fam,codigo,plu,des_plu1,sum(operacion) as operacion,sum(unidades) as unidades,sum(peso) as peso,sum(pesetas) as pesetas FROM totales " & a & " group by codi_ident,plu,codigo,des_plu1"
                sSQL = "select codi_ident as agrupador,'Periodo' as fecha, codi_fam,txt_fam,codigo,plu,des_plu1,sum(operacion) as operacion,sum(unidades) as unidades,sum(peso) as peso,sum(pesetas) as pesetas FROM totales " & a & " group by codi_ident,codi_fam,txt_fam,plu,codigo,des_plu1"
                a = CargaCadena(156) & " Tot. Per. " & a1 & "-->" & a2
            Else
                a = CargaCadena(156)
            End If
            ''''''''''''''''''''
            Informe_Totales a, sSQL, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
        Case 2
            'nTots = 4
            'ReDim Tots(4)
            'Longitudes(2) = 7
            'Longitudes(3) = 5
            'Longitudes(4) = 5
            'Longitudes(5) = 5
            'Longitudes(6) = 5
            'Longitudes(7) = 5
            'Longitudes(8) = 5
            'Tots(1) = "operacion"
            'Tots(2) = "unidad"
            'Tots(3) = "peso"
            'Tots(4) = "pesetas"
            'Informe_Totales CargaCadena(157), "select codi_ident as agrupador,fecha,ident_vend,nombre,tiquets,cancela,impor_neg,operacion,unidad,peso,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            nTots = 5
            ReDim Tots(5)
            Longitudes(2) = 7
            Longitudes(3) = 5
            Longitudes(4) = 5
            Longitudes(5) = 5
            Longitudes(6) = 5
            Longitudes(7) = 5
            Longitudes(8) = 5
            Tots(1) = "tiquets"
            'Tots(2) = "cancela"
            'Tots(3) = "impor_neg"
            Tots(2) = "operacion"
            Tots(3) = "unidad"
            Tots(4) = "peso"
            Tots(5) = "pesetas"
            'Informe_Totales CargaCadena(157), "select codi_ident as agrupador,fecha,ident_vend,nombre,tiquets,cancela,impor_neg,operacion,unidad,peso,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            Informe_Totales CargaCadena(157), "select codi_ident as agrupador,fecha,ident_vend,nombre,cancela,impor_neg,tiquets,operacion,unidad,peso,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            
        Case 3
            'nTots = 4
            'ReDim Tots(4)
            'Longitudes(0) = 5
            'Longitudes(1) = 5
            'Longitudes(2) = 5
            'Longitudes(3) = 5
            'Tots(0) = "agrupador"
            'Tots(1) = "codi_ident"
            'Tots(2) = "pesetas"
            'Tots(3) = Lista.ColumnHeaders(4)
            'c2f aquí llamar a posible nuevo informe
            'Informe_Totales CargaCadena(153) & " - N.G.T.", "select fecha as agrupador,codi_ident,pesetas from totales" & CadFiltro, Campos, Longitudes, Lista.ColumnHeaders.Count, Tots, nTots
            Dialogo.CancelError = True
            On Error Resume Next
            Dialogo.ShowPrinter
            If Err.Number = 0 Then
                Do_Events
                ImprimeLosNgt
            End If
            On Error GoTo 0
    End Select
End Sub
Private Sub impCab()
    Dim nFichCabe As Integer
    Dim sFichCabe(6) As String
    Dim nCont As Integer
    
    If Dir(sPathNGT & "cabngt.txt") <> "" Then
        nFichCabe = FreeFile()
        Open sPathNGT & "cabngt.txt" For Input As #nFichCabe
        On Error Resume Next
        For nCont = 1 To 6
            Line Input #nFichCabe, sFichCabe(nCont)
            If Err.Number <> 0 Then Exit For
        Next nCont
        On Error GoTo 0
        Close #nFichCabe
    Else
        sFichCabe(1) = "Listado Ordenado. Total Mostradores."
        sFichCabe(2) = Space(50)
        sFichCabe(3) = Space(50)
        sFichCabe(4) = Space(50)
        sFichCabe(5) = Space(50)
        sFichCabe(6) = Space(50)
    End If
    
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True

    'Printer.FontUnderline = True
    'Printer.Print "Tickets Sistema  " & Combo1.Text & " Listado el " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    'Printer.Print CargaCadena(1226) & " " & Combo1.Text & ". " & CargaCadena(1227) & "  " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    For nCont = 1 To 6
        Printer.Print sFichCabe(nCont)
    Next nCont

End Sub
Private Sub ImprimeLosNgt()
    Dim Bucle As Integer
    Dim Bucle2 As Integer
    Dim CuentaLineas As Long
    Dim nFichCabe As Integer
    Dim sFichCabe(6) As String
    Dim nCont As Integer
    Dim nSuperTot As Double
    Dim nContLin As Integer
    
    nSuperTot = 0
    
    Call impCab
    
    Printer.FontBold = False
    Printer.FontSize = 10
    
    Printer.Print ""
    Printer.FontUnderline = True
    Printer.Print "--- " & CargaCadena(434) & " ---    --- " & CargaCadena(47) & " ---   --- " & CargaCadena(132) & " ---   --- Nº     G.T ---"
    Printer.FontUnderline = False
    
    Printer.FontBold = True
    
    nCont = 0
    
    For Bucle = 1 To Lista.ListItems.Count
        sFichCabe(0) = " " & Format(Lista.ListItems(Bucle), "dd/mm/yyyy") & "            " & Format(Lista.ListItems(Bucle).SubItems(1), "00") & "             " & _
                    Format(Lista.ListItems(Bucle).SubItems(2), "0000000.00") & "            " & Format(Lista.ListItems(Bucle).SubItems(3), "000000")
        Printer.Print sFichCabe(0)
        nCont = nCont + 1
        If nCont > 54 Then
            Printer.NewPage
            Call impCab
            nCont = 0
            
        End If
        nSuperTot = nSuperTot + Lista.ListItems(Bucle).SubItems(2)
    Next Bucle
        
    Printer.Print ""
    Printer.Print CargaCadena(124) & " .....................:" & Format(nSuperTot, "0000000.00")
    
    Printer.EndDoc
End Sub

'\\\\\\
'//////
Private Sub Command1_Click()
    Dim nArch As Integer
    nArch = FreeFile()
    Open sPathNGT & "ngtord" For Output As #nArch
    Print #nArch, Format(nlastGT, "00000") 'Format(nlastGT + FrmTotales.Lista.ListItems.Count, "000000")
    Close #nArch
    
    ListaInforme
End Sub

Private Sub CorrerCortina()
    Dim vPaso As Integer
    Dim Bucle As Integer
    If Dir(App.Path & "\linux.txt") = "" Then
        Label2.Visible = True
        Do_Events
        If Cortina.Height = Lista.Height Then Exit Sub
        For Bucle = 0 To Lista.Height Step 400
            Cortina.Height = Bucle
            Do_Events
        Next Bucle
    End If
    Cortina.Height = Lista.Height
    Do_Events
End Sub
Private Sub IzarCortina()
    Dim Bucle As Integer
    Dim vPaso As Integer
    If Dir(App.Path & "\linux.txt") = "" Then
        Label2.Visible = False
        For Bucle = Lista.Height To 0 Step -400
            Cortina.Height = Bucle
            Do_Events
        Next Bucle
    End If
    Cortina.Height = 0
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub




Private Sub Command3_Click()
    Dim Arch As Integer
    Dim sParte As String
    'c2f aquí obtener número inicial G.T.
    If Dir(sPathNGT & "ngtord") = "" Then

        nlastGT = 1
        
    Else
        Arch = FreeFile()
        Open sPathNGT & "ngtord" For Input As #Arch
        Line Input #Arch, sParte
        Close #Arch
        If Val(sParte) > 9999 Then sParte = "0001"
        
        nlastGT = Val(sParte)
        Load frmnGT
        frmnGT.Visible = True
        Do While frmnGT.Visible = True
            Sleep (10)
            DoEvents
        Loop
    End If
    
    Consulta_Datos

End Sub

Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = Form2.Icon
    Me_Init
    Me.Width = Screen.Width * 0.97
    Me.Left = Screen.Width * 0.015
    Init_Filter
    Lista_Clear
End Sub
Private Function NomFichero(ByVal nfichero As Integer) As String
    Dim nRet As String
    If TiendaActual = 0 Then
        nRet = App.Path & "\01"
    Else
        nRet = App.Path & "\T" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00")
    End If
    Select Case nfichero
        Case 0
            nRet = nRet & "gtmaq.tot"
        Case 1
            nRet = nRet & "gtart.tot"
        Case 2
            nRet = nRet & "gtvend.tot"
    End Select
    NomFichero = nRet
End Function
Private Sub Lista_Clear()
    Lista.ListItems.Clear
    Lista.ColumnHeaders.Clear
End Sub
Private Sub Form_Resize()
    Me_Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Fin_Totales = True
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Ordenacion = ColumnHeader.Index
    If lngt Then
        Ordenacion = Ordenacion - 1
        If Ordenacion = 3 Then Exit Sub
    End If
    Consulta_Datos
End Sub

Private Sub OptFiltro_Click(Index As Integer)
    CmbFiltro(3).ListIndex = 0
    CmbFiltro(4).ListIndex = 0
    CmbFiltro(5).ListIndex = 0
    CmbFiltro(6).ListIndex = 0
    'CmbFiltro(7).ListIndex = 0
    Select Case Index
        Case 0
            'CmbFiltro(3).Enabled = False
            CmbFiltro(3).Enabled = True
            'CmbFiltro(4).Enabled = False
            CmbFiltro(4).Enabled = True
            CmbFiltro(5).Enabled = False
            CmbFiltro(6).Enabled = False
            'CmbFiltro(7).Enabled = False
        Case 1
            'If (Not HaySeguridad) Or (HaySeguridad And (UsuarioActual.Asociado = 0)) Then CmbFiltro(3).Enabled = True
            'CmbFiltro(4).Enabled = True
            'CmbFiltro(4).Enabled = False
            'CmbFiltro(5).Enabled = True
            CmbFiltro(5).Enabled = False
            'CmbFiltro(6).Enabled = False
            CmbFiltro(6).Enabled = True
            'If (Not HaySeguridad) Or (HaySeguridad And (UsuarioActual.Asociado = 0)) Then CmbFiltro(7).Enabled = True
'                CmbFiltro(7).Enabled = False
'            Else
'                CmbFiltro(7).Enabled = True
'            End If
        Case 2
            'CmbFiltro(3).Enabled = False
            'CmbFiltro(4).Enabled = False
            'CmbFiltro(5).Enabled = False
            CmbFiltro(5).Enabled = True
            'CmbFiltro(6).Enabled = True
            CmbFiltro(6).Enabled = False
            'CmbFiltro(7).Enabled = False
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim Bucle As Integer
    Dim B2 As Integer
    If NIndices = 0 Then Exit Sub
    For Bucle = 0 To Lista.ColumnHeaders.Count
        For B2 = 0 To NIndices - 1
            If Bucle = Indices(B2) Then
                LblTot(B2).Left = Lista.ColumnHeaders(Bucle + 1).Left + Lista.Left
                LblTot(B2).Width = Lista.ColumnHeaders(Bucle + 1).Width
                LblDesc(B2).Left = Lista.ColumnHeaders(Bucle + 1).Left + Lista.Left
                LblDesc(B2).Width = Lista.ColumnHeaders(Bucle + 1).Width
            End If
        Next B2
    Next Bucle
End Sub

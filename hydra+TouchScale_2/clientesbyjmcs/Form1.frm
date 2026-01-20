VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Listado de ventas por Cliente"
   ClientHeight    =   5010
   ClientLeft      =   2745
   ClientTop       =   5730
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":1E72
   ScaleHeight     =   5010
   ScaleMode       =   0  'User
   ScaleWidth      =   16880.74
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1800
      Picture         =   "Form1.frx":E176
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado de detalle de tiques por cliente"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "INTERVALO:"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "CLIENTE:"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Picture         =   "Form1.frx":13A7D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Listado de total por Cliente"
      Top             =   3240
      Width           =   495
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "Cliente"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":19384
      Left            =   120
      List            =   "Form1.frx":19386
      TabIndex        =   5
      ToolTipText     =   "Selección cliente o ""TODOS"""
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":19388
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3255
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      WordWrap        =   -1  'True
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   1
      GridLineWidth   =   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   3
   End
   Begin VB.ComboBox Combo2 
      CausesValidation=   0   'False
      DataField       =   "FECHA"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":1ECC8
      Left            =   120
      List            =   "Form1.frx":1ECCA
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Selección Fecha final con Totales o ""TODOS"""
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      DataField       =   "FECHA"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":1ECCC
      Left            =   120
      List            =   "Form1.frx":1ECCE
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Selección Fecha inicial con Totales  o ""TODOS"""
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":1ECD0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Listado totales por intervalo"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   9960
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu MnuListado 
      Caption         =   "&Listado"
      Index           =   0
      NegotiatePosition=   1  'Left
      Begin VB.Menu MnuIntervalo 
         Caption         =   "&General Intervalo"
      End
      Begin VB.Menu MnuPorcliente 
         Caption         =   "&Tiques por cliente"
         Index           =   0
      End
      Begin VB.Menu MnuUncliente 
         Caption         =   "&Total Cliente"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

'oculto los botones de impresión, porque los he puesto en menú
Command1.Visible = False
Command2.Visible = False
Command6.Visible = False
    
On Error GoTo Error
       
Set MyBase = OpenDatabase(App.Path & "\dbase.mdb")
Set MyRe = MyBase.OpenRecordset("SELECT DISTINCT FECHA from cabecera WHERE (cabecera.CODCLI <> '0')")
Set MyCli = MyBase.OpenRecordset("SELECT DISTINCT Cod_Cli, Nom_Cli FROM Cliente")
'Se usa SELECT DISTINCT para que no se repitan fechas y clientes....
'Si no hay ventas, lanzamos ventana de mensajes con advertencia
    
If MyRe.RecordCount = 0 Then
    MsgBox "No hay ventas almacenadas en Base de Datos", vbExclamation, "Mensaje"
    MSHFlexGrid1.Visible = False
    GoTo Error
End If
    
'Si no hay clientes programados, lanzamos ventana de mensajes con advertencia
If MyCli.RecordCount = 0 Then
    MsgBox "No se han programado Clientes", vbExclamation, "Mensaje"
    MSHFlexGrid1.Visible = False
End If
    

'Llenamos los combo con las fechas en las que ha habido ventas....
        

If Not MyRe.EOF Then
    With MyRe
        MyRe.MoveFirst
        Combo1.AddItem ("TODOS")
        Combo2.AddItem ("TODOS")
        Do Until MyRe.EOF
            Combo1.AddItem MyRe!Fecha
            Combo2.AddItem MyRe!Fecha
            MyRe.MoveNext
        Loop
    End With
End If
        
'Combo1.ListIndex = 0 'primera fecha
Combo1.ListIndex = -1
Combo1.Text = "Fecha Inicial"
       
'combo2.ListIndex = Combo2.ListCount - 2 'última fecha
 Combo2.ListIndex = -1
 Combo2.Text = "Fecha Final"
 MyRe.Close
 Set MyRe = Nothing
     
'llenamos el combo con clientes en base de datos

If Not MyCli.EOF Then
    With MyCli
        MyCli.MoveFirst
        Combo4.AddItem ("TODOS")
        Do Until MyCli.EOF
            If IsNull(MyCli!Cod_Cli) = True Or MyCli!Cod_Cli = 0 Then 'si cod_cli es nulo, me salto el registro
                MyCli.MoveNext
            Else
                clinom = MyCli!Cod_Cli & " - " & MyCli!Nom_cli 'asigno a la variable clinom el valor del campo cod_cli y nom_cli
                Combo4.AddItem clinom 'añado un item al combo
                Combo4.ItemData(Combo4.NewIndex) = MyCli!Cod_Cli 'asigno a la propiedad itemdata el valor del campo cod_cli
                MyCli.MoveNext
                clinom = "" 'vacio la variable clinom
            End If
        Loop
    End With
End If
        
Combo4.ListIndex = -1
Combo4.Text = "Selección de cliente"
MyCli.Close
Set MyCli = Nothing
MyBase.Close
Set MyBase = Nothing

' Llamamos a la función query pasando la cadena sql1 con un select de lo que nos interesa meter en el FlexGrid

query ("SELECT tickets.cliente AS Cliente, Cliente.Nom_Cli AS Texto,Cdate(tickets.`d_fecha`) AS FECHA, Cstr(SUM(tickets.AMOUNT)) AS Importe, Cstr(SUM(tickets.WEIGHT)) AS `Total Peso`, SUM(tickets.UNITS) AS Unidades,COUNT(tickets.UNITS) AS OPER FROM Cliente, tickets WHERE Cliente.Cod_Cli = tickets.cliente GROUP BY tickets.cliente, Cliente.Nom_Cli, Cdate(tickets.`d_fecha`)")

Error:
End Sub


Private Sub Combo1_Click()
fecha1 = -1
If Combo1.ListIndex <> fecha1 Or IsEmpty(fecha1) Then
    fecha1 = Combo1.ListIndex
    If fecha1 <> -1 Then
        query ("SELECT tickets.cliente AS Cliente, Cliente.Nom_Cli AS Texto,Cdate(tickets.d_fecha) AS FECHA, Cstr(SUM(tickets.AMOUNT))AS Importe, Cstr(SUM(tickets.WEIGHT)) AS `Total Peso`,SUM(tickets.UNITS) AS Unidades, COUNT(tickets.UNITS)AS OPER From cliente, tickets WHERE (Cliente.Cod_Cli = tickets.cliente AND Cliente.Cod_cli <> 0) GROUP BY tickets.cliente, Cliente.Nom_Cli, Cdate(tickets.d_fecha)")
    Else
    MsgBox "La fecha seleccionada no es valida"
    End If
Else
 
End If
Form1.Caption = "Listado de ventas por Cliente"
  
End Sub

Private Sub Combo2_Click()
fecha1 = -1
If Combo2.ListIndex <> fecha1 Or IsEmpty(fecha1) Then
    fecha1 = Combo2.ListIndex
    If fecha1 <> -1 Then
           query ("SELECT tickets.cliente AS Cliente, Cliente.Nom_Cli AS Texto,Cdate(tickets.d_fecha) AS FECHA, Cstr(SUM(tickets.AMOUNT))AS Importe, Cstr(SUM(tickets.WEIGHT)) AS `Total Peso`,SUM(tickets.UNITS) AS Unidades, COUNT(tickets.UNITS)AS OPER From cliente, tickets WHERE (Cliente.Cod_Cli  = tickets.cliente AND Cliente.Cod_Cli <> 0) GROUP BY tickets.cliente, Cliente.Nom_Cli, Cdate(tickets.d_fecha)")
    Else
    MsgBox "La fecha seleccionada no es valida"
    End If
Else
End If
Form1.Caption = "Listado de ventas por Cliente"
  
End Sub

Private Sub Combo4_Click()
 codcli = -1
    If Combo4.ListIndex <> codcli And Combo4.ListIndex <> 0 Then
        codcli = Combo4.ListIndex
        query ("SELECT cliente as Cliente, NUME as Número, NUMLIN as Línea, importe_bruto as Total, d_fecha as Fecha, Weight as Peso, UNITS as UND, VENDOR as Vendedor, PRICE as Precio, d_hora as Hora, CODE as Código, Item as Texto From tickets Where (cliente =   " & codcli & ")ORDER BY cliente, d_fecha, NUME")
        Form1.Caption = "Listado de ventas por Cliente  " & codcli
    Else
        codcli = Combo4.ListIndex
        'MsgBox "El nuevo cliente elegido es  " & codcli & " "
        query ("SELECT cliente as Cliente, NUME as Número, NUMLIN as Línea, importe_bruto as Total, d_fecha as Fecha, Weight as Peso, UNITS as UND, VENDOR as Vendedor, PRICE as Precio, d_hora as Hora, CODE as Código, Item as Texto From tickets where (cliente > 0) ORDER BY cliente, d_fecha, NUME")
        Form1.Caption = "Listado de ventas por Cliente  " & codcli & "  (TODOS)"
    End If
 


End Sub

Private Sub Command1_Click()
    
Dim a As Date
Dim b As String
On Error GoTo Error

If Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "Selecciona una fecha o TODOS", vbInformation, "Selecciona una Fecha"
Else
End If
        
If CStr(Combo1.Text) = "TODOS" Or CStr(Combo2.Text) = "TODOS" Or CStr(Combo1.Text) = "Fecha Inicial" Or CStr(Combo2.Text) = "Fecha Final" Then
    a = #1/1/1100#
    b = "99/99/9999"
Else
    a = CDate(Combo1.Text)
    b = CDate(Combo2.Text)
    a = Format(a, "dd/mm/yyyy")
    b = Format(b, "dd/mm/yyyy")
End If
        
If DataEnvironment1.rsCommand2.State = adStateOpen Then
    DataEnvironment1.rsCommand2.Close
Else
End If
        
DataEnvironment1.Command2 a, b
DataReport1.Refresh
Load DataReport1
DataReport1.Show

Error:
Exit Sub
End Sub

Private Sub Command2_Click()
Dim a As String
Dim b As String
Combo4.Refresh
'On Error GoTo Error
        
If Combo4.Text = "" Or Combo4.Text = "Selección de cliente" Then
    MsgBox "Selecciona un Cliente o TODOS", vbInformation, "Selecciona un Cliente"
Else
    If Combo4.Text = "TODOS" Then
        a = 1
        b = 99999
        If DataEnvironment1.rsCommand1.State = adStateOpen Then
            DataEnvironment1.rsCommand1.Close
        Else
        End If
    Else
        a = Combo4.ItemData(Combo4.ListIndex)
        b = Combo4.ItemData(Combo4.ListIndex)
        If DataEnvironment1.rsCommand1.State = adStateOpen Then
            DataEnvironment1.rsCommand1.Close
        Else
        End If
    
    End If
    DataEnvironment1.Command1 a, b
    DataReport3.Refresh
    Load DataReport3
    DataReport3.Show
End If
   

If DataEnvironment1.rsCommand1.State = adStateOpen Then
    DataEnvironment1.rsCommand1.Close
Else
End If

Error:
Exit Sub
End Sub

'Private Sub Command2_Click()
'       If DataEnvironment1.rsCommand2.State = adStateOpen Then
'        DataEnvironment1.rsCommand2.Close
'        End If
'        Dim a As Date
'        Dim b As Date
'        a = #1/1/1100#
'        b = Format(Now, ddmmyyyy)
'      DataEnvironment1.Command2 a, b
'    Load DataReport1
'    DataReport1.Show
'End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Command6_Click()
Dim a As String
Dim b As String
Combo4.Refresh
On Error GoTo Error
        
If Combo4.Text = "" Or Combo4.Text = "Selección de cliente" Then
    MsgBox "Selecciona un Cliente o TODOS", vbInformation, "Selecciona un Cliente"
Else
End If

If Combo4.Text = "TODOS" Then
    a = 1
    b = 99999
Else
    a = Combo4.ItemData(Combo4.ListIndex)
    b = Combo4.ItemData(Combo4.ListIndex)
End If
     
If DataEnvironment1.rsCommand3.State = adStateOpen Then
    DataEnvironment1.rsCommand3.Close
Else
End If
DataEnvironment1.Command3 a, b
DataReport2.Refresh
Load DataReport2
DataReport2.Show
    
DataEnvironment1.rsCommand3.Close

Error:
End Sub

' Manejando ADO montamos esta función para llenar el MSHFlexgrid...
'Private Sub Cargar_FlexGrid(path_Base As String, Consulta_SQL As String, FlexGrid As MSHFlexGrid)
Private Function query(sql1 As String)

Set BD = New ADODB.Connection
BD.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase.mdb"
Set rs = New Recordset
rs.CursorLocation = adUseClient
rs.Open sql1, BD, adOpenDynamic, adLockOptimistic

Set MSHFlexGrid1.DataSource = rs
    
On Error GoTo Error
                          
' -- Variables para las filas y columnas del Flexgrid
Dim Columna As Integer, Fila As Integer
' -- Matriz y variable para los anchos de columna
Dim Ancho_Columna() As Single, Ancho_Campo As Single
       
' -- Configurar el FlexGrid
With MSHFlexGrid1
' -- Deshabilitar el repintado del control ( Para que la carga sea mas veloz )
    .Redraw = False
' -- Seleccionar registros del Grid por Fila
        
    .SelectionMode = flexSelectionByRow
' -- Modo de encabezados
    .FixedRows = 1
    .FixedCols = 0
      
' -- Cantidad de filas y columnas
    .Rows = 1
    .Cols = rs.Fields.Count
      
' -- Redimensionar el Array a la cantidad de campos de la tabla
    ReDim Ancho_Columna(0 To rs.Fields.Count - 1)
      
' -- Recorrer los campos del recordset
    For Columna = 0 To rs.Fields.Count - 1
    ' -- Añade el título del campo al encabezado de columna
        .TextMatrix(0, Columna) = rs.Fields(Columna).Name
    ' -- Guardar el ancho del campo en la matriz
         Ancho_Columna(Columna) = TextWidth(rs.Fields(Columna).Name)
    Next Columna
          
    Fila = 1
' -- Recorrer todos los registros del recordset
    Do While Not rs.EOF
        .Rows = .Rows + 1 ' Añade una nueva fila
        For Columna = 0 To rs.Fields.Count - 1
                            
' -- Combobar que el valor no es nulo
        If Not IsNull(rs.Fields(Columna).Value) Then
' -- Agrega el registro en la fila y columna específica
            .TextMatrix(Fila, Columna) = rs.Fields(Columna).Value
' -- Almacena el ancho
            Ancho_Campo = TextWidth(rs.Fields(Columna).Value)
            If Ancho_Columna(Columna) < Ancho_Campo Then
                Ancho_Columna(Columna) = Ancho_Campo
            Else
            End If
        End If
' -- Establece los ancho de columna
        .ColWidth(Columna) = Ancho_Columna(Columna) + 300
        .FixedRows = 1
'colores de Fila fija y Background del flex
        .BackColorFixed = &H80000001
        .BackColorBkg = &H80000001
     Next
' -- Siguiente registro
     rs.MoveNext
     Fila = Fila + 1 'Incrementa la fila
Loop
' -- Volver a Habilitar el repintado del Grid
        
    .Redraw = True

' -- Cierra el recordset y la conexión abierta
        
    rs.Close
    BD.Close
    Set rs = Nothing
    Set BD = Nothing

End With
Error:
Exit Function
'-------------------------------------------------------------------------------------
End Function


Private Sub Form_Terminate()
On Error Resume Next
rs.Close
BD.Close
MyRe.Close
MyCli.Close
MyBase.Close
'MsgBox "Me cierran...", vbExclamation

On Error GoTo 0

End Sub

Private Sub Image1_Click()

Dim abreweb As Long
abreweb = ShellExecute(Me.hwnd, "Open", "http://www.grupoepelsa.com", "", "", 1)

End Sub


Private Sub MnuIntervalo_Click()

Dim a As Date
Dim b As String
On Error GoTo Error

If Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "Selecciona una fecha o TODOS", vbInformation, "Selecciona una Fecha"
Else
End If
        
If CStr(Combo1.Text) = "TODOS" Or CStr(Combo2.Text) = "TODOS" Or CStr(Combo1.Text) = "Fecha Inicial" Or CStr(Combo2.Text) = "Fecha Final" Then
    a = #1/1/1100#
    b = "99/99/9999"
Else
    a = CDate(Combo1.Text)
    b = CDate(Combo2.Text)
    a = Format(a, "dd/mm/yyyy")
    b = Format(b, "dd/mm/yyyy")
End If
        
If DataEnvironment1.rsCommand2.State = adStateOpen Then
    DataEnvironment1.rsCommand2.Close
Else
End If
        
DataEnvironment1.Command2 a, b
DataReport1.Refresh
Load DataReport1
DataReport1.Show

Error:
Exit Sub

End Sub

Private Sub MnuPorcliente_Click(Index As Integer)

Dim a As String
Dim b As String
Combo4.Refresh
'On Error GoTo Error
        
If Combo4.Text = "" Or Combo4.Text = "Selección de cliente" Then
    MsgBox "Selecciona un Cliente o TODOS", vbInformation, "Selecciona un Cliente"
Else
    If Combo4.Text = "TODOS" Then
        a = 1
        b = 99999
        If DataEnvironment1.rsCommand1.State = adStateOpen Then
            DataEnvironment1.rsCommand1.Close
        Else
        End If
    Else
        a = Combo4.ItemData(Combo4.ListIndex)
        b = Combo4.ItemData(Combo4.ListIndex)
        If DataEnvironment1.rsCommand1.State = adStateOpen Then
            DataEnvironment1.rsCommand1.Close
        Else
        End If
    
    End If
    DataEnvironment1.Command1 a, b
    DataReport3.Refresh
    Load DataReport3
    DataReport3.Show
End If
   

If DataEnvironment1.rsCommand1.State = adStateOpen Then
    DataEnvironment1.rsCommand1.Close
Else
End If

Error:
Exit Sub

End Sub

Private Sub MnuUncliente_Click()

Dim a As String
Dim b As String
Combo4.Refresh
On Error GoTo Error
        
If Combo4.Text = "" Or Combo4.Text = "Selección de cliente" Then
    MsgBox "Selecciona un Cliente o TODOS", vbInformation, "Selecciona un Cliente"
Else
End If

If Combo4.Text = "TODOS" Then
    a = 1
    b = 99999
Else
    a = Combo4.ItemData(Combo4.ListIndex)
    b = Combo4.ItemData(Combo4.ListIndex)
End If
     
If DataEnvironment1.rsCommand3.State = adStateOpen Then
    DataEnvironment1.rsCommand3.Close
Else
End If
DataEnvironment1.Command3 a, b
DataReport2.Refresh
Load DataReport2
DataReport2.Show
    
DataEnvironment1.rsCommand3.Close

Error:

End Sub




' Ordena al hacer clic en el encabezado de columna

Private Sub MSHFlexGrid1_Click()
        Static Modo  As Boolean
        If (MSHFlexGrid1.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If Modo Then
            MSHFlexGrid1.Col = MSHFlexGrid1.MouseCol
            MSHFlexGrid1.Sort = 2

            Modo = False
        ' Ordena en forma descendente
        Else
            MSHFlexGrid1.Col = MSHFlexGrid1.MouseCol
            MSHFlexGrid1.Sort = 1
            Modo = True
        End If
    End If
        Form1.Caption = "Listado de ventas por Cliente Ordenado por la columna:  " & MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)
End Sub

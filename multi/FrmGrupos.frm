VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmGrupos 
   Caption         =   "Grupos"
   ClientHeight    =   3930
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "FrmGrupos.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   2
      Text            =   "BAL.DAT"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "BAL.DAT"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "tiendas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "grupos"
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmGrupos.frx":1E72
      Height          =   2775
      Left            =   1200
      OleObjectBlob   =   "FrmGrupos.frx":1E86
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "FrmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Dim dB As Database
   Dim Rst As Recordset
   Set dB = OpenDatabase(App.Path & "\tiendas.mdb")
   Set Rst = dB.OpenRecordset("select * from grupos where grp=" & Format(Data1.Recordset.Fields(0)) & " order by tienda")
   With Rst
      .MoveFirst
      Do While Not Rst.EOF
         .Edit
         !Fichero = Trim(Text1.Text)
         !Patron = Trim(Text2.Text)
         .Update
         .MoveNext
      Loop
   End With
   dB.Close
   Data1.Refresh
   DBGrid1.Refresh
End Sub



Private Sub DBGrid1_Click()
Data1.Refresh
Text1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)

End Sub

Private Sub Form_Load()
Dim dB As Database
Dim Rst As Recordset

FrmGrupos.Caption = CargaCadena(1286)
FrmGrupos.DBGrid1.Columns(0).Caption = CargaCadena(1289)
FrmGrupos.DBGrid1.Columns(1).Caption = CargaCadena(1290)
FrmGrupos.DBGrid1.Columns(2).Caption = CargaCadena(1291)
FrmGrupos.Command1.Caption = CargaCadena(36)

Set dB = OpenDatabase(App.Path & "\tiendas.mdb")
Set Rst = dB.OpenRecordset("select * from grupos" & " order by tienda")
If Rst.EOF Then
   
   
   Exit Sub
End If


Data1.RecordSource = "select distinct(grp),fichero,patron from grupos"
Data1.Refresh
Text1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DescargaForm = True
End Sub


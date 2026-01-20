VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   12495
   Begin VB.Timer TmrInicio 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11400
      Top             =   240
   End
   Begin VB.ListBox ListTiCli 
      Height          =   3960
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   10455
   End
   Begin VB.ComboBox CmbClientes 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SalirPrograma()
    End
End Sub

Private Sub CmbClientes_Click()
Dim bBase As Database
Dim nRs As Recordset
    ListTiCli.Visible = False
    ListTiCli.Clear
    Set bBase = OpenDatabase(App.Path & "\dbase.mdb")
    Set nRs = bBase.OpenRecordset("select * from cabecera where codcli= '" & CStr(Val(CmbClientes.Text)) & "' order by nume")
    If Not nRs.EOF Then
        nRs.MoveFirst
        Do While Not nRs.EOF
            ListTiCli.AddItem CStr(nRs.Fields("nume"))
            nRs.MoveNext
        Loop
    
    End If
    nRs.Close
    bBase.Close
    Set nRs = Nothing
    Set bBase = Nothing
    ListTiCli.Visible = True
End Sub


Private Sub Form_Load()
Dim bBase As Database
Dim nRs As Recordset
    'Carga de datos en Combo
    ListTiCli.Visible = False
    Set bBase = OpenDatabase(App.Path & "\dbase.mdb")
    Set nRs = bBase.OpenRecordset("select distinct codcli from cabecera order by codcli")
    If Not nRs.EOF Then
        nRs.MoveFirst
        Do While Not nRs.EOF
           If nRs.Fields("codcli") <> 0 Then
               CmbClientes.AddItem CStr(nRs.Fields("codcli"))
           End If
           nRs.MoveNext
        Loop
        If CmbClientes.ListCount > 0 Then
            CmbClientes.ListIndex = 0
        Else
            MsgBox "No existen Clientes..."
            Call SalirPrograma
        End If
    Else
        MsgBox "No existen Clientes..."
        Call SalirPrograma
    End If
    nRs.Close
    bBase.Close
    Set nRs = Nothing
    Set bBase = Nothing
    TmrInicio.Enabled = True
End Sub

Private Sub TmrInicio_Timer()
Dim bBase As Database
Dim nRs As Recordset

    TmrInicio.Enabled = False
    If ListTiCli.Visible = False Then
        Set bBase = OpenDatabase(App.Path & "\dbase.mdb")
        Set nRs = bBase.OpenRecordset("select * from cabecera where codcli= '" & CStr(Val(CmbClientes.Text)) & "' order by nume")
        If Not nRs.EOF Then
            nRs.MoveFirst
            Do While Not nRs.EOF
                ListTiCli.AddItem CStr(nRs.Fields("nume"))
                nRs.MoveNext
            Loop
        
        End If
        nRs.Close
        bBase.Close
        Set nRs = Nothing
        Set bBase = Nothing
        ListTiCli.Visible = True
    End If
End Sub

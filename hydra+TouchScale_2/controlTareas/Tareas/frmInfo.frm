VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   Caption         =   "Tareas"
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   12030
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar Proceso."
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listar Procesos."
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' botón que enumera y lista los procesos _
 abiertos en el control Listview
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim Flag_Process As Long
Dim ret As Long
Dim Nombre As String

    ListView1.ListItems.Clear
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    proc.dwSize = Len(proc)
    Flag_Process = ProcessFirst(snap, proc)
    i = 0
    
    ' Bucle que enumera los procesos abiertos en windows
    While Flag_Process <> 0
        Nombre = proc.szExeFile ' nombre del proceso
        ListView1.ListItems.Add , "first" & CStr(i), Nombre ' loa grega
        ' agrega el número de id de proceso
        ListView1.ListItems("first" & CStr(i)).SubItems(1) = proc.th32ProcessID ' id
        i = i + 1
        ' siguiente proceso abierto
        Flag_Process = ProcessNext(snap, proc)
    Wend
    
    Call CloseHandle(snap)
End Sub
'Botón que Cierra el process indicado
Private Sub Command2_Click()
Dim i As Long
Dim hand As Long
    ' Obtiene el handle del proceso a partir de su id ( Valor de la segunda columna del LV)
    hand = OpenProcess(PROCESS_TERMINATE, True, CLng(ListView1.SelectedItem.SubItems(1)))
    ' le pasa el handle del proceso
    Call TerminateProcess(hand, 0)
    ' vuelve a listar los procesos
    Command1_Click
End Sub

Private Sub Form_Load()

Dim header As ColumnHeader
        
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    ' agrega dos columnas ( Proceso e Id de proceso)
    Set header = ListView1.ColumnHeaders.Add(, "first", "Nombre del Proceso", 4200)
    Set header = ListView1.ColumnHeaders.Add(, "second", "ID de Proceso", 1200)
    
    Command1.Caption = "Listar procesos"
    Command2.Caption = "Terminar proceso"
End Sub



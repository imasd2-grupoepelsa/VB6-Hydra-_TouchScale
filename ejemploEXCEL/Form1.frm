VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "EnlaExcel"
   ClientHeight    =   7845
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   12720
   Begin MSDBGrid.DBGrid DBGridLibro 
      Bindings        =   "Form1.frx":0000
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   8
      Top             =   480
      Width           =   12615
   End
   Begin VB.TextBox TextHoja 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "hoja1$"
      Top             =   7560
      Width           =   2655
   End
   Begin VB.ComboBox ComboVersion 
      Height          =   315
      ItemData        =   "Form1.frx":09EB
      Left            =   1680
      List            =   "Form1.frx":09FB
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin VB.FileListBox FileXls 
      Height          =   3015
      Left            =   3120
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.DirListBox DirXls 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   5880
      Y1              =   3240
      Y2              =   7800
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label LabelHoja 
      Caption         =   "Nombre Hoja:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Lblversion 
      Caption         =   "Versión Excel:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Lblxls 
      Caption         =   $"Form1.frx":0A2B
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   5775
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub DirXls_Change()
    FileXls.Path = DirXls.Path
    FileXls.Pattern = "*.xls"
    FileXls.Refresh

End Sub

Private Sub FileXls_Click()
    Dim sS As String
    Dim nErr As Long
    Dim sErr As String
    
    On Error GoTo TrataError
    Data1.Connect = Trim(ComboVersion.Text) & ";"
    Data1.DatabaseName = DirXls.Path & "\" & FileXls.FileName
    sS = Trim(TextHoja.Text)
    If Right(sS, 1) <> "$" Then
        sS = sS & "$"
    End If
    Data1.RecordSource = sS
    Data1.Refresh
    DBGridLibro.Refresh
    
TrataError:
    nErr = Err.Number
    sErr = Err.Description
    On Error GoTo 0
    If nErr <> 0 Then
    
        MsgBox "Error al Tratar el Archivo: " & DirXls.Path & "\" & FileXls.FileName & _
            Chr(13) & "Descripción: " & sErr & _
            Chr(13) & "Número de error: " & CStr(nErr), vbCritical
        
    Else
        Label1.Caption = "Vista de Fichero: " & DirXls.Path & "\" & FileXls.FileName
    End If
End Sub

Private Sub Form_Load()

    FileXls.Pattern = "*.xls"
    ComboVersion.ListIndex = 0
    DirXls.Path = App.Path
    
End Sub



Private Sub TextHoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FileXls.SetFocus
    End If
End Sub

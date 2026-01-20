VERSION 5.00
Begin VB.Form frmMARCAR 
   Caption         =   "Marcado De Artículos para enviar."
   ClientHeight    =   3090
   ClientLeft      =   3945
   ClientTop       =   1770
   ClientWidth     =   8895
   Icon            =   "frmMARCAR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DataPorFechaX 
      Caption         =   "DataPorFechaX"
      Connect         =   "Access"
      DatabaseName    =   "dbasetouch.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CommandEnvio 
      Caption         =   "Envio Modificaciones."
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CommandBaja 
      Caption         =   "Marcar Baja de Artículos."
      Height          =   495
      Left            =   6720
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox ComboPorSecSubFam 
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox ComboPorSeccion 
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Commandm 
      Caption         =   "Marcar para Envio"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Textn 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.Data DataPorFecha 
      Caption         =   "DataPorFecha"
      Connect         =   "Access"
      DatabaseName    =   "dbase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select distinct prc3 from articulo"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox ComboPorFecha 
      Height          =   315
      ItemData        =   "frmMARCAR.frx":000C
      Left            =   360
      List            =   "frmMARCAR.frx":000E
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Line Line6 
      X1              =   5280
      X2              =   9840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   5160
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   5280
      X2              =   9840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   5160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   5160
      X2              =   5160
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Label Labelz 
      Caption         =   "BAJAS."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Labelz 
      Caption         =   "MODIFICACIONES."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Labelx 
      Caption         =   "Sección - Subsección - Familia "
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Labely 
      Caption         =   "Por Número de Sección."
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Labeln 
      Caption         =   "Número de artículos:"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label LabelPorFecha 
      Caption         =   "Por Fecha De Modificación."
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmMARCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboPorFecha_GotFocus()
    Dim nMin As Long
    Dim nMax As Long
    Dim s As Variant
On Error GoTo FIN
    s = ComboPorFecha.TexT  'ComboPorFecha.Index

    s = Format(s, "yymmdd")
   
    nMin = CDbl(s) * 10000
    nMax = nMin + 2359
    DataPorFechaX.RecordSource = "select * from articulo where (prc3>=" & CStr(nMin) & " and prc3<=" & CStr(nMax) & ")"
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Por Fecha:"
Exit Sub
FIN:
On Error GoTo 0
End Sub

Private Sub ComboPorSeccion_GotFocus()
    Dim s As Variant

On Error GoTo FIN

    s = ComboPorSeccion.TexT  'ComboPorFecha.Index

    DataPorFechaX.RecordSource = "select * from articulo where codi_ident=" & s
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Por Sección:"

Exit Sub
FIN:
On Error GoTo 0

End Sub

Private Sub ComboPorSecSubFam_Click()
    Dim s As Variant
    
    s = ComboPorSecSubFam.TexT  'ComboPorFecha.Index

    DataPorFechaX.RecordSource = "select * from articulo where codi_ident=" & Mid(s, 1, 2) & " and codi_sub=" & Mid(s, 6, 2) & " and codi_fam=" & Mid(s, 11, 2)
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Para Baja:"
    
End Sub

Private Sub ComboPorSeccion_Click()
    Dim s As Variant
    
    s = ComboPorSeccion.TexT  'ComboPorFecha.Index

    DataPorFechaX.RecordSource = "select * from articulo where codi_ident=" & s
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Por Sección:"
    
End Sub

Private Sub ComboPorFecha_Click()
    Dim nMin As Long
    Dim nMax As Long
    Dim s As Variant
    
    s = ComboPorFecha.TexT  'ComboPorFecha.Index

    s = Format(s, "yymmdd")
   
    nMin = CDbl(s) * 10000
    nMax = nMin + 2359
    DataPorFechaX.RecordSource = "select * from articulo where (prc3>=" & CStr(nMin) & " and prc3<=" & CStr(nMax) & ")"
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Por Fecha:"
    
End Sub

Private Sub ComboPorsecsubfam_DblClick()
    Call ComboPorSecSubFam_Click
End Sub
Private Sub ComboPorFecha_DblClick()
    Call ComboPorFecha_Click
End Sub
Private Sub ComboPorSeccion_DblClick()
    Call ComboPorSeccion_Click
End Sub

Private Sub ComboPorSecSubFam_GotFocus()
    Dim s As Variant
On Error GoTo FIN

    s = ComboPorSecSubFam.TexT  'ComboPorFecha.Index

    DataPorFechaX.RecordSource = "select * from articulo where codi_ident=" & Mid(s, 1, 2) & " and codi_sub=" & Mid(s, 6, 2) & " and codi_fam=" & Mid(s, 11, 2)
    DataPorFechaX.Refresh
    DataPorFechaX.Recordset.MoveFirst
    DataPorFechaX.Recordset.MoveLast
    s = DataPorFechaX.Recordset.Recordcount
    Textn.TexT = CStr(s)
    Textn.Refresh
    
    Labeln.Caption = "Número de Artículos Para Baja:"
Exit Sub
FIN:
On Error GoTo 0

End Sub

Private Sub CommandBaja_Click()
    Dim n As Long
    CommandBaja.Enabled = False
    ComboPorFecha.Enabled = False
    ComboPorSeccion.Enabled = False
    ComboPorSecSubFam.Enabled = False
    Commandm.Enabled = False
    
    n = 0
    Textn.TexT = CStr(n)
    If Not DataPorFechaX.Recordset.EOF Then
    
        DataPorFechaX.Recordset.MoveFirst
        Do While Not DataPorFechaX.Recordset.EOF
            DataPorFechaX.Recordset.Edit
            
            DataPorFechaX.Recordset.Fields("borrado") = True
            DataPorFechaX.Recordset.Fields("tran_plu") = " "
            DataPorFechaX.Recordset.Fields("tran_plusc10") = " "
            DataPorFechaX.Recordset.Fields("tran_el") = " "
    
            DataPorFechaX.Recordset.Update
            n = n + 1
            Textn.TexT = CStr(n)
            Textn.Refresh
            Sleep (2)
            DataPorFechaX.Recordset.Movenext
        Loop
    
    End If
    
    CommandBaja.Enabled = True
    'CommandBaja.SetFocus

    ComboPorFecha.Enabled = True
    ComboPorSeccion.Enabled = True
    ComboPorSecSubFam.Enabled = True
    Commandm.Enabled = True

End Sub

'Private Sub CommandEnvio_Click()
'frmEpelsa.mnuEnvio_modificaciones_Click

'End Sub

Private Sub Commandm_Click()
    Dim n As Long
    Commandm.Enabled = False
    CommandBaja.Enabled = False
    
    ComboPorFecha.Enabled = False
    ComboPorSeccion.Enabled = False
    ComboPorSecSubFam.Enabled = False
    
    n = 0
    Textn.TexT = CStr(n)
    If Not DataPorFechaX.Recordset.EOF Then
    
        DataPorFechaX.Recordset.MoveFirst
        Do While Not DataPorFechaX.Recordset.EOF
            DataPorFechaX.Recordset.Edit
            DataPorFechaX.Recordset.Fields("tran_plu") = " "
            DataPorFechaX.Recordset.Fields("tran_plusc10") = " "
            DataPorFechaX.Recordset.Update
            n = n + 1
            Textn.TexT = CStr(n)
            Textn.Refresh
            Sleep (2)
            DataPorFechaX.Recordset.Movenext
        Loop
    
    End If
    Commandm.Enabled = True
    'Commandm.SetFocus
    
    ComboPorFecha.Enabled = True
    ComboPorSeccion.Enabled = True
    ComboPorSecSubFam.Enabled = True
    
    CommandBaja.Enabled = True
    
End Sub


Private Sub Form_Load()
    Dim n As Long
    Dim s As String
    Dim nx As Double
    
    nx = CDbl(Format(Now - 15, "yymmddhhmm"))
    
    n = 0
    DataPorFechaX.RecordSource = "select distinct codi_ident,codi_sub,codi_fam from articulo"
    DataPorFechaX.Refresh
    If Not DataPorFechaX.Recordset.EOF Then
        DataPorFechaX.Recordset.MoveFirst
        Do While Not DataPorFechaX.Recordset.EOF
            s = Format(DataPorFechaX.Recordset.Fields("codi_ident"), "00") & " - " & Format(DataPorFechaX.Recordset.Fields("codi_sub"), "00") & " - " & Format(DataPorFechaX.Recordset.Fields("codi_fam"), "00")
            ComboPorSecSubFam.AddItem s, n
            n = n + 1
            DataPorFechaX.Recordset.Movenext
        Loop
        ComboPorSecSubFam.ListIndex = 0
    End If
    
    n = 0
    DataPorFechaX.RecordSource = "select distinct codi_ident from articulo"
    DataPorFechaX.Refresh
    If Not DataPorFechaX.Recordset.EOF Then
        DataPorFechaX.Recordset.MoveFirst
        Do While Not DataPorFechaX.Recordset.EOF
            s = Format(DataPorFechaX.Recordset.Fields("codi_ident"), "00")
            ComboPorSeccion.AddItem s, n
            n = n + 1
            DataPorFechaX.Recordset.Movenext
        Loop
        ComboPorSeccion.ListIndex = 0
    End If
    
    n = 0
    DataPorFechaX.RecordSource = "select distinct val(prc3/10000) as prcx from articulo where (prc3 > " & CStr(nx) & ")"
    DataPorFechaX.Refresh
    If Not DataPorFechaX.Recordset.EOF Then
        DataPorFechaX.Recordset.MoveFirst
        Do While Not DataPorFechaX.Recordset.EOF
            s = Format(DataPorFechaX.Recordset.Fields("prcx"), "000000")
            s = Right(s, 2) & "/" & Mid(s, 3, 2) & "/" & "20" & left(s, 2)
            ComboPorFecha.AddItem s, n
            n = n + 1
            DataPorFechaX.Recordset.Movenext
        Loop
        ComboPorFecha.ListIndex = 0
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'DataPorFechaX.Recordset.Close
    Unload Me
    frmEpelsa.Enabled = True
    frmEpelsa.SetFocus
End Sub

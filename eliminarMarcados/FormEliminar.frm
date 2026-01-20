VERSION 5.00
Begin VB.Form FormEliminar 
   Caption         =   "Marcar /Eliminar Articulos Pendientes de envio."
   ClientHeight    =   2010
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   7200
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar Artículos Marcados para enviar."
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton CmdMarcar 
      Caption         =   "Marcar Como Enviados."
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "FormEliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sB As Database
Dim nR As Recordset

Private Sub CmdEliminar_Click()
    Do While Not nR.EOF
        Text1.Text = CStr(nR.Fields("codigo")) & " " & CStr(nR.Fields("plu")) & " " & nR.Fields("des_plu1")
        Text1.Refresh
        nR.Edit
        nR.Delete
        
        nR.MoveNext
    Loop

End Sub

Private Sub CmdMarcar_Click()
    Do While Not nR.EOF
        nR.Edit
        nR.Fields("tran_plu") = "*"
        nR.Fields("tran_plusc10") = "*"
        nR.Fields("tran_texto") = "*"
        nR.Fields("tran_textosc10") = "*"
        nR.Fields("tran_tx1") = "*"
        nR.Fields("tran_tx1sc10") = "*"
        nR.Update
        Text1.Text = CStr(nR.Fields("codigo")) & " " & CStr(nR.Fields("plu")) & " " & nR.Fields("des_plu1")
        Text1.Refresh
        nR.MoveNext
    Loop
End Sub

Private Sub Form_Load()

    
    On Error Resume Next
    Set sB = OpenDatabase(App.Path & "\dbase.mdb")
    If Err.Number <> 0 Then GoTo vete1
    On Error GoTo 0
    
    Set nR = sB.OpenRecordset("select * from articulo where (tran_plu<>chr(42) or tran_plu=chr(32) or isnull(tran_plu) or tran_plusc10<>chr(42) or tran_plusc10=chr(32) or isnull(tran_plusc10))")
    
    If nR.EOF Then
        MsgBox "No hay Artículos pendientes de Envio"
        GoTo fin
    End If
    nR.MoveFirst
    Text1.Text = CStr(nR.Fields("codigo")) & " " & CStr(nR.Fields("plu")) & " " & nR.Fields("des_plu1")
    Me.Show
    Do While 1
        Sleep (10)
        DoEvents
        Sleep (10)
    Loop
    
    
vete1:
    Set sB = Nothing
    MsgBox "No existe DBASE.MDB"
    End
    
fin:
    On Error GoTo 0
    Set nR = Nothing
    Set sB = Nothing
    End
End Sub

Private Sub Form_Terminate()
    Set nR = Nothing
    Set sB = Nothing
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set nR = Nothing
    Set sB = Nothing
    End

End Sub

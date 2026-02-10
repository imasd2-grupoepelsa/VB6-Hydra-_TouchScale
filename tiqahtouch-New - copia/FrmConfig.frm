VERSION 5.00
Begin VB.Form FrmConfig 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Información Configuración."
   ClientHeight    =   3465
   ClientLeft      =   5805
   ClientTop       =   4140
   ClientWidth     =   6030
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   3240
      Width           =   6015
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No marcar ""Captured"""
      Height          =   195
      Left            =   0
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CheckBox ChkEncargo 
      Caption         =   "No Recoger Tiques Encargo ni Aparcados"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CheckBox CheckFicha 
      Caption         =   "Ficha Vacuno"
      Height          =   315
      Left            =   4200
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox ChkFPago 
      Caption         =   "Cabecera F.Pago."
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox ChkCli 
      Caption         =   "Cabecera Cli./Lot."
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox CheckMonitor 
      Caption         =   "Visualizar Monitores."
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recoger Tiques sin Exportar."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Tiempo espera Generación Fichero (sg.):"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Hora de Borrado Totales:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Fichero General Tiques:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Ip/Port de la Servidora:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Ruta Ficheros de Tiques:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre Fichero Individual:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Intervalo de recogida Tiques:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check1_Click()
Dim nFich As Integer

    If Check1.Value = vbChecked Then
        nFich = FreeFile()
        Open App.Path & "\recosine.txt" For Output As #nFich
        Close #nFich
    Else
        If Dir(App.Path & "\recosine.txt") <> "" Then
            Kill App.Path & "\recosine.txt"
        End If
    End If
    
End Sub

Private Sub CheckMonitor_Click()
Dim n As Integer
If CheckMonitor.Value Then
    n = FreeFile()
    Open App.Path & "\nsmon.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\nsmon.txt") <> "" Then
        Kill App.Path & "\nsmon.txt"
    End If
End If
End Sub

Private Sub ChkCli_Click()
Dim n As Integer
If ChkCli.Value Then
    n = FreeFile()
    Open App.Path & "\nscli.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\nscli.txt") <> "" Then
        Kill App.Path & "\nscli.txt"
    End If
End If

End Sub

Private Sub ChkEncargo_Click()
Dim n As Integer
If ChkEncargo.Value Then
    n = FreeFile()
    Open App.Path & "\noencargo.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\noencargo.txt") <> "" Then
        Kill App.Path & "\noencargo.txt"
    End If
End If
End Sub

Private Sub Check2_Click()
Dim n As Integer
If Check2.Value Then
    n = FreeFile()
    Open App.Path & "\marcatime" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\marcatime") <> "" Then
        Kill App.Path & "\marcatime"
    End If
End If
End Sub

Private Sub ChkFPago_Click()
Dim n As Integer
If ChkFPago.Value Then
    n = FreeFile()
    Open App.Path & "\nsfpa.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\nsfpa.txt") <> "" Then
        Kill App.Path & "\nsfpa.txt"
    End If
End If

End Sub

Private Sub checkficha_Click()
Dim n As Integer
If CheckFicha.Value Then
    n = FreeFile()
    Open App.Path & "\fichalote.txt" For Output As #n
    Close #n
Else
    If Dir(App.Path & "\fichalote.txt") <> "" Then
        Kill App.Path & "\fichalote.txt"
    End If
End If

End Sub

Private Sub Form_Activate()
Me.Height = 3850
End Sub

Private Sub Form_Load()
    Dim n As Integer
    Dim s As String
    'marcatime
    If Dir(App.Path & "\recosine.txt") <> "" Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    If Dir(App.Path & "\marcatime") <> "" Then
        Check2.Value = vbChecked
    Else
        Check2.Value = vbUnchecked
    End If
    
    Text1.Text = sInftxt1
    Text2.Text = sInftxt2
    Text3.Text = sInftxt3
    Text4.Text = sInftxt4
    Text5.Text = sInftxt5 & "DDM"
    Text6.Text = sInftxt6
    If Dir(App.Path & "\ndplu.txt") <> "" Then
        n = FreeFile()
        Open App.Path & "\ndplu.txt" For Input As #n
        Line Input #n, s
        Close #n
    End If
    Text7.Text = s
    If Dir(App.Path & "\nsmon.txt") <> "" Then
        CheckMonitor.Value = vbChecked
    End If
    If Dir(App.Path & "\nscli.txt") <> "" Then
        ChkCli.Value = vbChecked
    End If
    If Dir(App.Path & "\nsfpa.txt") <> "" Then
        ChkFPago.Value = vbChecked
    End If
    If Dir(App.Path & "\fichalote.txt") <> "" Then
        CheckFicha.Value = vbChecked
    End If
    If Dir(App.Path & "\noencargo.txt") <> "" Then
        ChkEncargo.Value = vbChecked
    End If
    
    Me.Height = 3850
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
   If UnloadMode = 2 Then
       CadenadeLog "Salida: Detectado IPL..."
   Else
       If Cancel = 0 And UnloadMode = 1 Then
            CadenadeLog "MINIMIZADO Por Usuario"
       Else
            CadenadeLog "Salida: Detectado Kill por ejecución KILL( ó administrador Tareas)..."
       End If
   End If
   
   CadenadeLog "Solicitada Salida:" & Format(Cancel, "0000000000") & "     " & Format(UnloadMode, "0000000000")

   
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim n As Integer
    nDelayPluAsc = Val(Trim(Text7.Text))
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        If nDelayPluAsc = 0 Then
            nDelayPluAsc = 2
        End If
    End If
    n = FreeFile()
    Open App.Path & "\ndplu.txt" For Output As #n
    Print #n, CStr(nDelayPluAsc)
    Close #n
End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
    Dim n As Integer
    If KeyAscii < 48 And KeyAscii > 57 And KeyAscii <> 8 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii <> 8 Then
        nDelayPluAsc = Val(Trim(Text7.Text & Chr(KeyAscii)))
        n = FreeFile()
        Open App.Path & "\ndplu.txt" For Output As #n
        Print #n, CStr(nDelayPluAsc)
        Close #n
    End If
End Sub

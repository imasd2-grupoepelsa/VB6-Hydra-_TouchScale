VERSION 5.00
Begin VB.Form frmMonitor 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recogida contínua de Tiquets"
   ClientHeight    =   2520
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4845
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver Configura."
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprimir informe"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
      Height          =   372
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Minimizar"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apagar"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Contraseña Incorrecta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña :"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command5_Click()
    
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    Else
        FrmConfig.Visible = True
        FrmConfig.Show
    End If
    
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
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If

End Sub

'\\\\\\\\\\\\\\
'//////////////
Private Sub Form_Load()
    Me.Icon = frmControl.Icon
    Me.Height = 1200
End Sub

Private Sub Form_DblClick()
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    Unload Me
End Sub
'\\\\\\
'//////
Private Sub Command1_Click()
    Me.Height = 2832
    Me.Text1.SetFocus
    var.modo = 1  '///modo cierre
End Sub
'\\\\\\
'//////
Private Sub Command2_Click()
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    Unload Me
End Sub
'\\\\\\
'//////
Private Sub Command3_Click()
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    If Text1.Text = " epfab" Then
        If var.modo = 1 Then
            frmControl.cSysTray1.InTray = False
            'If Dir(App.Path & "\hydrarun.txt") <> "" Then
            '    Kill App.Path & "\hydrarun.txt"
            'End If
            
            Unload Me
            End
        ElseIf var.modo = 2 Then
            Me.Enabled = False
            Aux_Ejecuta "gtInfVen", "0"
            Me.Enabled = True
            var.modo = 0
            Me.Height = 1200
            Label2.Visible = False
            Text1.Text = ""
        End If
    Else
        Label2.Visible = True
        Text1.Text = ""
    End If
End Sub
'\\\\\\
'//////
Private Sub Command4_Click()
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    Me.Height = 2832
    Me.Text1.SetFocus
    var.modo = 2  '///modo impresion
End Sub


'\\\\\\
'//////
Private Sub Text1_GotFocus()
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    Label2.Visible = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If FrmConfig.Visible = True Then
        FrmConfig.Hide
        FrmConfig.Visible = False
    End If
    
    If KeyAscii = 13 Then Command3_Click
End Sub

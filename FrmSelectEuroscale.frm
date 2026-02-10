VERSION 5.00
Begin VB.Form FrmSelectEuroscale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   2205
   ClientLeft      =   6990
   ClientTop       =   2130
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   372
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   732
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3732
   End
End
Attribute VB_Name = "FrmSelectEuroscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedSys As Long

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    Else
        SelectedSys = Combo1.ListIndex
        Unload Me
    End If
End Sub

Public Sub LoadData()
    Dim bucle As Long
    Dim B2 As Long
    Dim Buf As String
    SelectedSys = -1
    Me.Caption = "Hydra +"
    Command1(1).Caption = CargaCadena(287)
    Command1(0).Caption = CargaCadena(288)
    Label1.Caption = CargaCadena(1203)
    Combo1.Clear
    Combo1.Enabled = False
    Combo1.AddItem "(no system)"
    Command1(1).Enabled = False
    MUE_Sistemas
    If SisEur.NSistemas > 0 Then
        Combo1.Clear
        For bucle = 0 To SisEur.NSistemas - 1
            Buf = ""
            For B2 = 0 To SisEur.Sistemas(bucle).NSecciones - 1
                Buf = Buf & Str(SisEur.Sistemas(bucle).Secciones(B2)) & " , "
            Next B2
            If Len(Buf) > 3 Then
                Buf = left(Buf, Len(Buf) - 3)
                Combo1.AddItem Buf
            End If
        Next bucle
        If Combo1.ListCount > 0 Then
            Combo1.Enabled = True
            Command1(1).Enabled = True

        Else
            Combo1.AddItem "(no system)"
        End If
    End If
    Combo1.ListIndex = 0
End Sub


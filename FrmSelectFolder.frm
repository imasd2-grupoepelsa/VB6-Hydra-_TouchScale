VERSION 5.00
Begin VB.Form FrmSelectFolder 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   7050
   ClientTop       =   2130
   ClientWidth     =   3000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2772
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2772
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1212
   End
End
Attribute VB_Name = "FrmSelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedPath As String
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////
Private Sub Form_Load()
    With Me
        .Icon = Form2.Icon
        .Dir1.Path = App.Path
        .Drive1.Drive = App.Path
        .SelectedPath = App.Path
        .Command1(0).Caption = CargaCadena(287)
        .Command1(1).Caption = CargaCadena(288)
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Command1_Click (1)
    End If
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////
Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        SelectedPath = Dir1.Path
    Else
        SelectedPath = ""
    End If
    Me.Hide
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////
Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    If Err.Number <> 0 Then
        MsgBox "Unidad no accesible", vbCritical
        Dir1.Path = App.Path
        Drive1.Drive = App.Path
    End If
    On Error GoTo 0
End Sub

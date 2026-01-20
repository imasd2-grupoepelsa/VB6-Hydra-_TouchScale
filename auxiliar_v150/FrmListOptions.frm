VERSION 5.00
Begin VB.Form FrmListOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydra+"
   ClientHeight    =   2160
   ClientLeft      =   5895
   ClientTop       =   2265
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3750
   Begin VB.CommandButton Command3 
      Caption         =   "Salir >"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Listar por Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visualizar en Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3492
   End
End
Attribute VB_Name = "FrmListOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Command1_Click()
    If Me.Tag = 1 Then
        Reporte.ExportReport rptKeyHTML, App.Path & "\dirtmptouch\report.htm", True
        ShellExecute Me.hwnd, "Open", App.Path & "\dirtmptouch\report.htm", "", "", 1
    Else
        Reporte2L.ExportReport rptKeyHTML, App.Path & "\dirtmptouch\report.htm", True
        ShellExecute Me.hwnd, "Open", App.Path & "\dirtmptouch\report.htm", "", "", 1
    End If
End Sub

Private Sub Command2_Click()
    If Me.Tag = 1 Then
        Reporte.PrintReport True
    Else
        Reporte2L.PrintReport True
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = Form2.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    L_WEB = True
End Sub

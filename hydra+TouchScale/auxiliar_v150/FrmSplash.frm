VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   5670
   ClientTop       =   2310
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Generando Informe, por favor, espere..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   972
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2772
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If id <> 0 Then
        Label1.Caption = "Generating Report, Please Wait..."
    End If
End Sub


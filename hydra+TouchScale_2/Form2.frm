VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   2880
   ClientTop       =   2100
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HYDRA+, GENERANDO NUEVA BASE DE DATOS, POR FAVOR, ESPERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Image Imagen 
      Height          =   825
      Left            =   0
      Picture         =   "Form2.frx":1E72
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5565
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Label1.Caption = CargaCadena(1369)
End Sub


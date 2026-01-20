VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Acerca de  Demo Tareas"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail: rmendez_26@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form2.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Elaborado Por: Ramon Alexander Mendez Mendez"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Punto Fijo - Estado Falcón - Venezuela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programa orientado a demostrar como se pueden utilizar algunas API's de windows y algo mas ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DEMOTAREAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Me
End Sub

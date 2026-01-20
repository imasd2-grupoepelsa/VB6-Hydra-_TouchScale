VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmfechahora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programación de fecha y hora"
   ClientHeight    =   2100
   ClientLeft      =   4410
   ClientTop       =   2850
   ClientWidth     =   2520
   Icon            =   "Frmfechahora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   2520
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdfechahora 
      Caption         =   "&Programar"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "HORA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmfechahora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdfechahora_Click()
    Dim LA_Fecha As Date
    Dim LA_Hora As Date
    LA_Fecha = Format(TxtFecha.Text, "dd/mm/yy")
    LA_Hora = Format(Label3.Caption, "HH:MM")
    
    AN_FechaHora LA_Fecha, LA_Hora
    Unload Me
End Sub

Private Sub Form_Load()
Dim Fecha As Date
Dim hora As Date
Me.Icon = Form2.Icon
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2


Fecha = Date
hora = Time


TxtFecha.Text = Format(Fecha, "dd/mm/yy")

Label3.Caption = Format(hora, "HH:MM")
frmfechahora.Caption = LoadResString(id + 101)
cmdfechahora.Caption = LoadResString(id + 102)
Label1.Caption = LoadResString(id + 434) & ":"
Label2.Caption = LoadResString(id + 435) & ":"



 'Do
 'do_events
 'fecha = Date
 'hora = Time
 'txtfecha.Text = Date
 
 'Loop
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Not StoreFlow Then frmEpelsa.Enabled = True
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then 'Permitir el uso de RETORNO
                Exit Sub
        End If
            'Los únicos caracteres válidos son los dígitos.
            If KeyAscii <> 13 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") Then
               KeyAscii = 0 'Poner a null si está fuera del intervalo.
                Beep
            End If
End Sub



VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPeriodo 
   Caption         =   "Resumen Impuestos Balanza."
   ClientHeight    =   2295
   ClientLeft      =   3405
   ClientTop       =   2460
   ClientWidth     =   5745
   Icon            =   "FrmPeriodo.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2295
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TextFin 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TextIni 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Introduzca Periodo a generar Resumen Impuestos"
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
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha Final:"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Final:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial:"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Fechas consolidadas disponibles Balanza"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "FrmPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim SALIR As Boolean
'Public Function LSALIR() As Boolean
'    LSALIR = SALIR
'End Function

Private Sub Command1_Click()
    If (CDate(TxtFecha(0).TexT) >= CDate(TextIni.TexT)) And (CDate(TxtFecha(1).TexT) <= CDate(TextFin.TexT)) _
    And (CDate(TxtFecha(0).TexT) <= CDate(TextFin.TexT)) And (CDate(TxtFecha(1).TexT) >= CDate(TextIni.TexT)) Then
        If (CDate(TxtFecha(0).TexT) >= CDate(TextIni.TexT)) And (CDate(TxtFecha(0).TexT) <= CDate(TextFin.TexT)) Then
            If (CDate(TxtFecha(1).TexT) >= CDate(TextIni.TexT)) And (CDate(TxtFecha(1).TexT) <= CDate(TextFin.TexT)) Then
            
                If (CDate(TxtFecha(0).TexT) <= CDate(TxtFecha(1).TexT)) Then
                    myPER1 = CDate(FrmPeriodo.TxtFecha(0).TexT)
                    myPER2 = CDate(FrmPeriodo.TxtFecha(1).TexT)
                
                    salir = 69
                    Unload Me
                End If
            End If
        End If
    End If
    If salir = 0 Then MsgBox "Periodo no válido...", vbExclamation
End Sub

Private Sub Command2_Click()
    salir = 0
    Unload Me
End Sub

Private Sub Form_Load()
    salir = 0
    If MyRecSet.State <> 0 Then
        MyRecSet.Close
        Set MyRecSet = Nothing
    End If
    MyRecSet.CursorLocation = adUseClient
    MyRecSet.Open "select min(date) as mind, max(date) as maxd from pcscale.totaldaily", MyConnObj, adOpenKeyset, adLockOptimistic
    If Not MyRecSet.EOF Then
        TextIni.TexT = Format(MyRecSet.Fields("mind"), "dd/mm/yyyy")
        TextFin.TexT = Format(MyRecSet.Fields("maxd"), "dd/mm/yyyy")
        TxtFecha(0).TexT = Format(MyRecSet.Fields("mind"), "dd/mm/yy")
        TxtFecha(1).TexT = Format(MyRecSet.Fields("maxd"), "dd/mm/yy")
        If MyRecSet.State <> 0 Then
            MyRecSet.Close
            Set MyRecSet = Nothing
        End If
    Else
        If MyRecSet.State <> 0 Then
            MyRecSet.Close
            Set MyRecSet = Nothing
        End If
        MsgBox "No existen totales consolidados por fechas en la Balanza...", vbExclamation
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SALIR = False
    Do_Events
    Sleep (10)
    'FrmConsultaTiquets.Enabled = True
    'FrmConsultaTiquets.SetFocus
End Sub

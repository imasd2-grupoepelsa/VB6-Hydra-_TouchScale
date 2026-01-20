VERSION 5.00
Begin VB.Form FrmIVA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Tipos de IVA"
   ClientHeight    =   2760
   ClientLeft      =   6885
   ClientTop       =   1650
   ClientWidth     =   4560
   ClipControls    =   0   'False
   Icon            =   "FrmIVA.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   1920
      TabIndex        =   11
      Top             =   1080
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   360
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "PLU 0"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Index           =   4
      Left            =   3120
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "%"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo 4"
      Height          =   252
      Index           =   4
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo 3"
      Height          =   252
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo 2"
      Height          =   252
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo 1"
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo 0"
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4575
   End
End
Attribute VB_Name = "FrmIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CambiarIdioma()
    Dim Bucle As Integer
    If UsaEcoPrint And TipoEcoPrint = 3 Then
        For Bucle = 1 To 5
            Label2(Bucle - 1).Caption = CargaCadena(55) & " " & Bucle
        Next Bucle
    
    Else
        For Bucle = 0 To 4
            Label2(Bucle).Caption = CargaCadena(55) & " " & Bucle
        Next Bucle
    End If
    Command1(0).Caption = CargaCadena(36)
    Command1(1).Caption = CargaCadena(7)
    FrmIVA.Caption = CargaCadena(8)
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Bucle As Integer
    Dim TIVAS(5) As Double
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    For Bucle = 0 To 4
        If Trim(Text1(Bucle).TexT) = "" Then Text1(Bucle).TexT = "0"
        TIVAS(Bucle) = ToDouble(Text1(Bucle).TexT)
    Next Bucle
    Alta_IVAS TIVAS
    MsgBox CargaCadena(439)
End Sub

Private Sub Form_Load()
   
    Dim Registro As New RecordNet
    Dim Bucle As Integer
    Me.Icon = Form2.Icon
    Label1.Visible = False
    CambiarIdioma
    Image1.Picture = Form2.Imagen.Picture
    For Bucle = 0 To 4
        Text1(Bucle).TexT = 0
    Next Bucle
   
    Registro.OpenRecordset ("select * from iva")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Text1(.Fields("Tipo")).TexT = .Fields("Valor")
                .Movenext
            Loop
        End If
    End With

End Sub



Private Sub Form_Unload(Cancel As Integer)
LockBase = False
    frmEpelsa.Enabled = True
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Bucle As Integer
    Dim Contador As Integer
    Dim HayComa As Boolean
    If KeyAscii = 46 Then KeyAscii = 44
    Text1(Index).Locked = Checktexto(KeyAscii, 5, Text1(Index).TexT, True, True)
    For Bucle = 1 To Len(Text1(Index).TexT)
        If Not IsNumeric(Mid(Text1(Index).TexT, Bucle, 1)) Then
            HayComa = True
        Else
            If HayComa = True Then Contador = Contador + 1
        End If
    Next Bucle
    If Contador >= 2 And KeyAscii <> 8 Then Text1(Index).Locked = True
    
    
End Sub

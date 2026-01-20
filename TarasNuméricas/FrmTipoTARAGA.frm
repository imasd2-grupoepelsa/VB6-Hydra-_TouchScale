VERSION 5.00
Begin VB.Form FrmTipoTARAGA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Euro-Scale."
   ClientHeight    =   3585
   ClientLeft      =   6990
   ClientTop       =   2130
   ClientWidth     =   6030
   Icon            =   "FrmTipoTARAGA.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command1"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   3000
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   1320
      TabIndex        =   21
      Top             =   3000
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   9
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   20
      Top             =   2280
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   8
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   19
      Top             =   1920
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   18
      Top             =   1560
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   17
      Top             =   1200
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   16
      Top             =   840
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   4
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2280
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   14
      Top             =   1920
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   13
      Top             =   1560
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1200
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   11
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   3120
      TabIndex        =   10
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   3120
      TabIndex        =   8
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   732
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6012
   End
End
Attribute VB_Name = "FrmTipoTARAGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim Base As Database
    Dim Registro As Recordset
    Dim Bucle As Integer
    For Bucle = 0 To 9
        If Not IsNumeric(Text1(Bucle).Text) Then
            Text1(Bucle).Text = "0"
        End If
    Next Bucle
    Set Base = OpenDatabase(App.Path & "\dbase.mdb", False, False)
    For Bucle = 0 To 9
        Set Registro = Base.OpenRecordset("select * from tipotara where tipo=" & (Bucle + 1))
        If Not Registro.EOF Then
            Registro.Edit
        Else
            Registro.AddNew
        End If
        Registro.Fields("tipo") = Bucle + 1
        Registro.Fields("valor") = CDbl(Text1(Bucle).Text)
        Registro.Fields("tran_el") = " "
        Registro.Update
    Next Bucle
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    'MsgBox CargaCadena(439), vbInformation
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CambiarIdioma
    CargaTaras
End Sub
Private Sub CargaTaras()
    Dim Base As Database
    Dim Registro As Recordset
    Dim Bucle As Integer
    Set Base = OpenDatabase(App.Path & "\dbase.mdb", False, False)
    For Bucle = 0 To 9
        Text1(Bucle).Text = "0"
    Next Bucle
    Set Registro = Base.OpenRecordset("select * from tipotara where tipo>0 and tipo<11")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Text1(.Fields("tipo") - 1).Text = .Fields("valor")
                .MoveNext
            Loop
        End If
    End With
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
End Sub
Private Sub CambiarIdioma()
    Dim Bucle As Integer
    For Bucle = 0 To 9
        Label2(Bucle).Caption = "Tipo Tara" & " " & Format(Bucle, "0")
    Next Bucle
    Me.Caption = "Edición Taras Numéricas."
    Label1.Caption = "Taras Numéricas para EuroScale."
    Command1.Caption = "Modificar."
    Command2.Caption = "Salir."
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case 48 To 59
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim MyNum As Integer
    Dim MyValor As Long
   
   If IsNumeric(Text1(Index).Text) Then
       If Val(Text1(Index).Text) > 9995 Then Text1(Index).Text = "9995"
       MyNum = Val(Right(Text1(Index).Text, 1))
       MyValor = Val(Left(Text1(Index).Text, Len(Text1(Index).Text) - 1)) * 10
       Select Case MyNum
            Case 0 To 2
                MyNum = 0
            Case 3 To 7
                MyNum = 5
            Case 8 To 9
                MyNum = 10
        End Select
        MyValor = MyValor + MyNum
        Text1(Index).Text = MyValor
    End If
    
End Sub

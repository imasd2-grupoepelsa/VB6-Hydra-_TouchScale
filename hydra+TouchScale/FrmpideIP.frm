VERSION 5.00
Begin VB.Form FrmpideIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP de la servidora"
   ClientHeight    =   3600
   ClientLeft      =   6915
   ClientTop       =   2115
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TmrConnect 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3840
      Top             =   2520
   End
   Begin VB.TextBox tXTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Por favor, indique el puerto ethernet de la Servidora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Por favor, indique la dirección IP de la balanza Servidora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmpideIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Fracaso As Boolean
Private Sub Command1_Click()
    Dim Buf As String
    Dim Bucle As Integer
    Dim MyVal As Integer
    Dim Arch As Integer
    Fracaso = False
    If Val(Text1.TexT) < 1024 Or Val(Text1.TexT) > 65535 Then
        MsgBox CargaCadena(999), vbCritical
        Exit Sub
    End If
    If Len(tXTip.TexT) = 0 Then
        MsgBox CargaCadena(957), vbCritical
        Exit Sub
    End If
    If Not IsNumeric(Right(tXTip.TexT, 1)) Then
        MsgBox CargaCadena(957), vbCritical
        Exit Sub
    End If
    tXTip.TexT = tXTip.TexT & "."
    For Bucle = 1 To Len(tXTip.TexT)
        If Asc(Mid(tXTip.TexT, Bucle, 1)) >= 48 And Asc(Mid(tXTip.TexT, Bucle, 1)) <= 59 Then
            MyVal = MyVal * 10 + Mid(tXTip.TexT, Bucle, 1)
        Else
            Buf = Buf & Format(MyVal, "000") & "."
            MyVal = 0
        End If
    Next Bucle
    Buf = Buf & Format(MyVal, "000")
    'bUF = Left(bUF, Len(bUF) - 1)
    tXTip.TexT = Buf
    If Len(tXTip.TexT) <> 15 Then
        MsgBox CargaCadena(957), vbCritical
        Exit Sub
    End If
    If IsNumeric(left(tXTip.TexT, 3)) And IsNumeric(Right(tXTip.TexT, 3)) And IsNumeric(Mid(tXTip.TexT, 5, 3)) And IsNumeric(Mid(tXTip.TexT, 9, 3)) Then
        If Val(left(tXTip.TexT, 3)) < 256 And Val(Right(tXTip.TexT, 3)) < 256 And Val(Mid(tXTip.TexT, 5, 3)) < 256 And Val(Mid(tXTip.TexT, 9, 3)) < 256 Then
            '
            ' Configurar "balanza"
            '
            
            lIp = True
            
            Arch = FreeFile()
            Open App.Path & "\servidoratouch.ini" For Output As #Arch
            Print #Arch, Val(left(tXTip.TexT, 3)) & "." & Val(Mid(tXTip.TexT, 5, 3)) & "." & Val(Mid(tXTip.TexT, 9, 3)) & "." & Val(Right(tXTip.TexT, 3))
            Close #Arch
            sAddIp = Val(left(tXTip.TexT, 3)) & "." & Val(Mid(tXTip.TexT, 5, 3)) & "." & Val(Mid(tXTip.TexT, 9, 3)) & "." & Val(Right(tXTip.TexT, 3))
            Me.Hide
        Else
            MsgBox CargaCadena(957), vbCritical
        End If
    Else
        MsgBox CargaCadena(957), vbCritical
    End If
End Sub

Private Sub Form_Load()
    Text1.TexT = 3306
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TmrConnect_Timer()
    Fracaso = True
End Sub

Private Sub tXTip_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 59
        Case Asc(".")
        Case vbKeyDelete
        Case vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

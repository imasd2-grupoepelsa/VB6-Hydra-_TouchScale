VERSION 5.00
Begin VB.Form FrmAviso 
   Caption         =   "Aviso Actividad TOUCH-SCALE."
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "FrmAviso.frx":0000
   LinkTopic       =   "FrmAviso"
   ScaleHeight     =   1230
   ScaleWidth      =   7200
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblAviso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "FrmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
 Dim vtData As Variant ' Data variable.

 Select Case State
 ' ... Other cases not shown.
 Case icError ' 11
     ' In case of error, return ResponseCode and

' ResponseInfo.
     vtData = Inet1.ResponseCode & ":" & _
     Inet1.ResponseInfo

 Case icResponseCompleted  ' 12
     Dim strData As String: strData = ""
     Dim bDone As Boolean: bDone = False

     ' Get first chunk.
     vtData = Inet1.GetChunk(1024, icString)
     Do_Events

     Do While Not bDone

         strData = strData & vtData
         ' Get next chunk.
         vtData = Inet1.GetChunk(1024, icString)
         Do_Events

         If Len(vtData) = 0 Then
             bDone = True
         End If
     Loop

    MsgBox strData
 End Select
End Sub


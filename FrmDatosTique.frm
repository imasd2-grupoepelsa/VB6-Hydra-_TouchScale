VERSION 5.00
Begin VB.Form FrmDatosTique 
   Caption         =   "Datos Generales Tique."
   ClientHeight    =   4410
   ClientLeft      =   3585
   ClientTop       =   1785
   ClientWidth     =   7200
   Icon            =   "FrmDatosTique.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtLey 
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
      Height          =   285
      Index           =   4
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   15
      Top             =   3120
      Width           =   5415
   End
   Begin VB.TextBox TxtLey 
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
      Height          =   285
      Index           =   3
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   14
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox TxtLey 
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
      Height          =   285
      Index           =   2
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox TxtLey 
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
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox TxtLey 
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
      Height          =   285
      Index           =   0
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox TxtCab 
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
      Height          =   285
      Index           =   2
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox TxtCab 
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
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   720
      Width           =   5415
   End
   Begin VB.TextBox TxtCab 
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
      Height          =   285
      Index           =   0
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label LblLey 
      Caption         =   "Leyenda  5:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LblLey 
      Caption         =   "Leyenda  4:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label LblLey 
      Caption         =   "Leyenda  3:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label LblLey 
      Caption         =   "Leyenda  2:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblLey 
      Caption         =   "Leyenda  1:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label LblCab 
      Caption         =   "Cabecera 3:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LblCab 
      Caption         =   "Cabecera 2:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label LblCab 
      Caption         =   "Cabecera 1:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDatosTique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdaceptar_Click()
    Dim Mybase As dao.Database
    Dim MyReg As dao.Recordset
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyReg = Mybase.OpenRecordset("select * from trade")
    If MyReg.EOF Then
        MyReg.AddNew
    Else
        MyReg.Edit
    End If

    MyReg.Fields("company") = Trim(TxtCab(0).TexT)
    MyReg.Fields("name") = Trim(TxtCab(1).TexT)
    MyReg.Fields("address") = Trim(TxtCab(2).TexT)

    MyReg.Fields("quarter") = Trim(TxtLey(0).TexT)
    MyReg.Fields("web") = Trim(TxtLey(1).TexT)
    MyReg.Fields("shopweb") = Trim(TxtLey(2).TexT)
    MyReg.Fields("description") = Trim(TxtLey(3).TexT)
    MyReg.Fields("slogan") = Trim(TxtLey(4).TexT)

    MyReg.Update
    MyReg.Close
    Set MyReg = Nothing
    Mybase.Close
    Set Mybase = Nothing

    CmdSalir_Click

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Mybase As dao.Database
    Dim MyReg As dao.Recordset
    Dim nN As Integer
    For nN = 0 To 2
        LblCab(nN).Caption = CargaCadena(317) & " " & CStr(nN + 1)
    Next nN
    For nN = 0 To 4
        LblLey(nN).Caption = CargaCadena(318) & " " & CStr(nN + 1)
    Next nN
    CmdAceptar.Caption = CargaCadena(287)
    CmdSalir.Caption = CargaCadena(288)
    Me.Caption = CargaCadena(1403)
    Set Mybase = dao.OpenDatabase(Base_General)
    Set MyReg = Mybase.OpenRecordset("select * from trade")
    If MyReg.EOF Then
        TxtCab(0).TexT = ""
        TxtCab(1).TexT = ""
        TxtCab(2).TexT = ""
        TxtLey(0).TexT = ""
        TxtLey(1).TexT = ""
        TxtLey(2).TexT = ""
        TxtLey(3).TexT = ""
        TxtLey(4).TexT = ""
    Else
        TxtCab(0).TexT = Trim(MyReg.Fields("company"))
        TxtCab(1).TexT = Trim(MyReg.Fields("name"))
        TxtCab(2).TexT = Trim(MyReg.Fields("address"))

        TxtLey(0).TexT = Trim(MyReg.Fields("quarter"))
        TxtLey(1).TexT = Trim(MyReg.Fields("web"))
        TxtLey(2).TexT = Trim(MyReg.Fields("shopweb"))
        TxtLey(3).TexT = Trim(MyReg.Fields("description"))
        TxtLey(4).TexT = Trim(MyReg.Fields("slogan"))
    End If


    MyReg.Close
    Set MyReg = Nothing
    Mybase.Close
    Set Mybase = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

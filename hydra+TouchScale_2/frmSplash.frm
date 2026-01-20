VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   Caption         =   "HydraTouchLite"
   ClientHeight    =   4455
   ClientLeft      =   6990
   ClientTop       =   2115
   ClientWidth     =   5190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   372
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1332
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "version componente 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "version componente 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "version app"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "HydraTouchLite  grupoEPELSA  (C)2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'\\\\\\\\\\\\\\
'//////////////
Private Sub Form_Load()
Dim cnt As Long
    With Me
        .Picture = Form2.Imagen.Picture
        For cnt = 0 To 3
            .lblTexto(cnt).Visible = False
        Next cnt
        .lblTexto(0).Visible = True
        .lblTexto(1).Visible = True
        .lblTexto(1).Caption = miver
        'If AhorraMas Then
        '    .lblTexto(1).Caption = .lblTexto(1).Caption & " AH"
        'Else
            'If Sabeco Then
            '    .lblTexto(1).Caption = .lblTexto(1).Caption & " SB"
            'Else
                'If AHOLD Then
                '    .lblTexto(1).Caption = .lblTexto(1).Caption & " Dinosol"
                'Else
                    'If StoreFlow Then
                    '    .lblTexto(1).Caption = .lblTexto(1).Caption & " SRV"
                    'Else
                    '    If Supeco Then
                    '        .lblTexto(1).Caption = .lblTexto(1).Caption & " AT"
                    '    End If
                    'End If
                'End If
            'End If
        'End If
        '.lblTexto(2).Visible = True
        '.lblTexto(2).Caption = "epelcom v" & frmControl.Balanza.Version
        .lblTexto(3).Visible = True
        .lblTexto(3).Caption = "import v" & frmControl.ControlImport1.Version
        If CDKEY_USER Then
            .lblTexto(0).Caption = .lblTexto(0).Caption & " KEY : " & ClavePublica
        End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub
'\\\\\\
'//////
Private Sub Command1_Click()
    Unload Me
End Sub

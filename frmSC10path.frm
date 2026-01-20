VERSION 5.00
Begin VB.Form frmSC10path 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elección de directorio de red"
   ClientHeight    =   3105
   ClientLeft      =   6900
   ClientTop       =   2130
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox DirSC10 
      Height          =   1890
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox drvSC10 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmSC10path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'\\\\\\\\\\\\\\
'//////////////
Private Sub Form_Load()
    With Me
        .drvSC10.Refresh
        .Caption = CargaCadena(266)
        .OKButton.Caption = CargaCadena(287)
        .CancelButton.Caption = CargaCadena(288)
    End With
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub OKButton_Click()
    FrmConfiguracion2.TxtPathSC10 = DirSC10.Path
    OKButton.Caption = CargaCadena(287)
    CancelButton.Caption = CargaCadena(288)
    frmSC10path.Hide
    'frmConfiguracion.Show
    'frmConfiguracion.Enabled = True
    Unload Me
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub CancelButton_Click()
    frmSC10path.Hide
    'frmConfiguracion.Show
    'frmConfiguracion.Enabled = True
    Unload Me
End Sub
'\\\\\\\\\\\\\\
'//////////////
Private Sub drvSC10_Change()
    On Error GoTo manejador
    DirSC10.Path = drvSC10.Drive
    On Error GoTo 0
    Exit Sub
manejador:
    MsgBox CargaCadena(1374) '"Unit not Accesible", vbCritical
End Sub

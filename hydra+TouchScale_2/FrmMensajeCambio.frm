VERSION 5.00
Begin VB.Form FrmMensajeCambio 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4425
   ClientLeft      =   6900
   ClientTop       =   2085
   ClientWidth     =   6045
   FillColor       =   &H00800000&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   2172
   End
   Begin VB.Image Image1 
      Height          =   372
      Index           =   1
      Left            =   2880
      Picture         =   "FrmMensajeCambio.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   612
   End
   Begin VB.Image Image1 
      Height          =   372
      Index           =   0
      Left            =   2160
      Picture         =   "FrmMensajeCambio.frx":030A
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmMensajeCambio.frx":0614
      ForeColor       =   &H00FFFFFF&
      Height          =   2652
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5052
   End
End
Attribute VB_Name = "FrmMensajeCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 1
            Label1.Caption = "Se ha detectado una base de datos con formato antiguo,anterior a la versión HydraTouch. En el cambio de formato a realizar, se eliminarán algunos campos obsoletos de la base para mejorar la eficiencia de ésta. Estos campos no tienen importancia para HydraTouch, no obstante, si su sistema de gestión ataca directamente a la base y utiliza estos campos para algún fin, habrá de crearlos de nuevo (la aplicación no los volverá a eliminar). Tras el cambio, quedará en la carpeta de programa una copia de la base antigua con el nombre dbase5.mdb. Antes de continuar asegurese de que ninguna otra aplicación está haciendo uso de la base, en otro caso fallará el proceso."
            Command2.Caption = "Cancelar"
            Command1.Caption = "Continuar"
            
        Case 0
            Label1.Caption = "An old database has been detected (format older than this version of HydraTouch). An automatic format adaptation will be performed, erasing some older fields in order to improve performance. These fields are not used by HydraTouch, however, if your backoffice uses them directly by any reason, you will need to create it again after this proccess (HydraTouch will not erase them again). After this proccess, a backup of the old database will be placed into program folder named dbase5.mdb. Before continue, please be sure that no any other application is using the database now, or the proccess will fail."
            Command2.Caption = "Cancel"
            Command1.Caption = "Continue"
    End Select
End Sub

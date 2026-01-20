VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFechaHistorico2 
   Caption         =   "Fecha historico"
   ClientHeight    =   3840
   ClientLeft      =   10095
   ClientTop       =   2145
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   3735
   End
   Begin MSComctlLib.TreeView tvwMO 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5953
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmFechaHistorico2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''
'''''''''''''''
Private Sub Form_Load()
'    Dim numMax As Long
'    Dim cnt As Long
'    Dim clave As String
'    Dim YYYYMM_MDB As String
'    With Me
'        .tvwMO.Nodes.Clear
'        'db.abre_conx_db var.rutadbSYS, conx
'        'numMax = db.carga_proyectos(conx, regProyectos)
'        'db.cierra_conx_db conx
'        'db.abre_conx_db var.rutadbLANG, conx
'        'strTxt = UCase(db.carga_str(conx, var.Idioma, 155))
'        'db.cierra_conx_db conx
'        YYYYMM_MDB = App.Path & "*.mdb"
'        Dir
'        .tvwMO.Nodes.Add , , "SYS", strTxt ', .imlProyec.ListImages.Item(1).Index
'        If numMax > 0 Then
'            For cnt = 0 To numMax - 1
'                clave = "P" & regProyectos(cnt).id
'                .tvwMO.Nodes.Add "SYS", tvwChild, clave, regProyectos(cnt).Nombre ', .imlProyec.ListImages.Item(2).Index
'            Next cnt
'        End If
'        .tvwMO.Nodes(1).Expanded = True
'    End With
End Sub


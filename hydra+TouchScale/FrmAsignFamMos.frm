VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAsignFamMos 
   Caption         =   "Asigna Familias a Mostradores y su Orden."
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   15015
   LinkTopic       =   "Form4"
   ScaleHeight     =   8235
   ScaleWidth      =   15015
   Begin MSComctlLib.ListView ListFam 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   12938
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TabDlg.SSTab TabCounters 
      Height          =   7335
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmAsignFamMos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "FrmAsignFamMos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Refresca_Familias
'Call Refresca_Mostradores
End Sub

Private Sub Form_Load()

End Sub

Private Sub Refresca_Familias()
Dim Registro As New RecordNet
Dim nC As Integer
Dim lvwItem As Variant
    nC = 0
    Registro.OpenRecordset ("select distinct(codi_fam),imagen,txt_fam from fam_code order by codi_fam")
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Set lvwItem = ListFam.ListItems.Add(nC, , Registro.Fields("txt_fam"), Registro.Fields("imagen"))
                .Movenext
                nC = nC + 1
            Loop
        End If
    End With
End Sub


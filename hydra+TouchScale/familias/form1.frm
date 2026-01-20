VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2190
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6480
      Top             =   7680
      Width           =   615
   End
   Begin VB.Line LineSecSepara 
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   5640
   End
   Begin VB.Line LineFamSepara 
      X1              =   120
      X2              =   7200
      Y1              =   5640
      Y2              =   5640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyBase As DAO.Database
Dim MyRs As DAO.Recordset
Dim MyRsAux As DAO.Recordset

Dim MyImgFam() As Image

Private Type FamData
    code As Long
    desc As String
    imag As String
    etiq As String
    vacu As Boolean
End Type
Dim Familias() As FamData
Dim nFam As Integer

Private Sub Form_Initialize()
Dim Token As Long
    Set MyBase = DAO.OpenDatabase(App.Path & "\dbasetouch.mdb")
    Set MyRs = MyBase.OpenRecordset("select * from fam_code order by codi_fam")
    If Not MyRs.EOF Then
        ReDim MyImgFam(MyRs.RecordCount)
        ReDim Familias(MyRs.RecordCount)
        nFam = 0
        MyRs.MoveFirst
        Do While Not MyRs.EOF
            Familias(nFam).code = MyRs.Fields("codi_fam")
            Familias(nFam).desc = MyRs.Fields("txt_fam")
            Familias(nFam).imag = "C:\hydra+TouchScale\images\" & Trim(MyRs.Fields("imagen"))
            Set MyImgFam(nFam) = Controls.Add("vb.image", "img" & CStr(MyRs.Fields("codi_fam")))
            MyImgFam(nFam).Width = 100 * Screen.TwipsPerPixelX
            MyImgFam(nFam).Height = 100 * Screen.TwipsPerPixelY
            MyImgFam(nFam).Stretch = True
            MyImgFam(nFam).left = nFam * MyImgFam(nFam).Width
            MyImgFam(nFam).top = LineFamSepara.Y1
            MyImgFam(nFam).BorderStyle = 1
            MyImgFam(nFam).ToolTipText = "Fam." & CStr(Familias(nFam).code) & ":" & Familias(nFam).desc
            MyImgFam(nFam).Visible = True
            
            On Error Resume Next
            Token = InitGDIPlus
            If Familias(nFam).imag <> "" Then
                MyImgFam(nFam).Picture = LoadPictureGDIPlus(Familias(nFam).imag)
            Else
                MyImgFam(nFam).Picture = LoadPictureGDIPlus(App.Path & "\vacio.jpg")
            End If
            FreeGDIPlus Token
            On Error GoTo 0
            
            nFam = nFam + 1
            
            If nFam = 10 Then Exit Do
            MyRs.MoveNext
        Loop
        MyRs.Close
        MyBase.Close
    Else
        MyRs.Close
        Set MyRs = Nothing
        MyBase.Close
        Set MyBase = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 1024 * Screen.TwipsPerPixelX
    Me.Height = 768 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_Resize()
    LineFamSepara.X1 = 0
    LineFamSepara.X2 = Me.Width
    LineFamSepara.Y1 = Me.Height - (200 * Screen.TwipsPerPixelY)
    LineFamSepara.Y2 = LineFamSepara.Y1
    LineSecSepara.X1 = Me.left + (100 * Screen.TwipsPerPixelX)
    LineSecSepara.X2 = LineSecSepara.X1
    LineSecSepara.Y1 = 0
    LineSecSepara.Y2 = LineFamSepara.Y2
End Sub


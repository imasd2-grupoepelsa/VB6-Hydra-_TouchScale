VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmImagen 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   1560
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommImage 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim s As String
    Dim e As Variant

    Do_Events
    Me.Visible = False
    Do_Events
    Select Case nQueImagen
    Case 0
        s = "Familias"
    Case 1
        s = "Artículos"
    Case 2
        s = "Vendedores"
    Case Else
        FrmImagen.Hide
    End Select
    CommImage.DialogTitle = "Imagenes " & s
    CommImage.Filter = "ALL Files|*.*|BMP Files|*.bmp|DIB Files|*.dib|GIF FILES|*.gif|JPG Files|*.jpg|WMF Files|*.wmf|EMF Files|*.emf|PNG Files|*.png"
    CommImage.DefaultExt = "*.*"
    CommImage.FileName = ""
    CommImage.FilterIndex = 0    '0=ALL , 5=jpg
    CommImage.InitDir = App.Path & "\images"

    SetDetailsView
    CommImage.ShowOpen
    e = CommImage.FileName
    If e = "" Then    'Or InStr(1, LCase(e), "vacio.jpg")<>0 Then
        '    On Error Resume Next
        '    FileCopy e, sQuePathImagen & sQueNombreImagen
        '    On Error GoTo 0
        'Else
        sQueNombreImagen = ""
    Else
        sQueNombreImagen = e
    End If
    FrmImagen.Hide

End Sub

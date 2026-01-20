Attribute VB_Name = "Var"
Option Explicit

Public Sub cargaImg(ByVal sImg As String, ByVal nIndex As Integer)
Dim Token As Long
    If InStr(1, LCase(sImg), ".png") <> 0 Then
        Token = InitGDIPlus
        Set FrmMain.Btn(nIndex).PictureNormal = LoadPictureGDIPlus(App.Path & "\" & sImg)
        FreeGDIPlus Token
    Else
        Set FrmMain.Btn(nIndex).PictureNormal = LoadPicture(App.Path & "\" & sImg)
    End If

End Sub

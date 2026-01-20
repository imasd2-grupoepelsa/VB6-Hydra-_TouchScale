Attribute VB_Name = "modCommon"
Option Explicit

'------------------------------------------------------------------------
Public Sub CenterForm(pobjForm As Form)
'------------------------------------------------------------------------

    With pobjForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With

End Sub

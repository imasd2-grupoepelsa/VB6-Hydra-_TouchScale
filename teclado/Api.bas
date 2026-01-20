Attribute VB_Name = "Api"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub Do_Events()
    Sleep (1)
    DoEvents
    Sleep (1)
End Sub

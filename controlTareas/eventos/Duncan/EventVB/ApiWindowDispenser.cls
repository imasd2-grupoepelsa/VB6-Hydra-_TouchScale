VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiWindowDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2407
Option Explicit

Private mIsCommonContolsInitialsed As Boolean

'\\ Creating new windows.....
Private Declare Function CreateWindowExApi Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Function CreateWindow(ByVal ExtendedWindowStyle As WindowStyleExtendedBits, ByVal TheClass As APIWndClass, ByVal WindowName As String, ByVal WindowStyle As WindowStyleBits, ByVal xPos As Long, ByVal yPos As Long, ByVal Width As Long, ByVal Height As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInst As Long, ByVal lpCreateStruct As Long) As ApiWindow
Attribute CreateWindow.VB_HelpID = 2408

Dim wndThis As ApiWindow
Dim lret As Long

lret = CreateWindowExApi(ExtendedWindowStyle, TheClass.Name, WindowName, WindowStyle, xPos, yPos, Width, Height, hWndParent, hMenu, hInst, lpCreateStruct)
If lret <> 0 And Err.LastDllError = 0 Then
    Set wndThis = New ApiWindow
    wndThis.hwnd = lret
Else
    Call ReportError(Err.LastDllError, "ApiWindowDispenser:CreateWindow", GetLastSystemError)
End If
Set CreateWindow = wndThis

End Function

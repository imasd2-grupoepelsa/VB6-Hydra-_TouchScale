VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSubclassedWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2297
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

'\\ If we subclass a window that is not owned by the current thread then we need
'\\ to attach that thread's input queue to this one
Private Declare Function AttachThreadInput Lib "kernel32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long


Public Sub Add(ByVal wnd As ApiWindow)
Attribute Add.VB_HelpID = 2298

Dim Item As Object
Dim lret As Long

On Local Error Resume Next

Set Item = mCol.Item(wnd.Key)

If Item Is Nothing Then
    mCol.Add wnd, wnd.Key
    '\\ If the window is from another thread, attach the input threads together...
    If wnd.ThreadParent <> APIDispenser.CurrentThreadId Then
        ReportError 0, "SubclassedWindows:Add", "Call failed - you cannot subclass windows in a thread other than the current one"
    End If
    '\\ Start subclassing it.....
    Call wnd.SetWindowLong(GWL_WNDPROC, AddressOf VB_WindowProc)
End If

End Sub

Public Function Item(ByVal hwnd As Long) As ApiWindow
Attribute Item.VB_HelpID = 2299

On Error Resume Next
Set Item = mCol.Item("HWND:" & hwnd)

End Function


Public Sub Remove(wnd As ApiWindow)
Attribute Remove.VB_HelpID = 2300

On Error Resume Next

If Not wnd Is Nothing Then
    If Not (mCol.Item(wnd.Key) Is Nothing) Then
        '\\ Stop subclassing it.....
        Call wnd.SetWindowLong(GWL_WNDPROC, wnd.OldProcAddress)
        
        '\\ and remove it from our collection
        mCol.Remove wnd.Key
        
    End If
End If

End Sub

Private Sub Class_Initialize()

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()


    Set mCol = Nothing

End Sub


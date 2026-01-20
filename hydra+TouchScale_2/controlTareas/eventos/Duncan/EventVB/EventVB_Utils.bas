Attribute VB_Name = "EventVB_Utils"
'\\ -----[EventVB_UtilityFuncs]----------------------------------------------------
'\\ Useful generic functions for use withing the EventVb project
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Option Explicit

'\\ For offsetting a collection of windows
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SetWindowPosFlags) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

'\\ For getting a window's base class
Private Const sWNDCLASS_BUTTON = "BUTTON"
Private Const sWNDCLASS_COMBO = "COMBOBOX"
Private Const sWNDCLASS_LIST = "LISTBOX"
Private Const sWNDCLASS_SCROLL = "SCROLLBAR"
Private Const sWNDCLASS_EDIT = "EDIT"
Private Const sWNDCLASS_STATIC = "STATIC"
Private Declare Function GetClassNameApi Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long



Public Function BaseClassName(ByVal wndcls As StandardControlTypes) As String

Select Case wndcls
Case Control_BUTTON
    BaseClassName = sWNDCLASS_BUTTON
Case Control_COMBOBOX
    BaseClassName = sWNDCLASS_COMBO
Case Control_LISTBOX
    BaseClassName = sWNDCLASS_LIST
Case Control_SCROLLBAR
    BaseClassName = sWNDCLASS_SCROLL
Case Control_EDIT
    BaseClassName = sWNDCLASS_EDIT
Case Control_STATIC
    BaseClassName = sWNDCLASS_STATIC
End Select

End Function
Public Function GetWindowBaseClass(ByVal hwnd As Long) As StandardControlTypes

Dim sClassName As String, lRet As Long

'\\ Get the class name
sClassName = String$(1024, 0)
lRet = GetClassNameApi(hwnd, sClassName, 1024)
If (Err.LastDllError = 0) And (lRet > 0) Then
    sClassName = Left$(sClassName, lRet)
End If

Select Case sClassName
Case sWNDCLASS_BUTTON, "ThunderCheckBox", "ThunderRTCheckBox", _
     "ThunderCommandButton", "ThunderRTCommandButton", _
     "ThunderFrame", "ThunderRTFrame", _
     "ThunderOptionButton", "ThunderRTOptionButton"
    '\\ Classes derived from the BUTTON base class
    GetWindowBaseClass = Control_BUTTON
Case sWNDCLASS_COMBO, "ThunderComboBox", "ThunderRTComboBox", _
     "ThunderDriveListBox", "ThunderRTDriveListBox"
    
    '\\ Classes derived from the COMBOBOX base class
    GetWindowBaseClass = Control_COMBOBOX
Case sWNDCLASS_LIST, _
     "ThunderDirListBox", "ThunderRTDirListBox", _
     "ThunderListBox", "ThunderRTListBox", _
     "ThunderFileListBox", "ThunderRTFileListBox"
     
    '\\ Classes derived from the LISTBOX base class
    GetWindowBaseClass = Control_LISTBOX
Case sWNDCLASS_SCROLL, _
    "ThunderVScrollBar", "ThunderRTVScrollBar", _
    "ThunderHScrollBar", "ThunderRTHScrollBar"
    
    '\\ Classes derived from the SCROLLBAR base class
    GetWindowBaseClass = Control_SCROLLBAR
Case sWNDCLASS_EDIT, _
     "ThunderTextBox", "ThunderRTTextBox"
    '\\ Classes derived from the SCROLLBAR base class
    GetWindowBaseClass = Control_EDIT
Case Else
    GetWindowBaseClass = Control_STATIC
End Select

End Function

Public Function GreatestCommonDenominator(ByVal x As Long, ByVal y As Long) As Long

Dim x2 As Long, y2 As Long

x2 = x
y2 = y

While x2 <> y2
    If x2 > y2 Then
        x2 = x2 - y2
    Else
        y2 = y2 - x2
    End If
Wend

GreatestCommonDenominator = x2

End Function

Public Sub OffsetWindows(ByVal colWindows As Collection, ByVal dx As Long, ByVal dy As Long, ByVal wndParent As ApiWindow)

Dim hDefPos As Long
Dim wndThis As ApiWindow
Dim ptWindow As POINTAPI

If Not colWindows Is Nothing Then
    If colWindows.Count > 0 Then
        hDefPos = BeginDeferWindowPos(colWindows.Count)
        If hDefPos <> 0 Then
            For Each wndThis In colWindows
                With wndThis
                    ptWindow.x = .RECT.Left
                    ptWindow.y = .RECT.Top
                    If Not wndParent Is Nothing Then
                        Call ScreenToClient(wndParent.hwnd, ptWindow)
                    End If
                    Call DeferWindowPos(hDefPos, .hwnd, 0, ptWindow.x + dx, ptWindow.y + dy, 0, 0, SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED)
                End With
            Next wndThis
            Call EndDeferWindowPos(hDefPos)
        End If
    End If
End If

End Sub


Public Function StandardControlClassName(ByVal ctlType As StandardControlTypes) As String

Select Case ctlType
Case Control_BUTTON
    StandardControlClassName = "BUTTON"
Case Control_EDIT
    StandardControlClassName = "EDIT"
Case Control_LISTBOX
    StandardControlClassName = "LISTBOX"
Case Control_SCROLLBAR
    StandardControlClassName = "SCROLLBAR"
Case Control_STATIC
    StandardControlClassName = "STATIC"
End Select

End Function



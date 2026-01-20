VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiScrollbarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2963
Attribute VB_Description="ApiScrollbarControl - Class for manipulating windows which are based on the SCROLLBAR class"
Option Explicit

' ##MODULE_DESCRIPTION ApiScrollbarControl - Class for manipulating windows which are based on the SCROLLBAR class
' ##MODULE_REMARKS

' ##ENUMERATION_DESCRIPTION ScrollbarStyles - Sets how the scrollbar is shown and behaves
Public Enum ScrollbarStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_HORZ - Scrollbar is horizontal
    SBS_HORZ = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_VERT - Scrollbar is vertical
    SBS_VERT = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_TOPALIGN - Scrollbar is aligned to the top edge of its window
    SBS_TOPALIGN = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_LEFTALIGN - Scrollbar is aligned to the left edge of its window
    SBS_LEFTALIGN = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_BOTTOMALIGN - Scrollbar is aligned to its bottom of its window
    SBS_BOTTOMALIGN = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_RIGHTALIGN - Scrollbar is aligned to the right of its window
    SBS_RIGHTALIGN = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_SIZEBOXTOPLEFTALIGN - The scrollbar size box is aligned to the top left
    SBS_SIZEBOXTOPLEFTALIGN = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_SIZEBOXBOTTOMRIGHTALIGN - The scrollbar is aligned to the bottom right
    SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_SIZEBOX - The scrollbar has a size box
    SBS_SIZEBOX = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBS_SIZEGRIP - The scrollbar has a size box with a raised edge
    SBS_SIZEGRIP = &H10
End Enum

' ##ENUMERATION_DESCRIPTION ScrollbarMessages -
Public Enum ScrollbarMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_SETPOS - Sets the scrollbar position
    SBM_SETPOS = &HE0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_GETPOS - Returns the scrollbar position
    SBM_GETPOS = &HE1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_SETRANGE - Sets the min and max of the scrollbar
    SBM_SETRANGE = &HE2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_SETRANGEREDRAW - Sets the min and max values and redraws the scrollbar to reflect the new values
    SBM_SETRANGEREDRAW = &HE6
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_GETRANGE - Gets the min and max of the scrollbar
    SBM_GETRANGE = &HE3
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_ENABLE_ARROWS - Sets the enabled state of the scrollbar arrows
    SBM_ENABLE_ARROWS = &HE4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_SETSCROLLINFO - Sets the scrollbar info data
    SBM_SETSCROLLINFO = &HE9
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_GETSCROLLINFO - Gets the scrollbar info data
    SBM_GETSCROLLINFO = &HEA
End Enum

' ##ENUMERATION_DESCRIPTION ScrollbarInfoIndexes - Describes the effect of a call to set the scroll bar information
Public Enum ScrollbarInfoIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_RANGE Set the scroll bar min/max positions
    SIF_RANGE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_PAGE Set the page size for proportional scrollbars
    SIF_PAGE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_POS Set the tracking (scrollbar grab handle) position when not tracking
    SIF_POS = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_DISABLENOSCROLL If this is set then the scrollbar is disabled rather than hidden if it is not needed
    SIF_DISABLENOSCROLL = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_TRACKPOS Set the tracking (scrollbar grab handle) position during tracking
    SIF_TRACKPOS = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION SIF_ALL Update all the information
    SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Const CLSNAME = "ApiScrollbarControl"

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As ScrollbarMessages, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mHwnd As Long
Private bChanged As Boolean

Public Property Let Max(ByVal newval As Long)
Attribute Max.VB_HelpID = 2965

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, SBM_SETRANGEREDRAW, Min, newval)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Max", GetLastSystemError
    End If
End If

End Property

Public Property Get Max() As Long

If mHwnd <> 0 Then
    Dim lMin As Long, lMax As Long
    Call SendMessageLong(mHwnd, SBM_GETRANGE, lMin, lMax)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Max", GetLastSystemError
    End If
    Max = lMax
End If

End Property

Public Property Let Min(ByVal newval As Long)
Attribute Min.VB_HelpID = 2967

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, SBM_SETRANGEREDRAW, newval, Max)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Min", GetLastSystemError
    End If
End If

End Property

Public Property Get Min() As Long

If mHwnd <> 0 Then
    Dim lMin As Long, lMax As Long
    Call SendMessageLong(mHwnd, SBM_GETRANGE, lMin, lMax)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Min", GetLastSystemError
    End If
    Min = lMin
End If

End Property


Public Property Let Value(ByVal newval As Long)
Attribute Value.VB_HelpID = 2969

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, SBM_SETPOS, newval, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Value", GetLastSystemError
    End If
End If

End Property

Public Property Get Value() As Long

If mHwnd <> 0 Then
    Value = SendMessageLong(mHwnd, SBM_GETPOS, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Value", GetLastSystemError
    End If
End If

End Property


Friend Property Set Window(ByVal newval As ApiWindow)

If newval Is Nothing Then
    mHwnd = 0
Else
    mHwnd = newval.hWnd
End If

End Property

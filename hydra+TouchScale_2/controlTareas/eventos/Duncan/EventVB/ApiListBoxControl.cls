VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiListBoxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2949
Attribute VB_Description = "ApiListboxControl - Class for manipulating windows which are based on the LIST class"
Option Explicit

Private Const CLSNAME = "ApiListboxControl"

' ##MODULE_DESCRIPTION ApiListboxControl - Class for manipulating windows which are based on the LIST class

' ##ENUMERATION_DESCRIPTION ListboxControlStyles - Window style bits specific to listboxes and derived windows
Public Enum ListboxControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_NOTIFY - Notify the parent window when the listbox is clicked
    LBS_NOTIFY = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_SORT - Listbox is sorted
    LBS_SORT = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_NOREDRAW - Listbox is not automatically repainted when content items are changed
    LBS_NOREDRAW = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_MULTIPLESEL - User can select more than one item
    LBS_MULTIPLESEL = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_OWNERDRAWFIXED - Control is drawn by its owner and each item is the same height
    LBS_OWNERDRAWFIXED = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_OWNERDRAWVARIABLE - Control is drawn by its owner and items can have different heights
    LBS_OWNERDRAWVARIABLE = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_HASSTRINGS - The listbox looks after the strings for the contents
    LBS_HASSTRINGS = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_USETABSTOPS - Allow the controls strings to be tab-aligned
    LBS_USETABSTOPS = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_NOINTEGRALHEIGHT - Allows the display of partial items rather than automatically resizing the control
    LBS_NOINTEGRALHEIGHT = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_MULTICOLUMN - Allows multiple columsn which can be scrolled horizontally
    LBS_MULTICOLUMN = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_WANTKEYBOARDINPUT - Allows special processing of keypresses when this control has the focus
    LBS_WANTKEYBOARDINPUT = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_EXTENDEDSEL - Allows multiple selection if the SHIFT button is down
    LBS_EXTENDEDSEL = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_DISABLENOSCROLL - Show a disabled scroll bar if there aren't sufficient members to force a scrollbar
    LBS_DISABLENOSCROLL = &O1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_NODATA - The listbox does not manage the data for its members, to save memory for very large lists
    LBS_NODATA = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBS_NOSEL - Items cannot be selected
    LBS_NOSEL = &H4000
End Enum

' ##ENUMERATION_DESCRIPTION - ListboxMessages are windows message specifically for listbox controls or controls derived from listboxes
Public Enum ListboxMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_ADDSTRING - Adds a string to the control
    LB_ADDSTRING = &H180
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_INSERTSTRING - Inserts a string into the control
    LB_INSERTSTRING = &H181
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_DELETESTRING - Removes a string from the control
    LB_DELETESTRING = &H182
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SELITEMRANGEEX
    LB_SELITEMRANGEEX = &H183
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_RESETCONTENT - Clears the listbox
    LB_RESETCONTENT = &H184
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETSEL
    LB_SETSEL = &H185
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETCURSEL
    LB_SETCURSEL = &H186
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETSEL
    LB_GETSEL = &H187
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETCURSEL
    LB_GETCURSEL = &H188
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETTEXT
    LB_GETTEXT = &H189
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETTEXTLEN
    LB_GETTEXTLEN = &H18A
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETCOUNT - Gets the listcount
    LB_GETCOUNT = &H18B
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SELECTSTRING
    LB_SELECTSTRING = &H18C
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_DIR
    LB_DIR = &H18D
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETTOPINDEX - Returns the index of the top visible item in the list box
    LB_GETTOPINDEX = &H18E
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_FINDSTRING
    LB_FINDSTRING = &H18F
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETSELCOUNT
    LB_GETSELCOUNT = &H190
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETSELITEMS
    LB_GETSELITEMS = &H191
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETTABSTOPS
    LB_SETTABSTOPS = &H192
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETHORIZONTALEXTENT
    LB_GETHORIZONTALEXTENT = &H193
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETHORIZONTALEXTENT
    LB_SETHORIZONTALEXTENT = &H194
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETCOLUMNWIDTH
    LB_SETCOLUMNWIDTH = &H195
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_ADDFILE - Adds a file to a listbox which has is a directory list
    LB_ADDFILE = &H196
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETTOPINDEX - Scrolls the listbox such that the given index is visible
    LB_SETTOPINDEX = &H197
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETITEMRECT
    LB_GETITEMRECT = &H198
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETITEMDATA
    LB_GETITEMDATA = &H199
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETITEMDATA
    LB_SETITEMDATA = &H19A
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SELITEMRANGE
    LB_SELITEMRANGE = &H19B
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETANCHORINDEX - Sets the start of a multiple selection
    LB_SETANCHORINDEX = &H19C
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETANCHORINDEX - Gets the starts of a multiple selection
    LB_GETANCHORINDEX = &H19D
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETCARETINDEX - Selects the given item
    LB_SETCARETINDEX = &H19E
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETCARETINDEX - Returns the currentl selected item
    LB_GETCARETINDEX = &H19F
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETITEMHEIGHT - Sets the row height of an item (or all the items if the style is not LBS_OWNERDRAWVARIABLE)
    LB_SETITEMHEIGHT = &H1A0
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETITEMHEIGHT - Gets the row height of an item
    LB_GETITEMHEIGHT = &H1A1
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_FINDSTRINGEXACT -
    LB_FINDSTRINGEXACT = &H1A2
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETLOCALE - Sets the locale used by this listbox to sort
    LB_SETLOCALE = &H1A5
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_GETLOCALE - Gets the locale used by this listbox to sort
    LB_GETLOCALE = &H1A6
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_SETCOUNT - Sets the list count
    LB_SETCOUNT = &H1A7
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_INITSTORAGE - Pre allocate memory for a listbox to use to make it faster
    LB_INITSTORAGE = &H1A8
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_ITEMFROMPOINT - Get the item at a given coordinates
    LB_ITEMFROMPOINT = &H1A9
    ' ##ENUMERATION_MEMBER_DESCRIPTION LB_MULTIPLEADDSTRING - Add a number of entries at once
    LB_MULTIPLEADDSTRING = &H1B1
End Enum

' ##ENUMERATION_DESCRIPTION ListBoxNotifications - Messages sent by a listbox to its parent window when events occur
Public Enum ListBoxNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_DBLCLK - The listbox was double-clicked
    LBN_DBLCLK = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_ERRSPACE - The listbox ran out of memory
    LBN_ERRSPACE = -2
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_KILLFOCUS - The listbox has lost the focus
    LBN_KILLFOCUS = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_SELCANCEL - The listbox selection was cancelled
    LBN_SELCANCEL = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_SELCHANGE - The listbox selection was changed
    LBN_SELCHANGE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION LBN_SETFOCUS - The listbox gained the focus
    LBN_SETFOCUS = 4
End Enum

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As ListboxMessages, ByVal wParam As Long, ByVal lParam As Long) As Long
Private mHwnd As Long
Private bChanged As Boolean

Public Property Let ItemData(ByVal Index As Long, ByVal Data As Long)
Attribute ItemData.VB_HelpID = 2951

If mHwnd <> 0 Then
    ItemData(Index) = SendMessageLong(mHwnd, LB_SETITEMDATA, Index, Data)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ItemData", GetLastSystemError
    End If
End If

End Property

Public Property Get ItemData(ByVal Index As Long) As Long

If mHwnd <> 0 Then
    ItemData = SendMessageLong(mHwnd, LB_GETITEMDATA, Index, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ItemData", GetLastSystemError
    End If
End If

End Property

Public Property Get ListCount() As Long
Attribute ListCount.VB_HelpID = 2952

If mHwnd <> 0 Then
    ListCount = SendMessageLong(mHwnd, LB_GETCOUNT, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListCount", GetLastSystemError
    End If
End If

End Property


Public Property Let ListIndex(ByVal newIndex As Long)
Attribute ListIndex.VB_HelpID = 2954

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, LB_SETCURSEL, newIndex, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListIndex", GetLastSystemError
    End If
End If

End Property

Public Property Get ListIndex() As Long

If mHwnd <> 0 Then
    ListIndex = SendMessageLong(mHwnd, LB_GETCURSEL, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListIndex", GetLastSystemError
    End If
End If

End Property


Public Property Let Selected(ByVal Index As Long, ByVal SelState As Boolean)
Attribute Selected.VB_HelpID = 2956

If mHwnd <> 0 Then
    Selected(Index) = SendMessageLong(mHwnd, LB_SETSEL, Index, SelState)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Selected", GetLastSystemError
    End If
End If

End Property

Public Property Get Selected(ByVal Index As Long) As Boolean

If mHwnd <> 0 Then
    Selected = SendMessageLong(mHwnd, LB_GETSEL, Index, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Selected", GetLastSystemError
    End If
End If

End Property

Public Property Let TopVisibleIndex(ByVal newTop As Long)
Attribute TopVisibleIndex.VB_Description = "TopVisibleIndex - The index at the top of the list  "
Attribute TopVisibleIndex.VB_HelpID = 2958

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, LB_SETTOPINDEX, newTop, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TopVisibleIndex", GetLastSystemError
    End If
End If

End Property

' ##BLOCK_DESCRIPTION TopVisibleIndex - The index at the top of the list
' ##BLOCK_REMARKS The scroll box will scroll to show the given index
Public Property Get TopVisibleIndex() As Long

If mHwnd <> 0 Then
    TopVisibleIndex = SendMessageLong(mHwnd, LB_GETTOPINDEX, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TopVisibleIndex", GetLastSystemError
    End If
End If

End Property
Public Sub Clear()
Attribute Clear.VB_HelpID = 2959

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, LB_RESETCONTENT, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Clear", GetLastSystemError
    End If
End If

End Sub
Public Property Set Locale(ByVal newLocale As ApiLocale)
Attribute Locale.VB_Description = "Locale indicates the locale and therefore the sort order that the combo uses if it is sorted  "
Attribute Locale.VB_HelpID = 2961

If Not Locale Is Nothing Then
    If mHwnd <> 0 Then
        Call SendMessageLong(mHwnd, LB_SETLOCALE, newLocale.LocaleId, 0)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":Locale", GetLastSystemError
        End If
    End If
End If

End Property

' ##BLOCK_DESCRIPTION Locale indicates the locale and therefore the sort order that the combo uses if it is sorted
Public Property Get Locale() As ApiLocale

Dim lcThis As ApiLocale
Set lcThis = New ApiLocale

If mHwnd <> 0 Then
    lcThis.LocaleId = SendMessageLong(mHwnd, CB_GETLOCALE, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Locale", GetLastSystemError
    End If
    Set Locale = lcThis
End If

End Property
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property

Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Friend Property Set Window(ByVal newval As ApiWindow)

If newval Is Nothing Then
    mHwnd = 0
Else
    mHwnd = newval.hwnd
End If

End Property

Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_HelpID = 2962

If mHwnd <> 0 Then
    Dim lLen As String
    Dim sList As String
    lLen = SendMessageLong(mHwnd, LB_GETTEXTLEN, Index, 0)
    If lLen > 0 Then
        sList = String$(lLen, 0)
        Call SendMessageLong(mHwnd, LB_GETTEXT, Index, StrPtr(sList))
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":List", GetLastSystemError
        End If
    End If
End If

End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiComboControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2932
Attribute VB_Description="ApiComboControl - Class for manipulating windows which are based on the COMBO class"
Option Explicit

Private Const CLSNAME = "ApiComboControl"

' ##MODULE_DESCRIPTION ApiComboControl - Class for manipulating windows which are based on the COMBO class

' ##ENUMERATION_DESCRIPTION
Public Enum ComboboxControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_SIMPLE
    CBS_SIMPLE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_DROPDOWN
    CBS_DROPDOWN = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_DROPDOWNLIST
    CBS_DROPDOWNLIST = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_OWNERDRAWFIXED
    CBS_OWNERDRAWFIXED = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_OWNERDRAWVARIABLE
    CBS_OWNERDRAWVARIABLE = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_AUTOHSCROLL
    CBS_AUTOHSCROLL = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_OEMCONVERT
    CBS_OEMCONVERT = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_SORT
    CBS_SORT = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_HASSTRINGS
    CBS_HASSTRINGS = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_NOINTEGRALHEIGHT
    CBS_NOINTEGRALHEIGHT = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_DISABLENOSCROLL Shows a disabled scroll bar if the list is not longer than the space available
    CBS_DISABLENOSCROLL = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_UPPERCASE Converts all strings to upper case
    CBS_UPPERCASE = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBS_LOWERCASE Converts all strings to lower case
    CBS_LOWERCASE = &H4000
End Enum

' ##ENUMERATION_DESCRIPTION Windows messages specific to COMBOBOX class windows and window classes derived from that class
Public Enum enComboBoxMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_ADDSTRING
    CB_ADDSTRING = &H143
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_DELETESTRING
    CB_DELETESTRING = &H144
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_DIR
    CB_DIR = &H145
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_FINDSTRING
    CB_FINDSTRING = &H14C
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_FINDSTRINGEXACT
    CB_FINDSTRINGEXACT = &H158
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETCOUNT
    CB_GETCOUNT = &H146
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETCURSEL
    CB_GETCURSEL = &H147
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETDROPPEDCONTROLRECT
    CB_GETDROPPEDCONTROLRECT = &H152
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETEDITSEL
    CB_GETEDITSEL = &H140
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETLBTEXT
    CB_GETLBTEXT = &H148
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETLBTEXTLEN
    CB_GETLBTEXTLEN = &H149
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_INSERTSTRING
    CB_INSERTSTRING = &H14A
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_LIMITTEXT
    CB_LIMITTEXT = &H141
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_RESETCONTENT
    CB_RESETCONTENT = &H14B
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SELECTSTRING
    CB_SELECTSTRING = &H14D
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETCURSEL
    CB_SETCURSEL = &H14E
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETEDITSEL
    CB_SETEDITSEL = &H142
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SHOWDROPDOWN
    CB_SHOWDROPDOWN = &H14F
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETITEMDATA
    CB_GETITEMDATA = &H150
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETITEMDATA
    CB_SETITEMDATA = &H151
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETITEMHEIGHT
    CB_SETITEMHEIGHT = &H153
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETITEMHEIGHT
    CB_GETITEMHEIGHT = &H154
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETEXTENDEDUI
    CB_SETEXTENDEDUI = &H155
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETEXTENDEDUI
    CB_GETEXTENDEDUI = &H156
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETDROPPEDSTATE
    CB_GETDROPPEDSTATE = &H157
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETLOCALE
    CB_SETLOCALE = &H159
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETLOCALE
    CB_GETLOCALE = &H15A
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETTOPINDEX
    CB_GETTOPINDEX = &H15B
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETTOPINDEX
    CB_SETTOPINDEX = &H15C
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETHORIZONTALEXTENT
    CB_GETHORIZONTALEXTENT = &H15D
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETHORIZONTALEXTENT
    CB_SETHORIZONTALEXTENT = &H15E
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_GETDROPPEDWIDTH
    CB_GETDROPPEDWIDTH = &H15F
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_SETDROPPEDWIDTH
    CB_SETDROPPEDWIDTH = &H160
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_INITSTORAGE
    CB_INITSTORAGE = &H161
    ' ##ENUMERATION_MEMBER_DESCRIPTION CB_MULTIPLEADDSTRING
    CB_MULTIPLEADDSTRING = &H163
End Enum

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As enComboBoxMessages, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mHWND As Long
Private bChanged As Boolean

Private mMaxlength As Long

Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_HelpID = 2933

If mHWND <> 0 Then
    Dim lLen As String
    Dim sList As String
    lLen = SendMessageLong(mHWND, CB_GETLBTEXTLEN, Index, 0)
    If lLen > 0 Then
        sList = String$(lLen, 0)
        Call SendMessageLong(mHWND, CB_GETLBTEXT, Index, StrPtr(sList))
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":List", GetLastSystemError
        End If
    End If
End If

End Property
Public Property Get ListCount() As Long
Attribute ListCount.VB_HelpID = 2934

If mHWND <> 0 Then
    ListCount = SendMessageLong(mHWND, CB_GETCOUNT, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListCount", GetLastSystemError
    End If
End If

End Property

Public Property Let ItemData(ByVal Index As Long, ByVal Data As Long)
Attribute ItemData.VB_HelpID = 2936

If mHWND <> 0 Then
    ItemData(Index) = SendMessageLong(mHWND, CB_SETITEMDATA, Index, Data)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ItemData", GetLastSystemError
    End If
End If

End Property

Public Property Get ItemData(ByVal Index As Long) As Long

If mHWND <> 0 Then
    ItemData = SendMessageLong(mHWND, CB_GETITEMDATA, Index, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ItemData", GetLastSystemError
    End If
End If

End Property
Public Sub Clear()
Attribute Clear.VB_HelpID = 2937

If mHWND <> 0 Then
    Call SendMessageLong(mHWND, CB_RESETCONTENT, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Clear", GetLastSystemError
    End If
End If

End Sub


' ##BLOCK_DESCRIPTION DroppedControlRectangle - The rectangle that is occupied by the dropped down section of the combo when it is extended
' ##BLOCK_REMARKS This rectangle is in screen coordinates
Public Property Get DroppedControlRectangle() As APIRect
Attribute DroppedControlRectangle.VB_HelpID = 2938
Attribute DroppedControlRectangle.VB_Description = "DroppedControlRectangle - The rectangle that is occupied by the dropped down section of the combo when it is extended"

Dim lret As Long

If mHWND <> 0 Then
    Dim rcThis As New APIRect
    Call SendMessageLong(mHWND, CB_GETDROPPEDCONTROLRECT, 0, lret)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":DroppedControlRectangle", GetLastSystemError
    End If
    If lret <> 0 Then
        rcThis.CreateFromPointer (lret)
    End If
    Set DroppedControlRectangle = rcThis
End If

End Property


Public Property Let DroppedDown(ByVal dropped As Boolean)
Attribute DroppedDown.VB_HelpID = 2940
Attribute DroppedDown.VB_Description = "DroppedDown - Indicates whether a drop down combo is dropped down or not  "

If mHWND <> 0 Then
    Call SendMessageLong(mHWND, CB_SHOWDROPDOWN, dropped, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":DroppedDown", GetLastSystemError
    End If
End If

End Property

' ##BLOCK_DESCRIPTION DroppedDown - Indicates whether a drop down combo is dropped down or not
' ##BLOCK_REMARKS This is only valid for drop-down combo controls
Public Property Get DroppedDown() As Boolean

If mHWND <> 0 Then
    DroppedDown = SendMessageLong(mHWND, CB_GETDROPPEDSTATE, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":DroppedDown", GetLastSystemError
    End If
End If

End Property


Public Property Let ListIndex(ByVal newIndex As Long)
Attribute ListIndex.VB_HelpID = 2942

If mHWND <> 0 Then
    Call SendMessageLong(mHWND, CB_SETCURSEL, newIndex, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListIndex", GetLastSystemError
    End If
End If

End Property

Public Property Get ListIndex() As Long

If mHWND <> 0 Then
    ListIndex = SendMessageLong(mHWND, CB_GETCURSEL, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":ListIndex", GetLastSystemError
    End If
End If

End Property

Public Property Set Locale(ByVal newLocale As ApiLocale)
Attribute Locale.VB_HelpID = 2944
Attribute Locale.VB_Description = "Locale indicates the locale and therefore the sort order that the combo uses if it is sorted  "

If Not Locale Is Nothing Then
    If mHWND <> 0 Then
        Call SendMessageLong(mHWND, CB_SETLOCALE, newLocale.LocaleId, 0)
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

If mHWND <> 0 Then
    lcThis.LocaleId = SendMessageLong(mHWND, CB_GETLOCALE, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Locale", GetLastSystemError
    End If
    Set Locale = lcThis
End If

End Property


' ##BLOCK_DESCRIPTION Maxlength sets the maximum length of the edit portion of a combo control
' ##BLOCK_REMARKS By default the max length is 30,000 characters
Public Property Let Maxlength(ByVal newmax As Long)
Attribute Maxlength.VB_HelpID = 2946
Attribute Maxlength.VB_Description = "Maxlength sets the maximum length of the edit portion of a combo controlMaxlength sets the maximum length of the edit portion of a combo control"

If mMaxlength <> newmax Then
    mMaxlength = newmax
    Dim lret As Long
    If mHWND <> 0 Then
        lret = SendMessageLong(mHWND, CB_LIMITTEXT, mMaxlength, 0)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":Maxlength", GetLastSystemError
        End If
    End If
End If

End Property

Public Property Get Maxlength() As Long

    Maxlength = mMaxlength
    
End Property


Public Property Let TopVisibleIndex(ByVal newTop As Long)
Attribute TopVisibleIndex.VB_HelpID = 2948
Attribute TopVisibleIndex.VB_Description = "TopVisibleIndex - The index at the top of the drop down list  "

If mHWND <> 0 Then
    Call SendMessageLong(mHWND, CB_SETTOPINDEX, newTop, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TopVisibleIndex", GetLastSystemError
    End If
End If

End Property

' ##BLOCK_DESCRIPTION TopVisibleIndex - The index at the top of the drop down list
' ##BLOCK_REMARKS The scroll box will scroll to show the given index
Public Property Get TopVisibleIndex() As Long

If mHWND <> 0 Then
    TopVisibleIndex = SendMessageLong(mHWND, CB_GETTOPINDEX, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TopVisibleIndex", GetLastSystemError
    End If
End If

End Property



Friend Property Set Window(ByVal newVal As ApiWindow)

If newVal Is Nothing Then
    mHWND = 0
Else
    mHWND = newVal.hwnd
End If

End Property

Private Sub Class_Initialize()

'\\ default max lenght is 30000
mMaxlength = 30000

End Sub



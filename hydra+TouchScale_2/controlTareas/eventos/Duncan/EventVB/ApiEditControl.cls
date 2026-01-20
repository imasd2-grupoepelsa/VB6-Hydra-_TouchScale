VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiEditControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2970
Attribute VB_Description = "ApiEditControl - Class for manipulating windows which are based on the EDIT class"
Option Explicit

Private Const CLSNAME = "ApiEditControl"

' ##MODULE_DESCRIPTION ApiEditControl - Class for manipulating windows which are based on the EDIT class

' ##ENUMERATION_DESCRIPTION EditMessages - Windows messages used specifically in edit windows
Public Enum EditMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_CANUNDO Sent to establish whether the undo buffer has any contents
    EM_CANUNDO = &HC6
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_EMPTYUNDOBUFFER Sent to clear the undo buffer
    EM_EMPTYUNDOBUFFER = &HCD
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_FMTLINES Sent to switch on or off the soft linebreak formatting
    EM_FMTLINES = &HC8
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETFIRSTVISIBLELINE Sent to find out the line number of the top visible line
    EM_GETFIRSTVISIBLELINE = &HCE
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETHANDLE Sent to get a handle to the memory
    EM_GETHANDLE = &HBD
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETLINE
    EM_GETLINE = &HC4
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETLINECOUNT Sent to find out how many lines a multiline edit contol has
    EM_GETLINECOUNT = &HBA
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETMODIFY Sent to find out if the edit control has changed
    EM_GETMODIFY = &HB8
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETPASSWORDCHAR Sent to get the password character
    EM_GETPASSWORDCHAR = &HD2
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETRECT
    EM_GETRECT = &HB2
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETSEL
    EM_GETSEL = &HB0
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETTHUMB
    EM_GETTHUMB = &HBE
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_LINEFROMCHAR
    EM_LINEFROMCHAR = &HC9
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_LINEINDEX
    EM_LINEINDEX = &HBB
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_LINELENGTH
    EM_LINELENGTH = &HC1
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_LINESCROLL
    EM_LINESCROLL = &HB6
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_REPLACESEL
    EM_REPLACESEL = &HC2
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SCROLL
    EM_SCROLL = &HB5
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SCROLLCARET
    EM_SCROLLCARET = &HB7
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETHANDLE
    EM_SETHANDLE = &HBC
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETMODIFY
    EM_SETMODIFY = &HB9
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETPASSWORDCHAR
    EM_SETPASSWORDCHAR = &HCC
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETREADONLY - Used to toggle read only state of the text box
    EM_SETREADONLY = &HCF
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETRECT
    EM_SETRECT = &HB3
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETRECTNP
    EM_SETRECTNP = &HB4
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETSEL
    EM_SETSEL = &HB1
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETTABSTOPS
    EM_SETTABSTOPS = &HCB
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_UNDO
    EM_UNDO = &HC7
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETWORDBREAKPROC
    EM_GETWORDBREAKPROC = &HD1
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETWORDBREAKPROC
    EM_SETWORDBREAKPROC = &HD0
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETMARGINS
    EM_SETMARGINS = &HD3&
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETMARGINS
    EM_GETMARGINS = &HD4&
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_SETLIMITTEXT
    EM_SETLIMITTEXT = &HC5&
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_GETLIMITTEXT
    EM_GETLIMITTEXT = &HD5&
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_POSFROMCHAR
    EM_POSFROMCHAR = &HD6
    ' ##ENUMERATION_MEMBER_DESCRIPTION EM_CHARFROMPOS
    EM_CHARFROMPOS = &HD7
End Enum

' ##ENUMERATION_DESCRIPTION EditControlStyles - Styles used to modify the standard windows edit control
Public Enum EditControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_AUTOHSCROLL - The edit control scrolls horizontally when the edge of the control is reached
    ES_AUTOHSCROLL = &H80&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_AUTOVSCROLL - The edit control scrolls vertically when the bottom of the control is reached
    ES_AUTOVSCROLL = &H40&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_CENTER - Text is centered in the edit box
    ES_CENTER = &H1&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_LEFT - Text is left aligned
    ES_LEFT = &H0&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_LOWERCASE - Text is forced to lower case
    ES_LOWERCASE = &H10&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_MULTILINE - The text box is multiline
    ES_MULTILINE = &H4&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_NOHIDESEL - The selection is not hidden when the control loses focus
    ES_NOHIDESEL = &H100&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_OEMCONVERT - The control converts to the OEM character set
    ES_OEMCONVERT = &H400&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_PASSWORD - The control hides text with the password character
    ES_PASSWORD = &H20&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_READONLY - The control is read only
    ES_READONLY = &H800&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_RIGHT - The control is right aligned
    ES_RIGHT = &H2&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_UPPERCASE - The control forces the text to be upper case
    ES_UPPERCASE = &H8&
    ' ##ENUMERATION_MEMBER_DESCRIPTION ES_WANTRETURN - If the user presses ENTER a carriage return will be added to the text
    ES_WANTRETURN = &H1000&
End Enum

' ##ENUMERATION_DESCRIPTION EditBoxNotifications - Notifications passed to the parent window by an edit control
Public Enum EditBoxParentNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_CHANGE - The text in the textbox control has changed
    EN_CHANGE = &H300
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_ERRSPACE - The textbox has run out of memory
    EN_ERRSPACE = &H500
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_HSCROLL - The textbox was scrolled horizontally
    EN_HSCROLL = &H601
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_KILLFOCUS - The textbox lost the focus
    EN_KILLFOCUS = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_MAXTEXT - The textbox has reached it's predefined maximum length
    EN_MAXTEXT = &H501
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_SETFOCUS - The textbox has gained the input focus
    EN_SETFOCUS = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_UPDATE - The text has changed but the change has not yet been drawn
    EN_UPDATE = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION EN_VSCROLL - The textbox has been scrolled horizontally
    EN_VSCROLL = &H602
End Enum

' ##ENUMERATION_DESCRIPTION EditBoxmarginIndexes - Used in getting or setting edit box margins
Public Enum EditBoxmarginIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION EC_LEFTMARGIN - Get or set the left hand margin
    EC_LEFTMARGIN = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION EC_RIGHTMARGIN - Get or set the right hand margin
    EC_RIGHTMARGIN = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION EC_USEFONTINFO - Set the margin to the default for the control font
    EC_USEFONTINFO = &HFFFF
End Enum


Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As EditMessages, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mHwnd As Long
Private bChanged As Boolean

Public Property Get LineCount() As Long
Attribute LineCount.VB_HelpID = 2971

If mHwnd <> 0 Then
    LineCount = SendMessageLong(mHwnd, EM_GETLINECOUNT, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":LineCount", GetLastSystemError
    End If
End If

End Property


Public Property Let MarginLeft(ByVal newval As Integer)
Attribute MarginLeft.VB_HelpID = 3214

Call SendMessageLong(mHwnd, EM_SETMARGINS, EC_LEFTMARGIN, MakeLong(newval, 0))

End Property

Public Property Get MarginLeft() As Integer

MarginLeft = LoWord(SendMessageLong(mHwnd, EM_GETMARGINS, 0, 0))

End Property

Public Property Let MarginRight(ByVal newval As Integer)
Attribute MarginRight.VB_HelpID = 3216

Call SendMessageLong(mHwnd, EM_SETMARGINS, EC_RIGHTMARGIN, MakeLong(0, newval))

End Property

Public Property Get MarginRight() As Integer

MarginRight = HiWord(SendMessageLong(mHwnd, EM_GETMARGINS, 0, 0))

End Property


Public Property Let Maxlength(ByVal newval As Long)
Attribute Maxlength.VB_HelpID = 3218

Call SendMessageLong(mHwnd, EM_SETLIMITTEXT, newval, 0)

End Property

Public Property Get Maxlength() As Long

Maxlength = SendMessageLong(mHwnd, EM_GETLIMITTEXT, 0, 0)

End Property

Public Property Let PasswordChar(ByVal char As String)
Attribute PasswordChar.VB_HelpID = 2973

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, EM_SETPASSWORDCHAR, Asc(char), 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":PasswordChar", GetLastSystemError
    End If
End If

End Property

Public Property Get PasswordChar() As String

If mHwnd <> 0 Then
    Dim lRet As Long
    lRet = SendMessageLong(mHwnd, EM_GETPASSWORDCHAR, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":PasswordChar", GetLastSystemError
    End If
    PasswordChar = Chr$(lRet)
End If

End Property

Public Property Let ReadOnly(ByVal newval As Boolean)
Attribute ReadOnly.VB_HelpID = 3219

Call SendMessageLong(mHwnd, EM_SETREADONLY, Abs(newval), 0)

End Property

Public Property Let SelLength(ByVal newlen As Long)
Attribute SelLength.VB_HelpID = 2975

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, EM_SETSEL, SelStart, SelStart + newlen)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelLength", GetLastSystemError
    End If
End If

End Property

Public Property Get SelLength() As Long

If mHwnd <> 0 Then
    Dim lStart As Long, lEnd As Long
    Call SendMessageLong(mHwnd, EM_GETSEL, lStart, lEnd)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelLength", GetLastSystemError
    End If
    SelLength = (lEnd - lStart)
End If

End Property

Public Property Let SelStart(ByVal newStart As Long)
Attribute SelStart.VB_HelpID = 2977

If mHwnd <> 0 Then
    Call SendMessageLong(mHwnd, EM_SETSEL, newStart, newStart + SelLength)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelStart", GetLastSystemError
    End If
End If

End Property

Public Property Get SelStart() As Long

If mHwnd <> 0 Then
    Dim lStart As Long, lEnd As Long
    Call SendMessageLong(mHwnd, EM_GETSEL, lStart, lEnd)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelStart", GetLastSystemError
    End If
    SelStart = lStart
End If

End Property

Friend Property Set Window(ByVal newval As ApiWindow)

If newval Is Nothing Then
    mHwnd = 0
Else
    mHwnd = newval.hwnd
End If

End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property

Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiButtonControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2922
Attribute VB_Description = "ApiButtonControl - Class for manipulating windows which are based on the BUTTON class"
Option Explicit

' ##MODULE_DESCRIPTION ApiButtonControl - Class for manipulating windows which are based on the BUTTON class
' ##MODULE_REMARKS
Private Const CLSNAME = "ApiButtonControl"

' ##ENUMERATION_DESCRIPTION ButtonControlStyles - Style bits that modify the appearance
Public Enum ButtonControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_PUSHBUTTON - A standard style push button
    BS_PUSHBUTTON = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_DEFPUSHBUTTON - A default push button
    BS_DEFPUSHBUTTON = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_CHECKBOX - A checkbox type button
    BS_CHECKBOX = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_AUTOCHECKBOX - A check box that deselects others in the same group when selected
    BS_AUTOCHECKBOX = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_RADIOBUTTON - A radio (option) button
    BS_RADIOBUTTON = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_3STATE - A checkbox that can be dimmed when disabled
    BS_3STATE = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_AUTO3STATE - A check box that can be dimmed and that deselects others in the same group when selected
    BS_AUTO3STATE = &H6
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_GROUPBOX - The button is a button group control
    BS_GROUPBOX = &H7
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_USERBUTTON - No longer used.
    BS_USERBUTTON = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_AUTORADIOBUTTON - A radio button that deselects others in the same group when selected
    BS_AUTORADIOBUTTON = &H9
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_OWNERDRAW - The button owner is responsible for drawing it
    BS_OWNERDRAW = &HB
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_LEFTTEXT - For check boxes this makes the check box on the left hand side of the text
    BS_LEFTTEXT = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_TEXT - Button has text
    BS_TEXT = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_ICON - Button contains an icon image
    BS_ICON = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_BITMAP - Button contains a bitmap (i.e is graphical)
    BS_BITMAP = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_LEFT - Left align the button text
    BS_LEFT = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_RIGHT - Right align the button text
    BS_RIGHT = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_CENTER - Center (horizontally) the button text
    BS_CENTER = &H300
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_TOP - Top align the button text
    BS_TOP = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_BOTTOM - Bottom align the button text
    BS_BOTTOM = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_VCENTER - Vertically centre the button text
    BS_vCENTER = &HC00
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_PUSHLIKE - Makes a checkbox style button behave as a push button
    BS_PUSHLIKE = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_MULTILINE - Button text wraps if it is too long for a single line
    BS_MULTILINE = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_NOTIFY - Send notification messages to the button owner
    BS_NOTIFY = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_FLAT - Draws the button in the flat style
    BS_FLAT = &H8000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_RIGHTBUTTON - For check boxes this makes the check box on the right hand side of the text
    BS_RIGHTBUTTON = &H20
End Enum

' ##ENUMERATION_DESCRIPTION ButtonControlMessages - Predefined windows messages that are specifically to do with controls derived from the BUTTON class
Public Enum ButtonControlMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_GETCHECK - Sent to return the check box state (checked or unchecked)
    BM_GETCHECK = &HF0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_SETCHECK - Sent to set the check box state (checked or unchecked)
    BM_SETCHECK = &HF1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_GETSTATE - Sent to return the button control state
    BM_GETSTATE = &HF2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_SETSTATE - Sent to set the button control state
    BM_SETSTATE = &HF3
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_SETSTYLE - Sent to set the new button style of this button
    BM_SETSTYLE = &HF4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_CLICK - Simulates the user clicking this button
    BM_CLICK = &HF5
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_GETIMAGE - Sent to get the picture displayed by this button, if it is a graphical style button
    BM_GETIMAGE = &HF6
    ' ##ENUMERATION_MEMBER_DESCRIPTION BM_SETIMAGE - Sent to set the picture displayed by this button, if it is a graphical style button
    BM_SETIMAGE = &HF7
End Enum

' ##ENUMERATION_DESCRIPTION ButtonControlStates - The state of an owner drawn button that needs painting
Public Enum ButtonControlStates
    ' ##ENUMERATION_MEMBER_DESCRIPTION BST_UNCHECKED - The button is unchecked (assuming it is a checkbox style button)
    BST_UNCHECKED = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BST_CHECKED - The button is  checked (assuming it is a checkbox style button)
    BST_CHECKED = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BST_INDETERMINATE
    BST_INDETERMINATE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BST_PUSHED - The button is pushed in
    BST_PUSHED = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BST_FOCUS - The button has the focus
    BST_FOCUS = &H8
End Enum

' ##ENUMERATION_DESCRIPTION ButtonParentNotifications - Messages a button sends to it's parent window when events occur with it
Public Enum ButtonParentNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_CLICKED - A button was clicked
    BN_CLICKED = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_PAINT - A button needs repainting - This notification is obsolete
    BN_PAINT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_HILITE - A button state is set to pushed (Also known as BN_PUSHED)
    BN_HILITE = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_UNHILITE - A button state is set to unpushed (Also known as BN_UNPUSHED)
    BN_UNHILITE = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_DISABLE - A button has been disabled
    BN_DISABLE = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_DOUBLECLICKED - A button has been double-clicked
    BN_DOUBLECLICKED = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_SETFOCUS - A button has gained the focus
    BN_SETFOCUS = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION BN_KILLFOCUS - A button has lost the focus
    BN_KILLFOCUS = 7
End Enum

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As ButtonControlMessages, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLongApi Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private mHwnd As Long
Private bChanged As Boolean

Friend Property Let Changed(ByVal newval As Boolean)

    bChanged = newval

End Property

Friend Property Get Changed() As Boolean

    Changed = bChanged

End Property

Public Property Let Check(ByVal newCheck As ButtonControlStates)
Attribute Check.VB_HelpID = 2924

Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_SETCHECK, newCheck, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Check", GetLastSystemError
    End If
End If

End Property

Public Property Get Check() As ButtonControlStates
Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_GETCHECK, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Check", GetLastSystemError
    End If
    Check = lRet
End If

End Property

Public Sub Click()
Attribute Click.VB_HelpID = 2925

Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_CLICK, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Click", GetLastSystemError
    End If
End If

End Sub

Public Property Set Icon(ByVal newIcon As ApiIcon)
Attribute Icon.VB_HelpID = 2927

Dim lRet As Long
Dim hIcon As Long

If mHwnd <> 0 Then
    If newIcon Is Nothing Then
        hIcon = 0
        UnsetButtonStyle BS_ICON
    Else
        hIcon = newIcon.hIcon
        SetButtonStyle BS_ICON
    End If
    lRet = SendMessageLong(mHwnd, BM_SETIMAGE, IMAGE_ICON, hIcon)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Icon", GetLastSystemError
    End If
End If

End Property

Public Property Get Icon() As ApiIcon

If mHwnd <> 0 Then
    Dim icoThis As ApiIcon
    Dim lRet As Long
    
    lRet = SendMessageLong(mHwnd, BM_GETIMAGE, IMAGE_ICON, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Icon", GetLastSystemError
    Else
        If lRet <> 0 Then
            Set icoThis = New ApiIcon
            icoThis.hIcon = lRet
            Set Icon = icoThis
        End If
    End If
End If

End Property


Public Property Set Image(ByVal newImage As ApiBitmap)
Attribute Image.VB_HelpID = 2929

Dim lRet As Long
Dim hPic As Long

If mHwnd <> 0 Then
    If newImage Is Nothing Then
        hPic = 0
        UnsetButtonStyle BS_BITMAP
    Else
        hPic = newImage.hBitmap
        SetButtonStyle BS_BITMAP
    End If
    lRet = SendMessageLong(mHwnd, BM_SETIMAGE, IMAGE_BITMAP, hPic)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Image", GetLastSystemError
    End If
End If

End Property

Public Property Get Image() As ApiBitmap

If mHwnd <> 0 Then
    Dim bmpThis As ApiBitmap
    Dim lRet As Long
    
    lRet = SendMessageLong(mHwnd, BM_GETIMAGE, IMAGE_BITMAP, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Image", GetLastSystemError
    Else
        If lRet <> 0 Then
            Set bmpThis = New ApiBitmap
            bmpThis.hBitmap = lRet
            Set Image = bmpThis
        End If
    End If
End If

End Property






Public Property Get IsButtonStyleSet(ByVal TestStyle As ButtonControlStyles)
Attribute IsButtonStyleSet.VB_HelpID = 3209

IsButtonStyleSet = (Style And TestStyle)

End Property

Public Sub SetButtonStyle(ByVal NewStyle As ButtonControlStyles)
Attribute SetButtonStyle.VB_HelpID = 3210

If Not IsButtonStyleSet(NewStyle) Then
    Style = (NewStyle Or Style)
End If

End Sub

Public Property Get Style() As Long
Attribute Style.VB_HelpID = 2930

Style = GetWindowLongApi(mHwnd, GWL_STYLE)


End Property

Public Property Let Style(ByVal Style As Long)

Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_SETSTYLE, Style, MakeLong(1, 0))
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Style", GetLastSystemError
    End If
End If

End Property


Public Sub UnsetButtonStyle(ByVal NewStyle As ButtonControlStyles)
Attribute UnsetButtonStyle.VB_HelpID = 3211

If IsButtonStyleSet(NewStyle) Then
    Style = Style Xor NewStyle
End If

End Sub

Friend Property Set Window(ByVal newval As ApiWindow)

If newval Is Nothing Then
    mHwnd = 0
Else
    mHwnd = newval.hwnd
End If

End Property

Public Property Get State() As ButtonControlStates
Attribute State.VB_HelpID = 2931

Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_GETSTATE, 0, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":State", GetLastSystemError
    Else
        State = lRet
    End If
End If

End Property

Public Property Let State(ByVal NewState As ButtonControlStates)

Dim lRet As Long

If mHwnd <> 0 Then
    lRet = SendMessageLong(mHwnd, BM_SETSTATE, NewState, 0)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":State", GetLastSystemError
    End If
End If

End Property


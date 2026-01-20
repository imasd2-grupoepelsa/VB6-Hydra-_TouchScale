VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiCommCtlPager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2540
Attribute VB_Description = "This class provides methods and properties for manipulating instances of the "
Option Explicit

' ##MODULE_DESCRIPTION This class provides methods and properties for manipulating _
instances of the "Pager" windows custom control

Private mChildWnd As Long
Private Const CLSNAME = "ApiCommCtlPager:"

Private mWND As ApiWindow
Attribute mWND.VB_VarHelpID = -1
Private WithEvents wndParent As ApiWindow
Attribute wndParent.VB_VarHelpID = -1

Private Const PGN_FIRST = -900

' ##ENUMERATION_DESCRIPTION PagerNotificationMessages - The notification messages for a pager control
Public Enum PagerNotificationMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGN_SCROLL - The pager was scrolled
    PGN_SCROLL = (PGN_FIRST - 1)
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGN_CALCSIZE - The pager size was recalculated
    PGN_CALCSIZE = (PGN_FIRST - 2)
End Enum

Private Const PGM_FIRST = &H1400

' ##ENUMERATION_DESCRIPTION PagerMessages - Special windows messages pertaining to the scroller control window
Public Enum PagerMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_SETCHILD - Set the window that is paged by this pager control
    PGM_SETCHILD = &H1401      '\\ lParam == hwnd
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_RECALCSIZE - Recalculate the pager control size
    PGM_RECALCSIZE = &H1402
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_FORWARDMOUSE - Allow or prevent mousemove messages being sent to the child window
    PGM_FORWARDMOUSE = &H1403  '\\ wParam = bForward
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_SETBKCOLOR - Sets the background colour
    PGM_SETBKCOLOR = &H1404    '\\ lParam = clr
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETBKCOLOR - Returns the background colour
    PGM_GETBKCOLOR = &H1405
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_SETBORDER - Sets the border size for the control
    PGM_SETBORDER = &H1406   '\\ lParam = iBorder
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETBORDER - Returns the border size for the control
    PGM_GETBORDER = &H1407
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_SETPOS - Sets the scroll position (in pixels)
    PGM_SETPOS = &H1408        '\\ lParam = iPos
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETPOS - Returns the scroll position (in pixels)
    PGM_GETPOS = &H1409
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_SETBUTTONSIZE - Sets the size of the scroller button
    PGM_SETBUTTONSIZE = &H1410 '\\lParam = Size
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETBUTTONSIZE - Returns the size of the scroller button
    PGM_GETBUTTONSIZE = &H1411
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETBUTTONSTATE - Returns the state of the selected button
    PGM_GETBUTTONSTATE = &H1412    '\\ lParam = iButton
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGM_GETDROPTARGET
    PGM_GETDROPTARGET = &H2004
End Enum

' ##ENUMERATION_DESCRIPTION PagerScrollDirections - The direction in which a scroll event took place
Public Enum PagerScrollDirections
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_SCROLLUP Pager was scrolled up
    PGF_SCROLLUP = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_SCROLLDOWN Pager was scrolled down
    PGF_SCROLLDOWN = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_SCROLLLEFT Pager was scrolled left
    PGF_SCROLLLEFT = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_SCROLLRIGHT Pager was scrolled right
    PGF_SCROLLRIGHT = 8
End Enum

' ##ENUMERATION_DESCRIPTION PagerScrollKeys - Keys pressed when a scroll event occured
Public Enum PagerScrollKeys
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGK_SHIFT - The shift key was down
    PGK_SHIFT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGK_CONTROL - The control key was down
    PGK_CONTROL = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGK_MENU - The ALT key was down
    PGK_MENU = 4
End Enum

' ##ENUMERATION_DESCRIPTION PagerCalcSizeDirection - Which dimension to calculate the pager size
Public Enum PagerCalcSizeDirection
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_CALCWIDTH - Calculate the width
    PGF_CALCWIDTH = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_CALCHEIGHT - Calculate the height
    PGF_CALCHEIGHT = 2
End Enum

' ##ENUMERATION_DESCRIPTION PagerControlStyles - Additional window style flags pertaining to the pager control
Public Enum PagerControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGS_VERT - Pager scrolls vertically
    PGS_VERT = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGS_HORZ - Pager scrolls horizontally
    PGS_HORZ = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGS_AUTOSCROLL - Pager scrolls automatically
    PGS_AUTOSCROLL = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGS_DRAGNDROP - Control supports drag and drop operations
    PGS_DRAGNDROP = &H4
End Enum

' ##ENUMERATION_DESCRIPTION PagerButtonStates - The state of the pager control buttons
Public Enum PagerButtonStates
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_INVISIBLE - Scroll button is not visible
    PGF_INVISIBLE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_NORMAL - Scroll button is in normal state
    PGF_NORMAL = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_GRAYED - Scroll button is in grayed state
    PGF_GRAYED = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_DEPRESSED - Scroll button is in depressed state
    PGF_DEPRESSED = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGF_HOT - Scroll button is in hot state
    PGF_HOT = 8
End Enum

' ##ENUMERATION_DESCRIPTION PagerButtonid - The button which is being queried
Public Enum PagerButtonid
    ' ##ENUMERATION_MEMBER_DESCRIPTION PGB_TOPORLEFT - The top or left button
    PGB_TOPORLEFT = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTIONPGB_BOTTOMORRIGHT - The bottom or right button
    PGB_BOTTOMORRIGHT = 1
End Enum

Private Type NMPGSCROLL
    NMHDR_hwndFrom As Long
    NMHDR_idfrom As Long
    NMHDR_code As Long
    fwKeys As Long              ' ##  Specifies which keys are down when this notification is send
    rcParent_Left As Long ' ##  Contains Parent Window Rect
    rcParent_Top As Long
    rcParent_Right As Long
    rcParent_Bottom As Long
    iDir As Long              ' ##  Scrolling Direction
    iXpos As Long            ' ##  Horizontal scroll position
    iYpos As Long             ' ##  Vertical scroll position
    iScroll As Long           ' ##  [in/out] Amount to scroll
End Type
Private Declare Sub CopyMemoryNMPGSCROLL Lib "kernel32" Alias "RtlMoveMemory" (Destination As NMPGSCROLL, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryToNMPGSCROLL Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As NMPGSCROLL, ByVal Length As Long)

' ##EVENT_DESCRIPTION PagerScroll - Triggered when the pager is being scrolled
' ##EVENT_REMARKS Direction - The direction of the scroll
' ##EVENT_REMARKS PrevX - The X position before the scroll
' ##EVENT_REMARKS PrevY - The Y position to scroll
' ##EVENT_REMARKS ScrollAmount - The number of pixels that will be scrolled - set this to zero to cancel the scroll
Public Event PagerScroll(ByVal Direction As PagerScrollDirections, ByVal PrevX As Long, ByVal PrevY As Long, ScrollAmount As Long)


Private Type NMPGCALCSIZE
    NMHDR_hwndFrom As Long
    NMHDR_idfrom As Long
    NMHDR_code As Long
    dwFlag As PagerCalcSizeDirection
    iWidth As Long
    iHeight As Long
End Type
Private Declare Sub CopyMemoryNMPGCALCSIZE Lib "kernel32" Alias "RtlMoveMemory" (Destination As NMPGCALCSIZE, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryToNMPGCALCSIZE Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As NMPGCALCSIZE, ByVal Length As Long)
Private mDefWidth As Long
Private mDefHeight As Long

' ##EVENT_DESCRIPTION CalculateSize - Triggered when the pager control wants to know it's size
' ##EVENT_REMARKS Direction - This depends on whether the pager control is vertical or horizontal
' ##EVENT_REMARKS Size - the size (in pixels) of the height or width
Public Event CalculateSize(ByVal Direction As PagerCalcSizeDirection, Size As Long)


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private bChanged As Boolean



Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Let BackgroundColour(ByVal newColour As ApiColour)
Attribute BackgroundColour.VB_HelpID = 2542

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_SETBKCOLOR, 0, ByVal newColour.ColourRef)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "BackgroundColour", GetLastSystemError
End If

End Property

Public Property Get BackgroundColour() As ApiColour

Dim lret As Long
Dim colThis As New ApiColour

lret = SendMessage(mWND.hwnd, PGM_GETBKCOLOR, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "BackgroundColour", GetLastSystemError
End If
colThis.ColourRef = lret
Set BackgroundColour = colThis

End Property

Public Property Let Border(ByVal newBorder As Long)
Attribute Border.VB_HelpID = 2544

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_SETBORDER, 0, ByVal newBorder)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "Border", GetLastSystemError
End If

End Property

Public Property Get Border() As Long

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_GETBORDER, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "Border", GetLastSystemError
End If
Border = lret

End Property


Public Property Let ButtonSize(ByVal newButtonSize As Long)
Attribute ButtonSize.VB_HelpID = 2546

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_SETBUTTONSIZE, 0, ByVal newButtonSize)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "ButtonSize", GetLastSystemError
End If
mWND.Refresh

End Property

Public Property Get ButtonSize() As Long

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_GETBUTTONSIZE, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "ButtonSize", GetLastSystemError
End If
ButtonSize = lret

End Property

Public Property Get ChildWindow() As ApiWindow
Attribute ChildWindow.VB_HelpID = 2547

Dim wndChild As ApiWindow

Set wndChild = New ApiWindow

wndChild.hwnd = mChildWnd
Set ChildWindow = wndChild

Set wndChild = Nothing

End Property

Public Property Set ChildWindow(ByVal hwndChild As ApiWindow)

Dim lret As Long

If hwndChild Is Nothing Then
    lret = SendMessage(mWND.hwnd, PGM_SETCHILD, 0, ByVal 0&)
    mChildWnd = 0
Else
    lret = SendMessage(mWND.hwnd, PGM_SETCHILD, 0, ByVal hwndChild.hwnd)
    mChildWnd = hwndChild.hwnd
    Set hwndChild.Parent = mWND
    With hwndChild.RECT
        mDefHeight = (.Bottom - .Top) * 2
        mDefWidth = (.Right - .Left) * 2
    End With
End If
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "ChildWindow", GetLastSystemError
End If
    
End Property

Public Property Get ClassName() As String
Attribute ClassName.VB_HelpID = 2548

    ClassName = "SysPager"
    
End Property




' ##BLOCK_DESCRIPTION CreateNew - Creates a new pager control
' ##BLOCK_REMARKS ParentWindow - The window that will contain this pager window
' ##BLOCK_REMARKS ChildWindowIn - The window that will be paged by this control
' ##BLOCK_REMARKS Position - Where the window will be positioned relative to the client area (in pixels)
' ##BLOCK_REMARKS Style - The control styles to apply - note that PGS_VERT ans PGS_HORZ are mutually exclusive
Public Sub CreateNew(ByVal ParentWindow As ApiWindow, ByVal ChildWindowIn As ApiWindow, _
                     ByVal Position As APIRect, ByVal Style As PagerControlStyles)


Dim clsThis As APIWndClass

Set clsThis = New APIWndClass
With clsThis
    .CreateFromClassname 0, ClassName
    .Name = ClassName
    .Style = Style
End With

With Position
    Set mWND = APIDispenser.System.WindowDispenser.CreateWindow(0, clsThis, "", WS_CHILD Or Style, .Left, .Top, .Right - .Left, .Bottom - .Top, ParentWindow.hwnd, 0, 0, 0)
    mWND.Show SW_NORMAL
End With
Set ChildWindow = ChildWindowIn

Set wndParent = ParentWindow
APIDispenser.SubclassedWindows.Add wndParent

End Sub

Public Property Let Window(ByVal newHwnd As ApiWindow)
Attribute Window.VB_HelpID = 2819

        Set mWND = newHwnd
    
End Property
Public Property Let Position(ByVal newPosition As Long)
Attribute Position.VB_HelpID = 2551


Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_SETPOS, 0, ByVal newPosition)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "Position", GetLastSystemError
End If

End Property


Public Property Get Position() As Long

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_GETPOS, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "Position", GetLastSystemError
End If
Position = lret

End Property


Public Sub RecalcSize()
Attribute RecalcSize.VB_HelpID = 2552

Dim lret As Long

lret = SendMessage(mWND.hwnd, PGM_RECALCSIZE, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & "RecalcSize", GetLastSystemError
End If

End Sub


Public Property Get Window() As ApiWindow

    Set Window = mWND
    
End Property

Private Sub Class_Initialize()

'\\ make sure the common controls class is initialised
Call APIDispenser.InitCommonControls(ICC_PAGESCROLLER_CLASS)

End Sub


Private Sub Class_Terminate()

Set wndParent = Nothing
Set mWND = Nothing

End Sub






Private Sub wndParent_WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)

Dim nmThis As ApiNotifyMessageHeader

If msg = WM_NOTIFY Then
    Set nmThis = New ApiNotifyMessageHeader
    nmThis.CreateFromPointer lParam
    If nmThis.Code = PGN_SCROLL Then
        '\\ A scroll notification was recieved
        Dim snThis As NMPGSCROLL
        Call CopyMemoryNMPGSCROLL(snThis, lParam, Len(snThis))
        With snThis
            RaiseEvent PagerScroll(.iDir, .iXpos, .iYpos, .iScroll)
        End With
        Call CopyMemoryToNMPGSCROLL(lParam, snThis, Len(snThis))
        
    ElseIf nmThis.Code = PGN_CALCSIZE Then
        '\\ a CalcSize notification was recieved
        Dim csThis As NMPGCALCSIZE
        Call CopyMemoryNMPGCALCSIZE(csThis, lParam, Len(csThis))
        With csThis
            If .dwFlag = PGF_CALCHEIGHT Then
                .iHeight = mDefHeight
                RaiseEvent CalculateSize(.dwFlag, .iHeight)
            Else
                .iWidth = mDefWidth
                RaiseEvent CalculateSize(.dwFlag, .iWidth)
            End If
        End With
        Call CopyMemoryToNMPGCALCSIZE(lParam, csThis, Len(csThis))
    End If
End If

End Sub



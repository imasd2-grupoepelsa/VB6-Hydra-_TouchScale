VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiStaticControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 3220
Option Explicit

Private Const CLSNAME = "ApiStaticControl"

' ##ENUMERATION_DESCRIPTION StaticControlStyles - The window style bits specific to static style windows and derived classes
Public Enum StaticControlStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_LEFT - Left aligned text
    SS_LEFT = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_CENTER - Centre aligned text
    SS_CENTER = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_RIGHT - Right aligned text
    SS_RIGHT = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ICON - Control has an icon
    SS_ICON = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_BLACKRECT - Control is a solid rectangle drawn in the window frame colour
    SS_BLACKRECT = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_GRAYRECT - Control is a solid rectangle drawn in the window background colour
    SS_GRAYRECT = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_WHITERECT - Control is a solid rectangle drawn in the window client area colour
    SS_WHITERECT = &H6
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_BLACKFRAME - Control is a frame drawn in the window frame colour
    SS_BLACKFRAME = &H7
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_GRAYFRAME  - Control is a frame drawn in the window background colour
    SS_GRAYFRAME = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_WHITEFRAME - Control is a frame drawn in the window foreground colour
    SS_WHITEFRAME = &H9
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_USERITEM - Control is a user defined item
    SS_USERITEM = &HA
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_SIMPLE - Control is a label
    SS_SIMPLE = &HB
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_LEFTNOWORDWRAP - Designates a simple rectangle and displays the given text flush-left in the rectangle
    SS_LEFTNOWORDWRAP = &HC
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_OWNERDRAW - Control is drawn by its parent
    SS_OWNERDRAW = &HD
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_BITMAP - Control has a bitmap picture
    SS_BITMAP = &HE
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ENHMETAFILE - Control has an enhanced metafile picture
    SS_ENHMETAFILE = &HF
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ETCHEDHORZ - Draws the top and bottom of the frame etched
    SS_ETCHEDHORZ = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ETCHEDVERT - Draws the left and right edges of the frame etched
    SS_ETCHEDVERT = &H11
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ETCHEDFRAME - Draws the control frame etched
    SS_ETCHEDFRAME = &H12
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_TYPEMASK
    SS_TYPEMASK = &H1F
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_REALSIZECONTROL
    SS_REALSIZECONTROL = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_NOPREFIX - Don't interpret & as underline prefix
    SS_NOPREFIX = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_NOTIFY - Notify the controls parent if it is clicked
    SS_NOTIFY = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_CENTERIMAGE - Centre the control picture
    SS_CENTERIMAGE = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_RIGHTJUST - Right justify the control
    SS_RIGHTJUST = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_REALSIZEIMAGE - Clips the control contents (icon or bitmap) rather than expanding the control
    SS_REALSIZEIMAGE = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_SUNKEN - Draws a half sunken border around the control
    SS_SUNKEN = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ENDELLIPSIS
    SS_ENDELLIPSIS = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_PATHELLIPSIS
    SS_PATHELLIPSIS = &H8000
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_WORDELLIPSIS - Truncates text that does not fit and adds ellipses
    SS_WORDELLIPSIS = &HC000
    ' ##ENUMERATION_MEMBER_DESCRIPTION SS_ELLIPSISMASK
    SS_ELLIPSISMASK = &HC000
End Enum

' ##ENUMERATION_DESCRIPTION
Public Enum StaticControlMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION STM_SETICON - Sets the icon for the control
    STM_SETICON = &H170
    ' ##ENUMERATION_MEMBER_DESCRIPTION STM_GETICON - Returns the icon for the control
    STM_GETICON = &H171
    ' ##ENUMERATION_MEMBER_DESCRIPTION STM_SETIMAGE - Sets the picture for the control
    STM_SETIMAGE = &H172
    ' ##ENUMERATION_MEMBER_DESCRIPTION STM_GETIMAGE - Returns the image for the control
    STM_GETIMAGE = &H173
End Enum

' ##ENUMERATION_DESCRIPTION StaticControlNotifications - The events a static control notifies its parent about
Public Enum StaticControlNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION STN_CLICKED - The static control was clicked
     STN_CLICKED = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION STN_DBLCLK - The static control was double-clicked
     STN_DBLCLK = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION STN_ENABLE - The static control was enabled
     STN_ENABLE = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION STN_DISABLE - The static control was disabled
     STN_DISABLE = 3
End Enum

Private mHwnd As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Property Let Bitmap(ByVal bmThis As ApiBitmap)
Attribute Bitmap.VB_HelpID = 3222

    If Not bmThis Is Nothing Then
        Call SendMessageLong(mHwnd, STM_SETIMAGE, bmThis.hBitmap, 0)
    Else
        Call SendMessageLong(mHwnd, STM_SETIMAGE, 0, 0)
    End If

End Property

Public Property Get Bitmap() As ApiBitmap

Dim lhBitmap As Long
lhBitmap = SendMessageLong(mHwnd, STM_GETIMAGE, 0, 0)
If lhBitmap <> 0 Then
    Dim icThis As New ApiBitmap
    icThis.hBitmap = lhBitmap
    Set Bitmap = icThis
End If

End Property

Public Property Let Icon(ByVal icThis As ApiIcon)
Attribute Icon.VB_HelpID = 3224

    If Not icThis Is Nothing Then
        Call SendMessageLong(mHwnd, STM_SETICON, icThis.hIcon, 0)
    Else
        Call SendMessageLong(mHwnd, STM_SETICON, 0, 0)
    End If
    
End Property


Public Property Get Icon() As ApiIcon

Dim lhIcon As Long
lhIcon = SendMessageLong(mHwnd, STM_GETICON, 0, 0)
If lhIcon <> 0 Then
    Dim icThis As New ApiIcon
    icThis.hIcon = lhIcon
    Set Icon = icThis
End If

End Property


Friend Property Set Window(ByVal newval As ApiWindow)

If newval Is Nothing Then
    mHwnd = 0
Else
    mHwnd = newval.hwnd
End If

End Property


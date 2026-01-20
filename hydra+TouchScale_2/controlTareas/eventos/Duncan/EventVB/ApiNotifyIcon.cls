VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiNotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2468
Attribute VB_Description = "This class provides the properties and methods for installing an icon in the system tray (the area next to the clock) and responding when the user interacts with that icon"
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for installing an icon in the system tray (the area next to the clock) _
and responding when the user interacts with that icon

' ##MODULE_DESCRIPTION This allows you to write applications which run in the _
background and allow the user to interact with them through this icon.

'\\ To put this app in the Task Tray:
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public CreatedOK As Boolean
Attribute CreatedOK.VB_VarHelpID = 2469

'\\ Private memory handling functions
Private Declare Sub CopyMemoryNotifyIconData Lib "kernel32" Alias "RtlMoveMemory" (Destination As NOTIFYICONDATA, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrNotifyIconData Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrNotifyIconData Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

' ##ENUMERATION_DESCRIPTION ShellNotifyMessages - The different messages that can be sent to the shell notification window
Public Enum ShellNotifyMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIM_ADD - Add an icon to the shell notification area
    NIM_ADD = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIM_MODIFY - Modify an entry in the shell notification area
    NIM_MODIFY = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIM_DELETE - Remove an icon from the shell notification area
    NIM_DELETE = &H2
End Enum

' ##ENUMERATION_DESCRIPTION ShellNotifyFormats - The make up of the message sent to the shell notify window
Public Enum ShellNotifyFormats
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIF_MESSAGE - Message contains the message identifier used to communicate between the notify icon and it's parent application
    NIF_MESSAGE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIF_ICON - Message contains icon information
    NIF_ICON = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION NIF_TIP - Message contains tooltip information
    NIF_TIP = &H4
End Enum

'\\ member variables
Private mszTip As String
Private mcbSize As Long
Private mHwnd As Long
Private muID As Long
Private muFlags As Long
Private mhIcon As Long
Private muCallbackMessage As Long
Private mAdded As Boolean

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private mMessage As Long
Private bChanged As Boolean

'\\ Each ApiNotifyIcon now has its own hidden message handling window
Private WithEvents wndThis As ApiWindow
Attribute wndThis.VB_VarHelpID = -1

' ##EVENT_DESCRIPTION MouseMove - The mouse pointer was moved over the icon in the shell notify area
Public Event MouseMove()
Attribute MouseMove.VB_Description = "MouseMove - The mouse pointer was moved over the icon in the shell notify area"
Attribute MouseMove.VB_HelpID = 3089

' ##EVENT_DESCRIPTION MouseDown - The mouse button was pressed over the icon in the shell notify area
Public Event MouseDown(ByVal Button As Integer)
Attribute MouseDown.VB_Description = "MouseDown - The mouse button was pressed over the icon in the shell notify area"
Attribute MouseDown.VB_HelpID = 3090

' ##EVENT_DESCRIPTION MouseUp - The mouse button was released over the icon in the shell notify area
Public Event MouseUp(ByVal Button As Integer)
Attribute MouseUp.VB_Description = "MouseUp - The mouse button was released over the icon in the shell notify area"
Attribute MouseUp.VB_HelpID = 3091

' ##EVENT_DESCRIPTION DblClick - The icon in the shell notify area was double-clicked
Public Event DblClick(ByVal Button As Integer)
Attribute DblClick.VB_Description = "DblClick - The icon in the shell notify area was double-clicked"
Attribute DblClick.VB_HelpID = 3092
Private Property Get NotifyWindow() As Long

    If mHwnd = 0 Then
        Dim clsHotkey As APIWndClass
        Set clsHotkey = New APIWndClass
        With clsHotkey
            .hInstance = App.hInstance
            .Name = "STATIC"
        End With
        '\\ Create the window to recieve the messages for this hotkey
        Set wndThis = APIDispenser.System.WindowDispenser.CreateWindow(0, clsHotkey, "MCL Hotkey", 0, 0, 0, 1, 1, 0, 0, 0, 0)
        APIDispenser.SubclassedWindows.Add wndThis
        mHwnd = wndThis.hwnd
    End If

    NotifyWindow = mHwnd
    
End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this NotifyIconData object from the location poiunted to by
'\\ lpPointAPI
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corPointAPIness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpNotifyIconData As Long) As Boolean

Dim niThis As NOTIFYICONDATA

CreatedOK = False

If Not IsBadReadPtrNotifyIconData(lpNotifyIconData, Len(niThis)) Then
    Call CopyMemoryNotifyIconData(niThis, lpNotifyIconData, Len(niThis))
    If Err.LastDllError = 0 Then
        With niThis
            
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Property Set Icon(ByVal newIcon As ApiIcon)
Attribute Icon.VB_Description = "Icon - The icon displayed in the shell notification window  "
Attribute Icon.VB_HelpID = 2471

If mhIcon <> newIcon.hIcon Then
    mhIcon = newIcon.hIcon
End If

End Property

' ##BLOCK_DESCRIPTION Icon - The icon displayed in the shell notification window
Public Property Get Icon() As ApiIcon

Dim myIcon As ApiIcon

Set myIcon = New ApiIcon
myIcon.hIcon = mhIcon
Set Icon = myIcon

End Property




Public Property Get NotifyWindowMessage() As Long
Attribute NotifyWindowMessage.VB_HelpID = 2473

    If mMessage = 0 Then
        mMessage = RegisterWindowMessage("APINOTIFYICONMSG")
        If Err.LastDllError > 0 Then
            ReportError Err.LastDllError, "ApiNotifyIcon:NotifyWindowMessage", GetLastSystemError
        End If
    End If
    NotifyWindowMessage = mMessage
    
End Property

Public Sub RefreshNotifyIcon()
Attribute RefreshNotifyIcon.VB_HelpID = 2474

Dim niThis As NOTIFYICONDATA
Dim lret As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = NotifyWindowMessage
End If

If mAdded Then
    With niThis
        .hIcon = mhIcon
        .hwnd = mHwnd
        .szTip = mszTip
        .uCallbackMessage = muCallbackMessage
        .uFlags = muFlags
        .uID = muID
        .cbSize = Len(niThis)
    End With
    
    lret = Shell_NotifyIcon(NIM_MODIFY, niThis)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, "ApiNotifyIcon:RefreshNotifyIcon", GetLastSystemError
    End If
End If

End Sub

'\\ --[SetNotifyIcon]--------------------------------------------------------
'\\ Sets the icon (and other info) from this ApiNotifyIcon
'\\ into the system tray
'\\ -------------------------------------------------------------------------
Public Sub Show()
Attribute Show.VB_HelpID = 2475

Dim niThis As NOTIFYICONDATA
Dim lret As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = NotifyWindowMessage
End If

With niThis
    .hIcon = mhIcon
    .hwnd = NotifyWindow
    .szTip = mszTip
    .uCallbackMessage = muCallbackMessage
    .uFlags = muFlags
    .uID = muID
    .cbSize = Len(niThis)
End With

lret = Shell_NotifyIcon(NIM_ADD, niThis)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiNotifyIcon:SetNotifyIcon", GetLastSystemError
Else
    mAdded = True
End If

End Sub

Public Property Get ToolTip() As String
Attribute ToolTip.VB_HelpID = 2476

    ToolTip = mszTip
    
End Property

Public Property Let ToolTip(ByVal NewTip As String)

    '\\ Limit tootltip to 63 chars + null terminator...
    If Len(NewTip) > 63 Then
        NewTip = Left$(NewTip, 63)
    End If
    
    '\\ Initialise all to nulls
    mszTip = NewTip & String$(64 - Len(NewTip), 0)
    RefreshNotifyIcon
    
End Property


Public Sub Hide()
Attribute Hide.VB_HelpID = 2477

Dim niThis As NOTIFYICONDATA
Dim lret As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = NotifyWindowMessage
End If

With niThis
    .hIcon = mhIcon
    .hwnd = NotifyWindow
    .szTip = mszTip
    .uCallbackMessage = muCallbackMessage
    .uFlags = muFlags
    .uID = muID
    .cbSize = Len(niThis)
End With

lret = Shell_NotifyIcon(NIM_DELETE, niThis)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiNotifyIcon:UnsetNotifyIcon", GetLastSystemError
Else
    mAdded = False
End If

End Sub


Private Sub Class_Initialize()

muFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE


End Sub


Private Sub Class_Terminate()

If Not wndThis Is Nothing Then
    wndThis.SendMessage WM_CLOSE, 0, 0
    Set wndThis = Nothing
End If

End Sub


Private Sub wndThis_WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)

If msg = NotifyWindowMessage Then
    Select Case lParam
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton)
    Case WM_MBUTTONDOWN
        RaiseEvent MouseDown(vbMiddleButton)
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton)
    Case WM_MBUTTONUP
        RaiseEvent MouseUp(vbMiddleButton)
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick(vbLeftButton)
    Case WM_RBUTTONDBLCLK
        RaiseEvent DblClick(vbRightButton)
    Case WM_MBUTTONDBLCLK
        RaiseEvent DblClick(vbMiddleButton)
    End Select
End If

End Sub



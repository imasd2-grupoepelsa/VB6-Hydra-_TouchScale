VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
'* The SetWindowsHookEx function installs an application-defined hook procedure into a hook chain. You would install a hook procedure to monitor the system for certain types of events. These events are associated either with a specific thread or with all threads in the system.
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'* The UnhookWindowsHookEx function removes a hook procedure installed in a hook chain by the SetWindowsHookEx function.
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'* The CallNextHookEx function passes the hook information to the next hook procedure in the current hook chain. A hook procedure can call this function either before or after processing the hook information.
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

'* The IsWindow function determines whether the specified window handle identifies an existing window.
Private Declare Function IsWindow Lib "user32" (ByVal hwnd&) As Long

'* The GetWindowText function copies the text of the specified window’s title bar (if it has one) into a buffer. If the specified window is a control, the text of the control is copied.
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'* The WindowFromPoint function retrieves the handle of the window that contains the specified point.
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'* The GetCursorPos function retrieves the cursor’s position, in screen coordinates.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'* Type to hold the mouse coordinates
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private point As POINTAPI

'******************************************************
'* System messages for Mouse that we want to trace
Private Const WM_MOUSEMOVE           As Long = &H200
Private Const WM_MOUSEWHEEL          As Long = &H20A
Private Const WM_LBUTTONDOWN         As Long = &H201
Private Const WM_LBUTTONUP           As Long = &H202
Private Const WM_RBUTTONDOWN         As Long = &H204
Private Const WM_RBUTTONUP           As Long = &H205
Private Const WM_MBUTTONDOWN         As Long = &H207
Private Const WM_MBUTTONUP           As Long = &H208
'******************************************************

'* System hook for mouse that we'll use
Private Const WH_MOUSE_LL            As Long = 14

Private mlngMouseHook               As Long
Private mstrWindowCaption           As String

'******************************************************
'* System events of Mouse that we can stop broadcasting
'* to the System and make only traceable from the caller application
Public Type EventThief
    LEFT_DOWN       As Boolean
    LEFT_UP         As Boolean
    RIGHT_DOWN      As Boolean
    RIGHT_UP        As Boolean
    MIDDLE_DOWN     As Boolean
    MIDDLE_UP       As Boolean
    WHEEL           As Boolean
    MOVE            As Boolean
End Type

Private mtypEventThief As EventThief

Private Const HC_ACTION = 0

'* Type to hold Mouse Hook information
Private Type MOUSELLHOOKSTRUCT
    point As POINTAPI
    data As Long
    flags As Long
    time As Long
    extra As Long
End Type

Private mousedata As MOUSELLHOOKSTRUCT

'******************************************************
'* Events that we'll be fired to the caller application
'* For different Mouse activities in the System
Public Event OnSystemMouseMove()
Public Event OnSystemMouseWheel()
Public Event OnSystemMouseLeftDown()
Public Event OnSystemMouseLeftUp()
Public Event OnSystemMouseRightDown()
Public Event OnSystemMouseRightUp()
Public Event OnSystemMouseMiddleDown()
Public Event OnSystemMouseMiddleUp()
'******************************************************

'* Get the mouse events that we want to steal from the System
Public Property Get StealMouseEvents() As EventThief
    StealMouseEvents = mtypEventThief
End Property

'* Set the mouse events that we want to steal from the System
Public Property Let StealMouseEvents(typEventThief As EventThief)
    mtypEventThief = typEventThief
End Property

'* Get the X coordinate of the Mouse in the current screen
Public Property Get CoordinateX() As Long
    CoordinateX = point.X
End Property

'* Get the Y coordinate of the Mouse in the current screen
Public Property Get CoordinateY() As Long
    CoordinateY = point.Y
End Property

'* Get the Caption (if any) of the window currently under the Mouse
Public Property Get WindowTextUnderMouse() As String
    WindowTextUnderMouse = GetWindowTitle(point.X, point.Y)
End Property

'* Get the handle (hWnd) of the window currently under the Mouse
Public Property Get WindowHandleUnderMouse() As Long
    WindowHandleUnderMouse = GetWindowHandle(point.X, point.Y)
End Property

'* Start tracing system wide mouse events
Public Sub StartMouseTracing(ByVal hwnd As Long)
    If IsWindow(hwnd) Then
        glnghWnd = hwnd
        
        '* If there is already a hook by the application, don't hook again
        If GetProp(hwnd, "MouseHook") Then
            Exit Sub
        End If
        
        '* Try to set the Object instance of the Tracer class in memory
        '* If successful, then set the hook
        If SetProp(hwnd, ByVal "MouseHook", ObjPtr(Me)) Then
            mlngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf modHook.MouseProc, App.hInstance, 0)
        End If
    End If
End Sub

'* Stop tracing system wide mouse events
Public Sub StopMouseTracing()
    '* First check if the hook was established
    If mlngMouseHook <> 0 Then
        RemoveProp glnghWnd, "MouseHook"
        UnhookWindowsHookEx mlngMouseHook
    End If
End Sub

Private Sub Class_Terminate()
    '* Make sure the hook is removed after the class is terminated
    StopMouseTracing
End Sub

'* Bypassed MouseProc from the modHook's MouseProc.
'* The trick is necessary here to raise the appropriate events to the caller
'* for particular System events of the Mouse
Friend Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (nCode = HC_ACTION) Then
        ' Mouse data not used in this example, but useful
        CopyMemory mousedata, ByVal lParam, Len(mousedata)
        Select Case wParam
            Case WM_MOUSEMOVE
                '* Trace the coordinate of the Mouse whenever it moves
                GetCursorPos point
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseMove
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.MOVE Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_MOUSEWHEEL
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseWheel
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.WHEEL Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_LBUTTONDOWN
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseLeftDown
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.LEFT_DOWN Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_LBUTTONUP
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseLeftUp
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.LEFT_UP Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_RBUTTONDOWN
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseRightDown
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.RIGHT_DOWN Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_RBUTTONUP
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseRightUp
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.RIGHT_UP Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_MBUTTONDOWN
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseMiddleDown
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.MIDDLE_DOWN Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case WM_MBUTTONUP
                '* Raise the event to the caller Window
                RaiseEvent OnSystemMouseMiddleUp
                '* Check if caller wanted this event as non traceable to the System
                If mtypEventThief.MIDDLE_UP Then
                    '* If so, steal the event from the system and make it only traceable
                    '* to the window of the caller application
                    MouseProc = -1
                Else
                    '* If not, release the event to the System
                    MouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
                End If
            Case Else
                ' not implemented yet
        End Select
    End If
End Function

'* Returns the Window Caption (if any) under the mouse
Private Function GetWindowTitle(CoordX As Long, CoordY As Long) As String
    Dim strTitle As String
    strTitle = String(255, Chr$(0))
    GetWindowText WindowFromPoint(CoordX, CoordY), strTitle, 255
    strTitle = Left$(strTitle, InStr(strTitle, Chr$(0)) - 1)
    GetWindowTitle = strTitle
End Function

'* Returns the Window Handle (hWnd) under the mouse
Private Function GetWindowHandle(CoordX As Long, CoordY As Long) As Long
    GetWindowHandle = WindowFromPoint(CoordX, CoordY)
End Function

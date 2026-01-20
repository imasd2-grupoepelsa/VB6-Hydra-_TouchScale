Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd&, ByVal lpString$) As Long

Public Declare Function IsWindow Lib "user32" (ByVal hWnd&) As Long

Public Const WM_RBUTTONDOWN         As Long = &H204
Public Const WM_RBUTTONUP           As Long = &H205
Public Const WM_RBUTTONDBLCLK       As Long = &H206
Public Const WM_MOUSEMOVE           As Long = &H200
Public Const WM_LBUTTONDOWN         As Long = &H201
Public Const WM_LBUTTONUP           As Long = &H202
Public Const WM_MBUTTONDOWN         As Long = &H207
Public Const WM_MBUTTONUP           As Long = &H208
Public Const WM_MOUSEWHEEL          As Long = &H20A

Public Const WH_MOUSE_LL            As Long = 14

Global glnghWnd            As Long


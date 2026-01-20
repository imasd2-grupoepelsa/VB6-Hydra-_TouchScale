Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Public Declare Function keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Long
Public Declare Function BlockInput Lib "user32" (ByVal fLock As Long) As Long
Public Const VK_SHIFT = &H10 'Shift Const
Public Const KEYEVENTF_KEYUP = &H2 ' Release key Const

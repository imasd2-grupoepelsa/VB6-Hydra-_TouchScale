Attribute VB_Name = "thumb"
Option Explicit
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const ID_VIEW_ICONS = 28713
Private Const ID_VIEW_SMALLICONS = 28714 '98 only
Private Const ID_VIEW_LIST = 28715
Private Const ID_VIEW_DETAILS = 28716
Private Const ID_VIEW_THUMBNAIL = 28717  'xp only
Private Const ID_VIEW_TILES = 28718      'xp only

Private Const WM_COMMAND = &H111&
Private Const ID_DETAILS = 28716
Private TimerId As Long

Public Sub SetDetailsView()
    TimerId = SetTimer(0, 0, 0, AddressOf TimerProc)
End Sub
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    hWnd = FindWindowEx(GetActiveWindow, 0, "SHELLDLL_DefView", vbNullString)
    SendMessage hWnd, WM_COMMAND, ID_VIEW_THUMBNAIL, ByVal 0&
    KillTimer 0, TimerId
End Sub


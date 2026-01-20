VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2441
Option Explicit

Private Const CLSNAME = "ApiThread"

Private Type LDT_BYTES  ' Defined for use in LDT_ENTRY Type
    BaseMid As Byte
    Flags1 As Byte
    Flags2 As Byte
    BaseHi As Byte
End Type

Private Type LDT_ENTRY
    LimitLow As Integer
    BaseLow As Integer
    HighWord As LDT_BYTES
End Type

Private Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long

Private Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Private Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long

Private Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long

Private Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Declare Function SetThreadToken Lib "advapi32" (ByVal Thread As Long, ByVal Token As Long) As Long
Private Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long

Private Declare Function CreateThreadApi Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

'\\ Member variables
Private mThreadId As Long
Private mThreadHandle As Long
Private mDesiredAccess As Long

Private mBaseAddress As Long

' ##ENUMERATION_DESCRIPTION The priority assigned to the thread
Public Enum enThreadPriorities
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_BASE_PRIORITY_IDLE Thread is idling
    THREAD_BASE_PRIORITY_IDLE = -15
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_BASE_PRIORITY_LOWRT Low real time priority
    THREAD_BASE_PRIORITY_LOWRT = 15
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_BASE_PRIORITY_MAX Maximum priority (2 points above normal)
    THREAD_BASE_PRIORITY_MAX = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_BASE_PRIORITY_MIN Minimum priority (2 points below normal)
    THREAD_BASE_PRIORITY_MIN = -2
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_PRIORITY_NORMAL Standard priority
    THREAD_PRIORITY_NORMAL = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_PRIORITY_ABOVE_NORMAL One point above normal priority
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_PRIORITY_BELOW_NORMAL One point below normal priority
    THREAD_PRIORITY_BELOW_NORMAL = -1
End Enum

Private Declare Function ResumeThreadApi Lib "kernel32" Alias "ResumeThread" (ByVal hThread As Long) As Long
Private Declare Function SuspendThreadApi Lib "kernel32" Alias "SuspendThread" (ByVal hThread As Long) As Long

Private Declare Function TerminateThreadApi Lib "kernel32" Alias "TerminateThread" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Private bChanged As Boolean

Private Declare Function CloseKernelObject Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInherit As Long, ByVal dwThreadId As Long) As Long

' ##ENUMERATION_DESCRIPTION The priviledge levels associated with thread access
Public Enum ThreadAccessPriviledges
    ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_TERMINATE Can terminate the thread
     THREAD_TERMINATE = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_SUSPEND_RESUME Can suspend and resume running of the thread
     THREAD_SUSPEND_RESUME = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_GET_CONTEXT Can get the context of the thread
     THREAD_GET_CONTEXT = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_SET_CONTEXT Can set the context of the thread
     THREAD_SET_CONTEXT = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_SET_INFORMATION Can set thread level information for the thread
     THREAD_SET_INFORMATION = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_QUERY_INFORMATION Can get thread level information for the thread
     THREAD_QUERY_INFORMATION = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_SET_THREAD_TOKEN Can set token information for the thread
     THREAD_SET_THREAD_TOKEN = &H80
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_IMPERSONATE
     THREAD_IMPERSONATE = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_DIRECT_IMPERSONATION
     THREAD_DIRECT_IMPERSONATION = &H200
     ' ##ENUMERATION_MEMBER_DESCRIPTION THREAD_SYNCHRONISE Can create synchronised events to the thread
     THREAD_SYNCHRONISE = &H100000
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUITHREADINFO
    cbSize As Long
    flags As Long
    hWndActive As Long
    hwndFocus As Long
    hwndCapture As Long
    hwndMenuOwner As Long
    hwndMoveSize As Long
    hwndCaret As Long
    rcCaret As RECT
End Type

' ##ENUMERATION_DESCRIPTION Various state information about a thread which has an application window
Public Enum GuiThreadStateFlags
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_CARETBLINKING The CARET is visible
    GUI_CARETBLINKING = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_INMOVESIZE The window is being moved or sized
    GUI_INMOVESIZE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_INMENUMODE The window has a menu open
    GUI_INMENUMODE = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_SYSTEMMENUMODE The window has a system menu open
    GUI_SYSTEMMENUMODE = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_POPUPMENUMODE The window has a popup menu open
    GUI_POPUPMENUMODE = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION GUI_16BITTASK The thread is a 16 bit application
    GUI_16BITTASK = &H20
End Enum

Private Declare Function GetGUIThreadInfo Lib "user32" (ByVal hThreadId As Long, pGuiThreadInfo As GUITHREADINFO) As Long

Private mThreadInfo As GUITHREADINFO


' ##BLOCK_DESCRIPTION ActiveWindow - The active window for this thread
' ##BLOCK_REMARKS This is not neccesarily the same as the focus window
Public Property Get ActiveWindow() As ApiWindow
Attribute ActiveWindow.VB_HelpID = 3082
Attribute ActiveWindow.VB_Description = "ActiveWindow - The active window for this thread"

Dim wndActive As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hWndActive <> 0 Then
    Set wndActive = New ApiWindow
    wndActive.hwnd = mThreadInfo.hWndActive
    Set ActiveWindow = wndActive
End If

End Property

' ##BLOCK_DESCRIPTION CaretWindow - The window that contains the visible caret
' ##BLOCK_REMARKS This will be set if the thread has a visible caret
Public Property Get CaretWindow() As ApiWindow
Attribute CaretWindow.VB_HelpID = 3083
Attribute CaretWindow.VB_Description = "CaretWindow - The window that contains the visible caret"

Dim wndCaret As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hwndCaret <> 0 Then
    Set wndCaret = New ApiWindow
    wndCaret.hwnd = mThreadInfo.hwndCaret
    Set CaretWindow = wndCaret
End If

End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Let BaseAddress(ByVal lProcAddress As Long)
Attribute BaseAddress.VB_HelpID = 2442

#If PRODUCTION Then
    On Error Resume Next
#End If

If Not IsBadCodePtr(lProcAddress) Then
    mBaseAddress = lProcAddress
Else
    ReportError vbError + 100, CLSNAME & ":BaseAddress", "Invalid code address"
End If
    
End Property


Public Property Let DesiredAccess(ByVal newAccess As ThreadAccessPriviledges)
Attribute DesiredAccess.VB_HelpID = 2875

#If PRODUCTION Then
    On Error Resume Next
#End If

If newAccess <> mDesiredAccess Then
    mDesiredAccess = newAccess
    If mThreadId <> 0 Then
        '\\ Refresh the open thread handle to get the new access
        If mThreadHandle <> 0 Then
            Call CloseKernelObject(mThreadHandle)
            If Err.LastDllError Then
                ReportError Err.LastDllError, CLSNAME & ":DesiredAccess", GetLastSystemError
            Else
                mThreadHandle = 0
            End If
            
        End If
        If mThreadHandle = 0 Then
            mThreadHandle = OpenThread(mDesiredAccess, False, mThreadId)
            If Err.LastDllError Then
                ReportError Err.LastDllError, CLSNAME & ":DesiredAccess (let)", GetLastSystemError
            End If
        End If
    End If
End If
End Property



' ##BLOCK_DESCRIPTION DeskTop - The desktop to which this thread is associated
' ##BLOCK_REMARKS Only Windows NT and the line support multiple desktops
' ##BLOCK_REMARKS On Windows 95/98/Me all threads are associated with the same desktop
Public Property Get Desktop() As ApiDesktop
Attribute Desktop.VB_HelpID = 3084
Attribute Desktop.VB_Description = "DeskTop - The desktop to which this thread is associated"

Dim dtThreadDesktop As ApiDesktop
Dim hDesk As Long

If mThreadId <> 0 Then
    hDesk = GetThreadDesktop(mThreadId)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Desktop", GetLastSystemError
    Else
        Set dtThreadDesktop = New ApiDesktop
        dtThreadDesktop.Handle = hDesk
        Set Desktop = dtThreadDesktop
    End If
End If

End Property

Public Property Get EndTime() As APIFileTime
Attribute EndTime.VB_HelpID = 2443
Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime


#If PRODUCTION Then
    On Error Resume Next
#End If

Set timeThis = New APIFileTime

lRet = GetThreadTimes(mThreadHandle, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":StartTime", GetLastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeEnd))
End If

Set EndTime = timeThis

End Property

' ##BLOCK_DESCRIPTION FocusWindow - The window that has the keyboard focus for this thread
' ##BLOCK_REMARKS This is not neccesarily the same as the active window
Public Property Get FocusWindow() As ApiWindow
Attribute FocusWindow.VB_HelpID = 3085
Attribute FocusWindow.VB_Description = "FocusWindow - The window that has the keyboard focus for this thread"

Dim wndFocus As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hwndFocus <> 0 Then
    Set wndFocus = New ApiWindow
    wndFocus.hwnd = mThreadInfo.hwndFocus
    Set FocusWindow = wndFocus
End If

End Property

Public Property Get KernelTime() As APIFileTime
Attribute KernelTime.VB_HelpID = 2444
Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime

#If PRODUCTION Then
    On Error Resume Next
#End If

lRet = GetThreadTimes(mThreadHandle, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":StartTime", GetLastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeKernel))
End If

Set KernelTime = timeThis

End Property

Public Property Get Key() As String
Attribute Key.VB_HelpID = 2876

    Key = "HTHREAD:" & mThreadId
    
End Property

' ##BLOCK_DESCRIPTION MenuWindow - The window that owns an open menu
' ##BLOCK_REMARKS This will be set if the thread has a system, popup or window menu open
Public Property Get MenuWindow() As ApiWindow
Attribute MenuWindow.VB_HelpID = 3086
Attribute MenuWindow.VB_Description = "MenuWindow - The window that owns an open menu"

Dim wndMenu As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hwndMenuOwner <> 0 Then
    Set wndMenu = New ApiWindow
    wndMenu.hwnd = mThreadInfo.hwndMenuOwner
    Set MenuWindow = wndMenu
End If

End Property

' ##BLOCK_DESCRIPTION MouseCaptureWindow - The window that has the mouse capture for this thread
' ##BLOCK_REMARKS This is not neccesarily the same as the focus window
Public Property Get MouseCaptureWindow() As ApiWindow
Attribute MouseCaptureWindow.VB_HelpID = 3087
Attribute MouseCaptureWindow.VB_Description = "MouseCaptureWindow - The window that has the mouse capture for this thread"

Dim wndCapture As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hwndCapture <> 0 Then
    Set wndCapture = New ApiWindow
    wndCapture.hwnd = mThreadInfo.hwndCapture
    Set MouseCaptureWindow = wndCapture
End If

End Property

' ##BLOCK_DESCRIPTION MovingSizingWindow - The window that is being moved or resized
' ##BLOCK_REMARKS This will be set if a window owned by this thread is being moved or resized
Public Property Get MovingSizingWindow() As ApiWindow
Attribute MovingSizingWindow.VB_HelpID = 3088
Attribute MovingSizingWindow.VB_Description = "MovingSizingWindow - The window that is being moved or resized"

Dim wndMoveSize As ApiWindow

Call RefreshThreadInfo
If mThreadInfo.hwndMoveSize <> 0 Then
    Set wndMoveSize = New ApiWindow
    wndMoveSize.hwnd = mThreadInfo.hwndMoveSize
    Set MovingSizingWindow = wndMoveSize
End If

End Property

Public Property Let Priority(ByVal newPriority As enThreadPriorities)
Attribute Priority.VB_HelpID = 2446

Dim lRet As Long

#If PRODUCTION Then
    On Error Resume Next
#End If

lRet = SetThreadPriority(mThreadHandle, newPriority)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Priority", GetLastSystemError
End If

End Property

Public Property Get Priority() As enThreadPriorities

Dim lRet As Long

#If PRODUCTION Then
    On Error Resume Next
#End If

lRet = GetThreadPriority(mThreadHandle)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Priority", GetLastSystemError
Else
    Priority = lRet
End If

End Property

' --[RefreshThreadInfo]--------------------------------------------------
' Fills the GUITHREADINFO structure for this thread...
' -----------------------------------------------------------------------
Private Sub RefreshThreadInfo()

#If PRODUCTION Then
    On Error Resume Next
#End If

With mThreadInfo
    .flags = 0
    .hWndActive = 0
    .hwndCapture = 0
    .hwndCaret = 0
    .hwndFocus = 0
    .hwndMenuOwner = 0
    .hwndMoveSize = 0
    With .rcCaret
        .Left = 0
        .Bottom = 0
        .Right = 0
        .Top = 0
    End With
    .cbSize = Len(mThreadInfo)
End With

Dim lRet As Long

If mThreadId <> 0 Then
    lRet = GetGUIThreadInfo(mThreadId, mThreadInfo)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":RefreshThreadInfo", GetLastSystemError
    End If
End If

End Sub


Public Function ResumeThread() As Long
Attribute ResumeThread.VB_HelpID = 2447


#If PRODUCTION Then
    On Error Resume Next
#End If

    ResumeThread = ResumeThreadApi(mThreadHandle)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, CLSNAME & ":ResumeThread", GetLastSystemError)
    End If
    
End Function

Public Property Get StartTime() As APIFileTime
Attribute StartTime.VB_HelpID = 2448

Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime


#If PRODUCTION Then
    On Error Resume Next
#End If

lRet = GetThreadTimes(mThreadHandle, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":StartTime", GetLastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeStart))
End If

Set StartTime = timeThis

End Property

Private Function SuspendThread() As Long


#If PRODUCTION Then
    On Error Resume Next
#End If

    SuspendThread = SuspendThreadApi(mThreadHandle)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, CLSNAME & ":SuspendThread", GetLastSystemError)
    End If

End Function

Public Sub TerminateThread(ByVal exitcode As Long)
Attribute TerminateThread.VB_HelpID = 2449


#If PRODUCTION Then
    On Error Resume Next
#End If

Call TerminateThreadApi(mThreadHandle, exitcode)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":TerminateThread", GetLastSystemError
End If

End Sub




Friend Property Let ThreadId(ByVal newId As Long)


#If PRODUCTION Then
    On Error Resume Next
#End If

If newId <> mThreadId Then
    mThreadId = newId
    If mThreadHandle <> 0 Then
        Call CloseKernelObject(mThreadHandle)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":Terminate", GetLastSystemError
        Else
            mThreadHandle = 0
        End If
    End If
    If mThreadHandle = 0 Then
        '\\ Def 20020211- Only Win2000 and beyond support this function
        If APIDispenser.System.IsRequiredSystem("OpenThread", ver_Win_Win2000) Then
            mThreadHandle = OpenThread(mDesiredAccess, False, mThreadId)
            If Err.LastDllError Then
                ReportError Err.LastDllError, CLSNAME & ":ThreadId (let)", GetLastSystemError
            End If
        End If
    End If
End If

End Property

Public Property Get ThreadId() As Long
Attribute ThreadId.VB_HelpID = 2451

    ThreadId = mThreadId
    
End Property
Public Property Get UserTime() As APIFileTime
Attribute UserTime.VB_HelpID = 2452

Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime


#If PRODUCTION Then
    On Error Resume Next
#End If

lRet = GetThreadTimes(mThreadHandle, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":StartTime", GetLastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeUser))
End If

Set UserTime = timeThis

End Property


Private Sub Class_Terminate()

#If PRODUCTION Then
    On Error Resume Next
#End If

If mThreadHandle <> 0 Then
    Call CloseKernelObject(mThreadHandle)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Terminate", GetLastSystemError
    End If
End If

End Sub



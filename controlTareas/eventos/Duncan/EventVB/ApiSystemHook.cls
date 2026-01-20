VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiSystemHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2002
Attribute VB_Description = "This class provides access to system hooks - system wide or application local"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' ##MODULE_DESCRIPTION This class provides access to system hooks - system wide or application local
' ##MODULE_REMARKS The following restrictions apply:
' ##MODULE_REMARKS WH_CALLWNDPROC - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_CBT - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_DEBUG - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_FOREGROUNDIDLE - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_GETMESSAGE - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_HARDWARE - This hook is not yet supported on any operating system
' ##MODULE_REMARKS WH_JOURNALPLAYBACK - HOOK_GLOBAL only
' ##MODULE_REMARKS WH_JOURNALRECORD - HOOK_GLOBAL only
' ##MODULE_REMARKS WH_MOUSE - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_MSGFILTER - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_SHELL - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_SYSMSGFILTER - HOOK_LOCAL_PROCESS only
' ##MODULE_REMARKS WH_KEYBOARD_LL - HOOK_GLOBAL only, Windows NT, 2000 or XP
' ##MODULE_REMARKS WH_MOUSE_LL - HOOK_GLOBAL only, Windows NT, 2000 or XP
' ##MODULE_REMARKS WH_KEYBOARD - HOOK_LOCAL_PROCESS only


Private Const CLSNAME = "ApiSystemHook"

Public Enum CallbackProcTypes
    DLGPROC = 1 'typedef BOOL (CALLBACK* DLGPROC)(HWND, UINT, WPARAM, LPARAM);
    TIMERPROC = 2 'typedef VOID (CALLBACK* TIMERPROC)(HWND, UINT, UINT, DWORD);
    GRAYSTRINGPROC = 3 'typedef BOOL (CALLBACK* GRAYSTRINGPROC)(HDC, LPARAM, int);
    HOOKPROC = 4 'typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
    SENDASYNCPROC = 5 'typedef VOID (CALLBACK* SENDASYNCPROC)(HWND, UINT, DWORD, LRESULT);
    PROPENUMPROC = 6 'typedef BOOL (CALLBACK* PROPENUMPROCA)(HWND, LPCSTR, HANDLE);
    PROPENUMPROCEX = 7 'typedef BOOL (CALLBACK* PROPENUMPROCEXA)(HWND, LPSTR, HANDLE, DWORD);
    EDITWORDBREAKPROC = 8 'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
    WNDENUMPROC = 9 'typedef BOOL (CALLBACK* WNDENUMPROC)(HWND, LPARAM);
    WNDPROC = 10 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
    WINSTATIONENUMPROC = 11
    DESKTOPENUMPROC = 12
    ENUMRESTYPEPROC = 13
    ENUMRESNAMEPROC = 14
    ENUMRESLANGPROC = 15
    '\\ DEJ 01/03/2001 - Added hook proc subtypes
    HOOKPROC_CALLWNDPROC = 16 'CallWndProc
    HOOKPROC_CALLWNDPROCRET = 17 'CallWndProcRet
    HOOKPROC_CBT = 18 'CBTProc
    HOOKPROC_DEBUG = 19 'DebugProc
    HOOKPROC_GETMESSAGE = 20 'GetMsgProc
    HOOKPROC_JOURNALPLAYBACK = 21 'JournalPlaybackProc
    HOOKPROC_JOURNALRECORD = 22 'JournalRecordProc
    HOOKPROC_KEYBOARD = 23 'KeyboardProc
    HOOKPROC_MOUSE = 24 'MouseProc
    HOOKPROC_MESSAGEFILTER = 25 'MessageProc
    HOOKPROC_SHELL = 26 'ShellProc
    HOOKPROC_SYSMESSAGEFILTER = 27 'SysMsgProc
    HOOKPROC_MOUSE_LL = 28
    HOOKPROC_KEYBOARD_LL = 29
    HOOKPROC_FOREGROUNDIDLE = 30
End Enum


'\\ API Calls to pass on WM_ message to previous windows proc...
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'\\ Existing HOOK id...
'\\ Note: current design is max one or each type of hook proc per ApiSystemHook
Private CallWndProcHookId As Long
Private CallWndProcRetHookId As Long
Private CBTHookId As Long
Private DebugHookId As Long
Private ForegroundidleHookId As Long
Private GetMessageHookid As Long
Private HardwareHookId As Long
Private MessageHookId As Long
Private JournalRecordHookId As Long
Private JournalPlaybackHookId As Long
Private MouseHookId As Long
Private MsgFilterHookId As Long
Private KeyboardHookId As Long
Private ShellHookId As Long
Private SysmessageHookId As Long
Private LowlevelMouseHookId As Long
Private LowLevelKeyboardHookId As Long

'\\ Hook codes....
Public Enum enHookCode
    HC_ACTION = 0
    HC_GETNEXT = 1
    HC_NOREMOVE = 3
    HC_SKIP = 2
    HC_SYSMODALOFF = 5
    HC_SYSMODALON = 4
End Enum

' ##ENUMERATION_DESCRIPTION HookTypes are the different kinds of hook that can be installed
Public Enum enHookTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_CALLWNDPROC Is triggered when a message is about to be added to a window message queue
    WH_CALLWNDPROC = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_CBT Is triggered when events that are useful in computer based training applications occur
    WH_CBT = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_DEBUG Is triggered when debugging other hooks
    WH_DEBUG = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_FOREGROUNDIDLE Is triggered when the foreground application is becoming idle - ideal for performing low priority tasks
    WH_FOREGROUNDIDLE = 11
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_GETMESSAGE Is triggered when a window gets a message from its message queue
    WH_GETMESSAGE = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_HARDWARE Not used at this stage.
    WH_HARDWARE = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_JOURNALPLAYBACK Used to post the events previously recorded by the Journal Record hook
    WH_JOURNALPLAYBACK = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_JOURNALRECORD Records mouse and keyboard events added to the system queue
    WH_JOURNALRECORD = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_MOUSE Is triggered when a mouse event occurs
    WH_MOUSE = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_MSGFILTER is triggered when an event happens in a message box, dialog, menu or scrollbar
    WH_MSGFILTER = (-1)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_SHELL is triggered by events pertaining to the shell
    WH_SHELL = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_SYSMSGFILTER is like the MsgFilter hook but on a desktop-wide basis
    WH_SYSMSGFILTER = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_KEYBOARD_LL is triggered when a keyboard event occurs (system wide).  This is NT, Win2K, WinXP only.
    WH_KEYBOARD_LL = 13
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_MOUSE_LL is triggered when a mouse event occurs (system wide).  This is NT, Win2K, WinXP only.
    WH_MOUSE_LL = 14
    ' ##ENUMERATION_MEMBER_DESCRIPTION WH_KEYBOARD is triggered when a keyboard event occurs
    WH_KEYBOARD = 2
End Enum

Public Enum enHookCBTCodes
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_ACTIVATE - The system is about to activate a window
    HCBT_ACTIVATE = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_CLICKSKIPPED - A mouse click has been removed from the queue for processing
    HCBT_CLICKSKIPPED = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_CREATEWND - A window is being created
    HCBT_CREATEWND = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_DESTROYWND - A window is being destroyed
    HCBT_DESTROYWND = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_KEYSKIPPED - A keystroke has been removed from the queue
    HCBT_KEYSKIPPED = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_MINMAX - A window is going to be minimised or maximised
    HCBT_MINMAX = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_MOVESIZE - A window is being moved or resized
    HCBT_MOVESIZE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_QS - A WM_QUEUESYNC message has been received
    HCBT_QS = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_SETFOCUS - A window is about to get a setfocus
    HCBT_SETFOCUS = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION HCBT_SYSCOMMAND - A system command is being executed
    HCBT_SYSCOMMAND = 8
End Enum

Public Enum PeekMessageFlags
    PM_NOREMOVE = &H0
    PM_NOYIELD = &H2
    PM_REMOVE = &H1
End Enum

' ##ENUMERATION_DESCRIPTION MessageFilterEventTypes are the different types of event that can trigger a message filter hook
Public Enum MessageFilterEventTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION MSGF_DDEMGR The message occured while a synchronous DDE operation was in progress
    MSGF_DDEMGR = &H8001
    ' ##ENUMERATION_MEMBER_DESCRIPTION MSGF_DIALOGBOX The message occured in a dialog box
    MSGF_DIALOGBOX = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION MSGF_MENU The message occured in a menu
    MSGF_MENU = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION MSGF_MESSAGEBOX The message occured in a message box
    MSGF_MESSAGEBOX = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION MSGF_SCROLLBAR The message occured in a scrollbar
    MSGF_SCROLLBAR = 5
End Enum


'\\ Windows hooks...
'SetWindowsHookEx
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public CurrentHookType As enHookTypes
Attribute CurrentHookType.VB_VarHelpID = 2042


Private bChanged As Boolean

' ##ENUMERATION_DESCRIPTION HookScopeSettings describe whether a hook is set system wide or just for the current application
Public Enum HookScopeSettings
    ' ##ENUMERATION_MEMBER_DESCRIPTION HOOK_LOCAL_PROCESS Only this application is hooked
    HOOK_LOCAL_PROCESS = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION HOOK_GLOBAL The entire system is hooked
    HOOK_GLOBAL = 1
End Enum

'----Mouse Hooking -------------------------------------------------------------
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    MouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Declare Sub CopyMemoryMouseHookStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As MOUSEHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromMOUSEHOOKSTRUCT Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As MOUSEHOOKSTRUCT, ByVal Length As Long)
Private Declare Sub CopyMemoryMSLLHOOKSTRUCT Lib "kernel32" Alias "RtlMoveMemory" (Destination As MSLLHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromMSLLHOOKSTRUCT Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As MSLLHOOKSTRUCT, ByVal Length As Long)

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'---Mouse Hook Messages ------------------------------------------------------------------------------------------------------------------------------------------------------
' ##EVENT_DESCRIPTION MouseButtonDown - One of the mouse buttons was depressed
' ##EVENT_REMARKS Triggered if a HOOKPROC_MOUSE or HOOKPROC_MOUSE_LL hook is active
' ##EVENT_REMARKS Button - The button that was pressed
' ##EVENT_REMARKS Location - The %cursor position:EventVB~ApiPoint% when the mouse button was depressed
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event MouseButtonDown(ByVal Button As Integer, ByVal Location As APIPoint, ByVal TargetWindow As ApiWindow, Cancel As Boolean)
Attribute MouseButtonDown.VB_Description = "MouseButtonDown - One of the mouse buttons was depressed"
Attribute MouseButtonDown.VB_HelpID = 2979

' ##EVENT_DESCRIPTION MouseButtonUp - One of the mouse buttons was released
' ##EVENT_REMARKS Triggered if a HOOKPROC_MOUSE or HOOKPROC_MOUSE_LL hook is active
' ##EVENT_REMARKS Button - The button that was released
' ##EVENT_REMARKS Location - The %cursor position:EventVB~ApiPoint% when the mouse button was released
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event MouseButtonUp(ByVal Button As Integer, ByVal Location As APIPoint, ByVal TargetWindow As ApiWindow, Cancel As Boolean)
Attribute MouseButtonUp.VB_Description = "MouseButtonUp - One of the mouse buttons was released"
Attribute MouseButtonUp.VB_HelpID = 2980

' ##EVENT_DESCRIPTION MouseMove - The mouse pointer moved
' ##EVENT_REMARKS Triggered if a HOOKPROC_MOUSE or HOOKPROC_MOUSE_LL hook is active
' ##EVENT_REMARKS Location - The %cursor position:EventVB~ApiPoint% that the mouse was moved to
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event MouseMove(ByVal Location As APIPoint, ByVal TargetWindow As ApiWindow, Cancel As Boolean)
Attribute MouseMove.VB_Description = "MouseMove - The mouse pointer moved"
Attribute MouseMove.VB_HelpID = 2981

' ##EVENT_DESCRIPTION MouseWheelMove - The mouse wheel (on the new type of mouse) was moved
' ##EVENT_REMARKS Triggered if a HOOKPROC_MOUSE or HOOKPROC_MOUSE_LL hook is active
' ##EVENT_REMARKS Rotation - The amount the wheel moved (positive is forward, negative is back)
' ##EVENT_REMARKS Location - The %cursor position:EventVB~ApiPoint% when the mouse wheel was moved
' ##EVENT_REMARKS TargetWindow - The %window:EventVB~ApiWindow% under the cursor location
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event MouseWheelMove(ByVal Rotation As Integer, ByVal Location As APIPoint, ByVal TargetWindow As ApiWindow, Cancel As Boolean)
Attribute MouseWheelMove.VB_Description = "MouseWheelMove - The mouse wheel (on the new type of mouse) was moved"
Attribute MouseWheelMove.VB_HelpID = 2982

'---Keyboard hooking-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Enum KeyboardHookFlags
    KF_ALTDOWN = &H2000
    KF_DLGMODE = &H800
    KF_EXTENDED = &H100
    KF_MENUMODE = &H1000
    KF_REPEAT = &H4000
    KF_UP = &H8000
End Enum

Private Declare Sub CopyMemoryKbdLLHookStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As KBDLLHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromKBDLLHOOKSTRUCT Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As KBDLLHOOKSTRUCT, ByVal Length As Long)

' ##ENUMERATION_DESCRIPTION LowLevelKeyboardHookFlags used to test the Flags returned by the low level keyboard hook (NT/2000/XP only)
Public Enum LowLevelKeyboardHookFlags
    ' ##ENUMERATION_MEMBER_DESCRIPTION LLKHF_EXTENDED Test if the key press is part of an extended key combination
     LLKHF_EXTENDED = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION LLKHF_INJECTED Test if the key event came fromthe keyboard or was injected by another process
     LLKHF_INJECTED = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION LLKHF_ALTDOWN Test if the ALT key is pressed
     LLKHF_ALTDOWN = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION LLKHF_UP Test if key is being released
     LLKHF_UP = &H80
End Enum

' ##EVENT_DESCRIPTION KeyDown - A keyboard key was pressed down
' ##EVENT_REMARKS Triggered if a HOOKPROC_KEYBOARD or HOOKPROC_KEYBOARD_LL hook is active
' ##EVENT_REMARKS VKey - The virtual key that was pressed down
' ##EVENT_REMARKS ScanCode - The keyboard specific scan code for the key that was pressed down
' ##EVENT_REMARKS ExtendedKey - True if this key press is part of an extended key conbination
' ##EVENT_REMARKS AltDown - True if the ALT key was down when this key was pressed
' ##EVENT_REMARKS Injected - True if the event was injected by another process rather than came from the keyboard
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event KeyDown(ByVal VKey As Long, ByVal scanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal Injected As Boolean, Cancel As Boolean)
Attribute KeyDown.VB_Description = "KeyDown - A keyboard key was pressed down"
Attribute KeyDown.VB_HelpID = 2983

' ##EVENT_DESCRIPTION KeyUp - A key was released
' ##EVENT_REMARKS Triggered if a HOOKPROC_KEYBOARD or HOOKPROC_KEYBOARD_LL hook is active
' ##EVENT_REMARKS VKey - The virtual key that was released
' ##EVENT_REMARKS ScanCode - The keyboard specific scan code for the key that was released
' ##EVENT_REMARKS ExtendedKey - True if this key press is part of an extended key combination
' ##EVENT_REMARKS AltDown - True if the ALT key was down when this key was released
' ##EVENT_REMARKS Injected - True if the event was injected by another process rather than came from the keyboard
' ##EVENT_REMARKS Cancel - Set this to true to prevent default processing of the message
Public Event KeyUp(ByVal VKey As Long, ByVal scanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal Injected As Boolean, Cancel As Boolean)
Attribute KeyUp.VB_Description = "KeyUp - A key was released"
Attribute KeyUp.VB_HelpID = 2984


'---Shell hooking-------------------------------------------------------------------------------------------------------------------------------------------------------
' ##ENUMERATION_DESCRIPTION ShellHookMessageTypes - The types of message trapped by a WH_SHELL hook
Public Enum ShellHookMessageTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_WINDOWCREATED A new application window was created
    HSHELL_WINDOWCREATED = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_ACTIVATESHELLWINDOW The shell is instructed to activate its main window
    HSHELL_ACTIVATESHELLWINDOW = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_WINDOWDESTROYED An application window has been destroyed
    HSHELL_WINDOWDESTROYED = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_WINDOWACTIVATED The focus has changed to a different application window
    HSHELL_WINDOWACTIVATED = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_GETMINRECT A window is being minimised or maximised and is requesting the size metrics
    HSHELL_GETMINRECT = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_REDRAW The title of a window in the task bar has changed and been redrawn
    HSHELL_REDRAW = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_TASKMAN The user has selected the task list
    HSHELL_TASKMAN = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_LANGUAGE The keyboard language has changed
    HSHELL_LANGUAGE = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_ACCESSIBILITYSTATE The accessibility state has changed
    HSHELL_ACCESSIBILITYSTATE = 11
    ' ##ENUMERATION_MEMBER_DESCRIPTION HSHELL_APPCOMMAND A system command has been issued but not handled by the window it was sent to
    HSHELL_APPCOMMAND = 12
End Enum

'---Shell hook messages -----------------------------------------------------------------------------------------------------------------------------------------------

' ##EVENT_DESCRIPTION TaskManagerRequested - Triggered when the user requests the task manager to be shown (WinNT4 or above)
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
' ##EVENT_REMARKS Cancel - Set this to true to prevent the task manager being shown
Public Event TaskManagerRequested(Cancel As Boolean)
Attribute TaskManagerRequested.VB_Description = "TaskManagerRequested - Triggered when the user requests the task manager to be shown (WinNT4 or above)"
Attribute TaskManagerRequested.VB_HelpID = 2985

' ##EVENT_DESCRIPTION WindowCreated - A new top level (application) window has been created
' ##EVENT_REMARKS wnd - The %window:EventVB~ApiWindow% which is being created
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event WindowCreated(ByVal wnd As ApiWindow)
Attribute WindowCreated.VB_Description = "WindowCreated - A new top level (application) window has been created"
Attribute WindowCreated.VB_HelpID = 2986

' ##EVENT_DESCRIPTION WindowDestroyed - A top level window is going to be closed
' ##EVENT_REMARKS wnd - The %window:EventVB~ApiWindow% which is being destroyed
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event WindowDestroyed(ByVal wnd As ApiWindow)
Attribute WindowDestroyed.VB_Description = "WindowDestroyed - A top level window is going to be closed"
Attribute WindowDestroyed.VB_HelpID = 2987

' ##EVENT_DESCRIPTION WindowTitleRedrawn - A top level window is redrawing it's title bar
' ##EVENT_REMARKS wnd - The %window:EventVB~ApiWindow% which is redrawing it's title bar
' ##EVENT_REMARKS Flashing - If true then the window is flashing to attract the user's attention
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event WindowTitleRedrawn(ByVal wnd As ApiWindow, ByVal Flashing As Boolean)
Attribute WindowTitleRedrawn.VB_Description = "WindowTitleRedrawn - A top level window is redrawing it's title bar"
Attribute WindowTitleRedrawn.VB_HelpID = 2988

' ##EVENT_DESCRIPTION KeyboardLanguageChanged - The keyboard language for a top level window was changed
' ##EVENT_REMARKS wnd - The %window:EventVB~ApiWindow% for which the keyboard language changed
' ##EVENT_REMARKS NewLayoutId - The new keyboard layout
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event KeyboardLanguageChanged(ByVal wnd As ApiWindow, ByVal NewLayoutId As Long)
Attribute KeyboardLanguageChanged.VB_Description = "KeyboardLanguageChanged - The keyboard language for a top level window was changed"
Attribute KeyboardLanguageChanged.VB_HelpID = 2989


Public Enum AccessibilityStateIndexes
    ACCESS_STICKYKEYS = &H1
    ACCESS_FILTERKEYS = &H2
    ACCESS_MOUSEKEYS = &H3
End Enum

' ##EVENT_DESCRIPTION AccessibilityStateChanged - The system accessibility state has changed
' ##EVENT_REMARKS Index - Indicates which accessibility state changed
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event AccessibilityStateChanged(ByVal Index As AccessibilityStateIndexes)
Attribute AccessibilityStateChanged.VB_Description = "AccessibilityStateChanged - The system accessibility state has changed"
Attribute AccessibilityStateChanged.VB_HelpID = 2990

' ##ENUMERATION_DESCRIPTION ApplicationCommandMessages - The different application command messages
Public Enum ApplicationCommandMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_BACKWARD Browse backwards in the history list
     APPCOMMAND_BROWSER_BACKWARD = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_FORWARD Browse forward in the history list
     APPCOMMAND_BROWSER_FORWARD = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_REFRESH Refresh the page in the browser
     APPCOMMAND_BROWSER_REFRESH = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_STOP Stop loading the current page in the browser
     APPCOMMAND_BROWSER_STOP = 4
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_SEARCH Show the internet search function
     APPCOMMAND_BROWSER_SEARCH = 5
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_FAVORITES Show the internet favourites
     APPCOMMAND_BROWSER_FAVORITES = 6
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BROWSER_HOME Browse to the system defined home page
     APPCOMMAND_BROWSER_HOME = 7
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_VOLUME_MUTE Mute the volume
     APPCOMMAND_VOLUME_MUTE = 8
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_VOLUME_DOWN Decrease the volume
     APPCOMMAND_VOLUME_DOWN = 9
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_VOLUME_UP Increase the volume
     APPCOMMAND_VOLUME_UP = 10
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_MEDIA_NEXTTRACK Skip to the next track in the media player
     APPCOMMAND_MEDIA_NEXTTRACK = 11
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_MEDIA_PREVIOUSTRACK Skip to the previous track in the media player
     APPCOMMAND_MEDIA_PREVIOUSTRACK = 12
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_MEDIA_STOP Stop the media player
     APPCOMMAND_MEDIA_STOP = 13
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_MEDIA_PLAY_PAUSE Pause the media player
     APPCOMMAND_MEDIA_PLAY_PAUSE = 14
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_LAUNCH_MAIL Launch the default mail program
     APPCOMMAND_LAUNCH_MAIL = 15
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_LAUNCH_MEDIA_SELECT
     APPCOMMAND_LAUNCH_MEDIA_SELECT = 16
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_LAUNCH_APP1
     APPCOMMAND_LAUNCH_APP1 = 17
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_LAUNCH_APP2
     APPCOMMAND_LAUNCH_APP2 = 18
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BASS_DOWN Decrease the base
     APPCOMMAND_BASS_DOWN = 19
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BASS_BOOST Toggle on/off the bass boost
     APPCOMMAND_BASS_BOOST = 20
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_BASS_UP Increase the bass
     APPCOMMAND_BASS_UP = 21
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_TREBLE_DOWN Decrease the treble
     APPCOMMAND_TREBLE_DOWN = 22
     ' ##ENUMERATION_MEMBER_DESCRIPTION APPCOMMAND_TREBLE_UP Increase the treble
     APPCOMMAND_TREBLE_UP = 23
End Enum

' ##EVENT_DESCRIPTION ApplicationCommand - An application command was sent to a window but not handled by it
' ##EVENT_REMARKS wnd - The %window:EventVB~ApiWindow% which was the target of the application command
' ##EVENT_REMARKS cmd - The application command message
' ##EVENT_REMARKS Triggered if a WH_SHELL hook is installed
Public Event ApplicationCommand(ByVal wnd As ApiWindow, ByVal cmd As ApplicationCommandMessages)
Attribute ApplicationCommand.VB_Description = "ApplicationCommand - An application command was sent to a window but not handled by it"
Attribute ApplicationCommand.VB_HelpID = 2991

'---CBT Hook events -----------------------------------------------------------------------------------------------------------------
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    Height As Long
    Width As Long
    Top As Long
    Left As Long
    Style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type

Private Declare Sub CopyMemoryCreateStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As CREATESTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromCREATESTRUCT Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As CREATESTRUCT, ByVal Length As Long)

Private Type CBT_CREATEWND
    csThis As Long '\\ Pointer to CREATESTRUCT
    hWndInsertAfter As Long
End Type

Private Declare Sub CopyMemoryCBT_CreateWnd Lib "kernel32" Alias "RtlMoveMemory" (Destination As CBT_CREATEWND, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromCBT_CreateWnd Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As CBT_CREATEWND, ByVal Length As Long)

Private Type CBTACTIVATESTRUCT
     fMouse As Long
     hWndActive As Long
End Type

Private Declare Sub CopyMemoryCBTActivateStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As CBTACTIVATESTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' ##EVENT_DESCRIPTION BeforeWindowCreate - Triggered when a new window is being created
' ##EVENT_REMARKS Name - The name of the new window.  If this is a top level window then this is its caption
' ##EVENT_REMARKS Height - The initial height of the window in pixels
' ##EVENT_REMARKS Width - The initial width of the window in pixels
' ##EVENT_REMARKS Left - The initial left position of the window
' ##EVENT_REMARKS Top - The initial top position of the window
' ##EVENT_REMARKS Style - Which window style bits are set for the window being created
' ##EVENT_REMARKS ExStyle - Which extended window style bits are set for the window being created
' ##EVENT_REMARKS ParentWindow - The window that this one is being created as a child of - this may be the desktop for a top level window
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being created
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeWindowCreate(ByVal Name As String, Height As Long, Width As Long, _
                                Left As Long, Top As Long, Style As Long, ExtendedStyle As Long, ByVal ParentWindow As ApiWindow, Cancel As Boolean)
Attribute BeforeWindowCreate.VB_Description = "BeforeWindowCreate - Triggered when a new window is being created"
Attribute BeforeWindowCreate.VB_HelpID = 3166

' ##EVENT_DESCRIPTION BeforeWindowActivate - The system is about to activate a window
' ##EVENT_REMARKS NewActiveWindow - The %window:EventVB~ApiWindow% about to be activated
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being activated
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeWindowActivate(ByVal NewActiveWindow As ApiWindow, Cancel As Boolean)
Attribute BeforeWindowActivate.VB_Description = "BeforeWindowActivate - The system is about to activate a window"
Attribute BeforeWindowActivate.VB_HelpID = 2993

' ##EVENT_DESCRIPTION BeforeWindowDestroy - The system is about to activate a window
' ##EVENT_REMARKS DestroyWindow - The %window:EventVB~ApiWindow% about to be destroyed
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being destroyed
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeWindowDestroy(ByVal DestroyWindow As ApiWindow, Cancel As Boolean)
Attribute BeforeWindowDestroy.VB_Description = "BeforeWindowDestroy - The system is about to activate a window"
Attribute BeforeWindowDestroy.VB_HelpID = 2994


' ##EVENT_DESCRIPTION BeforeWindowMinmaxRestore - The system is about to change the show state of a window
' ##EVENT_REMARKS ShowWindow - The %window:EventVB~ApiWindow% about to be minimised, maximised or restored
' ##EVENT_REMARKS ShowCommand - The state change to be affected
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being changed
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeWindowMinmaxRestore(ByVal ShowWindow As ApiWindow, ByVal ShowCommand As ShowWindowCommands, Cancel As Boolean)
Attribute BeforeWindowMinmaxRestore.VB_Description = "BeforeWindowMinmaxRestore - The system is about to change the show state of a window"
Attribute BeforeWindowMinmaxRestore.VB_HelpID = 2995

' ##EVENT_DESCRIPTION BeforeWindowMoveSize - The window is being moved or resized
' ##EVENT_REMARKS MoveSizeWindow - The %window:EventVB~ApiWindow% about to be moved or resized
' ##EVENT_REMARKS NewPosition - The rectangle bounding the new position of the window.  You can alter the rectangle to change the size/move operation
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being moved or resized
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeWindowMoveSize(ByVal MoveSizeWindow As ApiWindow, ByVal newPosition As APIRect, Cancel As Boolean)
Attribute BeforeWindowMoveSize.VB_Description = "BeforeWindowMoveSize - The window is being moved or resized"
Attribute BeforeWindowMoveSize.VB_HelpID = 2996

' ##EVENT_DESCRIPTION BeforeWindowMoveSize - The window is being moved or resized
' ##EVENT_REMARKS GainFocus - The %window:EventVB~ApiWindow% that is gaining the keyboard focus
' ##EVENT_REMARKS LoseFocus - The %window:EventVB~ApiWindow% that is losing the focus
' ##EVENT_REMARKS Cancel - Set this to true to prevent the focus changing
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event BeforeFocusChanged(ByVal GainFocus As ApiWindow, ByVal LoseFocus As ApiWindow, Cancel As Boolean)
Attribute BeforeFocusChanged.VB_Description = "BeforeWindowMoveSize - The window is being moved or resized"
Attribute BeforeFocusChanged.VB_HelpID = 2997

' ##ENUMERATION_DESCRIPTION SystemCommandActivationMethods - Indicates how a system command was initiated
Public Enum SystemCommandActivationMethods
    ' ##ENUMERATION_MEMBER_DESCRIPTION SCAM_BYMNEMONIC A mnemonic triggered the system command
    SCAM_BYMNEMONIC = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SCAM_BYMOUSE A system command was triggered by a mouse click on a menu
    SCAM_BYMOUSE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SCAM_BYSYSACCELERATOR A system accelerator key combination triggered the system command
    SCAM_BYSYSACCELERATOR = 2
End Enum

' ##EVENT_DESCRIPTION MonitorPowerOff - The monitor is being switched off by the power saving system
' ##EVENT_REMARKS ActivatedBy - How this command was triggered
' ##EVENT_REMARKS CursorPosition - The coordinates of the mouse when the menu was clicked (if ActivatedBy is SCAM_BYMOUSE)
' ##EVENT_REMARKS Cancel - Set this to true to prevent the monitor switching off
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event MonitorPowerOff(ByVal AcivatedBy As SystemCommandActivationMethods, ByVal CursorPosition As APIPoint, Cancel As Boolean)
Attribute MonitorPowerOff.VB_Description = "MonitorPowerOff - The monitor is being switched off by the power saving system"
Attribute MonitorPowerOff.VB_HelpID = 2998

' ##EVENT_DESCRIPTION MonitorLowPower - The monitor is being switched to low power by the power saving system
' ##EVENT_REMARKS ActivatedBy - How this command was triggered
' ##EVENT_REMARKS CursorPosition - The coordinates of the mouse when the menu was clicked (if ActivatedBy is SCAM_BYMOUSE)
' ##EVENT_REMARKS Cancel - Set this to true to prevent the monitor power save
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event MonitorLowPower(ByVal AcivatedBy As SystemCommandActivationMethods, ByVal CursorPosition As APIPoint, Cancel As Boolean)
Attribute MonitorLowPower.VB_Description = "MonitorLowPower - The monitor is being switched to low power by the power saving system"
Attribute MonitorLowPower.VB_HelpID = 2999


' ##EVENT_DESCRIPTION ActivateWindowByHotkey - A window is being activated in response to an application hotkey
' ##EVENT_REMARKS WindowToActivate - The window that is going to be activated as a result of the hotkey
' ##EVENT_REMARKS Cancel - Set this to true to prevent the window being activated
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event ActivateWindowByHotkey(ByVal WindowToActivate As ApiWindow, Cancel As Boolean)
Attribute ActivateWindowByHotkey.VB_Description = "ActivateWindowByHotkey - A window is being activated in response to an application hotkey"
Attribute ActivateWindowByHotkey.VB_HelpID = 3000

' ##EVENT_DESCRIPTION InitiateContextHelp - Triggered when the pointer switches to context sensitive help mode
' ##EVENT_REMARKS ActivatedBy - How this command was triggered
' ##EVENT_REMARKS CursorPosition - The coordinates of the mouse when the menu was clicked (if ActivatedBy is SCAM_BYMOUSE)
' ##EVENT_REMARKS Cancel - Set this to prevent the context help being initiated
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event InitiateContextHelp(ByVal AcivatedBy As SystemCommandActivationMethods, ByVal CursorPosition As APIPoint, Cancel As Boolean)
Attribute InitiateContextHelp.VB_Description = "InitiateContextHelp - Triggered when the pointer switches to context sensitive help mode"
Attribute InitiateContextHelp.VB_HelpID = 3001

' ##EVENT_DESCRIPTION StartMenu - Triggered when the start menu is opened
' ##EVENT_REMARKS ActivatedBy - How this command was triggered
' ##EVENT_REMARKS CursorPosition - The coordinates of the mouse when the menu was clicked (if ActivatedBy is SCAM_BYMOUSE)
' ##EVENT_REMARKS Cancel - Set this to prevent the context help being initiated
' ##EVENT_REMARKS Triggered if a WH_CBT hook is installed
Public Event StartMenu(ByVal AcivatedBy As SystemCommandActivationMethods, ByVal CursorPosition As APIPoint, Cancel As Boolean)
Attribute StartMenu.VB_Description = "StartMenu - Triggered when the start menu is opened"
Attribute StartMenu.VB_HelpID = 3002


'-- WH_FOREGROUNDIDLE hook event -----------------------------------------------------------------------------------------------------------
' ##EVENT_DESCRIPTION ForegroundIdle - Triggered when the foreground thread becomes idle
' ##EVENT_REMARKS Triggered if a WH_FOREGROUNDIDLE hook is installed
' ##EVENT_REMARKS This hook type is not supported on Windows 95/98/Me
Public Event ForegroundIdle()
Attribute ForegroundIdle.VB_Description = "ForegroundIdle - Triggered when the foreground thread becomes idle"
Attribute ForegroundIdle.VB_HelpID = 3003

'-- WH_JOURNALRECORD hook events -------------------------------------------------------------------------------
Private Type EVENTMSG
    Message As Long
    paramL As Long
    paramH As Long
    time As Long
    hwnd As Long
End Type


Private Declare Sub CopyMemoryToEventMsg Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As EVENTMSG, ByVal Length As Long)
Private Declare Sub CopyMemoryEventMsg Lib "kernel32" Alias "RtlMoveMemory" (Destination As EVENTMSG, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long

' ##EVENT_DESCRIPTION JournalMouseMessage - Triggered by a mouse event being recorded
' ##EVENT_REMARKS msg - The actual mouse message that occured
' ##EVENT_REMARKS x - The x position of the cursor when the message occured
' ##EVENT_REMARKS y - The y position of the cursor when the message occured
' ##EVENT_REMARKS Time - The time when the message occured, measured in system clock ticks
' ##EVENT_REMARKS TargetWindow - The %window:EventVB~ApiWindow% that will recieve this mouse message
' ##EVENT_REMARKS Triggered if a WH_JOURNALRECORD hook is installed
Public Event JournalMouseMessage(ByVal msg As WindowMessages, ByVal x As Long, ByVal y As Long, ByVal time As Long, ByVal TargetWindow As ApiWindow)
Attribute JournalMouseMessage.VB_Description = "JournalMouseMessage - Triggered by a mouse event being recorded"
Attribute JournalMouseMessage.VB_HelpID = 3004


' ##EVENT_DESCRIPTION JournalKeyboardMessage - Triggered by a keyboard event being recorded
' ##EVENT_REMARKS msg - The actual keyboard message that occured
' ##EVENT_REMARKS VKey - The virtual key code of the key that triggered the event
' ##EVENT_REMARKS ScanCode - The hardware specific scan code of the key that triggered the message
' ##EVENT_REMARKS Time - The time when the message occured, measured in system clock ticks
' ##EVENT_REMARKS TargetWindow - The %window:EventVB~ApiWindow% that will recieve this keyboard message
' ##EVENT_REMARKS Triggered if a WH_JOURNALRECORD hook is installed
Public Event JournalKeyboardMessage(ByVal msg As WindowMessages, ByVal VKey As Long, ByVal scanCode As Long, ByVal time As Long, ByVal TargetWindow As ApiWindow)
Attribute JournalKeyboardMessage.VB_Description = "JournalKeyboardMessage - Triggered by a keyboard event being recorded"
Attribute JournalKeyboardMessage.VB_HelpID = 3005

'-- WH_JOURNALPLAYBACK hook events ------------------------------------------------------------------------------------------------------------

' ##EVENT_DESCRIPTION PlaybackMouseMessage - Triggered by a playback hook to request a mouse message
' ##EVENT_REMARKS If no message is passed in and the playback is not cancelled, a PlaybackKeyboardMessage event will be triggered
' ##EVENT_REMARKS PauseTime - Number of ticks to wait before processing the mouse message.  Set to zero to process it immediately
' ##EVENT_REMARKS msg - Set this to the mouse message that is played back, or zero to indicate no mouse message
' ##EVENT_REMARKS x - The x position of the cursor when the message is played back
' ##EVENT_REMARKS y - The y position of the cursor when the message is played back
' ##EVENT_REMARKS TargetWindow - The %window:EventVB~ApiWindow% that will recieve this mouse message
' ##EVENT_REMARKS Triggered if a WH_JOURNALPLAYBACK hook is installed
Public Event PlaybackMouseMessage(PauseTime As Long, msg As WindowMessages, x As Long, y As Long, TargetWindow As ApiWindow)
Attribute PlaybackMouseMessage.VB_Description = "PlaybackMouseMessage - Triggered by a playback hook to request a mouse message"
Attribute PlaybackMouseMessage.VB_HelpID = 3006


' ##EVENT_DESCRIPTION PlaybackKeyboardMessage - Triggered by a playback hook to request a keyboard message
' ##EVENT_REMARKS PauseTime - Number of ticks to wait before processing the keyboard message.  Set to zero to process it immediately
' ##EVENT_REMARKS msg - Set this to the keyboard message that is played back, or zero to indicate no keyboard message
' ##EVENT_REMARKS VKey - The virtual key code of the key that triggers the event
' ##EVENT_REMARKS ScanCode - The hardware specific scan code of the key that triggers the message
' ##EVENT_REMARKS TargetWindow - The %window:EventVB~ApiWindow% that will recieve this mouse message
' ##EVENT_REMARKS Triggered if a WH_JOURNALPLAYBACK hook is installed
Public Event PlaybackKeyboardMessage(PauseTime As Long, msg As WindowMessages, VKey As Long, scanCode As Long, TargetWindow As ApiWindow)
Attribute PlaybackKeyboardMessage.VB_Description = "PlaybackKeyboardMessage - Triggered by a playback hook to request a keyboard message"
Attribute PlaybackKeyboardMessage.VB_HelpID = 3007

'-- WH_DEBUG --------------------------------------------------------------------------------------------------------------------------------------------------


Private Type DebugHookInfo
    hModuleHook As Long
    hInstaller As Long
    lParam As Long
    wParam As Long
    Code As Long
End Type

Private Declare Sub CopyMemoryDebugHookInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DebugHookInfo, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrDebugHookInfo Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long

' ##EVENT_DESCRIPTION DebugHookTriggered - Triggered when a hook being debugged is called
' ##EVENT_REMARKS HookType - What kind of hook proc is going to be called
' ##EVENT_REMARKS Cancel - Prevents the message being passed on to the hook being called
' ##EVENT_REMARKS HookThread - The %thread:EventVB~ApiThread% that contains the code for the hook that will be called
' ##EVENT_REMARKS HookInstaller - The %thread:EventVB~ApiThread% that initaited the hook that will be called
' ##EVENT_REMARKS wParam - The wParam value that will be passed to the hook procedure
' ##EVENT_REMARKS lParam - The lParam value that will be passed to the hook procedure
' ##EVENT_REMARKS Code - The hook code that will be passed to the hook procedure
' ##EVENT_REMARKS Triggered if a WH_DEBUG hook is installed
' ##EVENT_REMARKS This hook type is only supported in Windows NT/2000/XP
Public Event DebugHookTriggered(ByVal HookType As enHookTypes, ByVal HookThread As ApiThread, ByVal HookInstaller As ApiThread, ByVal wParam As Long, ByVal lParam As Long, ByVal Code As enHookCode, Cancel As Boolean)
Attribute DebugHookTriggered.VB_Description = "DebugHookTriggered - Triggered when a hook being debugged is called"
Attribute DebugHookTriggered.VB_HelpID = 3008


'-- WH_GETMESSAGE --------------------------------------------------------------------------------------
Private Type msg
    hwnd As Long
    Message As WindowMessages
    wParam As Long
    lParam As Long
    time As Long
    x As Long
    y As Long
End Type

Private Declare Sub CopyMemoryMSG Lib "kernel32" Alias "RtlMoveMemory" (Destination As msg, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryToMSG Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As msg, ByVal Length As Long)



' ##EVENT_DESCRIPTION AnyWindowMessageHandled is triggered when a window message is looked at in the message queue while a WH_GETMSG hook is set
' ##EVENT_REMARKS TargetWnd - The window for which this context menu is to be created
' ##EVENT_REMARKS msg - The window message that is being passed to the target window
' ##EVENT_REMARKS wParam - The message's wParam member, the meaning of which depends on the message
' ##EVENT_REMARKS lParam - The message's lParam member, the meaning of which depends on the message
' ##EVENT_REMARKS Triggered if a WH_GETMESSAGE hook is installed
Public Event AnyWindowMessageHandled(ByVal TargetWnd As ApiWindow, ByVal Removed As Boolean, msg As WindowMessages, wParam As Long, lParam As Long)
Attribute AnyWindowMessageHandled.VB_Description = "AnyWindowMessageHandled is triggered when a window message is looked at in the message queue while a WH_GETMSG hook is set"
Attribute AnyWindowMessageHandled.VB_HelpID = 3184

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Let HookIdByType(ByVal HookType As enHookTypes, ByVal newHookId As Long)
Attribute HookIdByType.VB_HelpID = 2044

Select Case HookType
Case WH_CALLWNDPROC
    If newHookId <> CallWndProcHookId Then
        CallWndProcHookId = newHookId
    End If

Case WH_CBT
    If newHookId <> CBTHookId Then
        CBTHookId = newHookId
    End If
    
Case WH_DEBUG
    If newHookId <> DebugHookId Then
        DebugHookId = newHookId
    End If
    
Case WH_FOREGROUNDIDLE
    If newHookId <> ForegroundidleHookId Then
        ForegroundidleHookId = newHookId
    End If
    
Case WH_GETMESSAGE
    If newHookId <> GetMessageHookid Then
        GetMessageHookid = newHookId
    End If
    
Case WH_HARDWARE
    If newHookId <> HardwareHookId Then
        HardwareHookId = newHookId
    End If
    
Case WH_JOURNALPLAYBACK
    If newHookId <> JournalPlaybackHookId Then
        JournalPlaybackHookId = newHookId
    End If
    
Case WH_JOURNALRECORD
    If newHookId <> JournalRecordHookId Then
        JournalRecordHookId = newHookId
    End If
    
Case WH_MOUSE
    If newHookId <> MouseHookId Then
        MouseHookId = newHookId
    End If
    
Case WH_MSGFILTER
    If newHookId <> MsgFilterHookId Then
        MsgFilterHookId = newHookId
    End If
    
Case WH_SHELL
    If newHookId <> ShellHookId Then
        ShellHookId = newHookId
    End If
    
Case WH_SYSMSGFILTER
    If newHookId <> SysmessageHookId Then
        SysmessageHookId = newHookId
    End If
    
Case WH_KEYBOARD_LL
    If newHookId <> LowLevelKeyboardHookId Then
        LowLevelKeyboardHookId = newHookId
    End If
    
Case WH_MOUSE_LL
    If newHookId <> LowlevelMouseHookId Then
        LowlevelMouseHookId = newHookId
    End If
    
End Select

End Property

Public Property Get HookIdByType(ByVal HookType As enHookTypes) As Long

Select Case HookType
Case WH_CALLWNDPROC
    HookIdByType = CallWndProcHookId
Case WH_CBT
    HookIdByType = CBTHookId
Case WH_DEBUG
    HookIdByType = DebugHookId
Case WH_FOREGROUNDIDLE
    HookIdByType = ForegroundidleHookId
Case WH_GETMESSAGE
    HookIdByType = GetMessageHookid
Case WH_HARDWARE
    HookIdByType = HardwareHookId
Case WH_JOURNALPLAYBACK
    HookIdByType = JournalPlaybackHookId
Case WH_JOURNALRECORD
    HookIdByType = JournalRecordHookId
Case WH_MOUSE
    HookIdByType = MouseHookId
Case WH_MSGFILTER
    HookIdByType = MsgFilterHookId
Case WH_SHELL
    HookIdByType = ShellHookId
Case WH_SYSMSGFILTER
    HookIdByType = SysmessageHookId
Case WH_KEYBOARD_LL
    HookIdByType = LowLevelKeyboardHookId
Case WH_MOUSE_LL
    HookIdByType = LowlevelMouseHookId
End Select

End Property

Public Sub StartHook(ByVal HookType As enHookTypes, ByVal Scope As HookScopeSettings)
Attribute StartHook.VB_HelpID = 2045

Dim lRet As Long
Dim hMod As Long
Dim ModuleHandle As Long
Dim ThreadId As Long

If Scope = HOOK_LOCAL_PROCESS Then
    ThreadId = App.ThreadId
Else
    ModuleHandle = APIDispenser.AppModuleHandle
End If

'\\ If a hook of this type is already set, unhook this first
If HookIdByType(HookType) > 0 Then
    Call UnhookWindowsHookEx(HookIdByType(HookType))
End If

'\\ Start the appropriate hook procedure
Select Case HookType
Case WH_CALLWNDPROC
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKCALLWNDPROC, ModuleHandle, ThreadId)
Case WH_CBT
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKCBTPROC, ModuleHandle, ThreadId)
Case WH_DEBUG
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKDEBUGPROC, ModuleHandle, ThreadId)
Case WH_FOREGROUNDIDLE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKFOREGROUNDIDLEPROC, ModuleHandle, ThreadId)
Case WH_GETMESSAGE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKGETMESSAGEPROC, ModuleHandle, ThreadId)
Case WH_JOURNALPLAYBACK
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKJOURNALPLAYBACKPROC, ModuleHandle, ThreadId)
Case WH_JOURNALRECORD
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKJOURNALRECORDPROC, ModuleHandle, ThreadId)
Case WH_MOUSE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKMOUSEPROC, ModuleHandle, ThreadId)
Case WH_MSGFILTER
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKMESSAGEFILTERPROC, ModuleHandle, ThreadId)
Case WH_SHELL
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKSHELLPROC, ModuleHandle, ThreadId)
Case WH_SYSMSGFILTER
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKSYSMESSAGEFILTERPROC, ModuleHandle, ThreadId)
Case WH_KEYBOARD_LL
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKLOWLEVELKEYBOARDPROC, ModuleHandle, ThreadId)
Case WH_MOUSE_LL
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKLOWLEVELMOUSEPROC, ModuleHandle, ThreadId)
Case WH_KEYBOARD
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKKEYBOARDPROC, ModuleHandle, ThreadId)
Case Else
    '\\ Non specific hook type
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKPROC, ModuleHandle, ThreadId)
End Select


If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiSystemHook:StartHook", GetLastSystemError)
End If

If lRet > 0 Then
    HookIdByType(HookType) = lRet
    CurrentHookType = HookType
End If

End Sub

Public Sub StopHook(ByVal HookType As enHookTypes)
Attribute StopHook.VB_HelpID = 2046

Dim lRet As Long

'\\ If a hook of this type is already set, unhook this first
If HookIdByType(HookType) > 0 Then
    lRet = UnhookWindowsHookEx(HookIdByType(HookType))
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiSystemHook:StopHook", GetLastSystemError)
    Else
        HookIdByType(HookType) = 0
    End If
End If

End Sub


Friend Sub TriggerEvent(ByVal ProcType As CallbackProcTypes, Arguments() As Variant)
Attribute TriggerEvent.VB_HelpID = 2047

Dim hwnd As Long, wMsg As Long, wParam As Long, lParam As Long
Dim nEventId As Long, dWord As Long, dwTime As Long
Dim hdc As Long, nInt As Long
Dim Code As Long
Dim lResult As Long
Dim LPCSTR As String, Handle As Long
Dim lpch As String, ichCurrent As Long, cch As Long
Dim lRet As Long, lWndProcRet As Long
Dim bDiscardMsg As Boolean
Dim ftTime As APIFileTime
Dim stTime As APISystemTime
Dim wndThis As ApiWindow
Dim dtTime As Date
Dim lpstrName As String
Dim hModule As Long, lpType As String, lpName As String, wLanguage As Long
Dim bCurrentProc As Boolean
Dim ptThis As APIPoint
Dim Cancel As Boolean
Dim TargetWnd As ApiWindow
Dim x As Long, y As Long

Select Case ProcType

Case WNDPROC

    hwnd = Arguments(1)
    wMsg = Arguments(2)
    wParam = Arguments(3)
    lParam = Arguments(4)
    
    On Error Resume Next
    Call AllSubclassedWindows.Item(hwnd).TriggerEvent(wMsg, wParam, lParam, bDiscardMsg, lWndProcRet)
    
    If Not bDiscardMsg Then
        If AllSubclassedWindows.Item(hwnd).OldProcAddress > 0 Then
            lRet = CallWindowProc(AllSubclassedWindows.Item(hwnd).OldProcAddress, hwnd, wMsg, wParam, lParam)
        Else
            lRet = DefWindowProc(hwnd, wMsg, wParam, lParam)
        End If
    End If
    '\\ DEJ 2003-02-26 - If a WM_CLOSE message is recieved, unsubclass the window safely..
    If wMsg = WM_CLOSE Then
        AllSubclassedWindows.Remove AllSubclassedWindows.Item(hwnd)
    End If
    
    If lWndProcRet = 0 Then
        Arguments(5) = lRet
    Else
        Arguments(5) = lWndProcRet
    End If
    
'\\ DEJ 01/03/2001 - New specific hook procedures
Case HOOKPROC_CALLWNDPROC
    Code = Arguments(1)
    bCurrentProc = Arguments(2)
    lParam = Arguments(3)
    '\\ Pass lResult back...
    Arguments(4) = lResult
    
Case HOOKPROC_CALLWNDPROCRET
    Code = Arguments(1)
    bCurrentProc = Arguments(2)
    lParam = Arguments(3)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_CBT
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    '\\ Create a new window for these to use...
    Set wndThis = New ApiWindow
    
    '\\ Different events raised according to which message this is...
    Select Case True
    Case Code = HCBT_ACTIVATE
        wndThis.hwnd = wParam
        RaiseEvent BeforeWindowActivate(wndThis, Cancel)
        If Cancel Then
            Arguments(4) = 1
        End If
        
    Case Code = HCBT_CREATEWND
        Dim CreateWndStruct As CBT_CREATEWND
        Dim csThis As CREATESTRUCT
        Dim sName As String
        Call CopyMemoryCBT_CreateWnd(CreateWndStruct, lParam, Len(CreateWndStruct))
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":HCBT_CREATEWND", GetLastSystemError
        Else
            Call CopyMemoryCreateStruct(csThis, CreateWndStruct.csThis, Len(csThis))
            With csThis
                sName = StringFromPointer(.lpszName, 32)
                If .hWndParent = 0 Then
                    wndThis.hwnd = GetDesktopWindow
                Else
                    wndThis.hwnd = .hWndParent
                End If
                RaiseEvent BeforeWindowCreate(sName, .Height, .Width, .Left, .Top, .Style, .ExStyle, wndThis, Cancel)
            End With
            If Not Cancel Then
                Call CopyMemoryFromCREATESTRUCT(CreateWndStruct.csThis, csThis, Len(csThis))
            Else
                Arguments(4) = 1
            End If
        End If
    
    Case Code = HCBT_DESTROYWND
        wndThis.hwnd = wParam
        RaiseEvent BeforeWindowDestroy(wndThis, Cancel)
        If Cancel Then
            Arguments(4) = 1
        End If
        
    Case Code = HCBT_MINMAX
        wndThis.hwnd = wParam
        RaiseEvent BeforeWindowMinmaxRestore(wndThis, LoWord(lParam), Cancel)
        If Cancel Then
            Arguments(4) = 1
        End If
        
    Case Code = HCBT_MOVESIZE
        Dim SizeRect As APIRect
        Set SizeRect = New APIRect
        SizeRect.CreateFromPointer lParam
        wndThis.hwnd = wParam
        RaiseEvent BeforeWindowMoveSize(wndThis, SizeRect, Cancel)
        '\\ The sizerect can be changed :. save it back..
        If SizeRect.Changed Then
            SizeRect.SaveToPointer lParam
        End If
        Set SizeRect = Nothing
        If Cancel Then
            Arguments(4) = 1
        End If
        
    Case Code = HCBT_SETFOCUS
        Dim wndLostFocus As ApiWindow
        Set wndLostFocus = New ApiWindow
        wndThis.hwnd = wParam
        wndLostFocus.hwnd = lParam
        RaiseEvent BeforeFocusChanged(wndThis, wndLostFocus, Cancel)
        If Cancel Then
            Arguments(4) = 1
        End If
        
    Case Code = HCBT_SYSCOMMAND
        Dim ptMouse As New APIPoint
        Dim activeBy As SystemCommandActivationMethods
        Select Case HiWord(lParam)
        Case 0
            activeBy = SCAM_BYMNEMONIC
        Case -1
            activeBy = SCAM_BYSYSACCELERATOR
        Case Else
            activeBy = SCAM_BYMOUSE
            ptMouse.y = HiWord(lParam)
            ptMouse.x = LoWord(lParam)
        End Select
        '\\ Only the upper 3 bytes are used - the lower 4 are reserved for the system...
        wParam = (wParam Or &HFFFFF000)
        Select Case wParam
        Case SC_ARRANGE
        Case SC_CLOSE
        Case SC_HOTKEY
            '\\ A window is being activated by a hotkey
            wndThis.hwnd = lParam
            RaiseEvent ActivateWindowByHotkey(wndThis, Cancel)
        Case SC_HSCROLL
        Case SC_KEYMENU
        Case SC_MAXIMIZE
        Case SC_MINIMIZE
        Case SC_MOUSEMENU
        Case SC_MOVE
        Case SC_NEXTWINDOW
        Case SC_PREVWINDOW
        Case SC_RESTORE
        Case SC_SCREENSAVE
        Case SC_SIZE
        Case SC_TASKLIST
            RaiseEvent StartMenu(activeBy, ptMouse, Cancel)
        Case SC_VSCROLL
        Case SC_MONITORPOWER
            If lParam = 1 Then
                RaiseEvent MonitorLowPower(activeBy, ptMouse, Cancel)
            ElseIf lParam = 2 Then
                RaiseEvent MonitorPowerOff(activeBy, ptMouse, Cancel)
            End If
            If Cancel Then
                Arguments(4) = 1
            End If
            
        Case SC_CONTEXTHELP
            RaiseEvent InitiateContextHelp(activeBy, ptMouse, Cancel)
        End Select
        
    End Select
    If Cancel Then
        Arguments(4) = 1
    End If
        
Case HOOKPROC_DEBUG
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Dim DebugInfo As DebugHookInfo
    Dim thrdHook As New ApiThread
    Dim thrCaller As New ApiThread
    If Not IsBadReadPtrDebugHookInfo(lParam, Len(DebugInfo)) Then
        Call CopyMemoryDebugHookInfo(DebugInfo, lParam, Len(DebugInfo))
    End If
    With DebugInfo
        thrdHook.ThreadId = .hModuleHook
        thrCaller.ThreadId = .hInstaller
        RaiseEvent DebugHookTriggered(wParam, thrdHook, thrCaller, .wParam, .lParam, .Code, Cancel)
    End With
    '\\ Pass lResult back...
    If Cancel Then
        Arguments(4) = True
    End If
    
Case HOOKPROC_GETMESSAGE
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    '\\ Read the MSG
    Dim msgIn As msg
    Call CopyMemoryMSG(msgIn, lParam, Len(msgIn))
    '\\ Raise events according to the message....
    Set TargetWnd = New ApiWindow
    With msgIn
        TargetWnd.hwnd = .hwnd
        RaiseEvent AnyWindowMessageHandled(TargetWnd, (wParam = PM_REMOVE), .Message, .wParam, .lParam)
    End With
    Call CopyMemoryToMSG(lParam, msgIn, Len(msgIn))

Case HOOKPROC_JOURNALPLAYBACK
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    
    Static lTimeOut As Long
    Dim msg As WindowMessages
    Dim xPos As Long
    Dim yPos As Long
    Static emOut As EVENTMSG
       
    Select Case Code
    Case HC_GETNEXT
       If emOut.Message <> 0 Then
         If Not IsBadWritePtr(lParam, Len(emOut)) Then
           Call CopyMemoryToEventMsg(lParam, emOut, Len(emOut))
         End If
       End If
       
       '\\ If last message required a timeout then this will be called again when that timeout occurs...
       If lTimeOut > 0 Then
          lResult = lTimeOut
          lTimeOut = 0
       End If
       
    Case HC_SYSMODALON
        '\\ A system modal message box has opened :. playback must be paused
        
    Case HC_SYSMODALOFF
        '\\ A system modal message box has closed :. playback can resume
    
    Case HC_SKIP
       '\\ Get the next message....
       With emOut
          RaiseEvent PlaybackMouseMessage(lTimeOut, msg, .paramH, .paramL, TargetWnd)
       End With
       If msg > 0 And (Not Cancel) Then
          '\\ Copy the message to lParam
          With emOut
            If Not TargetWnd Is Nothing Then
                .hwnd = TargetWnd.hwnd
            End If
            .Message = msg
          End With
       Else
          '\\ if no message returned try for a keyboard message
          With emOut
            RaiseEvent PlaybackKeyboardMessage(lTimeOut, msg, .paramH, .paramL, TargetWnd)
          End With
          If msg > 0 And (Not Cancel) Then
            '\\ Copy the keybaord message to lParam
            With emOut
              If Not TargetWnd Is Nothing Then
                  .hwnd = TargetWnd.hwnd
              End If
              .Message = msg
            End With
          End If
        End If
    End Select
    '\\ Pass lResult back...
    Arguments(4) = lResult

    
Case HOOKPROC_MOUSE_LL, HOOKPROC_MOUSE
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Set ptThis = New APIPoint
    '\\ Convert the pointer into a mouse hook structure object
    If ProcType = HOOKPROC_MOUSE Then
        Dim mhData As MOUSEHOOKSTRUCT
        Call CopyMemoryMouseHookStruct(mhData, lParam, Len(mhData))
        Call ptThis.CreateFromPointer(VarPtr(mhData.pt))
    Else
        Dim llmData As MSLLHOOKSTRUCT
        Call CopyMemoryMSLLHOOKSTRUCT(llmData, lParam, Len(llmData))
        Call ptThis.CreateFromPointer(VarPtr(llmData.pt))
    End If
    '\\ Get the window under the cursor
    Set wndThis = New ApiWindow
    wndThis.hwnd = WindowFromPoint(ptThis.x, ptThis.y)
    If wParam = WM_LBUTTONDOWN Then
        RaiseEvent MouseButtonDown(vbLeftButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_RBUTTONDOWN Then
        RaiseEvent MouseButtonDown(vbRightButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_MBUTTONDOWN Then
        RaiseEvent MouseButtonDown(vbMiddleButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_LBUTTONUP Then
        RaiseEvent MouseButtonUp(vbLeftButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_RBUTTONUP Then
        RaiseEvent MouseButtonUp(vbRightButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_MBUTTONUP Then
        RaiseEvent MouseButtonUp(vbMiddleButton, ptThis, wndThis, Cancel)
    ElseIf wParam = WM_MOUSEMOVE Then
        RaiseEvent MouseMove(ptThis, wndThis, Cancel)
    ElseIf wParam = WM_MOUSEWHEEL Then
        RaiseEvent MouseWheelMove(HiWord(llmData.MouseData), ptThis, wndThis, Cancel)
    End If
    '\\ Pass lResult back...
    If Cancel Then
        Arguments(4) = 1
    End If

Case HOOKPROC_MESSAGEFILTER
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_SHELL
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    Set wndThis = New ApiWindow
    wndThis.hwnd = wParam
    If Code = HSHELL_WINDOWCREATED Then
        '\\ An application main window is being created
        RaiseEvent WindowCreated(wndThis)
    ElseIf Code = HSHELL_WINDOWDESTROYED Then
        '\\ An application main window is being destroyed
        RaiseEvent WindowDestroyed(wndThis)
    ElseIf Code = HSHELL_GETMINRECT Then
        '\\ System wants to know the min or max size to make a window....
    ElseIf Code = HSHELL_REDRAW Then
        '\\ A title bar of an application has changed....
        RaiseEvent WindowTitleRedrawn(wndThis, CBool(lParam))
    ElseIf Code = HSHELL_TASKMAN Then
        RaiseEvent TaskManagerRequested(Cancel)
        If Cancel Then
            lResult = True
        End If
    ElseIf Code = HSHELL_LANGUAGE Then
        RaiseEvent KeyboardLanguageChanged(wndThis, lParam)
    ElseIf Code = HSHELL_ACCESSIBILITYSTATE Then
        RaiseEvent AccessibilityStateChanged(wParam)
    ElseIf Code = HSHELL_APPCOMMAND Then
        Dim cmdApp As ApplicationCommandMessages
        cmdApp = (HiWord(lParam) And (Not &HF000))
        RaiseEvent ApplicationCommand(wndThis, cmdApp)
    End If
    Set wndThis = Nothing
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_SYSMESSAGEFILTER
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_KEYBOARD
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Dim lRepetition As Long
    Dim lScanCode As Long
    Dim bExtendedKey As Boolean
    Dim bAltDown As Boolean
    Dim bAlreadyPressed As Boolean
    Dim bBeingPressed As Boolean
    With APIDispenser
        lRepetition = .GetBits(lParam, 0, 15)
        lScanCode = .GetBits(lParam, 16, 23)
        bAlreadyPressed = .GetBits(lParam, 30, 30)
        bBeingPressed = .GetBits(lParam, 31, 31)
    End With
    bExtendedKey = (lParam Or KF_EXTENDED) = lParam
    bAltDown = (lParam Or KF_ALTDOWN) = lParam
    If bBeingPressed Then
        RaiseEvent KeyUp(wParam, lScanCode, bExtendedKey, bAltDown, False, Cancel)
    Else
        RaiseEvent KeyDown(wParam, lScanCode, bExtendedKey, bAltDown, False, Cancel)
    End If
    '\\ Pass lResult back...
    If Cancel Then
        Arguments(4) = 1
    End If

Case HOOKPROC_KEYBOARD_LL
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)

    Dim kbdll As KBDLLHOOKSTRUCT
    Call CopyMemoryKbdLLHookStruct(kbdll, lParam, Len(kbdll))
    With kbdll
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(.vkCode, .scanCode, (.flags And LLKHF_EXTENDED), (.flags And LLKHF_ALTDOWN), (.flags And LLKHF_INJECTED), Cancel)
        ElseIf wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            RaiseEvent KeyUp(.vkCode, .scanCode, (.flags And LLKHF_EXTENDED), (.flags And LLKHF_ALTDOWN), (.flags And LLKHF_INJECTED), Cancel)
        End If
    End With
    '\\ Pass lResult back...
    If Cancel Then
        Arguments(4) = 1
    End If
    
Case HOOKPROC_FOREGROUNDIDLE
    '\\ Other argumenst aren't used
    RaiseEvent ForegroundIdle

Case HOOKPROC_JOURNALRECORD
    '\\ NOTE: Do not breakpoint this hook - it is unrecoverable...
    Code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    Dim emThis As EVENTMSG
    If Not IsBadReadPtr(lParam, Len(emThis)) Then
        Call CopyMemoryEventMsg(emThis, lParam, Len(emThis))
        Set wndThis = New ApiWindow
        wndThis.hwnd = emThis.hwnd
        With emThis
            Select Case .Message
            Case WM_MOUSEMOVE, _
                 WM_LBUTTONUP, WM_LBUTTONDOWN, WM_LBUTTONDBLCLK, _
                 WM_MBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONDBLCLK, _
                 WM_RBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONDBLCLK
                RaiseEvent JournalMouseMessage(.Message, .paramH, .paramL, .time, wndThis)
            Case Else
                RaiseEvent JournalKeyboardMessage(.Message, LoByte(LoWord(.paramL)), .paramH, .time, wndThis)
            End Select
        End With
    End If

End Select


End Sub

Private Sub Class_Terminate()

Dim lRet As Long

'\\ If ANY hook id is set, it must be closed
Call StopHook(WH_CALLWNDPROC)
Call StopHook(WH_CBT)
Call StopHook(WH_DEBUG)
Call StopHook(WH_FOREGROUNDIDLE)
Call StopHook(WH_GETMESSAGE)
Call StopHook(WH_HARDWARE)
Call StopHook(WH_JOURNALPLAYBACK)
Call StopHook(WH_JOURNALRECORD)
Call StopHook(WH_KEYBOARD_LL)
Call StopHook(WH_MOUSE)
Call StopHook(WH_MOUSE_LL)
Call StopHook(WH_MSGFILTER)
Call StopHook(WH_SHELL)
Call StopHook(WH_SYSMSGFILTER)
Call StopHook(WH_KEYBOARD)

End Sub



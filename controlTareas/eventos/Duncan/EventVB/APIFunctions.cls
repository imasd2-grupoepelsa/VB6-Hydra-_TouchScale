VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2048
Attribute VB_Description = "This class is the head class of the EventVB project.  It provides a set of generic functions useful for dealing with the API and also the API error event which is raised whenever an internal error occurs when dealing with the API."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ##MODULE_DESCRIPTION This class is the head class of the EventVB project.  It _
provides a set of generic functions useful for dealing with the API and also the _
API error event which is raised whenever an internal error occurs when dealing with _
the API.

Private Const CLSNAME = "APIFunctions"



'\\ Get the window text....
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'\\ Get Window Long Indexes...
Public Enum enGetWindowLong
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum



'\\ Child window enumerations....
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long



Public Enum enExtendedEditMessages
    '\\ RichEdit extended messages
    EM_GETLIMITTEXT = (&H400 + 37)
    EM_POSFROMCHAR = (&H400 + 38)
    EM_CHARFROMPOS = (&H400 + 39)
    EM_CANPASTE = (&H400 + 50)
    EM_DISPLAYBAND = (&H400 + 51)
    EM_EXGETSEL = (&H400 + 52)
    EM_EXLIMITTEXT = (&H400 + 53)
    EM_EXLINEFROMCHAR = (&H400 + 54)
    EM_EXSETSEL = (&H400 + 55)
    EM_FINDTEXT = (&H400 + 56)
    EM_FORMATRANGE = (&H400 + 57)
    EM_GETCHARFORMAT = (&H400 + 58)
    EM_GETEVENTMASK = (&H400 + 59)
    EM_GETOLEINTERFACE = (&H400 + 60)
    EM_GETPARAFORMAT = (&H400 + 61)
    EM_GETSELTEXT = (&H400 + 62)
    EM_HIDESELECTION = (&H400 + 63)
    EM_PASTESPECIAL = (&H400 + 64)
    EM_REQUESTRESIZE = (&H400 + 65)
    EM_SELECTIONTYPE = (&H400 + 66)
    EM_SETBKGNDCOLOR = (&H400 + 67)
    EM_SETCHARFORMAT = (&H400 + 68)
    EM_SETEVENTMASK = (&H400 + 69)
    EM_SETOLECALLBACK = (&H400 + 70)
    EM_SETPARAFORMAT = (&H400 + 71)
    EM_SETTARGETDEVICE = (&H400 + 72)
    EM_STREAMIN = (&H400 + 73)
    EM_STREAMOUT = (&H400 + 74)
    EM_GETTEXTRANGE = (&H400 + 75)
    EM_FINDWORDBREAK = (&H400 + 76)
    EM_SETOPTIONS = (&H400 + 77)
    EM_GETOPTIONS = (&H400 + 78)
    EM_FINDTEXTEX = (&H400 + 79)
    EM_GETWORDBREAKPROCEX = (&H400 + 80)
    EM_SETWORDBREAKPROCEX = (&H400 + 81)
'/* Richedit v2.0 messages */
    EM_SETUNDOLIMIT = (&H400 + 82)
    EM_REDO = (&H400 + 84)
    EM_CANREDO = (&H400 + 85)
    EM_GETUNDONAME = (&H400 + 86)
    EM_GETREDONAME = (&H400 + 87)
    EM_STOPGROUPTYPING = (&H400 + 88)
    EM_SETTEXTMODE = (&H400 + 89)
    EM_GETTEXTMODE = (&H400 + 90)
    EM_AUTOURLDETECT = (&H400 + 91)
    EM_GETAUTOURLDETECT = (&H400 + 92)
    EM_SETPALETTE = (&H400 + 93)
    EM_GETTEXTEX = (&H400 + 94)
    EM_GETTEXTLENGTHEX = (&H400 + 95)
'/* Far East specific messages */
    EM_SETPUNCTUATION = (&H400 + 100)
    EM_GETPUNCTUATION = (&H400 + 101)
    EM_SETWORDWRAPMODE = (&H400 + 102)
    EM_GETWORDWRAPMODE = (&H400 + 103)
    EM_SETIMECOLOR = (&H400 + 104)
    EM_GETIMECOLOR = (&H400 + 105)
    EM_SETIMEOPTIONS = (&H400 + 106)
    EM_GETIMEOPTIONS = (&H400 + 107)
    EM_CONVPOSITION = (&H400 + 108)
    EM_SETLANGOPTIONS = (&H400 + 120)
    EM_GETLANGOPTIONS = (&H400 + 121)
    EM_GETIMECOMPMODE = (&H400 + 122)
    EM_FINDTEXTW = (&H400 + 123)
    EM_FINDTEXTEXW = (&H400 + 124)
'/* BiDi specific messages */
    EM_SETBIDIOPTIONS = (&H400 + 200)
    EM_GETBIDIOPTIONS = (&H400 + 201)
End Enum

Public Enum enRichEditTextModes
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2            '    /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8      '    /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32       '    /* default behavior */
End Enum

Public Enum enEditControlOperations
    ECO_AUTOWORDSELECTION = &H1
    ECO_AUTOVSCROLL = &H40
    ECO_AUTOHSCROLL = &H80
    ECO_NOHIDESEL = &H100
    ECO_READONLY = &H800
    ECO_WANTRETURN = &H1000
    ECO_SAVESEL = &H8000
    ECO_SELECTIONBAR = &H1000000
    ECO_VERTICAL = &H400000
End Enum

Public Enum enEditControlOperationOptions
    ECOOP_SET = &H1
    ECOOP_OR = &H2
    ECOOP_AND = &H3
    ECOOP_XOR = &H4
End Enum

Public Enum enCharformatMasks
    CFM_BOLD = &H1
    CFM_ITALIC = &H2
    CFM_UNDERLINE = &H4
    CFM_STRIKEOUT = &H8
    CFM_PROTECTED = &H10
    CFM_LINK = &H20   '/* Exchange hyperlink extension */
    CFM_SIZE = &H80000000
    CFM_COLOR = &H40000000
    CFM_FACE = &H20000000
    CFM_OFFSET = &H10000000
    CFM_CHARSET = &H8000000
End Enum

Public Enum enCharFormatEffects
    CFE_BOLD = &H1
    CFE_ITALIC = &H2
    CFE_UNDERLINE = &H4
    CFE_STRIKEOUT = &H8
    CFE_PROTECTED = &H10
    CFE_LINK = &H20
    CFE_AUTOCOLOR = &H40000000      '/* NOTE: this corresponds to CFM_COLOR*/
End Enum

Public Enum enSetCharFormatModifiers
    SCF_SELECTION = &H1
    SCF_WORD = &H2
    SCF_DEFAULT = &H0     ' ##  set the default charformat or paraformat
    SCF_ALL = &H4         ' ##  not valid with SCF_SELECTION or SCF_WORD
    SCF_USEUIRULES = &H8  ' ##  modifier for SCF_SELECTION; says that the format came from a toolbar
End Enum

'\\ Edit word break code constants
Public Enum enWordBreak
    WB_ISDELIMITER = 2
    WB_LEFT = 0
    WB_RIGHT = 1
End Enum

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Enum enDialogWindowLong
    DWL_DLGPROC = 4
    DWL_MSGRESULT = 0
    DWL_USER = 8
End Enum

Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetProcessWindowStation Lib "user32" () As Long
Private Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long) As Long
Private Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long


' ##ENUMERATION_DESCRIPTION These are the standard predefined resource types
Public Enum ResourceTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_CURSOR Cursor resource (Device dependent)
    RT_CURSOR = 1&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_BITMAP Bitmap resource
    RT_BITMAP = 2&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_ICON Icon resource (Device dependent)
    RT_ICON = 3&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_MENU Menu resource
    RT_MENU = 4&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_DIALOG Dialog box resource
    RT_DIALOG = 5&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_STRING Text string resource
    RT_STRING = 6&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_FONTDIR Font directory resource
    RT_FONTDIR = 7&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_FONT Font resource
    RT_FONT = 8&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_ACCELERATOR Accelerator table resource
    RT_ACCELERATOR = 9&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_RCDATA Raw application defined data resource
    RT_RCDATA = 10&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_MESSAGETABLE Message table resource
    RT_MESSAGETABLE = 11&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_GROUP_CURSOR Cursor resource (Device independent)
    RT_GROUP_CURSOR = 12&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_GROUP_ICON Icon resource (Device independent)
    RT_GROUP_ICON = 14&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_VERSION Version reource
    RT_VERSION = 16&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_DLGINCLUDE Dialog box include resource
    RT_DLGINCLUDE = 17&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_PLUGPLAY Plug and play resource
    RT_PLUGPLAY = 19&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_VXD VXD Resource
    RT_VXD = 20&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_ANICURSOR Animated Cursor resource (Device dependent)
    RT_ANICURSOR = 21&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_ANIICON Animated Icon resource (Device dependent)
    RT_ANIICON = 22&
    ' ##ENUMERATION_MEMBER_DESCRIPTION RT_HTML HTML Resource (Windows NT 4 and above)
    RT_HTML = 23&
End Enum

'Private Declare Function SendMessageMove Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As RECT) As Long
Private Declare Function SendMessageProc Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long


Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SendMessageRepaint Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Enum enWindowActive
    WA_ACTIVE = 1
    WA_INACTIVE = 0
    WA_CLICKACTIVE = 2
End Enum

Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long



'\\ Windows Menu related API calls
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Private Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long


Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long


'\\ Sending messages to a window....
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lpstrParam As String) As Long

'\\ Shell execute - running external applications
'\\ Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'\\ Scrollbar messages
Public Enum enScrollMessages
    SB_ENDSCROLL = 8
    SB_TOP = 6
    SB_LINELEFT = 0
    SB_LINERIGHT = 1
    SB_LINEDOWN = 1
    SB_LINEUP = 0
    SB_PAGEDOWN = 3
    SB_PAGELEFT = 2
    SB_PAGERIGHT = 3
    SB_PAGEUP = 2
    SB_RIGHT = 7
    SB_LEFT = 6
    SB_THUMBPOSITION = 4
    SB_THUMBTRACK = 5
End Enum

' ##ENUMERATION_DESCRIPTION SystemCommands - The types of command sent to a %window:EventVB~ApiWindow% when a command is selected from the system menu (control box)
Public Enum SystemCommands
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_ARRANGE Arrange the windows
    SC_ARRANGE = &HF110
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_CLOSE Close the window
    SC_CLOSE = &HF060
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_HOTKEY Activate the window associated with the application-specified hot key
    SC_HOTKEY = &HF150
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_HSCROLL Scroll the window horizontally
    SC_HSCROLL = &HF080
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_KEYMENU Retrieve the window menu as a result of a keystroke
    SC_KEYMENU = &HF100
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_MAXIMIZE Maximise the window
    SC_MAXIMIZE = &HF030
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_MINIMIZE Minimise the window
    SC_MINIMIZE = &HF020
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_MOUSEMENU Retrieve the window menu as a result of a mouse click.
    SC_MOUSEMENU = &HF090
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_MOVE Move the window
    SC_MOVE = &HF010
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_NEXTWINDOW Select the next window in the tab order
    SC_NEXTWINDOW = &HF040
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_PREVWINDOW Select the previous window in the tab order
    SC_PREVWINDOW = &HF050
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_RESTORE Restore the window size
    SC_RESTORE = &HF120
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_SCREENSAVE Execute the screen saver application
    SC_SCREENSAVE = &HF140
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_SIZE Size the window
    SC_SIZE = &HF000
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_TASKLIST Activate the Start menu
    SC_TASKLIST = &HF130
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_VSCROLL Scroll the window vertically
    SC_VSCROLL = &HF070
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_MONITORPOWER Set the power saving state of the display
    SC_MONITORPOWER = &HF170
    ' ##ENUMERATION_MEMBER_DESCRIPTION SC_CONTEXTHELP Changes the cursor to a question mark with a pointer for context sensitive help
    SC_CONTEXTHELP = &HF180
End Enum


' ##ENUMERATION_DESCRIPTION ShowWindowCommands - Changes to the window show state
Public Enum ShowWindowCommands
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_HIDE Hide the window and activates another window
    SW_HIDE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_MAX Show the window maximised
    SW_MAX = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_MAXIMIZE Show the window maximised
    SW_MAXIMIZE = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_MINIMIZE Show the window minimised
    SW_MINIMIZE = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_NORMAL Show the window in normal size
    SW_NORMAL = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_RESTORE Activate and display the window
    SW_RESTORE = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOW Activate the window and display it in its current size and position
    SW_SHOW = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWDEFAULT Set the show state based on the default specified by the calling process
    SW_SHOWDEFAULT = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWMAXIMIZED Activate the window and display it as a maximized window
    SW_SHOWMAXIMIZED = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWMINIMIZED Activate the window and display it as a minimized window
    SW_SHOWMINIMIZED = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWMINNOACTIVE Display the window as a minimized window
    SW_SHOWMINNOACTIVE = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWNA Display the window in its current size and position without activating it
    SW_SHOWNA = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWNOACTIVATE Display the window in its most recent size and position without activating it
    SW_SHOWNOACTIVATE = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SHOWNORMAL Activate and display a window
    SW_SHOWNORMAL = 1
End Enum

' ##ENUMERATION_DESCRIPTION ShowWindowFlags - Describe the cause of a WM_SHOWWINDOW message
Public Enum ShowWindowFlags
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_OTHERUNZOOM The window is being uncovered because a maximize window was restored or minimized
    SW_OTHERUNZOOM = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_OTHERZOOM The window is being covered by another window that has been maximized
    SW_OTHERZOOM = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_PARENTCLOSING The window's owner window is being minimized
    SW_PARENTCLOSING = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_PARENTOPENING The window's owner window is being restored
    SW_PARENTOPENING = 3
End Enum

Public Enum enHWND
    HWND_BOTTOM = 1
    HWND_BROADCAST = &HFFFF&
    HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum



'\\ Module handles....
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long



Public Enum enHitTestResult
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTERROR = (-2)
    HTGROWBOX = 4
    HTHSCROLL = 6
    HTLEFT = 10
    HTMAXBUTTON = 9
    HTMENU = 5
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTRIGHT = 11
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTVSCROLL = 7
    HTTRANSPARENT = (-1)
    HTOBJECT = 19
    HTCLOSE = 20
    HTHELP = 21
End Enum


' ##EVENT_DESCRIPTION ApiError - Triggered when an error occurs in an API call
' ##EVENT_REMARKS Number - The error number
' ##EVENT_REMARKS Source - The class that experienced the error
' ##EVENT_REMARKS Description - The text explanation of the error
Public Event ApiError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Attribute ApiError.VB_Description = "ApiError - Triggered when an error occurs in an API call"
Attribute ApiError.VB_HelpID = 3167


' ##EVENT_DESCRIPTION StartUp - Triggered when an instance of the ApiFunctions class is created
Public Event StartUp()
Attribute StartUp.VB_Description = "StartUp - Triggered when an instance of the ApiFunctions class is created"
Attribute StartUp.VB_HelpID = 3168


Public Enum enPrinterActions
     PRINTACTION_OPEN = 0
     PRINTACTION_PROPERTIES = 1
     PRINTACTION_NETINSTALL = 2
     PRINTACTION_NETINSTALLLINK = 3
     PRINTACTION_TESTPAGE = 4
     PRINTACTION_OPENNETPRN = 5
     PRINTACTION_DOCUMENTDEFAULTS = 6
     PRINTACTION_SERVERPROPERTIES = 7
End Enum

'SHSTDAPI_(BOOL)  SHInvokePrinterCommandA(HWND hwnd, UINT uAction, LPCSTR lpBuf1, LPCSTR lpBuf2, BOOL fModal);
Private Declare Function SHInvokePrinterCommand Lib "shell32.dll" Alias "SHInvokePrinterCommandA" (ByVal hwnd As Long, ByVal uAction As enPrinterActions, ByVal Buffer1 As String, ByVal Buffer2 As String, ByVal Modal As Long) As Long

Private Declare Function InitCommonControlsExApi Lib "comctl32.dll" Alias "InitCommonControlsEx" (ByRef lpInitCtrls As INITCOMMONCONTROLSEX) As Long

Private Type INITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type

Public Enum enInitCommonControlClasses
    ICC_LISTVIEW_CLASSES = &H1       ' listview, header
    ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
    ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
    ICC_TAB_CLASSES = &H8            ' tab, tooltips
    ICC_UPDOWN_CLASS = &H10          ' updown
    ICC_PROGRESS_CLASS = &H20        ' progress
    ICC_HOTKEY_CLASS = &H40          ' hotkey
    ICC_ANIMATE_CLASS = &H80         ' animate
    ICC_WIN95_CLASSES = &HFF
    ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
    ICC_USEREX_CLASSES = &H200       ' comboex
    ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
    ICC_INTERNET_CLASSES = &H800
    ICC_PAGESCROLLER_CLASS = &H1000      ' page scroller
    ICC_NATIVEFNTCTL_CLASS = &H2000       ' native font control
End Enum



Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private bInternalRef As Boolean

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Sub About()
Attribute About.VB_HelpID = 2050

    MsgBox "(c) 2000-2002 Merrion Computing Ltd.", vbInformation, "EventVB " & App.Major & "." & App.Minor

End Sub


Public Property Get CurrentThreadId() As Long
Attribute CurrentThreadId.VB_HelpID = 2051

    CurrentThreadId = App.ThreadId
    
End Property


Public Sub DisplayDocumentDefaults(ByVal PrinterName As String, ByVal hwnd As Long)
Attribute DisplayDocumentDefaults.VB_HelpID = 2052

Dim lret As Long

#If PRODUCTION Then
    On Error Resume Next
#End If

lret = SHInvokePrinterCommand(hwnd, PRINTACTION_DOCUMENTDEFAULTS, PrinterName, "", 0)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":DisplayDocumentDefaults", GetLastSystemError
End If

End Sub


Public Sub EndLink()
Attribute EndLink.VB_HelpID = 2053

'\\ Kill all internally created objects
If Not (AllSubclassedWindows Is Nothing) Then
    Set AllSubclassedWindows = Nothing
End If
If Not (APIDispenser Is Nothing) Then
    Set APIDispenser = Nothing
End If
If Not (Eventhandler Is Nothing) Then
    Set Eventhandler = Nothing
End If
If Not (RegisteredHotkeys Is Nothing) Then
    Set RegisteredHotkeys = Nothing
End If

End Sub



'\\ --[GetBits]---------------------------------------------------------------------------------------
'\\ Returns the requested bits of a long
'\\ Parameters:
'\\   lInput - The long we're splitting up
'\\   nStart (0-31) - The start bit to get
'\\   nEnd (0-31) - The end bit to get
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function GetBits(ByVal lInput As Long, ByVal nStart As Integer, ByVal nEnd As Integer) As Long
Attribute GetBits.VB_HelpID = 2055

Dim lMask As Long
Dim lMaskPos As Long

If nEnd < nStart Then
    nEnd = nStart
End If

For lMaskPos = (nStart + 1) To (nEnd + 1)
    lMask = lMask + (lMaskPos ^ 2)
Next lMaskPos

GetBits = (lInput And lMask)

End Function



Public Sub InitCommonControls(ByVal ClassToInit As enInitCommonControlClasses)
Attribute InitCommonControls.VB_HelpID = 2057

Dim lret As Long
Dim iccx As INITCOMMONCONTROLSEX

#If PRODUCTION Then
    On Error Resume Next
#End If

With iccx
    .dwSize = Len(iccx)
    .dwICC = ClassToInit
End With

lret = InitCommonControlsExApi(iccx)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":InitCommonControls", LastSystemError
End If


End Sub

Friend Property Let InternalRef(ByVal ref As Boolean)
    bInternalRef = ref
End Property



Public Property Get ModuleHandle() As Long
Attribute ModuleHandle.VB_HelpID = 2060

#If PRODUCTION Then
    On Error Resume Next
#End If

    ModuleHandle = GetModuleHandle(App.exename)
    
End Property

Public Property Get ModuleInstance() As Long
Attribute ModuleInstance.VB_HelpID = 2061

    ModuleInstance = App.hInstance
    
End Property


Friend Sub RaiseError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

'\\ 2001-11-08
'\\ Do not raise zero error
If Number <> 0 Then
    RaiseEvent ApiError(Number, Source, Description)
End If

End Sub

Friend Sub RunProgram(ByVal exename As String)

Dim lret As Long

lret = ShellExecute(0, "Open", exename, "", "", SW_NORMAL)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":RunProgram", GetLastSystemError
End If

End Sub

Public Sub StartLink()
Attribute StartLink.VB_HelpID = 2065

#If PRODUCTION Then
    On Error Resume Next
#End If

Static bStarted As Boolean
If Not bStarted Then
    Set Eventhandler = New ApiSystemHook
    Set AllSubclassedWindows = New colSubclassedWindows
    Set APIDispenser = Me
    '\\ Initialise the collection classes for start..
    Set RegisteredHotkeys = New Collection
    bStarted = True
End If

End Sub

Public Property Get System() As ApiSystem
Attribute System.VB_HelpID = 2066

Dim mSystem As ApiSystem

Set mSystem = New ApiSystem

Set System = mSystem


End Property

Public Property Get SubclassedWindows() As colSubclassedWindows
Attribute SubclassedWindows.VB_HelpID = 2075

    Set SubclassedWindows = AllSubclassedWindows

End Property



'\\-[PrintDocument]-----------------------------------------------------
'\\ Prints the document using the app. which created it.
'\\ (Similar to selecting "Print" from explorer context menu)
'\\ --------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Sub PrintDocument(ByVal sDocumentName As String)
Attribute PrintDocument.VB_HelpID = 2076
Dim lret As Long

lret = ShellExecute(0, "Print", sDocumentName, "", "", SW_HIDE)

End Sub

'\\ -- [ LastSystemError ]------------------------------------------------------------------
'\\ Returns the message from the system which describes the last dll error to occur, as
'\\ held in Err.LastDllError.  This function should be called as soon after the API call
'\\ which might have errored, as this member can be reset to zero by subsequent API calls.
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function LastSystemError() As String
Attribute LastSystemError.VB_HelpID = 2077

LastSystemError = GetLastSystemError

End Function
Public Function AppModuleHandle() As Long
Attribute AppModuleHandle.VB_HelpID = 2078

    AppModuleHandle = GetModuleHandle(App.exename)

    If Err.LastDllError > 0 Then
        If Err.LastDllError = 126 Then
            '\\ Running in debug mode....module handle is visual basic
            AppModuleHandle = GetModuleHandle("vb5.exe")
        Else
            Call ReportError(Err.LastDllError, "AppModuleHandle", LastSystemError)
        End If
    End If
    
End Function

Private Sub Class_Initialize()

Call StartLink
'\\ Notify the class has started
RaiseEvent StartUp

End Sub


Private Sub Class_Terminate()

If bInternalRef Then
    Call EndLink
End If

End Sub

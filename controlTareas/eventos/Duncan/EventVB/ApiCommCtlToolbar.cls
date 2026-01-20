VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiCommCtlToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2771
Attribute VB_Description = "This class provides methods and properties for manipulating instances of the "
Option Explicit

' ##MODULE_DESCRIPTION This class provides methods and properties for manipulating _
instances of the "Toolbar" windows custom control

Private mWND As ApiWindow

Public Enum ToolbarButtonStates
    TBSTATE_CHECKED = &H1
    TBSTATE_PRESSED = &H2
    TBSTATE_ENABLED = &H4
    TBSTATE_HIDDEN = &H8
    TBSTATE_INDETERMINATE = &H10
    TBSTATE_WRAP = &H20
    TBSTATE_ELLIPSES = &H40
    TBSTATE_MARKED = &H80
End Enum

Public Enum ToolbarButtonStyles
    TBSTYLE_BUTTON = &H0
    TBSTYLE_SEP = &H1
    TBSTYLE_CHECK = &H2
    TBSTYLE_GROUP = &H4
    TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
    TBSTYLE_DROPDOWN = &H8
    TBSTYLE_AUTOSIZE = &H10    ' ##  If set, automatically calculate the width of the button
    TBSTYLE_NOPREFIX = &H20    ' ##  If set, this button should not have accelerator prefix for the &
    TBSTYLE_TOOLTIPS = &H100
    TBSTYLE_WRAPABLE = &H200
    TBSTYLE_ALTDRAG = &H400
    TBSTYLE_FLAT = &H800
    TBSTYLE_LIST = &H1000
    TBSTYLE_CUSTOMERASE = &H2000
    TBSTYLE_REGISTERDROP = &H4000
    TBSTYLE_TRANSPARENT = &H8000
End Enum


Private Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(2) As Byte
    dwData As Long
    iString As Long
End Type

Public Enum ToolbarMessages
    TB_ENABLEBUTTON = (&H400 + 1)
    TB_CHECKBUTTON = (&H400 + 2)
    TB_PRESSBUTTON = (&H400 + 3)
    TB_HIDEBUTTON = (&H400 + 4)
    TB_INDETERMINATE = (&H400 + 5)
    TB_MARKBUTTON = (&H400 + 6)
    TB_ISBUTTONENABLED = (&H400 + 9)
    TB_ISBUTTONCHECKED = (&H400 + 10)
    TB_ISBUTTONPRESSED = (&H400 + 11)
    TB_ISBUTTONHIDDEN = (&H400 + 12)
    TB_ISBUTTONINDETERMINATE = (&H400 + 13)
    TB_ISBUTTONHIGHLIGHTED = (&H400 + 14)
    TB_SETSTATE = (&H400 + 17)
    TB_GETSTATE = (&H400 + 18)
    TB_ADDBITMAP = (&H400 + 19)
    TB_ADDBUTTONSA = (&H400 + 20)
    TB_INSERTBUTTONA = (&H400 + 21)
    TB_ADDBUTTONS = (&H400 + 20)
    TB_INSERTBUTTON = (&H400 + 21)
    TB_DELETEBUTTON = (&H400 + 22)
    TB_GETBUTTON = (&H400 + 23)
    TB_BUTTONCOUNT = (&H400 + 24)
    TB_COMMANDTOINDEX = (&H400 + 25)
    TB_SAVERESTOREA = (&H400 + 26)
    TB_SAVERESTOREW = (&H400 + 76)
    TB_CUSTOMIZE = (&H400 + 27)
    TB_ADDSTRINGA = (&H400 + 28)
    TB_ADDSTRINGW = (&H400 + 77)
    TB_GETITEMRECT = (&H400 + 29)
    TB_BUTTONSTRUCTSIZE = (&H400 + 30)
    TB_SETBUTTONSIZE = (&H400 + 31)
    TB_SETBITMAPSIZE = (&H400 + 32)
    TB_AUTOSIZE = (&H400 + 33)
    TB_GETTOOLTIPS = (&H400 + 35)
    TB_SETTOOLTIPS = (&H400 + 36)
    TB_SETPARENT = (&H400 + 37)
    TB_SETROWS = (&H400 + 39)
    TB_GETROWS = (&H400 + 40)
    TB_GETBITMAPFLAGS = (&H400 + 41)
    TB_SETCMDID = (&H400 + 42)
    TB_CHANGEBITMAP = (&H400 + 43)
    TB_GETBITMAP = (&H400 + 44)
    TB_GETBUTTONTEXTA = (&H400 + 45)
    TB_GETBUTTONTEXTW = (&H400 + 75)
    TB_REPLACEBITMAP = (&H400 + 46)
    TB_SETINDENT = (&H400 + 47)
    TB_SETIMAGELIST = (&H400 + 48)
    TB_GETIMAGELIST = (&H400 + 49)
    TB_LOADIMAGES = (&H400 + 50)
    TB_GETRECT = (&H400 + 51)             ' ##  wParam is the Cmd instead of index
    TB_SETHOTIMAGELIST = (&H400 + 52)
    TB_GETHOTIMAGELIST = (&H400 + 53)
    TB_SETDISABLEDIMAGELIST = (&H400 + 54)
    TB_GETDISABLEDIMAGELIST = (&H400 + 55)
    TB_SETSTYLE = (&H400 + 56)
    TB_GETSTYLE = (&H400 + 57)
    TB_GETBUTTONSIZE = (&H400 + 58)
    TB_SETBUTTONWIDTH = (&H400 + 59)
    TB_SETMAXTEXTROWS = (&H400 + 60)
    TB_GETTEXTROWS = (&H400 + 61)
    TB_GETOBJECT = (&H400 + 62)            ' ##  wParam == IID, lParam void **ppv
    TB_GETHOTITEM = (&H400 + 71)
    TB_SETHOTITEM = (&H400 + 72)           ' ##  wParam == iHotItem
    TB_SETANCHORHIGHLIGHT = (&H400 + 73)   ' ##  wParam == TRUE/FALSE
    TB_GETANCHORHIGHLIGHT = (&H400 + 74)
    TB_MAPACCELERATORA = (&H400 + 78)     ' ##  wParam == ch, lParam int * pidBtn
    TB_GETINSERTMARK = (&H400 + 79)        ' ##  lParam == LPTBINSERTMARK
    TB_SETINSERTMARK = (&H400 + 80)        ' ##  lParam == LPTBINSERTMARK
    TB_INSERTMARKHITTEST = (&H400 + 81)    ' ##  wParam == LPPOINT lParam == LPTBINSERTMARK
    TB_MOVEBUTTON = (&H400 + 82)
    TB_GETMAXSIZE = (&H400 + 83)           ' ##  lParam == LPSIZE
    TB_SETEXTENDEDSTYLE = (&H400 + 84)     ' ##  For TBSTYLE_EX_*
    TB_GETEXTENDEDSTYLE = (&H400 + 85)     ' ##  For TBSTYLE_EX_*
    TB_GETPADDING = (&H400 + 86)
    TB_SETPADDING = (&H400 + 87)
    TB_SETINSERTMARKCOLOR = (&H400 + 88)
    TB_GETINSERTMARKCOLOR = (&H400 + 89)
    TB_MAPACCELERATORW = (&H400 + 90)
    TB_SETCOLORSCHEME = &H2002   ' ##  lParam is color scheme
    TB_GETCOLORSCHEME = &H2003 ' ##  fills in COLORSCHEME pointed to by lParam
    TB_SETUNICODEFORMAT = &H2005
    TB_GETUNICODEFORMAT = &H2006
End Enum

Public Enum StandardToolbarBitmaps
    STD_CUT = 0
    STD_COPY = 1
    STD_PASTE = 2
    STD_UNDO = 3
    STD_REDO = 4
    STD_DELETE = 5
    STD_FILENEW = 6
    STD_FILEOPEN = 7
    STD_FILESAVE = 8
    STD_PRINTPRE = 9
    STD_PROPERTIES = 10
    STD_HELP = 11
    STD_FIND = 12
    STD_REPLACE = 13
    STD_PRINT = 14
End Enum

Public Enum ViewToolbarBitmaps
    VIEW_LARGEICONS = 0
    VIEW_SMALLICONS = 1
    VIEW_LIST = 2
    VIEW_DETAILS = 3
    VIEW_SORTNAME = 4
    VIEW_SORTSIZE = 5
    VIEW_SORTDATE = 6
    VIEW_SORTTYPE = 7
    VIEW_PARENTFOLDER = 8
    VIEW_NETCONNECT = 9
    VIEW_NETDISCONNECT = 10
    VIEW_NEWFOLDER = 11
    VIEW_VIEWMENU = 12
End Enum

Public Enum HistoryToolbarBitmaps
    HIST_BACK = 0
    HIST_FORWARD = 1
    HIST_FAVORITES = 2
    HIST_ADDTOFAVORITES = 3
    HIST_VIEWTREE = 4
End Enum

Private Const TBN_FIRST = -700

Public Enum ToolbarNotificationMessages
     TBN_GETBUTTONINFOA = (TBN_FIRST - 0)
     TBN_GETBUTTONINFOW = (TBN_FIRST - 20)
     TBN_BEGINDRAG = (TBN_FIRST - 1)
     TBN_ENDDRAG = (TBN_FIRST - 2)
     TBN_BEGINADJUST = (TBN_FIRST - 3)
     TBN_ENDADJUST = (TBN_FIRST - 4)
     TBN_RESET = (TBN_FIRST - 5)
     TBN_QUERYINSERT = (TBN_FIRST - 6)
     TBN_QUERYDELETE = (TBN_FIRST - 7)
     TBN_TOOLBARCHANGE = (TBN_FIRST - 8)
     TBN_CUSTHELP = (TBN_FIRST - 9)
     TBN_DROPDOWN = (TBN_FIRST - 10)
     TBN_GETOBJECT = (TBN_FIRST - 12)
     TBN_HOTITEMCHANGE = (TBN_FIRST - 13)
     TBN_DRAGOUT = (TBN_FIRST - 14)            ' ##  this is sent when the user clicks down on a button then drags off the button
     TBN_DELETINGBUTTON = (TBN_FIRST - 15)     ' ##  uses TBNOTIFY
     TBN_GETDISPINFOA = (TBN_FIRST - 16)       ' ##  This is sent when the  toolbar needs  some display information
     TBN_GETDISPINFOW = (TBN_FIRST - 17)       ' ##  This is sent when the  toolbar needs  some display information
     TBN_GETINFOTIPA = (TBN_FIRST - 18)
     TBN_GETINFOTIPW = (TBN_FIRST - 19)
End Enum

Private bChanged As Boolean

Private Property Get m_hwnd() As Long

m_hwnd = mWND.hwnd

End Property
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get ClassName() As String
Attribute ClassName.VB_HelpID = 2772

    ClassName = "ToolbarWindow32"
    
End Property

Public Property Get Window() As ApiWindow
Attribute Window.VB_HelpID = 2828

    Set Window = mWND
    
End Property

Private Sub Class_Initialize()

'\\ make sure the common controls class is initialised
Call APIDispenser.InitCommonControls(ICC_COOL_CLASSES)

Dim lStyle As Long
Dim lExStyle As Long

If mWND Is Nothing Then
    Dim clsPager As New APIWndClass
    With clsPager
        .CreateFromClassname App.hInstance, ClassName
    End With
    lStyle = clsPager.Style
    Set mWND = APIDispenser.System.WindowDispenser.CreateWindow(lExStyle, clsPager, "", lStyle, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
End If

End Sub

Private Sub Class_Terminate()

Set mWND = Nothing

End Sub



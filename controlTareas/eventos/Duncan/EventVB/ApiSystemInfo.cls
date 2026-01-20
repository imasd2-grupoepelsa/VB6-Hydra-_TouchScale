VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2574
Attribute VB_Description = "ApiSystem - Top level class for all operations for the current system"
'\\ --[APISystem] - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'\\ Top level object that represents a computer system
'\\ (Functionally equivalent to the "Application" object in traditonal DOM
'\\ implementations...
'\\ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Const CLSNAME = "ApiSystem"

' ##MODULE_DESCRIPTION ApiSystem - Top level class for all operations for the current system


Public Enum ProcessorTypes
    PROCESSOR_INTEL_386 = 386
    PROCESSOR_INTEL_486 = 486
    PROCESSOR_INTEL_PENTIUM = 586
    PROCESSOR_INTEL_860 = 860
    PROCESSOR_MIPS_R2000 = 2000
    PROCESSOR_MIPS_R3000 = 3000
    PROCESSOR_MIPS_R4000 = 4000      ' ##  incl R4101 & R3910 for Windows CE
    PROCESSOR_ALPHA_21064 = 21064
    PROCESSOR_PPC_601 = 601
    PROCESSOR_PPC_603 = 603
    PROCESSOR_PPC_604 = 604
    PROCESSOR_PPC_620 = 620
    PROCESSOR_HITACHI_SH3 = 10003   ' ##  Windows CE
    PROCESSOR_HITACHI_SH3E = 10004   ' ##  Windows CE
    PROCESSOR_HITACHI_SH4 = 10005    ' ##  Windows CE
    PROCESSOR_MOTOROLA_821 = 821     ' ##  Windows CE
    PROCESSOR_SHx_SH3 = 103
    PROCESSOR_SHx_SH4 = 104
    PROCESSOR_STRONGARM = 2577
    PROCESSOR_ARM720 = 1824
    PROCESSOR_ARM820 = 2080
    PROCESSOR_ARM920 = 2336
    PROCESSOR_ARM_7TDMI = 70001      ' ##  Windows CE
End Enum

Public Enum ProcessorArchitectures
    PROCESSOR_ARCHITECTURE_INTEL = 0
    PROCESSOR_ARCHITECTURE_MIPS = 1
    PROCESSOR_ARCHITECTURE_ALPHA = 2
    PROCESSOR_ARCHITECTURE_PPC = 3
    PROCESSOR_ARCHITECTURE_SHX = 4
    PROCESSOR_ARCHITECTURE_ARM = 5
    PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF
End Enum

Public Enum ProcessorLevels
   '\\ Intel Processors
   PROC_INTEL_80386
   PROC_INTEL_80486
   PROC_INTEL_PENMTIUM
   PROC_INTEL_PENTIUM_PRO 'also pentium 2
   PROC_MIPS_8400
   PROC_ALPHA_21064
   PROC_ALPHA_21066
   PROC_ALPHA_21164
   PROC_PPC_601
   PROC_PPC_602
   PROC_PPC_603
   PROC_PPC_604
   PROC_PPC_603_PLUS
   PROC_PPC_604_PLUS
   PROC_PPC_620
   '\\ ANY NEW PROCESSOR REVISIONS WILL BE ADDED HERE...
End Enum


Private Type SYSTEM_INFO
   OEMId As Long
   PageSize As Long
   lpMinAddress As Long
   lpMaxAddress As Long
   ActiveProcessorMask As Long
   NumberOfProcessors As Long
   ProcessorType As Long
   MemoryAllocationGranularity As Long
   ProcessorLevel As Integer
   ProcessorRevision As Integer
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private mSysInfo As SYSTEM_INFO
                  
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVerInfo As OSVERSIONINFO) As Long
Private mVerInfo As OSVERSIONINFO

Public Enum WindowsOSPlatforms
     VER_PLATFORM_WIN32s = 0
     VER_PLATFORM_WIN32_WINDOWS = 1
     VER_PLATFORM_WIN32_NT = 2
End Enum

Public Enum WindowsOSVersions
    ver_Win_NT351 = 1
    ver_Win_95 = 2
    ver_Win_98 = 3
    ver_Win_Me = 4
    ver_win_nt4 = 5
    ver_Win_Win2000 = 6
    ver_win_Whistler = 7
End Enum


Public Enum ProcessorFeatures
     PF_FLOATING_POINT_PRECISION_ERRATA = 0
     PF_FLOATING_POINT_EMULATED = 1
     PF_COMPARE_EXCHANGE_DOUBLE = 2
     PF_MMX_INSTRUCTIONS_AVAILABLE = 3
     PF_PPC_MOVEMEM_64BIT_OK = 4
     PF_ALPHA_BYTE_INSTRUCTIONS = 5
End Enum
Private Declare Function IsProcessorFeaturePresent Lib "kernel32.dll" (ByVal ProcessorFeature As ProcessorFeatures) As Long

'\\ To get/set the local computer na..
Private Const MAX_COMPUTERNAME_LENGTH = 15
Private Declare Function GetComputername Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputername Lib "kernel32.dll" Alias "SetComputerNameA" (ByVal lpBuffer As String) As Long

Private Type LARGE_INTEGER
   HiInt As Long
   LoInt As Long
End Type

Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum SystemMetricsIndexes
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_SECURE = 44
    SM_CXEDGE = 45
    SM_CYEDGE = 46
    SM_CXMINSPACING = 47
    SM_CYMINSPACING = 48
    SM_CXSMICON = 49
    SM_CYSMICON = 50
    SM_CYSMCAPTION = 51
    SM_CXSMSIZE = 52
    SM_CYSMSIZE = 53
    SM_CXMENUSIZE = 54
    SM_CYMENUSIZE = 55
    SM_ARRANGE = 56
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
    SM_NETWORK = 63
    SM_CLEANBOOT = 67
    SM_CXDRAG = 68
    SM_CYDRAG = 69
    SM_SHOWSOUNDS = 70
    SM_CXMENUCHECK = 71           '/* Use instead of GetMenuCheckMarkDimensions()! */
    SM_CYMENUCHECK = 72
    SM_SLOWMACHINE = 73
    SM_MIDEASTENABLED = 74
    SM_MOUSEWHEELPRESENT = 75
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    SM_CMONITORS = 80
    SM_SAMEDISPLAYFORMAT = 81
End Enum

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Public Enum enSystemParametersInfo
    SPI_NOMESSAGE = 0 '\\ To cope with Win95/WinNT differences!
    SPI_GETBEEP = 1
    SPI_SETBEEP = 2
    SPI_GETMOUSE = 3
    SPI_SETMOUSE = 4
    SPI_GETBORDER = 5
    SPI_SETBORDER = 6
    SPI_GETKEYBOARDSPEED = 10
    SPI_SETKEYBOARDSPEED = 11
    SPI_LANGDRIVER = 12
    SPI_ICONHORIZONTALSPACING = 13
    SPI_GETSCREENSAVETIMEOUT = 14
    SPI_SETSCREENSAVETIMEOUT = 15
    SPI_GETSCREENSAVEACTIVE = 16
    SPI_SETSCREENSAVEACTIVE = 17
    SPI_GETGRIDGRANULARITY = 18
    SPI_SETGRIDGRANULARITY = 19
    SPI_SETDESKWALLPAPER = 20
    SPI_SETDESKPATTERN = 21
    SPI_GETKEYBOARDDELAY = 22
    SPI_SETKEYBOARDDELAY = 23
    SPI_ICONVERTICALSPACING = 24
    SPI_GETICONTITLEWRAP = 25
    SPI_SETICONTITLEWRAP = 26
    SPI_GETMENUDROPALIGNMENT = 27
    SPI_SETMENUDROPALIGNMENT = 28
    SPI_SETDOUBLECLKWIDTH = 29
    SPI_SETDOUBLECLKHEIGHT = 30
    SPI_GETICONTITLELOGFONT = 31
    SPI_SETDOUBLECLICKTIME = 32
    SPI_SETMOUSEBUTTONSWAP = 33
    SPI_SETICONTITLELOGFONT = 34
    SPI_GETFASTTASKSWITCH = 35
    SPI_SETFASTTASKSWITCH = 36
    SPI_SETDRAGFULLWINDOWS = 37
    SPI_GETDRAGFULLWINDOWS = 38
    SPI_GETNONCLIENTMETRICS = 41
    SPI_SETNONCLIENTMETRICS = 42
    SPI_GETMINIMIZEDMETRICS = 43
    SPI_SETMINIMIZEDMETRICS = 44
    SPI_GETICONMETRICS = 45
    SPI_SETICONMETRICS = 46
    SPI_SETWORKAREA = 47
    SPI_GETWORKAREA = 48
    SPI_SETPENWINDOWS = 49
    SPI_GETHIGHCONTRAST = 66
    SPI_SETHIGHCONTRAST = 67
    SPI_GETKEYBOARDPREF = 68
    SPI_SETKEYBOARDPREF = 69
    SPI_GETSCREENREADER = 70
    SPI_SETSCREENREADER = 71
    SPI_GETANIMATION = 72
    SPI_SETANIMATION = 73
    SPI_GETFONTSMOOTHING = 74
    SPI_SETFONTSMOOTHING = 75
    SPI_SETDRAGWIDTH = 76
    SPI_SETDRAGHEIGHT = 77
    SPI_SETHANDHELD = 78
    SPI_GETLOWPOWERTIMEOUT = 79
    SPI_GETPOWEROFFTIMEOUT = 80
    SPI_SETLOWPOWERTIMEOUT = 81
    SPI_SETPOWEROFFTIMEOUT = 82
    SPI_GETLOWPOWERACTIVE = 83
    SPI_GETPOWEROFFACTIVE = 84
    SPI_SETLOWPOWERACTIVE = 85
    SPI_SETPOWEROFFACTIVE = 86
    SPI_SETCURSORS = 87
    SPI_SETICONS = 88
    SPI_GETDEFAULTINPUTLANG = 89
    SPI_SETDEFAULTINPUTLANG = 90
    SPI_SETLANGTOGGLE = 91
    SPI_GETWINDOWSEXTENSION = 92
    SPI_SETMOUSETRAILS = 93
    SPI_GETMOUSETRAILS = 94
    SPI_SETSCREENSAVERRUNNING = 97
    SPI_GETFILTERKEYS = 50
    SPI_SETFILTERKEYS = 51
    SPI_GETTOGGLEKEYS = 52
    SPI_SETTOGGLEKEYS = 53
    SPI_GETMOUSEKEYS = 54
    SPI_SETMOUSEKEYS = 55
    SPI_GETSHOWSOUNDS = 56
    SPI_SETSHOWSOUNDS = 57
    SPI_GETSTICKYKEYS = 58
    SPI_SETSTICKYKEYS = 59
    SPI_GETACCESSTIMEOUT = 60
    SPI_SETACCESSTIMEOUT = 61
    SPI_GETSERIALKEYS = 62
    SPI_SETSERIALKEYS = 63
    SPI_GETSOUNDSENTRY = 64
    SPI_SETSOUNDSENTRY = 65
    SPI_GETMOUSEHOVERWIDTH = 98
    SPI_SETMOUSEHOVERWIDTH = 99
    SPI_GETMOUSEHOVERHEIGHT = 100
    SPI_SETMOUSEHOVERHEIGHT = 101
    SPI_GETMOUSEHOVERTIME = 102
    SPI_SETMOUSEHOVERTIME = 103
    SPI_GETWHEELSCROLLLINES = 104
    SPI_SETWHEELSCROLLLINES = 105
    SPI_GETSHOWIMEUI = 110
    SPI_SETSHOWIMEUI = 111
    SPI_GETMOUSESPEED = 112
    SPI_SETMOUSESPEED = 113
    SPI_GETSCREENSAVERRUNNING = 114
    SPI_GETACTIVEWINDOWTRACKING = &H1000
    SPI_SETACTIVEWINDOWTRACKING = &H1001
    SPI_GETMENUANIMATION = &H1002
    SPI_SETMENUANIMATION = &H1003
    SPI_GETCOMBOBOXANIMATION = &H1004
    SPI_SETCOMBOBOXANIMATION = &H1005
    SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
    SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
    SPI_GETGRADIENTCAPTIONS = &H1008
    SPI_SETGRADIENTCAPTIONS = &H1009
    SPI_GETMENUUNDERLINES = &H100A
    SPI_SETMENUUNDERLINES = &H100B
    SPI_GETACTIVEWNDTRKZORDER = &H100C
    SPI_SETACTIVEWNDTRKZORDER = &H100D
    SPI_GETHOTTRACKING = &H100E
    SPI_SETHOTTRACKING = &H100F
    SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
    SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
    SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
    SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
    SPI_GETFOREGROUNDFLASHCOUNT = &H2004
    SPI_SETFOREGROUNDFLASHCOUNT = &H2005
End Enum

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SystemParametersInfoByLong Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

'\\ Used to decide whether to update system wide/ini file...
'\\ If true, update ini and send a message
Public Enum SystemParametersUpdateFlags
    SPIF_NOUPDATE = &H0
    SPIF_UPDATEINIFILE = &H1
    SPIF_SENDWININICHANGE = &H2
End Enum
Public SystemWideUpdate As SystemParametersUpdateFlags
Attribute SystemWideUpdate.VB_VarHelpID = 2575

Private bChanged As Boolean

'\\ Used to get the Volumes collection
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'\\ Used to get the keyboard layouts...
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'\\ Used to get the desktop window....
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'\\ Used to enumerate the running processes...
Private Const MAX_PATH = 260
Private Type PROCESSENTRY32
    dwSize As Long '\\ Size of the PROCESSENTRY32 structure
    dwUsageCount As Long
    hProcessId As Long '\\ Unique process handle
    hDefaultHeap As Long
    hModuleId As Long
    dwThreadCount As Long
    hParentProcessId As Long
    pcBaseClassPriority As Long
    dwFlags As Long '\\ Reserved value
    szExeFilename As String * MAX_PATH
End Type
Private Declare Function CloseKernelObject Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal hProcess As Long) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

' ##ENUMERATION_DESCRIPTION Used to decide which information is included in the toolhelp snapshot
Public Enum ToolhelpSnapshotIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_SNAPHEAPLIST Snapshot heaps information
     TH32CS_SNAPHEAPLIST = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_SNAPPROCESS Snapshot process information
     TH32CS_SNAPPROCESS = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_SNAPTHREAD Snapshot thread information
     TH32CS_SNAPTHREAD = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_SNAPMODULE Snapshot module information
     TH32CS_SNAPMODULE = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_SNAPALL Snapshot all information
     TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
     ' ##ENUMERATION_MEMBER_DESCRIPTION TH32CS_INHERIT Snapshot handle is inheritable
     TH32CS_INHERIT = &H80000000
End Enum

'\\ NT4 (PSAPI.dll) version of getting running processes....
Private Declare Function EnumProcesses Lib "psapi.dll" (lpidProcesses As Long, ByVal cb As Long, cbNeeded As Long) As Long

' ##BLOCK_DESCRIPTION Provides access to global or local windows hooking
' ##BLOCK_REMARKS Excessive system hooking can degrade system performance, and should be used sparinglt
Public Property Get Hooks() As ApiSystemHook
Attribute Hooks.VB_Description = "Provides access to global or local windows hooking"
Attribute Hooks.VB_HelpID = 3097

    Set Hooks = Eventhandler
    
End Property
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get CurrentHardwareProfile() As ApiHardwareProfile
Attribute CurrentHardwareProfile.VB_HelpID = 2576

Dim pThis As ApiHardwareProfile

Set pThis = New ApiHardwareProfile

Set CurrentHardwareProfile = pThis

Set pThis = Nothing

End Property

' ##BLOCK_DESCRIPTION The %window:EventVB~ApiWindow% that is the active desktop
' ##BLOCK_REMARKS The desktop window is the parent of all application windows
Public Property Get DesktopWindow() As ApiWindow
Attribute DesktopWindow.VB_Description = "The window:EventVB~ApiWindow that is the active desktop"
Attribute DesktopWindow.VB_HelpID = 2881

Dim wndThis As New ApiWindow

wndThis.hwnd = GetDesktopWindow()
Set DesktopWindow = wndThis

End Property

' ##BLOCK_DESCRIPTION The %graphical device interface:EventVB~ApiGraphicalDeviceInterface% provides access to the various graphical objects
Public Property Get GraphicalDeviceInterface() As ApiGraphicalDeviceInterface
Attribute GraphicalDeviceInterface.VB_Description = "The graphical device interface:EventVB~ApiGraphicalDeviceInterface provides access to the various graphical objects"
Attribute GraphicalDeviceInterface.VB_HelpID = 2577

Dim gdi As ApiGraphicalDeviceInterface

Set gdi = New ApiGraphicalDeviceInterface

Set GraphicalDeviceInterface = gdi

End Property

' ##BLOCK_DESCRIPTION The collection of %locales:EventVB~ApiLocale% installed on the system
' ##BLOCK_REMARKS The locale affects regional settings, sort orders and keyboard settings
Public Property Get InstalledLocales() As Collection
Attribute InstalledLocales.VB_Description = "The collection of locales:EventVB~ApiLocale installed on the system"
Attribute InstalledLocales.VB_HelpID = 2578

Call APICallbackProcs.RefreshInstalledLocales
Set InstalledLocales = AllInstalledLocales

End Property

' ##BLOCK_DESCRIPTION The collection of keyboard layouts installed on the system
Public Property Get KeyboardLayouts() As Collection
Attribute KeyboardLayouts.VB_Description = "The collection of keyboard layouts installed on the system"
Attribute KeyboardLayouts.VB_HelpID = 2821

Dim lret As Long
Dim lpList() As Long
Dim colThis As New Collection

ReDim lpList(1) As Long
lret = GetKeyboardLayoutList(0, lpList(1))
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":KeyboardLayouts", GetLastSystemError
Else
    If lret > 0 Then
        ReDim Preserve lpList(1 To lret - 1) As Long
        lret = GetKeyboardLayoutList(UBound(lpList), lpList(1))
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":KeyboardLayouts", GetLastSystemError
        End If
    End If
End If

Set KeyboardLayouts = colThis

End Property

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an ampty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function
Public Property Let ComputerName(ByVal newname As String)
Attribute ComputerName.VB_HelpID = 2580

Dim lret As Long

If LenB(newname) > (MAX_COMPUTERNAME_LENGTH) Then
    ReportError 0, CLSNAME & ":ComputerName", "Computer length exceeds MAX_COMPUTERNAME_LEMGTH"
    Exit Property
End If

lret = SetComputername(newname)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, CLSNAME & ":ComputerName", GetLastSystemError
End If

End Property

Public Property Get ComputerName() As String

    Dim sTmp As String
    Dim lret As Long
    
    sTmp = String(MAX_COMPUTERNAME_LENGTH, 0)
    lret = GetComputername(sTmp, Len(sTmp))
    If Err.LastDllError <> 0 And Err.LastDllError <> 203 Then
        ReportError Err.LastDllError, CLSNAME & ":ComputerName", GetLastSystemError
    Else
        ComputerName = Left$(sTmp, InStr(sTmp, Chr$(0)) - 1)
    End If
    
End Property


'\\ --[IsRequiredSystem]- - - - - - - - - - - - - - - - - - - - - - - - - - - -
'\\ Returns TRUE if the OS is at or above the level required by the function,
'\\ or raises a meaningful error if it isn't
'\\ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - -
Friend Function IsRequiredSystem(ByVal AttemptedFunction As String, _
                                 ByVal RequiredVersion As WindowsOSVersions) As Boolean



    If (WindowsOSVersion >= RequiredVersion) Then
        IsRequiredSystem = True
    Else
        IsRequiredSystem = False
        ReportError 0, "Api Function - " & AttemptedFunction, "The function requires a different version of windows that you have installed"
    End If


End Function


Public Property Get Metrics(ByVal Index As SystemMetricsIndexes) As Long
Attribute Metrics.VB_HelpID = 2581

    Metrics = GetSystemMetrics(Index)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Metrics (Get)", GetLastSystemError
    End If
    
End Property

Public Property Get PerformanceCounter() As Currency
Attribute PerformanceCounter.VB_HelpID = 2582

Dim lret As Long
Dim lCounter As LARGE_INTEGER

lret = QueryPerformanceCounter(lCounter)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, CLSNAME & ":PerformanceCounter", GetLastSystemError
Else
    PerformanceCounter = LargeIntToCurrency(lCounter)
End If

End Property

Public Property Get PerformanceFrequency() As Currency
Attribute PerformanceFrequency.VB_HelpID = 2583

Dim lret As Long
Dim lFrequency As LARGE_INTEGER

lret = QueryPerformanceFrequency(lFrequency)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, CLSNAME & ":PerformanceFrequency", GetLastSystemError
Else
    PerformanceFrequency = LargeIntToCurrency(lFrequency)
End If

End Property

' ##BLOCK_DESCRIPTION The collection of %printers:EventVB~ApiPrinter% installed on the system
Public Property Get InstalledPrinters() As Collection
Attribute InstalledPrinters.VB_Description = "The collection of printers:EventVB~ApiPrinter installed on the system"
Attribute InstalledPrinters.VB_HelpID = 2584

Dim printerthis As ApiPrinter
Dim vbPrint As Printer
Dim colPrinters As Collection

If Printers.Count > 0 Then
    Set colPrinters = New Collection
End If

For Each vbPrint In Printers
    Set printerthis = New ApiPrinter
    printerthis.DeviceName = vbPrint.DeviceName
    colPrinters.Add printerthis, vbPrint.DeviceName
Next vbPrint

Set InstalledPrinters = colPrinters

End Property

Public Property Get Processes(ByVal DesiredAccess As ProcessAccessPriviledges) As Collection
Attribute Processes.VB_HelpID = 2882

Dim colProcesses As New Collection
Dim hSnapshot As Long
Dim lret As Long
Dim procHandles() As Long
Dim cbNeeded As Long
    
Dim peThis As PROCESSENTRY32
peThis.dwSize = Len(peThis)
Dim procThis As ApiProcess

'\\ If windows version is win2K or beyond we can use CreateToolhelp32Snapshot..
On Error GoTo WrongOSLevel

'\\ Open the process snapshot
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Processes", GetLastSystemError
Else
    '\\ If the snapshot handle opened OK...
    If hSnapshot Then
        lret = Process32First(hSnapshot, peThis)
        While lret
            If Err.LastDllError Then
                ReportError Err.LastDllError, CLSNAME & ":Processes", GetLastSystemError
            Else
                Set procThis = New ApiProcess
                With procThis
                    .DesiredAccess = DesiredAccess
                    .ProcessId = peThis.hProcessId
                End With
                colProcesses.Add procThis, procThis.Key
            End If
            lret = Process32Next(hSnapshot, peThis)
        Wend
    End If
End If


ProcessesEnd:
    Set Processes = colProcesses
    Exit Property

WrongOSLevel:
    On Error GoTo 0
    ReDim procHandles(0 To 64000) As Long
    lret = EnumProcesses(procHandles(0), UBound(procHandles) * Len(procHandles(0)), cbNeeded)
    If Err.LastDllError Then
        Call ReportError(Err.LastDllError, CLSNAME & ":Processes", GetLastSystemError)
    Else
        If cbNeeded > 0 Then
            ReDim Preserve procHandles(0 To cbNeeded / Len(procHandles(0))) As Long
            For lret = 0 To UBound(procHandles)
                If procHandles(lret) Then
                Set procThis = New ApiProcess
                With procThis
                    .DesiredAccess = DesiredAccess
                    .ProcessId = procHandles(lret)
                End With
                colProcesses.Add procThis, procThis.Key
                End If
            Next lret
        End If
    End If
    Resume ProcessesEnd
    
End Property

Public Property Get ProcessorAlphaByteInstructions() As Boolean
Attribute ProcessorAlphaByteInstructions.VB_HelpID = 2585

If IsRequiredSystem("Processor.AlphaByteInstructions", ver_win_nt4) Then
    ProcessorAlphaByteInstructions = IsProcessorFeaturePresent(PF_ALPHA_BYTE_INSTRUCTIONS)
End If

End Property

Public Property Get ProcessorArchitecture() As ProcessorArchitectures
Attribute ProcessorArchitecture.VB_HelpID = 2586

    'processorarchitecture = msysinfo.p
End Property

Public Property Get ProcessorCompareExchangeDouble() As Boolean
Attribute ProcessorCompareExchangeDouble.VB_HelpID = 2587

If IsRequiredSystem("Processor.CompareExchangeDouble", ver_win_nt4) Then
    ProcessorCompareExchangeDouble = IsProcessorFeaturePresent(PF_COMPARE_EXCHANGE_DOUBLE)
End If

End Property

Public Property Get ProcessorCount() As Long
Attribute ProcessorCount.VB_HelpID = 2588

    ProcessorCount = mSysInfo.NumberOfProcessors
    
End Property

Public Property Get ProcessorFloatingPointEmulation() As Boolean
Attribute ProcessorFloatingPointEmulation.VB_HelpID = 2589

If IsRequiredSystem("Processor.FloatingPointEmulation", ver_win_nt4) Then
    ProcessorFloatingPointEmulation = IsProcessorFeaturePresent(PF_FLOATING_POINT_EMULATED)
End If

End Property


Public Property Get ProcessorFloatingPointError() As Boolean
Attribute ProcessorFloatingPointError.VB_HelpID = 2590

If IsRequiredSystem("Processor.FloatingPointError", ver_win_nt4) Then
    ProcessorFloatingPointError = IsProcessorFeaturePresent(PF_FLOATING_POINT_PRECISION_ERRATA)
End If

End Property

Public Property Get ProcessorLevel() As ProcessorLevels
Attribute ProcessorLevel.VB_HelpID = 2591

Select Case True
Case ProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL
    '\\ Intel (and compatible) processor levels
    Select Case mSysInfo.ProcessorLevel
    Case 3
        ProcessorLevel = PROC_INTEL_80386
    Case 4
        ProcessorLevel = PROC_INTEL_80486
    Case 5
        ProcessorLevel = PROC_INTEL_PENMTIUM
    Case 6
        ProcessorLevel = PROC_INTEL_PENTIUM_PRO
    End Select
    
Case ProcessorArchitecture = PROCESSOR_ARCHITECTURE_ALPHA
    '\\ DEC Alpha chipsets
    Select Case mSysInfo.ProcessorLevel
    Case 21064
        ProcessorLevel = PROC_ALPHA_21064
    Case 21066
        ProcessorLevel = PROC_ALPHA_21066
    Case 21164
        ProcessorLevel = PROC_ALPHA_21164
    End Select

Case ProcessorArchitecture = PROCESSOR_ARCHITECTURE_PPC
    '\\ Power PC chipset architecture
    Select Case mSysInfo.ProcessorLevel
    Case 1  'PPC 601
        ProcessorLevel = PROC_PPC_601
    Case 3  'PPC 603
        ProcessorLevel = PROC_PPC_603
    Case 4  'PPC 604
        ProcessorLevel = PROC_PPC_604
    Case 6  'PPC 603+
        ProcessorLevel = PROC_PPC_603_PLUS
    Case 9  'PPC 604+
        ProcessorLevel = PROC_PPC_604_PLUS
    Case 20 'PPC 620
        ProcessorLevel = PROC_PPC_620
    End Select
    
Case ProcessorArchitecture = PROCESSOR_ARCHITECTURE_MIPS
    '\\ MIPS (or clone) chip
        ProcessorLevel = PROC_MIPS_8400
End Select

End Property

Public Property Get ProcessorMMXInstructions() As Boolean
Attribute ProcessorMMXInstructions.VB_HelpID = 2592

If IsRequiredSystem("Processor.MMXInstructions", ver_win_nt4) Then
    ProcessorMMXInstructions = IsProcessorFeaturePresent(PF_MMX_INSTRUCTIONS_AVAILABLE)
End If

End Property

Public Property Get ProcessorMoveMem64Bit() As Boolean
Attribute ProcessorMoveMem64Bit.VB_HelpID = 2593

If IsRequiredSystem("Processor.MoveMem64Bit", ver_win_nt4) Then
    ProcessorMoveMem64Bit = IsProcessorFeaturePresent(PF_PPC_MOVEMEM_64BIT_OK)
End If

End Property


Public Property Get ProcessorType() As ProcessorTypes
Attribute ProcessorType.VB_HelpID = 2594

    ProcessorType = mSysInfo.ProcessorType
    
End Property


Public Property Get TopLevelWindows() As Collection
Attribute TopLevelWindows.VB_HelpID = 2595

Call RefreshTopLevelWindows

Set TopLevelWindows = AllTopLevelWindows

End Property

'\\ API calls to populate the volume collection....
Public Property Get Volumes() As Collection
Attribute Volumes.VB_HelpID = 2596

Dim colThis As Collection
Dim volThis As ApiVolume
Set colThis = New Collection

Dim sLogicalDrives As String
Dim sThisDrive As String
Dim lBufflen As Long, lret As Long

lBufflen = GetLogicalDriveStrings(0, sLogicalDrives)
If lBufflen > 0 Then
    sLogicalDrives = String$(lBufflen, 0)
    lret = GetLogicalDriveStrings(lBufflen, sLogicalDrives)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Volumes", GetLastSystemError
    Else
        While Len(sLogicalDrives) > 0
            sThisDrive = Left$(sLogicalDrives, InStr(sLogicalDrives, Chr$(0)) - 1)
            If sThisDrive <> "" Then
                Set volThis = New ApiVolume
                volThis.RootName = sThisDrive
                colThis.Add volThis, volThis.RootName
                sLogicalDrives = Mid$(sLogicalDrives, Len(sThisDrive) + 2)
            Else
                sLogicalDrives = ""
            End If
        Wend
    End If
End If
Set Volumes = colThis


End Property

Public Property Get WindowDispenser() As ApiWindowDispenser
Attribute WindowDispenser.VB_HelpID = 2822

Dim wdThis As ApiWindowDispenser

Set wdThis = New ApiWindowDispenser
Set WindowDispenser = wdThis

End Property

Public Property Get WindowsOSBuild() As Long
Attribute WindowsOSBuild.VB_HelpID = 2597

    WindowsOSBuild = mVerInfo.dwBuildNumber
    
End Property

Public Property Get WindowsOSCSDVersion() As String
Attribute WindowsOSCSDVersion.VB_HelpID = 2598

WindowsOSCSDVersion = mVerInfo.szCSDVersion
End Property

Public Property Get WindowsOSPlatform() As WindowsOSPlatforms
Attribute WindowsOSPlatform.VB_HelpID = 2599

WindowsOSPlatform = mVerInfo.dwPlatformId

End Property

Public Property Get WindowsOSVersion() As WindowsOSVersions
Attribute WindowsOSVersion.VB_HelpID = 2600

With mVerInfo
    Select Case .dwMajorVersion
    Case 3
        WindowsOSVersion = ver_Win_NT351
    Case 4
        If .dwMinorVersion = 0 Then
            If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
                WindowsOSVersion = ver_Win_95
            Else
                WindowsOSVersion = ver_win_nt4
            End If
        ElseIf .dwMinorVersion = 10 Then
            WindowsOSVersion = ver_Win_98
        ElseIf .dwMinorVersion = 90 Then
            WindowsOSVersion = ver_Win_Me
        End If
    Case 5
        If .dwMinorVersion = 0 Then
            WindowsOSVersion = ver_Win_Win2000
        Else
            WindowsOSVersion = ver_win_Whistler
        End If
    End Select
End With

End Property

Public Property Get WindowStations() As colWinstations
Attribute WindowStations.VB_HelpID = 2883

Dim dwAccess As Long

dwAccess = WINSTA_FULLACCESS

Call APICallbackProcs.RefreshWinstations(dwAccess)
Set WindowStations = APICallbackProcs.AllWinstations

End Property

Private Sub Class_Initialize()

'\\ Get the current version information at startup...
Dim lret As Long
mVerInfo.dwOSVersionInfoSize = Len(mVerInfo)
lret = GetVersionEx(mVerInfo)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "ApiSystem", GetLastSystemError
End If

'\\ Get the system info at startup...
Call GetSystemInfo(mSysInfo)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "ApiSystem", GetLastSystemError
End If

End Sub




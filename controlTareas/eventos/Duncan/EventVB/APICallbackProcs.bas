Attribute VB_Name = "APICallbackProcs"
'\\ --[APICallbackProcs]---------------------------------------------------------------
'\\ Series of visual basic functions whose addresses can be passed as lpfnProcAddress
'\\ parameter of windows API callback functions using the AddressOf operator.
'\\ NOTE:
'\\ When creating a new callback proc, don't forget to declare the parameters ByVal,
'\\ or VB's type conversion will fail with GPF consequences
'\\ -----------------------------------------------------------------------------------

'typedef BOOL (CALLBACK* GRAYSTRINGPROC)(HDC, LPARAM, int);
'typedef VOID (CALLBACK* SENDASYNCPROC)(HWND, UINT, DWORD, LRESULT);
'typedef BOOL (CALLBACK* PROPENUMPROCA)(HWND, LPCSTR, HANDLE);
'typedef BOOL (CALLBACK* PROPENUMPROCEXA)(HWND, LPSTR, HANDLE, DWORD);
'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
'typedef BOOL (CALLBACK* NAMEENUMPROCA)(LPSTR, LPARAM);
'typedef BOOL (CALLBACK* ENUMRESTYPEPROC)(HMODULE hModule, LPTSTR lpType, LONG lParam);
'typedef BOOL (CALLBACK* ENUMRESNAMEPROC)(HMODULE hModule, LPCTSTR lpType, LPTSTR lpName, LONG lParam);
'typedef BOOL (CALLBACK* ENUMRESLANGPROC)(HMODULE hModule, LPCTSTR lpType, LPCTSTR lpName, WORD  wLanguage, LONG lParam);

Option Explicit

Private Const AUTHOR_NAME = "Duncan@merrioncomputing.com"

'\\ Application global variables....
Public Eventhandler As ApiSystemHook
Public APIDispenser As APIFunctions
Public AllSubclassedWindows As colSubclassedWindows
Public AllTopLevelWindows As Collection
Public AllInstalledLocales As Collection
Public RegisteredHotkeys As Collection
Public AllWinstations As colWinstations
Public Desktops As colDesktops
Public DesktopWindows As Collection

'\\ Windows hooks...
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

'\\ Enumerating windows....
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowApi Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long

'\\ Enumerating system locales
Private Enum SystemLocaleEnumerationTypes
     LCID_INSTALLED = &H1        ' ##  installed locale ids
     LCID_SUPPORTED = &H2        ' ##  supported locale ids
     LCID_ALTERNATE_SORTS = &H4        ' ##  alternate sort locale ids
End Enum
Private Declare Function EnumSystemLocales Lib "kernel32" Alias "EnumSystemLocalesA" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long

'\\ Speeding up error reporting
'\\ API Error decoding
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const BROADCAST_QUERY_DENY = &H424D5144

'\\ Hotkey and printer communication windows
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'\\ Enumerating window stations
Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Private mWindowBeingEnumerated As ApiWindow
Private Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long

Public Sub EnumerateProperties(ByVal wndEnum As ApiWindow)

Dim lret As Long

Set mWindowBeingEnumerated = wndEnum

lret = EnumProps(wndEnum.hWnd, AddressOf VB_ENUMPROP)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiWindow:EnumerateProperties", GetLastSystemError
End If


End Sub


Public Function GetLastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Dim sError As String * 500 '\\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long

lErrNum = Err.LastDllError

lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

GetLastSystemError = Trim(sError)
Debug.Print sError

End Function


Public Sub Main()

Set APIDispenser = New APIFunctions

End Sub


Public Sub RefreshDesktops(ByVal hWinsta As Long)

Set Desktops = New colDesktops
Dim lret As Long

lret = EnumDesktops(hWinsta, AddressOf VB_EnumDesktops, hWinsta)
If Err.LastDllError Then
    ReportError Err.LastDllError, "RefreshDesktops", GetLastSystemError
End If

End Sub

Public Sub RefreshDesktopWindows(ByVal hDesk As Long)

Set DesktopWindows = New Collection
Dim lret As Long
lret = EnumDesktopWindows(hDesk, AddressOf VB_EnumDesktopWindowsProc, hDesk)
If Err.LastDllError Then
    ReportError Err.LastDllError, "RefreshDesktopWindows", GetLastSystemError
End If

End Sub

Public Sub RefreshInstalledLocales()

Dim lret As Long

Set AllInstalledLocales = New Collection

lret = EnumSystemLocales(AddressOf VB_ENUMLOCALESPROC, LCID_INSTALLED)


End Sub

Public Sub RefreshTopLevelWindows()

Dim lret As Long

Set AllTopLevelWindows = New Collection

lret = EnumWindows(AddressOf TopLevelWndEnumProc, 0)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "System:TopLevelWindows", GetLastSystemError
End If

End Sub

Public Sub RefreshWinstations(ByVal DesiredAccess As Long)

Dim lret As Long

Set AllWinstations = New colWinstations

lret = EnumWindowStations(AddressOf VB_EnumWinstations, DesiredAccess)
If Err.LastDllError Then
    ReportError Err.LastDllError, "RefreshWinStations", GetLastSystemError
End If

End Sub

Public Sub ReportError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

If APIDispenser Is Nothing Then
    Err.Raise Number, Source, Description
Else
    APIDispenser.RaiseError Number, Source, Description
End If

Err.Clear

End Sub
'\\ --[VB_DLGPROC]----------------------------------------------------------------------------
'\\ typedef BOOL (CALLBACK* DLGPROC)(HWND, UINT, WPARAM, LPARAM)
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_DLGPROC(ByVal hWnd As Long, ByVal uint As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 5) As Variant
Params(1) = hWnd
Params(2) = uint
Params(3) = wParam
Params(4) = lParam
Params(5) = 0

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent DLGPROC, Params()
End If

VB_DLGPROC = Params(5)

End Function

'\\ --[VB_EDITWORDBREAKPROCA]------------------------------------------------------------
'\\ 'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
'\\ This gets called by an edit control when a line of text has filled up the available
'\\ space.
'\\ By default, a text edit box breaks on spaces.
'\\ (This version prevents numbers being broken up if the digit grouping sepeartor is a space.)
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EDITWORDBREAKPROCA(ByVal lpch As Long, _
                                      ByVal ichCurrent As Long, _
                                      ByVal cch As Long, _
                                      ByVal Code As Long) As Long
                                      
On Local Error Resume Next

Dim sCharacters As String

Dim lCharPos As Long

sCharacters = StringFromPointer(lpch, 1024)

Select Case Code
Case WB_ISDELIMITER
    '\\ Edit control is asking if this character is a wordbreak char...
    '\\ Reply FALSE is it is not a space, or if the characters either side of it
    '\\ are numbers....
    If Mid$(sCharacters, ichCurrent, 1) = " " Then
        VB_EDITWORDBREAKPROCA = 1
        If (ichCurrent > 0) And (ichCurrent < Len(sCharacters)) Then
            If IsNumeric(Mid$(sCharacters, ichCurrent - 1, 1)) And IsNumeric(Mid$(sCharacters, ichCurrent + 1, 1)) Then
                VB_EDITWORDBREAKPROCA = 0
            End If
        End If
    Else
        VB_EDITWORDBREAKPROCA = 0
    End If

Case WB_LEFT
  '\\ Find the begining of a word to the left of this position....
  For lCharPos = ichCurrent To 1 Step -1
    If Mid$(sCharacters, lCharPos, 1) = " " Then
        If Not (IsNumeric(Mid$(sCharacters, lCharPos - 1, 1)) And IsNumeric(Mid$(sCharacters, lCharPos + 1, 1))) Then
            VB_EDITWORDBREAKPROCA = lCharPos
            Exit For
        End If
    End If
  Next lCharPos
  
Case WB_RIGHT
'\\ Find the begining of a word to the right of this position....
  For lCharPos = ichCurrent To Len(sCharacters)
    If Mid$(sCharacters, lCharPos, 1) = " " Then
        If Not (IsNumeric(Mid$(sCharacters, lCharPos - 1, 1)) And IsNumeric(Mid$(sCharacters, lCharPos + 1, 1))) Then
            VB_EDITWORDBREAKPROCA = lCharPos
            Exit For
        End If
    End If
  Next lCharPos
End Select

End Function


'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EnumDesktops(ByVal lpstrName As Long, ByVal lParam As Long) As Long

Dim sName As String
sName = StringFromPointer(lpstrName, 1024)

Dim dtThis As ApiDesktop
Set dtThis = New ApiDesktop

dtThis.WinstationHandle = lParam
dtThis.Name = sName

Desktops.AddItem dtThis

VB_EnumDesktops = 1

End Function

'\\ --[VB_ENUMLOCALESPROC]-----------------------------------------------------------
'\\ BOOL CALLBACK EnumLocalesProc(
'\\   LPTSTR lpLocaleString   // locale identifier string
'\\   };
'\\ ---------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing Ltd.  All rights reserved
'\\ ---------------------------------------------------------------------------------
Public Function VB_ENUMLOCALESPROC(ByVal lpLocales As Long) As Long

Dim sLocale As String
Dim LCid As Long

Dim lcThis As New ApiLocale

sLocale = StringFromPointer(lpLocales, 9)

LCid = Val("&H" & sLocale)
lcThis.LocaleId = LCid
AllInstalledLocales.Add lcThis, "&H" & sLocale

VB_ENUMLOCALESPROC = 1

End Function

'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMPROP(ByVal hWnd As Long, ByVal lpStrPropName As String, ByVal hHandle As Long) As Long


mWindowBeingEnumerated.AddProperty lpStrPropName

VB_ENUMPROP = True

End Function



'\\ --[VB_ENUMRESLANGPROC]---------------------------------------------
'\\ Decl:
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESLANGPROC(ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 5) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lpName
Params(4) = wLanguage
Params(5) = lParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESLANGPROC, Params()
End If

End Function

'\\ --[VB_ENUMRESNAMEPROC]------------------------------------------------------------
'\\ Decl:
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESNAMEPROC(ByVal hModule As Long, ByVal lpType As String, _
                                    ByVal lpName As String, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 4) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lpName
Params(4) = lParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESNAMEPROC, Params()
End If

End Function

'\\ --[VB_ENUMRESTYPEPROC]----------------------------------------------
'\\ Enumerates the resource types in a module
'\\ Decl:
'\\ BOOL (CALLBACK* ENUMRESTYPEPROC)(HMODULE hModule, LPTSTR lpType, LONG lParam);
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESTYPEPROC(ByVal hModule As Long, ByVal lpType As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 3) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lParam

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESTYPEPROC, Params()
End If

End Function

'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EnumWinstations(ByVal lpstrName As Long, ByVal lParam As Long) As Long

Dim sName As String
sName = StringFromPointer(lpstrName, 1024)

Dim wsThis As ApiWinstation
Set wsThis = New ApiWinstation

wsThis.DesiredAccess = lParam
wsThis.Name = sName
AllWinstations.AddItem wsThis

VB_EnumWinstations = 1

End Function

'\\ -[VB_TimerProc]------------------------------------------------------
'\\ 'typedef VOID (CALLBACK* TIMERPROC)(HWND, UINT, UINT, DWORD);
'\\ parameters:
'\\ hWnd - The window handle to which the timer is attached...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Sub VB_TIMERPROC(ByVal hWnd As Long, _
                        ByVal uint1 As Long, _
                        ByVal nEventId As Long, _
                        ByVal dwParam As Long)

On Error Resume Next

Dim Params() As Variant

ReDim Params(1 To 4) As Variant
Params(1) = hWnd
Params(2) = uint1
Params(3) = nEventId
Params(4) = dwParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent TIMERPROC, Params()
End If

End Sub

'\\ --[VB_WindowProc]-------------------------------------------------------------------
'\\ 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
'\\ Parameters:
'\\   hwnd - window handle receiving message
'\\   wMsg - The window message (WM_..etc.)
'\\   wParam - First message parameter
'\\   lParam - Second message parameter
'\\ Note:
'\\    When subclassing a window proc using this, set the eventhandler's
'\\    hOldWndProc property to the window's previous window proc address.
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next
Dim lret As Long
Dim Params() As Variant
ReDim Params(1 To 5) As Variant

Params(1) = hWnd
Params(2) = wMsg
Params(3) = wParam
Params(4) = lParam
Params(5) = lret
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent WNDPROC, Params()
End If
lret = Params(5)

VB_WindowProc = lret

End Function

'\\ [VB_WndEnumProc]---------------------------------------------------------------------------
'\\ 'typedef BOOL (CALLBACK* WNDENUMPROC)(HWND, LPARAM);
'\\ Used in EnumWindows and EnumChildWindows
'\\ hwnd - Window handle of the enumerated window,
'\\ lparam - passed into the enumwindows proc by programmer...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_WndEnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

'\\ 1 - Pack the param array.....
ReDim Params(1 To 2) As Variant
Params(1) = hWnd
Params(2) = lParam

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent WNDENUMPROC, Params()
End If

VB_WndEnumProc = 1

End Function

Public Function VB_EnumDesktopWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim wndThis As New ApiWindow
wndThis.hWnd = hWnd
DesktopWindows.Add wndThis

VB_EnumDesktopWindowsProc = 1

End Function
Private Function TopLevelWndEnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long


Dim Window As ApiWindow

If IsWindowApi(hWnd) Then
    '\\ If its already subclassed, return that instance...
    On Error Resume Next
    Set Window = AllSubclassedWindows.Item(hWnd)
    On Error GoTo 0
    If Window Is Nothing Then
        Set Window = New ApiWindow
        Window.hWnd = hWnd
    End If
    AllTopLevelWindows.Add Window, "hwnd:" & Window.hWnd
    TopLevelWndEnumProc = True
Else
    '\\ A bad hwnd has been returned - do not continue
    TopLevelWndEnumProc = False
End If


End Function

'\\ [VB_HOOKCALLWNDPROC]----------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ NOTE:
'\\  This code has been kept for backwards compatibility only.
'\\  You should use the specific CBTHookProc, ShellHookProc etc...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_HOOKCALLWNDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKCALLWNDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_CALLWNDPROC), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_CALLWNDPROC, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_CALLWNDPROC), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKCALLWNDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKCALLWNDPROC = lret
Else
    VB_HOOKCALLWNDPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKCBTPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKCBTPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKCBTPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_CBT), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_CBT, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_CBT), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKCBTPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKCBTPROC = lret
Else
    VB_HOOKCBTPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKDEBUGPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKDEBUGPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKDEBUGPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_DEBUG), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_DEBUG, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_DEBUG), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKDEBUGPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKDEBUGPROC = lret
Else
    VB_HOOKDEBUGPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKFOREGROUNDIDLEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKFOREGROUNDIDLEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKFOREGROUNDIDLEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_FOREGROUNDIDLE), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_FOREGROUNDIDLE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_FOREGROUNDIDLE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKFOREGROUNDIDLEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKFOREGROUNDIDLEPROC = lret
Else
    VB_HOOKFOREGROUNDIDLEPROC = lMsgRet
End If

End Function


'\\ [VB_HOOKGETMESSAGEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKGETMESSAGEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKGETMESSAGEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_GETMESSAGE), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_GETMESSAGE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_GETMESSAGE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKGETMESSAGEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKGETMESSAGEPROC = lret
Else
    VB_HOOKGETMESSAGEPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKJOURNALPLAYBACKPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKJOURNALPLAYBACKPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKJOURNALPLAYBACKPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALPLAYBACK), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_JOURNALPLAYBACK, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALPLAYBACK), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKJOURNALPLAYBACKPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKJOURNALPLAYBACKPROC = lret
Else
    VB_HOOKJOURNALPLAYBACKPROC = lMsgRet
End If


End Function

'\\ [VB_HOOKJOURNALRECORDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKJOURNALRECORDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKJOURNALRECORDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALRECORD), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_JOURNALRECORD, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALRECORD), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKJOURNALRECORDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKJOURNALRECORDPROC = lret
Else
    VB_HOOKJOURNALRECORDPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKMOUSEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKMOUSEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKMOUSEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MOUSE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKMOUSEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKMOUSEPROC = lret
Else
    VB_HOOKMOUSEPROC = lMsgRet
End If

End Function


'\\ [VB_HOOKLOWLEVELMOUSEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKLOWLEVELMOUSEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKLOWLEVELMOUSEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE_LL), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MOUSE_LL, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE_LL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKLOWLEVELMOUSEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKLOWLEVELMOUSEPROC = lret
Else
    VB_HOOKLOWLEVELMOUSEPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKMESSAGEFILTERPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKMESSAGEFILTERPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKMESSAGEFILTERPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MSGFILTER), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MESSAGEFILTER, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MSGFILTER), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKMESSAGEFILTERPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKMESSAGEFILTERPROC = lret
Else
    VB_HOOKMESSAGEFILTERPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKSHELLPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKSHELLPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKSHELLPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_SHELL), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_SHELL, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_SHELL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKSHELLPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKSHELLPROC = lret
Else
    VB_HOOKSHELLPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKSYSMESSAGEFILTERPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKSYSMESSAGEFILTERPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKSYSMESSAGEFILTERPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_SYSMSGFILTER), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_SYSMESSAGEFILTER, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_SYSMSGFILTER), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKSYSMESSAGEFILTERPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKSYSMESSAGEFILTERPROC = lret
Else
    VB_HOOKSYSMESSAGEFILTERPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKLOWLEVELKEYBOARDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKLOWLEVELKEYBOARDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKLOWLEVELKEYBOARDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD_LL), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam 'virtual key (WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP)
Params(3) = lParam 'pointer to a KBDLLHOOKSTRUCT object
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_KEYBOARD_LL, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD_LL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKLOWLEVELKEYBOARDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKLOWLEVELKEYBOARDPROC = lret
Else
    VB_HOOKLOWLEVELKEYBOARDPROC = lMsgRet
End If

End Function

'\\ [VB_HOOKKEYBOARDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKKEYBOARDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKKEYBOARDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_KEYBOARD, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKKEYBOARDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKKEYBOARDPROC = lret
Else
    VB_HOOKKEYBOARDPROC = lMsgRet
End If

End Function

'\\ [VB_HookProc]----------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ NOTE:
'\\  This code has been kept for backwards compatibility only.
'\\  You should use the specific CBTHookProc, ShellHookProc etc...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_HOOKPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKPROC = CallNextHookEx(Eventhandler.HookIdByType(Eventhandler.CurrentHookType), Code, wParam, lParam)
    Exit Function
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    If Eventhandler.CurrentHookType = WH_MOUSE Then
        Eventhandler.TriggerEvent HOOKPROC_MOUSE, Params()
    ElseIf Eventhandler.CurrentHookType = WH_MOUSE_LL Then
        Eventhandler.TriggerEvent HOOKPROC_MOUSE_LL, Params()
    Else
        Eventhandler.TriggerEvent HOOKPROC, Params()
    End If
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(Eventhandler.CurrentHookType), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKPROC", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If lMsgRet = 0 Then
    '\\ Return value to calling code....
    VB_HOOKPROC = lret
Else
    VB_HOOKPROC = lMsgRet
End If

End Function

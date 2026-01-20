VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIWndClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2079
Option Explicit

Private Const CLSNAME = "ApiWndClass"

Private Type WndClass
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long '\\ String conversion after api calls
    lpszClassName As Long '\\ String conversion afer api calls
End Type

Private Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
    hIconSm As Long
End Type

'\\ Window CLASS information....
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WndClass) As Long

'\\ Get Class Long Index constants....
' ##ENUMERATION_DESCRIPTION ClassLongIndexes - The offsets for the different class information types
Public Enum ClassLongIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_MENUNAME Get or Set the address of the menu name string. The string identifies the menu resource associated with the class
    GCL_MENUNAME = (-8)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_HBRBACKGROUND Get or Set a handle to the background brush associated with the class
    GCL_HBRBACKGROUND = (-10)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_HCURSOR Get or Set a handle to the cursor associated with the class
    GCL_HCURSOR = (-12)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_HICON Get or Set a handle to the icon associated with the class
    GCL_HICON = (-14)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_HMODULE Get or Set a handle to the module that registered the class
    GCL_HMODULE = (-16)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_CBWNDEXTRA Get or Set Sets the size, in bytes, of the extra window memory associated with each window in the class
    GCL_CBWNDEXTRA = (-18)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_CBCLSEXTRA Get or Set the size, in bytes, of the extra memory associated with the class
    GCL_CBCLSEXTRA = (-20)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_WNDPROC Get or Set the window procedure address
    GCL_WNDPROC = (-24)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_STYLE Get or Set the class style bits
    GCL_STYLE = (-26)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCW_ATOM
    GCW_ATOM = (-32)
    ' ##ENUMERATION_MEMBER_DESCRIPTION GCL_HICONSM Get or Set a handle to the small icon associated with the class
    GCL_HICONSM = (-34)
End Enum

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As ClassLongIndexes) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As ClassLongIndexes, ByVal dwNewLong As Long) As Long


'\\ Class style constants....
Public Enum ClassStyleConstants
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_BYTEALIGNCLIENT the client area of windows belonging to the class are aligned on a byte boundary
    CS_BYTEALIGNCLIENT = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_BYTEALIGNWINDOW the windows belonging to the class are aligned on a byte boundary
    CS_BYTEALIGNWINDOW = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_CLASSDC all the windows in that class share the same device context
    CS_CLASSDC = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_DBLCLKS causes Windows to detect a double-click on any window created with this class
    CS_DBLCLKS = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_HREDRAW causes a window to be completely redrawn whenever its horizontal size changes
    CS_HREDRAW = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_INSERTCHAR
    CS_INSERTCHAR = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_KEYCVTWINDOW
    CS_KEYCVTWINDOW = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_NOCLOSE does not have the Close command in its System menu
    CS_NOCLOSE = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_NOKEYCVT
    CS_NOKEYCVT = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_NOMOVECARET
    CS_NOMOVECARET = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_OWNDC each window in that class has its own private device context
    CS_OWNDC = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_PARENTDC each window in that class uses its parent window device context
    CS_PARENTDC = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_PUBLICCLASS
    CS_PUBLICCLASS = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_SAVEBITS
    CS_SAVEBITS = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION CS_VREDRAW causes a window to be completely redrawn whenever its vertical size changes
    CS_VREDRAW = &H1
End Enum

'\\ Creation successful?
Public CreatedOK As Boolean

'\\ Members
Public Style As Long
Public cbClsExtra As Long
Public cbWndExtra2 As Long
Public hInstance As Long
'Public hIcon As Long
Public hCursor As Long
Public hbrBackground As Long


Private mlpszMenuName As String
Private mlpszClassName As String

Private m_lpfnwndproc As Long
Private mIcon As ApiIcon

'\\ Private APIs
Private Declare Function RegisterClassExApi Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClassApi Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long


'\\ Private memory handling functions
Private Declare Sub CopyMemoryWndClass Lib "kernel32" Alias "RtlMoveMemory" (Destination As WndClass, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrWndclass Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrWndclass Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long


Public Enum enCommonControlClasses
    CCC_BUTTONSLISTBOX = 1
    CCC_HOTKEY_CLASS = 2
    CCC_PROGRESS_CLASS = 3
    CCC_STATUSCLASSNAME = 4
    CCC_RICHEDIT = 5
    CCC_TOOLBARCLASSNAME = 6
    CCC_TOOLTIPS_CLASS = 7
    CCC_TRACKBARCLASS = 8
    CCC_UPDOWN_CLASS = 9
    CCC_WC_HEADER = 10
    CCC_WC_LISTVIEW = 11
    CCC_WC_TABCONTROL = 12
    CCC_WC_TREEVIEW = 13
End Enum

Private bChanged As Boolean

Attribute CreatedOK.VB_VarHelpID = 2080
Attribute Style.VB_VarHelpID = 2081
Attribute cbClsExtra.VB_VarHelpID = 2082
Attribute cbWndExtra2.VB_VarHelpID = 2083
Attribute hInstance.VB_VarHelpID = 2084
Attribute hCursor.VB_VarHelpID = 2085
Attribute hbrBackground.VB_VarHelpID = 2086
Attribute mlpszMenuName.VB_VarHelpID = 2087
Attribute mlpszClassName.VB_VarHelpID = 2088
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Private Function CommonControlClassname(ByVal ClassType As enCommonControlClasses) As String

Select Case ClassType
Case CCC_BUTTONSLISTBOX
    CommonControlClassname = "BUTTONSLISTBOX"
Case CCC_HOTKEY_CLASS
    CommonControlClassname = "HOTKEY_CLASS"
Case CCC_PROGRESS_CLASS
    CommonControlClassname = "PROGRESS_CLASS"
Case CCC_STATUSCLASSNAME
    CommonControlClassname = "STATUSCLASSNAME"
Case CCC_RICHEDIT
    CommonControlClassname = "RICHEDIT"
Case CCC_TOOLBARCLASSNAME
    CommonControlClassname = "TOOLBARCLASSNAME"
Case CCC_TOOLTIPS_CLASS
    CommonControlClassname = "TOOLTIPS_CLASS"
Case CCC_TRACKBARCLASS
    CommonControlClassname = "TRACKBARCLASS"
Case CCC_UPDOWN_CLASS
    CommonControlClassname = "UPDOWN_CLASS"
Case CCC_WC_HEADER
    CommonControlClassname = "WC_HEADER"
Case CCC_WC_LISTVIEW
    CommonControlClassname = "WC_LISTVIEW"
Case CCC_WC_TABCONTROL
    CommonControlClassname = "WC_TABCONTROL"
Case CCC_WC_TREEVIEW
    CommonControlClassname = "WC_TREEVIEW"
End Select

End Function


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this WNDCLASS object from the class name identified to by
'\\ the instance and class name combination
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromClassname(ByVal hInst As Long, ByVal lpClassName As String) As Boolean
Attribute CreateFromClassname.VB_HelpID = 2089

Dim lpClass As WndClass
Dim lret As Long

If hInst > 0 Then
    lret = GetClassInfo(hInst, lpClassName, lpClass)
Else
    lret = GetClassInfo(ByVal 0&, lpClassName, lpClass)
End If
If Err.LastDllError = 0 And lret > 0 Then
    CreateFromClassname = CreateFromPointer(VarPtr(lpClass))
ElseIf Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CreateFromClassname", GetLastSystemError
End If

End Function

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this WNDCLASS object from the location poiunted to by
'\\ lpWndClass
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpWndClass As Long) As Boolean

Dim wcThis As WndClass

CreatedOK = False

If Not IsBadReadPtrWndclass(lpWndClass, Len(wcThis)) Then
    Call CopyMemoryWndClass(wcThis, lpWndClass, Len(wcThis))
    If Err.LastDllError = 0 Then
        With wcThis
            Style = .Style
            lpfnWndProc = .lpfnWndProc
            cbClsExtra = .cbClsExtra
            cbWndExtra2 = .cbWndExtra2
            hInstance = .hInstance
            Set Icon = New ApiIcon
            Icon.hIcon = .hIcon
            hCursor = .hCursor
            hbrBackground = .hbrBackground
            If .lpszClassName > &HC00 Then
                mlpszClassName = StringFromPointer(.lpszClassName, 1024)
            End If
            If .lpszMenuName > 0 Then
                mlpszMenuName = StringFromPointer(.lpszMenuName, 1024)
            End If
            If Err.LastDllError = 0 Then
                CreatedOK = True
            Else
                ReportError Err.LastDllError, CLSNAME & ":CreateFromPointer", GetLastSystemError
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Property Set Icon(ByVal newIcon As ApiIcon)
Attribute Icon.VB_HelpID = 2091

If newIcon Is Nothing Then
    Set mIcon = Nothing
Else
    If newIcon.hIcon <> Icon.hIcon Then
        Set mIcon = newIcon
    End If
End If

End Property

Public Property Get Icon() As ApiIcon

    If mIcon Is Nothing Then
        Set mIcon = New ApiIcon
    End If
    
    Set Icon = mIcon
    
End Property

Public Property Let lpfnWndProc(ByVal newLPfnwndproc As Long)
Attribute lpfnWndProc.VB_HelpID = 2093

    m_lpfnwndproc = newLPfnwndproc
    
End Property

Public Property Get lpfnWndProc() As Long

    lpfnWndProc = m_lpfnwndproc
    
End Property


Public Property Let Name(ByVal newval As String)
Attribute Name.VB_HelpID = 3170

mlpszClassName = newval

End Property

Public Property Get Name() As String

    Name = mlpszClassName

End Property

Public Function RegisterClassEx() As Boolean
Attribute RegisterClassEx.VB_HelpID = 2094

Dim lret As Long

Dim wndclassThis As WNDCLASSEX

With wndclassThis
    .cbClsExtra = cbClsExtra
    .cbWndExtra = cbWndExtra2
    .cbSize = LenB(wndclassThis)
    .hbrBackground = hbrBackground
    .hCursor = hCursor
    .hInstance = hInstance
    .hIcon = Icon.hIcon
    '.hIconSm = IconSmall.hIcon
    .lpfnWndProc = lpfnWndProc
    If mlpszClassName <> "" Then
        .lpszClassName = StrPtr(mlpszClassName)
    End If
    If mlpszMenuName <> "" Then
        .lpszMenuName = StrPtr(mlpszMenuName)
    End If
    .Style = Style
End With

lret = RegisterClassExApi(wndclassThis)
If Err.LastDllError = 0 Then
    RegisterClassEx = True
Else
    Call ReportError(Err.LastDllError, "ApiWndClass:RegisterClassEx", GetLastSystemError)
End If

End Function

Public Sub SetDefaultProcAddress()
Attribute SetDefaultProcAddress.VB_HelpID = 2095

   Call SetProcAddress(AddressOf VB_WindowProc)
    
End Sub

Private Sub SetProcAddress(ByVal lpAddress As Long)

    lpfnWndProc = lpAddress
    
End Sub


'\\ --[StandardClassName]-------------------------------------------
'\\ Converts between the exported StandarWindowClass type
'\\ and the actual string class name used by windows
'\\ This is done because VB doesn't allow fixed length string
'\\ constants to be exported from a class.
'\\ ----------------------------------------------------------------
Private Function StandardClassName(ByVal ClassType As StandardControlTypes) As String

StandardClassName = StandardControlClassName(ClassType)

End Function

Private Sub Class_Terminate()

Set mIcon = Nothing

End Sub



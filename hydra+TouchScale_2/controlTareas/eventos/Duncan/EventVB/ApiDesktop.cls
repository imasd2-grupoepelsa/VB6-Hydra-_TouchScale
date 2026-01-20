VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2892
Attribute VB_Description="The ApiDesktop class provides the functions and properties to deal with desktop objects.  A desktop is a logical object that contains processes, threads, windows, menus and hooks."
Option Explicit

'\\ --[ApiDesktop]------------------------------------------------------
' ##MODULE_DESCRIPTION The ApiDesktop class provides the functions and _
properties to deal with desktop objects.  A desktop is a logical object that _
contains processes, threads, windows, menus and hooks.

' ##MODULE_DESCRIPTION Desktops are only supported on systems running Windows NT, _
Windows 2000 or Windows XP and beyond.  They are not supported on Windows 95, _
Windows 98 or Windows ME

'\\ --------------------------------------------------------------------

Private mhWinsta As Long '\\ handle to the window station on which this ApiDesktop exists
Private mhDesk As Long '\\ handle to this desktop
Private msName As String
Private dwDesiredAccess As Long


Private Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Long

Private Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long

Private Const DF_ALLOWOTHERACCOUNTHOOK = &H1

Private Const CLSNAME = "ApiDesktop"

Public Property Let DesiredAccess(ByVal dwDes As DesktopAccessRights)
Attribute DesiredAccess.VB_HelpID = 2893

If dwDesiredAccess <> dwDes Then
    dwDesiredAccess = dwDes
    Call RefreshHandle
End If

End Property


Public Property Let Name(ByVal newName As String)
Attribute Name.VB_HelpID = 2894

If newName <> msName Then
    msName = newName
    Call RefreshHandle
End If

End Property

Public Property Get Name() As String

    Name = msName
    
End Property


Private Sub RefreshHandle()

    If mhDesk <> 0 Then
        Call CloseDesktop(mhDesk)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":RefreshHandle", GetLastSystemError
        End If
        mhDesk = 0
    End If
    If mhDesk = 0 Then
        If msName <> "" Then
            Dim lRet As Long
            lRet = OpenDesktop(msName, DF_ALLOWOTHERACCOUNTHOOK, True, dwDesiredAccess)
            If Err.LastDllError Then
                ReportError Err.LastDllError, CLSNAME & ":RefreshHandle", GetLastSystemError
            Else
                mhDesk = lRet
            End If
        End If
    End If


End Sub

Public Property Get Windows() As Collection
Attribute Windows.VB_HelpID = 2895

Call APICallbackProcs.RefreshDesktopWindows(mhDesk)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Windows", GetLastSystemError
End If
Set Windows = APICallbackProcs.DesktopWindows

End Property

Friend Property Let WinstationHandle(ByVal newHandle As Long)

mhWinsta = newHandle

End Property

Public Property Get Handle() As Long
Attribute Handle.VB_HelpID = 2896

    Handle = mhDesk
    
End Property

Friend Property Let Handle(ByVal newHandle As Long)

    mhDesk = newHandle
    
End Property

' ##BLOCK_DESCRIPTION Makes this desktop visible and available for user interaction
Public Sub Activate()
Attribute Activate.VB_HelpID = 2897
Attribute Activate.VB_Description = "Makes this desktop visible and available for user interaction"

Dim lRet As Long

lRet = SwitchDesktop(mhDesk)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiDesktop:Activate", GetLastSystemError
End If

End Sub

Private Sub Class_Terminate()

If mhDesk <> 0 Then
    Call CloseDesktop(mhDesk)
End If

End Sub



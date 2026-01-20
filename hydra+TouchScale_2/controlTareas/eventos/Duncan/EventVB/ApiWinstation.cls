VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiWinstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2898
Attribute VB_Description="The ApiWinstation class provides the functions and properties to deal with window station objects.  A window station is a logical object that contains a clipboard, a set of global atoms and a group of desktop:EventVB~ApiDesktop objects.."
Option Explicit

'\\ --[ApiWinstation]--------------------------------------------------
' ##MODULE_DESCRIPTION The ApiWinstation class provides the functions and _
properties to deal with window station objects.  A window station is a logical object that _
contains a clipboard, a set of global atoms and a group of %desktop:EventVB~ApiDesktop% objects..

' ##MODULE_DESCRIPTION Desktops are only supported on systems running Windows NT, _
Windows 2000 or Windows XP and beyond.  They are not supported on Windows 95, _
Windows 98 or Windows ME

'\\ -------------------------------------------------------------------



Private mhWinsta As Long '\\ Window station handle
Private msName As String
Private mlDesiredAccess As Long

Private Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long

Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type
Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
End Type


Private Declare Function GetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long

' ##ENUMERATION_DESCRIPTION Specifies what type of information will be returned by GetUserObjectInformation API call
Public Enum UserObjectInformationIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION UOI_FLAGS Get the object flags (whose meaning is different depending on the object type)
     UOI_FLAGS = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION UOI_NAME Get the object's name
     UOI_NAME = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION UOI_TYPE Get the object's type name
     UOI_TYPE = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION UOI_USER_SID Get the SID for the object's associated user
     UOI_USER_SID = 4
End Enum

Private Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseWindowStation Lib "user32" (ByVal hWinsta As Long) As Long

Public Property Let DesiredAccess(ByVal dwnew As Long)
Attribute DesiredAccess.VB_HelpID = 2899

    mlDesiredAccess = dwnew

End Property


Public Property Get Desktops() As colDesktops
Attribute Desktops.VB_HelpID = 2900

Call APICallbackProcs.RefreshDesktops(mhWinsta)
Set Desktops = APICallbackProcs.Desktops

End Property

Public Property Get Handle() As Long
Attribute Handle.VB_HelpID = 2901

    Handle = mhWinsta
    
End Property

Friend Property Let Handle(ByVal newHandle As Long)

    mhWinsta = newHandle
    
End Property
Public Property Let Name(ByVal newName As String)
Attribute Name.VB_HelpID = 2903

If newName <> msName Then
    msName = newName
    If mhWinsta <> 0 Then
        Call CloseWindowStation(mhWinsta)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiWinstation:Name", GetLastSystemError
        End If
        mhWinsta = 0
    End If
    If mhWinsta = 0 Then
        Dim lret As Long
        lret = OpenWindowStation(msName, True, mlDesiredAccess)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiWinstation:Name", GetLastSystemError
        Else
            mhWinsta = lret
        End If
    End If
End If

End Property

Public Property Get Name() As String

    Name = msName
    
End Property


Private Sub Class_Terminate()

If mhWinsta <> 0 Then
    Call CloseWindowStation(mhWinsta)
End If

End Sub



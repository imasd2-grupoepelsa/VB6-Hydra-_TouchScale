VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiHardwareProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2651
Option Explicit

'\\ --[ ApiHardwareProfile ] - - - - - - - - - - - - - - - - - - - - - - -
'\\ Code for reading the current hardware profile (plug and play info)
'\\ from the Api - Requires NT4 or Win95
'\\ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Const HW_PROFILE_GUIDLEN = 39
Private Const MAX_HW_PROFILE_LEN = 80

Public Enum DOCKINFO_DOCKSTATUS
    DOCKSTATE_UNDOCKED = &H1
    DOCKSTATE_DOCKED = &H2
    DOCKSTATE_USER_SUPPLIED = &H4
    DOCKSTATE_USER_UNDOCKED = &H5  '&H4 | &H1
    DOCKSTATE_USER_DOCKED = &H6    '&H4 | &H2
End Enum

Private Type HARDWARE_PROFILE
    DockState As DOCKINFO_DOCKSTATUS
    ProfileGuid As String * HW_PROFILE_GUIDLEN
    ProfileName As String * MAX_HW_PROFILE_LEN
End Type

Private Declare Function GetCurrentHardwareProfile Lib "advapi32.dll" Alias "GetCurrentHwProfileA" (hwProfile As HARDWARE_PROFILE) As Long
Private mProfile As HARDWARE_PROFILE

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get DockState() As DOCKINFO_DOCKSTATUS
Attribute DockState.VB_HelpID = 2652

Call RefreshProfile
DockState = mProfile.DockState

End Property

Private Sub RefreshProfile()

Dim lRet As Long
lRet = GetCurrentHardwareProfile(mProfile)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "ApiHardwareProfile:RefreshProfile", GetLastSystemError
End If

End Sub

Public Property Get Guid() As String
Attribute Guid.VB_HelpID = 2653

Call RefreshProfile
Guid = mProfile.ProfileGuid

End Property

Public Property Get Name() As String
Attribute Name.VB_HelpID = 2654

Call RefreshProfile
Name = mProfile.ProfileName

End Property

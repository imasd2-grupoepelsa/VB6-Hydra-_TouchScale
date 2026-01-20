VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2489
Attribute VB_Description = "This class provides properties and methods for installing system wide hotkeys."
Option Explicit

' ##MODULE_DESCRIPTION This class provides properties and methods for installing system wide _
hotkeys.

' ##MODULE_DESCRIPTION A system wide hotkey is a special key combination that triggers an _
event when pressed regardless of which application has the focus when the hotkey is pressed. _
For example, pressing the Windows key + E will launch an instance of the windows explorer and _
pressing Windows Key + M minimises all the open windows on a system.

' ##MODULE_DESCRIPTION Using this class allows you to create your own global hotkeys for your _
applications.

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

' ##ENUMERATION_DESCRIPTION enHotkeyModifiers are the keys that must be _
pressed along with the selected hotkey to trigger the hotkey event
Public Enum enHotkeyModifiers
    ' ##ENUMERATION_MEMBER_DESCRIPTION MOD_ALT The ALT key is included in the hotkey combination
    MOD_ALT = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION MOD_CONTROL The CTRL key is included in the hotkey combination
    MOD_CONTROL = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION MOD_SHIFT The SHIFT key is included in the hotkey combination
    MOD_SHIFT = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION MOD_WIN The WINDOWS key is included in the hotkey combination.
    MOD_WIN = &H8
End Enum

Private mModifier As Long
Private mvKey As Long
Private mHwnd As Long

Private mhAtom As Long

Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private bChanged As Boolean

' ##EVENT_DESCRIPTION HotkeyPressed Triggered whenever the registered hotkey combination is pressed
' ##EVENT_REMARKS HotkeyPressed can be triggered whether your application has the input focus or not.
' ##EVENT_REMARKS Note that if you show a message box in response to this event while the _
application is not the foreground application it will not be visible until the user switches _
back to your application.
Public Event HotkeyPressed()
Attribute HotkeyPressed.VB_Description = "HotkeyPressed Triggered whenever the registered hotkey combination is pressed"
Attribute HotkeyPressed.VB_HelpID = 2813

Private WithEvents wndThis As ApiWindow
Attribute wndThis.VB_VarHelpID = -1
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Let AltKey(ByVal newValue As Boolean)
Attribute AltKey.VB_HelpID = 2491

If newValue Then
    mModifier = mModifier Or MOD_ALT
Else
    mModifier = Not MOD_ALT And mModifier
End If

End Property

Public Property Get AltKey() As Boolean

    AltKey = (mModifier And MOD_ALT)
    
End Property


Public Property Let ControlKey(ByVal newValue As Boolean)
Attribute ControlKey.VB_HelpID = 2493

If newValue Then
    mModifier = mModifier Or MOD_CONTROL
Else
    mModifier = Not MOD_CONTROL And mModifier
End If

End Property


Public Property Get ControlKey() As Boolean

    ControlKey = (mModifier And MOD_CONTROL)
    
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_HelpID = 2494

    If mHwnd = 0 Then
        Dim clsHotkey As APIWndClass
        Set clsHotkey = New APIWndClass
        With clsHotkey
            .hInstance = App.hInstance
            .Name = "STATIC"
        End With
        '\\ Create the window to recieve the messages for this hotkey
        Set wndThis = APIDispenser.System.WindowDispenser.CreateWindow(0, clsHotkey, "MCL Hotkey", 0, 0, 0, 1, 1, 0, 0, 0, 0)
        APIDispenser.SubclassedWindows.Add wndThis
        mHwnd = wndThis.hwnd
    End If

    hwnd = mHwnd
    
End Property



Public Property Let Modifier(ByVal newval As Long)
Attribute Modifier.VB_HelpID = 2495

    mModifier = newval
    
End Property

Public Sub ReadFromAtom(ByVal hAtom As Long)
Attribute ReadFromAtom.VB_HelpID = 2496

Dim sAtom As String
Dim lret As Long

sAtom = String$(255, 0)

lret = GlobalGetAtomName(hAtom, sAtom, Len(sAtom))
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiHotkey:ReadFromAtom", GetLastSystemError
    Exit Sub
End If
If lret > 0 Then
    sAtom = Left$(sAtom, lret)
End If

If InStr(sAtom, "WIN+") Then
    Winkey = True
End If
If InStr(sAtom, "SHIFT+") Then
    ShiftKey = True
End If
If InStr(sAtom, "CONTROL+") Then
    ControlKey = True
End If
If InStr(sAtom, "ALT+") Then
    AltKey = True
End If

If IsNumeric("0" & Right$(sAtom, 3)) Then
    VKey = Val("0" & Right$(sAtom, 3))
End If

End Sub

Public Sub Register()
Attribute Register.VB_HelpID = 2497

    If mhAtom <> 0 Then
        GlobalDeleteAtom mhAtom
    End If
    mhAtom = GlobalAddAtom(UniqueAtomName())
    
    Call RegisterHotKey(hwnd, mhAtom, mModifier, mvKey)
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, "ApiHotkey:Register", GetLastSystemError
    End If
    
End Sub

Public Property Let ShiftKey(ByVal newValue As Boolean)
Attribute ShiftKey.VB_HelpID = 2499

If newValue Then
    mModifier = mModifier Or MOD_SHIFT
Else
    mModifier = Not MOD_SHIFT And mModifier
End If

End Property

Public Property Get ShiftKey() As Boolean

    ShiftKey = (mModifier And MOD_SHIFT)
    
End Property


Public Sub TriggerEvent()
Attribute TriggerEvent.VB_HelpID = 2814

If mhAtom <> 0 Then
    RaiseEvent HotkeyPressed
End If

End Sub

'\\ --[UniqueAtomName]-------------------------------------------------
'\\ Returns a string which uniquely identifies this hotkey combination
'\\ for use in the .id member when regitsering a new hotkey
'\\ -------------------------------------------------------------------
Private Function UniqueAtomName() As String

Dim sAtom As String

If (mModifier And MOD_WIN) Then
    sAtom = sAtom & "WIN+"
End If

If (mModifier And MOD_SHIFT) Then
    sAtom = sAtom & "SHIFT+"
End If

If (mModifier And MOD_CONTROL) Then
    sAtom = sAtom & "CONTROL+"
End If

If (mModifier And MOD_ALT) Then
    sAtom = sAtom & "ALT+"
End If

sAtom = sAtom & Format$(VKey, "000")

UniqueAtomName = sAtom

End Function

Public Sub Unregister()
Attribute Unregister.VB_HelpID = 2500

If mhAtom <> 0 Or mvKey <> 0 Then
    Call UnregisterHotKey(mHwnd, mhAtom)
    GlobalDeleteAtom mhAtom
End If

End Sub

Public Property Let VKey(ByVal newValue As Long)
Attribute VKey.VB_HelpID = 2502

    If newValue <> mvKey Then
        mvKey = newValue
    End If
    
End Property

Public Property Get VKey() As Long

    VKey = mvKey
    
End Property

Public Property Let Winkey(ByVal newValue As Boolean)
Attribute Winkey.VB_HelpID = 2504

If newValue Then
    mModifier = mModifier Or MOD_WIN
Else
    mModifier = Not MOD_WIN And mModifier
End If

End Property

Public Property Get Winkey() As Boolean

    Winkey = (mModifier And MOD_WIN)
    
End Property


Private Sub Class_Terminate()

'\\ Remove from the subclassed windows collection
APIDispenser.SubclassedWindows.Remove wndThis
'\\ Destroy the owned window
If Not wndThis Is Nothing Then
    wndThis.Destroy
    Set wndThis = Nothing
End If

End Sub



Private Sub wndThis_Hotkey(ByVal Hotkey As ApiHotkey)

RaiseEvent HotkeyPressed

End Sub


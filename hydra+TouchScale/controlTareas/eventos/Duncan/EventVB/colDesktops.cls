VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colDesktops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2909
Option Explicit

Private mCol As Collection


Private Declare Function CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, ByVal lpDevMode As Long, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, ByVal lpsa As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

' ##ENUMERATION_DESCRIPTION Specifies the access rights to a created %desktop:EventVB~ApiDesktop%
Public Enum DesktopAccessRights
    ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_READOBJECTS Can read objects on the desktop
     DESKTOP_READOBJECTS = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_CREATEWINDOW Can create windows on the desktop
     DESKTOP_CREATEWINDOW = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_CREATEMENU Can create menus on the desktop
     DESKTOP_CREATEMENU = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_HOOKCONTROL Can install windows hooks on the desktop
     DESKTOP_HOOKCONTROL = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_JOURNALRECORD Can install a journal record hook
     DESKTOP_JOURNALRECORD = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_JOURNALPLAYBACK Can install a journal playback hook
     DESKTOP_JOURNALPLAYBACK = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_ENUMERATE Can enumerate this desktop
     DESKTOP_ENUMERATE = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_WRITEOBJECTS Can write and modify objects on the desktop
     DESKTOP_WRITEOBJECTS = &H80
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_SWITCHDESKTOP Can switch between desktops
     DESKTOP_SWITCHDESKTOP = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION DESKTOP_FULLACCESS Full access to the desktop
     DESKTOP_FULLACCESS = &H1FF
End Enum

Private Const DF_ALLOWOTHERACCOUNTHOOK = &H1

' ##BLOCK_DESCRIPTION Adds a new %desktop:EventVB~ApiDesktop%
' ##BLOCK_REMARKS Name must be unique and must not contain \
Public Function Add(ByVal Name As String, _
                    ByVal AccessRights As DesktopAccessRights) As ApiWinstation
Attribute Add.VB_HelpID = 2910
Attribute Add.VB_Description = "Adds a new desktop:EventVB~ApiDesktop"
                    
Dim dtThis As ApiDesktop

Set dtThis = New ApiDesktop

Dim lret As Long
lret = CreateDesktop(Name, vbNullString, ByVal 0&, DF_ALLOWOTHERACCOUNTHOOK, AccessRights, ByVal 0&)
If Err.LastDllError Then
    ReportError Err.LastDllError, "colDesktops:Add", GetLastSystemError
End If

dtThis.WinstationHandle = lret

mCol.Add dtThis, Name
Set Add = dtThis

End Function


Friend Sub AddItem(ByVal Item As ApiDesktop)

    If Not Item Is Nothing Then
        mCol.Add Item
    End If
    
End Sub
Public Property Get Item(vntIndexKey As Variant) As ApiDesktop
Attribute Item.VB_HelpID = 2911
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2912
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property




Friend Sub RemoveItem(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2913
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub



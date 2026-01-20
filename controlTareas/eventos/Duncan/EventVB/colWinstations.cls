VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colWinstations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2904
Option Explicit

Private mCol As Collection

' ##ENUMERATION_DESCRIPTION The possible access rights for a %window station:EventVB~ApiWinstation%
Public Enum WindowStationAccessRights
    ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_ENUMDESKTOPS Can list the %desktops:EventVB~ApiDesktop% on the window station
     WINSTA_ENUMDESKTOPS = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_READATTRIBUTES Can read the window station attributes
     WINSTA_READATTRIBUTES = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_ACCESSCLIPBOARD Can copy/paste from the system clipboard
     WINSTA_ACCESSCLIPBOARD = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_CREATEDESKTOP Can create %desktops:EventVB~ApiDesktop%
     WINSTA_CREATEDESKTOP = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_WRITEATTRIBUTES Can update the window station attributes
     WINSTA_WRITEATTRIBUTES = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_ACCESSGLOBALATOMS Can read/write to the global atom table
     WINSTA_ACCESSGLOBALATOMS = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_EXITWINDOWS Can exit or shutdown windows
     WINSTA_EXITWINDOWS = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_ENUMERATE Can list the %window stations:EventVB~ApiWinstation%
     WINSTA_ENUMERATE = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_READSCREEN Can access the screen contents
     WINSTA_READSCREEN = &H200
     ' ##ENUMERATION_MEMBER_DESCRIPTION WINSTA_FULLACCESS Can access all aspects of the window station
     WINSTA_FULLACCESS = &H37F
End Enum

Private Declare Function CreateWindowStation Lib "user32" Alias "CreateWindowStationA" (ByVal lpName As String, ByVal dwReserved As Long, ByVal dwAccessMask As Long, ByVal lpSecurityAttributes As Long) As Long



' ##BLOCK_DESCRIPTION Adds a new %window station:EventVB~ApiWinstation%
' ##BLOCK_REMARKS Name must be unique and must not contain \
' ##BLOCK_REMARKS AccessRights is a combination of %WindowStationAccessRights:EventVB~WindowStationAccessRights_EN% values
Public Function Add(ByVal Name As String, _
                    ByVal AccessRights As Long) As ApiWinstation
Attribute Add.VB_HelpID = 2905
Attribute Add.VB_Description = "Adds a new window station:EventVB~ApiWinstation"
                    
Dim wsThis As ApiWinstation
Set wsThis = New ApiWinstation

wsThis.Handle = CreateWindowStation(Name, 0, AccessRights, ByVal 0&)
mCol.Add wsThis, Name
Set Add = wsThis

End Function


Friend Sub AddItem(ByVal Item As ApiWinstation)

    If Not Item Is Nothing Then
        mCol.Add Item
    End If
    
End Sub
Public Property Get Item(vntIndexKey As Variant) As ApiWinstation
Attribute Item.VB_HelpID = 2906
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2907
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
Attribute NewEnum.VB_HelpID = 2908
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cMenu"
Attribute VB_Ext_KEY = "Member0" ,"cMenu"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2301
Option Explicit

'local variable to hold collection
Private mCol As Collection


Friend Function AddInternal(ByVal hMenu As Long, ByVal WindowHandle As Long, ByVal Index As Long) As ApiMenu
Attribute AddInternal.VB_HelpID = 2302
    'create a new object
    Dim objNewMember As ApiMenu
    Set objNewMember = New ApiMenu


    'set the properties passed into the method
    With objNewMember
        .hMenu = hMenu
        .WindowHandle = WindowHandle
        .Index = Index
    End With
    mCol.Add objNewMember, objNewMember.Key

    'return the object created
    Set AddInternal = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ApiMenu
Attribute Item.VB_HelpID = 2303
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2304
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property



Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 2305
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2306
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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


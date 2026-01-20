VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colRegValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2533
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mKey As Long

Public Sub Add(ByVal newValue As ApiRegistryValue)
Attribute Add.VB_HelpID = 2534

    mCol.Add newValue, IIf(newValue.name = "", "[Default]", newValue.name)
    
End Sub

Public Property Get Item(vntIndexKey As Variant) As ApiRegistryValue
Attribute Item.VB_HelpID = 2535
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2536
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property



'\\ --[NewValue]---------------------------------------------------------
'\\ Creates and adds a new registry value entry to this collection of subkeys.
'\\ -------------------------------------------------------------------
Public Function newValue(ByVal name As String, ByVal ValueType As enRegistryKeyValueTypes) As ApiRegistryValue
Attribute newValue.VB_HelpID = 2537

Dim tmpValue As ApiRegistryValue


Set newValue = tmpValue

End Function

Friend Property Let ParentKey(ByVal NewKey As Long)

    mKey = NewKey
    
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 2538
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2539
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRegKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2526
Option Explicit


'local variable to hold collection
Private mCol As Collection
Private mKey As Long

Public Sub Add(ByVal NewKey As ApiRegKey)
Attribute Add.VB_HelpID = 2527

    mCol.Add NewKey, NewKey.Name
    
End Sub

Public Property Get Item(vntIndexKey As Variant) As ApiRegKey
Attribute Item.VB_HelpID = 2528
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2529
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property



'\\ --[NewKey]---------------------------------------------------------
'\\ Creates and adds a new registry to this collection of subkeys.
'\\ -------------------------------------------------------------------
Public Function NewKey(ByVal Name As String) As ApiRegKey
Attribute NewKey.VB_HelpID = 2530

Dim tmpKey As ApiRegKey

With tmpKey
    .Name = Name
    .ParentKey = ParentKey
    .OpenKey
End With

Add tmpKey

Set NewKey = tmpKey

End Function

Friend Property Get ParentKey() As Long

    ParentKey = mKey
    
End Property

Friend Property Let ParentKey(ByVal NewKey As Long)

    mKey = NewKey
    
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 2531
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2532
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

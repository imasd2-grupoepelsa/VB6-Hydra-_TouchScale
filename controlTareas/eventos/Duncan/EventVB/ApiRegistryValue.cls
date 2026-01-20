VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiRegistryValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2518
Attribute VB_Description="This class contains the properties and methods for manipulating a single value stored in a registry key:EventVB~ApiRegKey in the system registry."
Option Explicit

' ##MODULE_DESCRIPTION This class contains the properties and methods _
for manipulating a single value stored in a %registry key:EventVB~ApiRegKey% _
in the system registry.

' ##MODULE_DESCRIPTION The system registry is the prefered method of storing any _
configuration settings that your application uses.


'\\ Describes the data type held in a registry entry...
Public Enum enRegistryKeyValueTypes
    REG_BINARY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
    REG_LINK = 6                       ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7                   ' Multiple Unicode strings
    REG_NONE = 0                       ' No value type
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    REG_RESOURCE_LIST = 8              ' Resource list in the resource map
End Enum


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetStringValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryStringValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long          ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'\\ Memory manipulation routines
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private mParentKey As Long
Private mType As enRegistryKeyValueTypes
Private mName As String

Private mBinaryValue() As Byte
Private mSize As Long
Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get Name() As String
Attribute Name.VB_HelpID = 2519

    Name = mName
    
End Property

Public Property Let Name(ByVal newname As String)

    If mName <> newname Then
        mName = newname
    End If
    
End Property


Friend Property Let ParentKey(ByVal NewKey As Long)

    If NewKey <> mParentKey Then
        mParentKey = NewKey
    End If
    
End Property

Public Property Get ParentKey() As Long
Attribute ParentKey.VB_HelpID = 2520

    ParentKey = mParentKey
    
End Property

Friend Property Let RawData(ByVal newData As String)

If Len(newData) <> Size Then
    Size = Len(newData)
End If
CopyMemory ByVal VarPtr(mBinaryValue(0)), ByVal StrPtr(newData), ByVal Size
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiRegistryValue:RawData", GetLastSystemError
End If

End Property

Public Sub Read()
Attribute Read.VB_HelpID = 2521

Dim lret As Long
Dim lLen As Long

'\\ 1st, get the size requirted....
lret = RegQueryValueEx(mParentKey, mName, 0, mType, ByVal 0&, lLen)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "ApiRegistryValue:Read (Size)", GetLastSystemError
End If
If lLen > 0 Then
   ReDim mBinaryValue(0 To (lLen - 1)) As Byte
   lret = RegQueryValueEx(mParentKey, mName, 0, mType, mBinaryValue(0), lLen)
End If
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "ApiRegistryValue:Read (Value)", GetLastSystemError
End If

End Sub

Public Property Get Size() As Long
Attribute Size.VB_HelpID = 2522

    Size = mSize
    
End Property

Public Property Let Size(ByVal newSize As Long)

    If newSize <> mSize Then
        mSize = newSize
        ReDim mBinaryValue(0 To mSize) As Byte
    End If
    
End Property





Public Property Let ValueType(ByVal newType As enRegistryKeyValueTypes)
Attribute ValueType.VB_HelpID = 2524

    If newType <> mType Then
        mType = newType
    End If
    
End Property

Public Property Get ValueType() As enRegistryKeyValueTypes
    
    ValueType = mType

End Property

Public Sub Save()
Attribute Save.VB_HelpID = 2525

Dim lret As Long

lret = RegSetValueEx(mParentKey, mName, 0, mType, mBinaryValue(0), mSize)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiRegistryValue:Save", GetLastSystemError
End If

End Sub

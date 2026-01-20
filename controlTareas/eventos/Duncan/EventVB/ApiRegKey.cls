VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiRegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2505
Attribute VB_Description = "This class contains the properties and methods for manipulating a key in the registry."
Option Explicit

' ##MODULE_DESCRIPTION This class contains the properties and methods _
for manipulating a key in the registry.

' ##MODULE_DESCRIPTION The system registry is the prefered method of storing any _
configuration settings that your application uses.  A registry key can be considered akin _
to a directory in a file structure and the %registry values:EventVB~ApiRegistryValue% are _
analogous to files.


Public Enum enStandardRegistryKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum


Public Enum enRegsitryKeyOptions
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_OPTION_NON_VOLATILE The key is preserved when system is rebooted
    REG_OPTION_NON_VOLATILE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_OPTION_VOLATILE The key is not preserved when system is rebooted
    REG_OPTION_VOLATILE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_OPTION_CREATE_LINK The created key is a symbolic link
    REG_OPTION_CREATE_LINK = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_OPTION_BACKUP_RESTORE The key is open for backup or restore
    REG_OPTION_BACKUP_RESTORE = 4
End Enum

Public Enum enStandardRights
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_EXECUTE = &H20000
    STANDARD_RIGHTS_READ = &H20000
    STANDARD_RIGHTS_REQUIRED = &HF0000
    STANDARD_RIGHTS_WRITE = &H20000
End Enum

Public Enum enRegsitryKeyPermissions
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000))
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_EVENT = &H1     '  Event contains key event record
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_READ = ((STANDARD_RIGHTS_READ Or &H1 Or &H8 Or &H10) And (Not &H100000))
    KEY_SET_VALUE = &H2
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or &H1 Or &H4) And (Not &H100000))
End Enum

Public Enum enRegistryKeyDispositions
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_CREATED_NEW_KEY A New Registry Key created
    REG_CREATED_NEW_KEY = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION REG_OPENED_EXISTING_KEY An Existing Key opened
    REG_OPENED_EXISTING_KEY = &H2
End Enum

Public ParentKey As Long
Attribute ParentKey.VB_VarHelpID = 2506
Private mKey As Long '\\ Unique handle of this key object
Private mName As String
Private bOpen As Boolean

Public DesiredAccess As enRegsitryKeyPermissions
Attribute DesiredAccess.VB_VarHelpID = 2507
Public Options As enRegsitryKeyOptions
Attribute Options.VB_VarHelpID = 2508

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Dim cSubkeys As colRegKeys
Dim cValues As colRegValues

Private mParentPath As String

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Sub CloseKey()
Attribute CloseKey.VB_HelpID = 2509

Dim lret As Long

If bOpen Then
    lret = RegCloseKey(Key)
    If Err.LastDllError = 0 Then
        Key = 0
    Else
        ReportError Err.LastDllError, "ApiRegKey:CloseKey", GetLastSystemError
    End If
End If

End Sub

Public Property Get FullName() As String
Attribute FullName.VB_HelpID = 2510

    FullName = mParentPath & "\" & Name
    
End Property

Public Property Let Key(ByVal NewKey As Long)
Attribute Key.VB_HelpID = 2512

If NewKey <> mKey Then
    mKey = NewKey
End If

End Property

Public Property Get Key() As Long

    If mKey = 0 Then
        Call OpenKey
    End If
    Key = mKey
    
End Property

Public Property Let Name(ByVal newname As String)
Attribute Name.VB_HelpID = 2514

    mName = newname
    
End Property

Public Property Get Name() As String

Name = mName

End Property


Public Sub OpenKey()
Attribute OpenKey.VB_HelpID = 2515

Dim lNewKey As Long
Dim lRes As Long

lRes = RegOpenKeyEx(ParentKey, Name, Options, DesiredAccess, lNewKey)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiRegKey:OpenKey", GetLastSystemError
End If
If lRes = 0 Then
    bOpen = True
    mKey = lNewKey
End If

End Sub



Friend Property Let ParentPath(ByVal thePath As String)

    mParentPath = thePath
    
End Property

Public Property Get Subkeys() As colRegKeys
Attribute Subkeys.VB_HelpID = 2516

Dim sKeyName As String, sClassName As String
Dim lKeyLen As Long, lClassLen As Long
Dim lIndex As Long, lReserved As Long
Dim lpftLastWriteTime As FILETIME
Dim lret As Long


Dim cRegKey As ApiRegKey

Set cSubkeys = New colRegKeys

If Not bOpen Then
    Call OpenKey
End If


'\\ Set the parent key
cSubkeys.ParentKey = Key

'\\ Initialise the string that are going to hold the values...
lKeyLen = 2000
lClassLen = 2000
sKeyName = String$(lKeyLen, 0)
sClassName = String$(lClassLen, 0)

'\\ Enumerate through the keys adding each one...
While lret = 0
    lret = RegEnumKeyEx(Key, lIndex, sKeyName, lKeyLen, lReserved, vbNullString, lClassLen, lpftLastWriteTime)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, "ApiRegKey:Subkeys", GetLastSystemError
        lret = -1
    End If
    If lret = 0 Then
    lIndex = lIndex + 1
    Set cRegKey = New ApiRegKey
    With cRegKey
        .Name = Left$(sKeyName, lKeyLen)
        .ParentKey = Key
        .ParentPath = FullName
    End With
    Call cSubkeys.Add(cRegKey)
    Set cRegKey = Nothing
    End If
Wend

Set Subkeys = cSubkeys

End Property

Public Property Get Values() As colRegValues
Attribute Values.VB_HelpID = 2517

Dim dwIndex As Long, lpValueName As String, lpcbValueName As Long
Dim lpType As enRegistryKeyValueTypes
Dim lpData() As Byte, lpcbData As Long
Dim lret As Long

Dim cRegValue As ApiRegistryValue

Set cValues = New colRegValues

If Not bOpen Then
    Call OpenKey
End If

cValues.ParentKey = Key

lpcbValueName = 2048
lpValueName = String$(lpcbValueName, 0)

While lret = 0
    '\\ 1st get the size that will be required....
    lpcbData = 0
    lret = RegEnumValue(Key, dwIndex, lpValueName, lpcbValueName, ByVal 0&, lpType, ByVal 0&, lpcbData)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, "ApiRegKey:Values", GetLastSystemError
        lret = -1
    End If
    If lret = 0 Or lret = 234 Then
        If lpcbData > 0 Then
            ReDim lpData(0 To lpcbData + 1) As Byte
            lret = RegEnumValue(Key, dwIndex, lpValueName, lpcbValueName, ByVal 0, lpType, lpData(0), lpcbData)
            If Err.LastDllError > 0 Then
                ReportError Err.LastDllError, "ApiRegKey:Values", GetLastSystemError
                lret = -1
            End If
            If lret = 0 Then
                lpValueName = Left$(lpValueName, lpcbValueName)
                Set cRegValue = New ApiRegistryValue
                With cRegValue
                    .Name = lpValueName
                    .ValueType = lpType
                    .Size = lpcbData
                    .RawData = CStr(lpData)
                End With
                cValues.Add cRegValue
                Set cRegValue = Nothing
            End If
        End If
        Debug.Print "lret = " & lret & " , lpcbdata = " & lpcbData
    End If
    '\\ Move on to next value
    dwIndex = dwIndex + 1
Wend

Set Values = cValues

End Property


Private Sub Class_Terminate()

'\\ Close the key if it has been left open...
If bOpen Then
    Call CloseKey
End If

Set cValues = Nothing
Set cSubkeys = Nothing

End Sub



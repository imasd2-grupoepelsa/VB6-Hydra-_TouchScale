Attribute VB_Name = "modRegistry"
' modRegistry.Bas - Declares, Subroutines and Functions for the Registry
' 1997/05/03 Copyright 1994-1997, Larry Rebich, The Bridge, Inc.
'
' Functions and Subroutines included in this module:
'
'   Function RegRead    Returns a value using the Supplied Key and Value Name
'   Function RegWrite   Write a value using the Supplied Key, Value Name and Value
'                       Remove the Value Name if value is null [""]
'   Function RegCreate  Create a Key, open the Key, then close the Key.
'                       This function is called by RegWrite if the key does not exist.
'                       There should be no reason to call this function directly.
'
' Only string data [REG_SZ] is process by these routines.
'
' For details see the comments associated with each function.
'
    
    Option Explicit
    DefLng A-Z

    Const HKEY_CURRENT_USER = &H80000001
    Const ERROR_FILE_NOT_FOUND& = 2
    Const ERROR_BADKEY& = 1010
    Public Const ERROR_SUCCESS& = 0
    Public Const NO_ERROR& = 0

    Const REG_SZ = 1                    ' Unicode nul terminated string

    Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
'
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescription As Long   'SECURITY_DESCRIPTOR
        bInheritHandle As Boolean
    End Type

    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
         ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
         lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
         lpdwDisposition As Long) As Long

    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, ByVal lpValueName As String) As Long

    Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" _
        (ByVal hKey As Long, ByVal lpValueName As String, phkResult As Long) As Long

    Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal ulOptions As Long, _
         ByVal samDesired As Long, phkResult As Long) As Long

    Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, _
         lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

    Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
         ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

    Const KEY_QUERY_VALUE = &H1&
    Const KEY_SET_VALUE = &H2&
    Const KEY_CREATE_SUB_KEY = &H4&
    Const KEY_ENUMERATE_SUB_KEYS = &H8&
    Const KEY_NOTIFY = &H10&
    Const READ_CONTROL = &H20000
    Const SYNCHRONIZE = &H100000
    Const STANDARD_RIGHTS_READ = READ_CONTROL
    Const STANDARD_RIGHTS_WRITE = READ_CONTROL
    Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
    Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
    Const REG_OPTION_NON_VOLATILE = 0&
    Public Const REG_CREATED_NEW_KEY& = 1
    Public Const REG_OPENED_EXISTING_KEY& = 2
' End of Declarations ----------------------------------------------------------------
    
Private Function RegCreate(sKey As String) As Long
' Create a key
' Returns:
'    False if Fails to Create the Key
' or lDisposition:
'    REG_CREATED_NEW_KEY& = 1& or   'created a new key
'    REG_OPENED_EXISTING_KEY& = 2&  'key already exists
'
    Dim lRtn As Integer
    Dim lHKey As Long           'return handle to opened key
    Dim lDisposition As Long    'disposition
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES

    lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, sKey, 0&, "", _
        REG_OPTION_NON_VOLATILE, KEY_WRITE, lpSecurityAttributes, _
        lHKey, lDisposition)

    If lRtn = ERROR_SUCCESS Then
        RegCreate = lDisposition    'tell 'em if it existed or was created
        lRtn = RegCloseKey(lHKey)   'close the Registry
    End If
End Function
Public Function RegRead(sKey As String, sValueName As String, Optional vntOptionalHKey As Variant) As String
' Returns the Value found for this Key and ValueName
' Input:        Sample:
'   sKey        "Software\Microsoft\File Manager\Settings"
'   sValueName  "Face"
' Return:
'               "FixedSys" or
'               "" [null] if not found
'-----------------------------------------------------------------------------------
' 96/09/18 Add support for different root level key. Needed to find DAO3032.DLL in class registry. Larry.
    Dim lOptionalHKey As Long   '96/09/18 Can open a different area key.
    Dim lKeyType As Long
    Dim lHKey As Long       'return handle to opened key
    Dim lpcbData As Long    'length of data in returned string
    Dim sReturnedString As String   'returned string value
    Dim sTemp As String     'temp string
    Dim lRtn As Long        'success or not success
    
    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_CURRENT_USER   'Use current user
    Else
        lOptionalHKey = vntOptionalHKey     'Use the one supplied
    End If

    lKeyType = REG_SZ       'data type is string
    lRtn = RegOpenKeyEx(lOptionalHKey, sKey, 0&, KEY_READ, lHKey)
    If lRtn = ERROR_SUCCESS Then
        lpcbData = 1024                     'get this many characters
        sReturnedString = Space$(lpcbData)  'setup the buffer
        lRtn = RegQueryValueEx(lHKey, sValueName, ByVal 0&, lKeyType, sReturnedString, lpcbData)
        If lRtn = ERROR_SUCCESS Then
            sTemp = Left$(sReturnedString, lpcbData - 1)
        End If
        RegCloseKey lHKey
    End If
    RegRead = sTemp
End Function
Public Function RegWrite(sKey As String, sValueName As String, sValue As String, Optional vntOptionalHKey As Variant) As Integer
' Input:        Sample:
'   sKey        "Software\Microsoft\File Manager\Settings"
'   sValueName  "Face"
'   sValue      "FixedSys"
' Return:
'   True if successful
'
' If the current setting is the same as the new setting then the update is bypassed.
'
' Note: If sValue = "" then sValueName is removed [deleted].
'-----------------------------------------------------------------------------------
    Dim lOptionalHKey As Long   '10/14/96 Can open a different area key(to register fonts). Boris
    Dim lRtn        As Long
    Dim lKeyType    As Long 'returns the key type.  This function expects REG_SZ
    Dim lHKey       As Long 'return handle to opened key
    Dim iSuccessCount As Integer
    lKeyType = REG_SZ       'these routines support only string types

    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_CURRENT_USER   'Use current user
    Else
        lOptionalHKey = vntOptionalHKey     'Use the one supplied
    End If
    
    If Trim$(sValue) <> "" Then             'if there is a value then update it
RegWriteTryAgain:
        lRtn = RegOpenKeyEx(lOptionalHKey, sKey, 0&, KEY_SET_VALUE, lHKey)  'open the Registry for update
        If lRtn = ERROR_SUCCESS Then
            lRtn = RegSetValueEx(lHKey, sValueName, 0&, lKeyType, ByVal sValue, CLng(Len(sValue) + 1))   'update the value
            If lRtn = ERROR_SUCCESS Then
                iSuccessCount = iSuccessCount + 1
            End If
            lRtn = RegCloseKey(lHKey)       'close the Registry
        ElseIf lRtn = ERROR_FILE_NOT_FOUND Or lRtn = ERROR_BADKEY Then 'create it
            If RegCreate(sKey) Then         'Create it, was it successful?
                GoTo RegWriteTryAgain       'Yes, go try writing again
            End If
        End If
    Else                                    'Value is null, delete the key
        lRtn = RegOpenKeyEx(lOptionalHKey, sKey, 0&, KEY_SET_VALUE, lHKey)  'open the Registry for update
        If lRtn = ERROR_SUCCESS Then
            lRtn = RegDeleteValue(lHKey, sValueName)
            If lRtn = ERROR_SUCCESS Then
                iSuccessCount = iSuccessCount + 1
            End If
            lRtn = RegCloseKey(lHKey)       'close the Registry
        End If
    End If
    If iSuccessCount > 0 Then
        RegWrite = True                     'OK, changed
    End If
End Function













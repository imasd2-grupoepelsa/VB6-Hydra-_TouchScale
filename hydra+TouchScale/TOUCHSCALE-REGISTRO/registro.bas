Attribute VB_Name = "registro"
Option Explicit
Public Const REG_DWORD = 4   ' Número de 32 bits
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long


Public Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegSetValueExLong Ret, strValue, 0&, REG_DWORD, CByte(strData), 4&

    RegCloseKey Ret
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2040
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OReg As Registro

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Const API_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const ERROR_NOT_REGISTRY_FILE = 1017&
Const ERROR_KEY_DELETED = 1018&
Const ERROR_NO_LOG_SPACE = 1019&
Const ERROR_KEY_HAS_CHILDREN = 1020&
Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
Const ERROR_RXACT_INVALID_STATE = 1369&

Public Function GetRegErrorText(lStatus As Long) As String
Select Case lStatus
  Case ERROR_BADDB
    GetRegErrorText = "The configuration registry database " & _
                      "is corrupt."
       
  Case ERROR_BADKEY
    GetRegErrorText = "The configuration registry key is " & _
                      "invalid."
       
  Case ERROR_CANTOPEN
    GetRegErrorText = "The configuration registry key could " & _
                      "not be opened."
       
  Case ERROR_CANTREAD
    GetRegErrorText = "The configuration registry key could " & _
                      "not be read."
       
  Case ERROR_CANTWRITE
    GetRegErrorText = "The configuration registry key could " & _
                      "not be written."
       
  Case ERROR_REGISTRY_RECOVERED
    GetRegErrorText = "One of the files in the Registry " & _
                      "database had to be recovered " & _
                      "by use of a log or alternate copy. " & _
                      "The recovery was successful."
                      
  Case ERROR_REGISTRY_CORRUPT
    GetRegErrorText = "The Registry is corrupt. The structure " & _
                      "of one of the files that contains " & _
                      "Registry data is corrupt, or the " & _
                      "system's image of the file in memory " & _
                      "is corrupt, or the file could not be " & _
                      "recovered because the alternate " & _
                      "copy or log was absent or corrupt."
                      
  Case ERROR_REGISTRY_IO_FAILED
    GetRegErrorText = "An I/O operation initiated by the " & _
                      "Registry failed unrecoverably. " & _
                      "The Registry could not read in, or " & _
                      "write out, or flush, one of the files " & _
                      "that contain the system's image of " & _
                      "the Registry."
                      
  Case ERROR_NOT_REGISTRY_FILE
    GetRegErrorText = "The system has attempted to load or " & _
                      "restore a file into the Registry, but the " & _
                      "specified file is not in a Registry " & _
                      "file format."
                      
  Case ERROR_KEY_DELETED
    GetRegErrorText = "Illegal operation attempted on a " & _
                      "Registry key which has been marked " & _
                      "for deletion."
       
  Case ERROR_NO_LOG_SPACE
    GetRegErrorText = "System could not allocate the required " & _
                      "space in a Registry log."
       
  Case ERROR_KEY_HAS_CHILDREN
    GetRegErrorText = "Cannot create a symbolic link in a " & _
                      "Registry key that already " & _
                      "has subkeys or values."
                      
  Case ERROR_CHILD_MUST_BE_VOLATILE
    GetRegErrorText = "Cannot create a stable subkey under a " & _
                      "volatile parent key."
       
  Case ERROR_RXACT_INVALID_STATE
    GetRegErrorText = "The transaction state of a Registry " & _
                      "subtree is incompatible with the " & _
                      "requested operation."
                      
End Select
  
End Function

Private Sub Form_Load()

End Sub

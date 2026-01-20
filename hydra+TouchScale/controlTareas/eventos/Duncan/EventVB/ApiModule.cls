VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 3140
Attribute VB_Description="Provides the functions for getting information about a module"
Option Explicit

' ##MODULE_DESCRIPTION Provides the functions for getting information about a module
' ##MODULE_REMARKS Requires Windows NT4, 2000 or XP
' ##MODULE_REMARKS PSAPI.dll exports used: EmptyWorkingSet  GetMappedFileNameA  GetModuleBaseNameA  GetModuleFileNameExA  GetModuleInformation GetProcessMemoryInfo GetWsChanges InitializeProcessForWsWatch QueryWorkingSet

Private Const CLSNAME = "ApiModule"

Private mhModule As Long
Private mhProcess As Long

Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpszBaseName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFilenameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpszFileName As String, ByVal nSize As Long) As Long

Friend Property Let Handle(ByVal newval As Long)

mhModule = newval

End Property

Friend Property Let ProcessHandle(ByVal newval As Long)

    mhProcess = newval
    
End Property

' ##BLOCK_DESCRIPTION BaseName - The base name of the module
' ##BLOCK_REMARKS This is often, but not always, the filename without the path
Public Property Get BaseName() As String
Attribute BaseName.VB_HelpID = 3141
Attribute BaseName.VB_Description = "BaseName - The base name of the module"

Dim sBuff As String
Dim lret As Long

sBuff = String$(64, 0)
lret = GetModuleBaseName(mhProcess, mhModule, sBuff, Len(sBuff))
If lret > 0 Then
    sBuff = Left$(sBuff, lret)
Else
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":BaseName", GetLastSystemError
    End If
End If

BaseName = sBuff

End Property

' ##BLOCK_DESCRIPTION BaseName - The file name of the module
' ##BLOCK_REMARKS This includes the full path
Public Property Get Filename() As String
Attribute Filename.VB_HelpID = 3142
Attribute Filename.VB_Description = "BaseName - The file name of the module"
Dim sBuff As String
Dim lret As Long

sBuff = String$(512, 0)
lret = GetModuleFilenameEx(mhProcess, mhModule, sBuff, Len(sBuff))
If lret > 0 Then
    sBuff = Left$(sBuff, lret)
Else
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Filename", GetLastSystemError
    End If
End If

Filename = sBuff

End Property


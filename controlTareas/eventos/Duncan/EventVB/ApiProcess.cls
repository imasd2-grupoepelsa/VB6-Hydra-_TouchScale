VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2914
Attribute VB_Description="Object wrapper for api functions pertaining to a process"
Option Explicit

'\\ --[ApiProcess]---------------------------------------------------------------
' ##MODULE_DESCRIPTION Object wrapper for api functions pertaining to a process
'\\ -----------------------------------------------------------------------------

Private Const CLSNAME = "ApiProcess"

'\\ Private member variables
Private mhProcess As Long '\\ process handle

Private mProcid As Long
Private mDesiredAccess As Long
Private bInheritableHandle As Boolean

'\\ API calls used
Private Declare Function CloseKernelObject Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Type IO_COUNTERS
    ReadOperationCount As Long
    WriteOperationCount As Long
    OtherOperationCount As Long
    ReadTransferCount As Long
    WriteTransferCount As Long
    OtherTransferCount As Long
End Type
Private mIOCounters As IO_COUNTERS

Private Declare Function GetProcessIoCounters Lib "kernel32" (ByVal hProcess As Long, ioCount As IO_COUNTERS) As Long

' ##ENUMERATION_DESCRIPTION The desired security access priviledge for the process
Public Enum ProcessAccessPriviledges
    ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_TERMINATE Can kill the process
     PROCESS_TERMINATE = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_CREATE_THREAD Can create threads under the process
     PROCESS_CREATE_THREAD = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_SET_SESSIONID Can update the session id
     PROCESS_SET_SESSIONID = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_VM_OPERATION Can update the processes virtual memory
     PROCESS_VM_OPERATION = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_VM_READ Can read the processes virtual memory
     PROCESS_VM_READ = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_VM_WRITE Can write to the processes virtual memory (use with caution)
     PROCESS_VM_WRITE = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_DUP_HANDLE Can duplicate the process handle
     PROCESS_DUP_HANDLE = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_CREATE_PROCESS (Not used by client apps)
     PROCESS_CREATE_PROCESS = &H80
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_SET_QUOTA Can set the process quotas and working set size
     PROCESS_SET_QUOTA = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_SET_INFORMATION Can set the priority class
     PROCESS_SET_INFORMATION = &H200
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_QUERY_INFORMATION Can get the exit codes and priority class
     PROCESS_QUERY_INFORMATION = &H400
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_SYNCHRONISE Can synchronise with the process
     PROCESS_SYNCHRONISE = &H100000
     ' ##ENUMERATION_MEMBER_DESCRIPTION PROCESS_ALL_ACCESS All process access rights
     PROCESS_ALL_ACCESS = &H100FFF
End Enum

'\\ tO ENUMERATE THE THREADS
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal hProcess As Long) As Long
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, teThis As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, teThis As THREADENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, modArray As Long, ByVal cbLen As Long, cbRequired As Long) As Long

'\\ Used to enumerate the process
Friend Property Let DesiredAccess(ByVal newAccess As ProcessAccessPriviledges)

If newAccess <> mDesiredAccess Then
    mDesiredAccess = newAccess
    '\\ Refresh the process handle to get the desired access
    If mProcid <> 0 Then
        If mhProcess <> 0 Then
            Call CloseKernelObject(mhProcess)
            If Err.LastDllError Then
                ReportError Err.LastDllError, "ApiProcess:DesiredAccess", GetLastSystemError
            Else
                mhProcess = 0
            End If
        End If
        If mhProcess = 0 Then
            mhProcess = OpenProcess(mDesiredAccess, bInheritableHandle, mProcid)
            If Err.LastDllError Then
                ReportError Err.LastDllError, "ApiProcess:DesiredAccess", GetLastSystemError
            End If
        End If
    End If
End If

End Property

Friend Property Get Key() As String

    Key = "HPROC:" & mProcid
    
End Property

Public Property Get Modules() As Collection
Attribute Modules.VB_HelpID = 3112

Dim lMods() As Long
Dim lret As Long
Dim lBytesRequired As Long
Dim colMods As New Collection
Dim modThis As ApiModule

ReDim lMods(0 To 30) As Long
lret = EnumProcessModules(mhProcess, lMods(0), (UBound(lMods) - LBound(lMods) * Len(lMods(1))), lBytesRequired)
If lret = 0 And Err.LastDllError = 0 Then
    '\\ The function failed - increase the buffer size
    ReDim lMods(0 To (lBytesRequired / 4) + 1)
    lret = EnumProcessModules(mhProcess, lMods(0), (UBound(lMods) - LBound(lMods) * Len(lMods(1))), lBytesRequired)
End If
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Modules", GetLastSystemError
Else
    For lret = 0 To UBound(lMods)
        Set modThis = New ApiModule
        With modThis
            .Handle = lMods(lret)
            .ProcessHandle = mhProcess
        End With
        colMods.Add modThis
    Next lret
End If

Set Modules = colMods

End Property

Public Property Get ProcessHandle() As Long
Attribute ProcessHandle.VB_HelpID = 2916

    ProcessHandle = mhProcess
    
End Property

Friend Property Let ProcessId(ByVal newId As Long)

If newId <> mProcid Then
    mProcid = newId
    '\\ If process handle was open, close it
    If mhProcess <> 0 Then
        Call CloseKernelObject(mhProcess)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiProcess:ProcessId (set)", GetLastSystemError
        Else
            mhProcess = 0
        End If
    End If
    If mhProcess = 0 Then
        mhProcess = OpenProcess(mDesiredAccess, bInheritableHandle, mProcid)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiProcess:ProcessId (set)", GetLastSystemError
        End If
    End If
End If

End Property
Public Property Get ReadOperations() As Long
Attribute ReadOperations.VB_HelpID = 2917

Call RefreshIOCounters
ReadOperations = mIOCounters.ReadOperationCount

End Property

Public Property Get ReadTransfers() As Long
Attribute ReadTransfers.VB_HelpID = 2918

Call RefreshIOCounters
ReadTransfers = mIOCounters.ReadTransferCount

End Property

Private Sub RefreshIOCounters()

Static bInsufficientVersion As Boolean
Dim lret As Long

If mhProcess <> 0 Then
    If Not bInsufficientVersion Then
        If Not APIDispenser.System.IsRequiredSystem("RefreshIOCounters", ver_Win_Win2000) Then
            bInsufficientVersion = True
            Exit Sub
        End If
        lret = GetProcessIoCounters(mhProcess, mIOCounters)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiProcess:RefreshIOCounters", GetLastSystemError
        End If
    End If
End If

End Sub

Public Property Get Threads() As Collection
Attribute Threads.VB_HelpID = 2919

Dim colThreads As New Collection
Dim teThis As THREADENTRY32
Dim threadThis As ApiThread
Dim hSnapshot As Long
Dim lret As Long

hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, mhProcess)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiProcess:Threads", GetLastSystemError
Else
    If hSnapshot <> 0 Then
        teThis.dwSize = Len(teThis)
        lret = Thread32First(hSnapshot, teThis)
        While lret <> 0
            Set threadThis = New EventVB.ApiThread
            With threadThis
                .DesiredAccess = THREAD_SYNCHRONISE
                .ThreadId = teThis.th32ThreadID
            End With
            colThreads.Add threadThis, threadThis.Key
            lret = Thread32Next(hSnapshot, teThis)
        Wend
        lret = CloseHandle(hSnapshot)
    End If
End If
Set Threads = colThreads

End Property

Public Property Get WriteOperations() As Long
Attribute WriteOperations.VB_HelpID = 2920

Call RefreshIOCounters
WriteOperations = mIOCounters.WriteOperationCount

End Property

Public Property Get WriteTransfers() As Long
Attribute WriteTransfers.VB_HelpID = 2921

    Call RefreshIOCounters
    WriteTransfers = mIOCounters.WriteTransferCount
    
End Property

Private Sub Class_Terminate()

    If mhProcess <> 0 Then
        Call CloseKernelObject(mhProcess)
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiProcess:ProcessId (set)", GetLastSystemError
        Else
            mhProcess = 0
        End If
    End If
    
End Sub



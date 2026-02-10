Attribute VB_Name = "Process"
Option Explicit

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Public Const MAX_PATH = 260
Private Declare Function CloseHandle Lib "kernel32" ( _
                                     ByVal hObject As Long) As Long
Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long    ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long    ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long    ' This process's parent process
    pcPriClassBase As Long    ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260    ' MAX_PATH
End Type

'===============
'
''''''''''''''''
Private Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String

    Dim lngCounter As Long

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Calculate the offset for the item required based on the number of columns the list
    ' 'strList' has i.e. 'lngNumColumns' and from which row the element is to be
    ' selected i.e. 'lngRow'.
    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn)

    ' Search for the 'lngColumn' item from the list 'strList'.
    For lngCounter = 0 To lngColumn - 1

        ' Remove each item from the list.
        strList = Mid$(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))

        ' If list becomes empty before 'lngColumn' is found then just
        ' return an empty string.
        If Len(strList) = 0 Then
            GetElement = ""
            Exit Function
        End If

    Next lngCounter

    ' Return the sought list element.
    GetElement = left$(strList, InStr(strList, strDelimiter) - 1)

End Function


Private Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer

    Dim intElementCount As Integer

    ' If no elements in the list 'strList' then just return 0.
    If Len(strList) = 0 Then
        GetNumElements = 0
        Exit Function
    End If

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Count the number of elements in 'strlist'
    While InStr(strList, strDelimiter) > 0
        intElementCount = intElementCount + 1
        strList = Mid$(strList, InStr(strList, strDelimiter) + 1, Len(strList))
    Wend

    ' Return the number of elements in 'strList'.
    GetNumElements = intElementCount

End Function
Public Function GetProcesses(ByVal sAppName As String) As Boolean
    Dim booResult As Boolean
    Dim lngLength As Long
    Dim lngProcessID As Long
    Dim strProcessName As String
    Dim lngSnapHwnd As Long
    Dim udtProcEntry As PROCESSENTRY32
    Dim lngCBSize As Long    'Specifies the size, In bytes, of the lpidProcess array
    Dim lngCBSizeReturned As Long    'Receives the number of bytes returned
    Dim lngNumElements As Long
    Dim lngProcessIDs() As Long
    Dim lngCBSize2 As Long
    Dim lngModules(1 To 200) As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim lngSize As Long
    Dim lngHwndProcess As Long
    Dim lngLoop As Long
    Dim b As Long
    Dim c As Long
    Dim e As Long
    Dim d As Long
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim lRet As Long
    Dim strProcName2 As String
    Dim strProcName As String
    Dim nf As Integer
    Dim nF1 As Integer
    Dim MemStat As MEMORYSTATUS
    Dim MemOcupied As Long
    Dim sS As String
    Dim nF2 As Integer
    Dim MyIP As String
    Dim nPing As Integer
    Dim nFX As Integer

    'Turn on Error handler
    On Error GoTo Error_handler

    booResult = False


    MemOcupied = 0

    lngCBSize = 8    ' Really needs To be 16, but Loop will increment prior to calling API
    lngCBSizeReturned = 96

    Do While lngCBSize <= lngCBSizeReturned
        DoEvents
        'Increment Size
        lngCBSize = lngCBSize * 2
        'Allocate Memory for Array
        ReDim lngProcessIDs(lngCBSize / 4) As Long
        'Get Process ID's
        lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
    Loop

    'Count number of processes returned
    lngNumElements = lngCBSizeReturned / 4
    'Loop thru each process

    For lngLoop = 1 To lngNumElements
        DoEvents

        'Get a handle to the Process and Open
        lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))

        If lngHwndProcess <> 0 Then
            'Get an array of the module handles for the specified process
            lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)

            'If the Module Array is retrieved, Get the ModuleFileName
            If lngReturn <> 0 Then

                'Buffer with spaces first to allocate memory for byte array
                strModuleName = Space(MAX_PATH)

                'Must be set prior to calling API
                lngSize = 500

                'Get Process Name
                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, lngSize)

                'Remove trailing spaces
                strProcessName = left(strModuleName, lngReturn)

                'Check for Matching Upper case result
                strProcessName = UCase$(Trim$(strProcessName))

                strProcName2 = GetElement(Trim(Replace(strProcessName, Chr$(0), "")), "\", 0, 0, GetNumElements(Trim(Replace(strProcessName, Chr$(0), "")), "\") - 1)

                If (LCase(strProcName2) = LCase(sAppName)) Then

                    booResult = True

                End If
            End If
        End If
        'Close the handle to this process
        lngReturn = CloseHandle(lngHwndProcess)
        DoEvents
    Next

Error_handler:
    On Error GoTo 0
    GetProcesses = booResult

End Function


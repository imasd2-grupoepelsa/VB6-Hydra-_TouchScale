VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiPrintJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2683
Attribute VB_Description = "This class provides the properties for a getting information about a job in a printer queue."
Option Explicit

Private Const CLSNAME = "ApiPrintJob"

' ##MODULE_DESCRIPTION This class provides the properties _
for a getting information about a job in a printer queue.

' ##ENUMERATION_DESCRIPTION PrintJobStatuses are the various different statuses that it is possible for _
a job on a print queue to be in.
Public Enum PrintJobStatuses
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_PAUSED The print job has been paused during printing
    JOB_STATUS_PAUSED = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_ERROR The print job is stalled by an error
    JOB_STATUS_ERROR = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_DELETING The print job is being deleted from the queue
    JOB_STATUS_DELETING = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_SPOOLING The print job is being added to the spooler queue
    JOB_STATUS_SPOOLING = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_PRINTING The print job is in the midst of being printed
    JOB_STATUS_PRINTING = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_OFFLINE The print job is stalled because the printer is off line
    JOB_STATUS_OFFLINE = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_PAPEROUT The print job is stalled because the printer has run out of paper
    JOB_STATUS_PAPEROUT = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_PRINTER The print job is waiting on the printer
    JOB_STATUS_PRINTED = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_DELETED The print job has been deleted
    JOB_STATUS_DELETED = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_BLOCKED_DEVICEQUEUE The device queue is blocked so the print job is also blocked
    JOB_STATUS_BLOCKED_DEVICEQUEUE = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_USER_INTERVENTION The queued job is paused by user intervention
    JOB_STATUS_USER_INTERVENTION = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_STATUS_RESTART The queued print job is restarting after being paused
    JOB_STATUS_RESTART = &H800
End Enum

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type JOB_INFO_1
    JobId As Long
    lpPrinterName As String
    lpMachinename As String
    lpUserName As String
    lpDocumentName As String
    lpDataType As String
    lpStatus As String
    Status As Long
    Priority As Long
    Position As Long
    TotalPages As Long
    PagesPrinted As Long
    Submitted As SYSTEMTIME
End Type

Private Type JOB_INFO_2
    JobId As Long
    lpPrinterName As String
    lpMachinename As String
    lpUserName As String
    lpDocumentName As String
    lpNotifyName As String
    lpDataType As String
    lpPrintProcessor As String
    lpParameters As String
    lpDriverName As String
    lpDevMode As Long 'Pointer to DevMode
    lpStatus As String
    lpSecurityDescriptor As Long 'Pointer to SECURITY_DESCRIPTOR
    Status As Long
    Priority As Long
    Position As Long
    StartTime As Long
    UntilTime As Long
    TotalPages As Long
    JobSize As Long
    Submitted As SYSTEMTIME
    time As Long
    PagesPrinted As Long
End Type

'\\ Included in WINNT4 and beyond
Private Type JOB_INFO_3
    JobId As Long
    NextJobId As Long
    Reserved As Long 'must be set to zero
End Type

Private Declare Function GetJob Lib "winspool.drv" Alias "GetJobA" _
                          (ByVal hPrinter As Long, _
                           ByVal JobId As Long, _
                           ByVal Level As Long, _
                           buffer As Long, _
                           ByVal pbSize As Long, _
                           pbSizeNeeded As Long) As Long
                           
Private Declare Function SetJob Lib "winspool.drv" Alias _
                     "SetJobA" (ByVal hPrinter As Long, _
                                ByVal JobId As Long, _
                                ByVal Level As Long, _
                                pJob As Long, _
                                ByVal Command As Long) As Long


'\\ Private member variables
Private mJobId As Long 'Unique id of this job
Private mhPrinter As Long 'Printer this job is on...

Private mJOB_INFO_1 As JOB_INFO_1
Private mJOB_INFO_2 As JOB_INFO_2
Private mJOB_INFO_3 As JOB_INFO_3

'\\ Private variable to prevent unneccessary calls to the dll...
Private mLastApiCall(1 To 3) As Date

Private bChanged As Boolean

' ##ENUMERATION_DESCRIPTION PrintJobControlCommands are the different commands that can be sent to the _
print spooler for a specific print job
Public Enum PrintJobControlCommands
    ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_PAUSE Pause printing of the current job
     JOB_CONTROL_PAUSE = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_RESUME Resume printing a paused job
     JOB_CONTROL_RESUME = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_CANCEL Cancel a queued or printing job
     JOB_CONTROL_CANCEL = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_RESTART Restart a paused job that was printing
     JOB_CONTROL_RESTART = 4
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_DELETE Remove a job from the queue
     JOB_CONTROL_DELETE = 5
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_SENT_TO_PRINTER End a print job
     JOB_CONTROL_SENT_TO_PRINTER = 6
     ' ##ENUMERATION_MEMBER_DESCRIPTION JOB_CONTROL_LAST_PAGE_EJECTED End a print job
     JOB_CONTROL_LAST_PAGE_EJECTED = 7
     ' ##ENUMERATION_MEMBER_DESCRIPTION
End Enum

'\\ When changing information other than the position for JOB_INFO_1 and
'\\ JOB_INFO_2 you must set the Position member to JOB_POSITION_UNSPECIFIED
Private Const JOB_POSITION_UNSPECIFIED = 0

Private buffer() As Long
Private WithEvents dmThis As ApiDeviceMode
Attribute dmThis.VB_VarHelpID = -1

Private Sub dmThis_Changed()

Call RefreshJobInfo(2)
If dmThis.CopyToPointer(mJOB_INFO_2.lpDevMode) Then
    Call SavePrintJobInfo(2)
End If

End Sub
Public Property Get DeviceMode() As ApiDeviceMode
Attribute DeviceMode.VB_HelpID = 3107

Set dmThis = New ApiDeviceMode

Call RefreshJobInfo(2)
If dmThis.CreateFromPointer(mJOB_INFO_2.lpDevMode) Then
    Set DeviceMode = dmThis
End If

End Property
'  ##BLOCK_DESCRIPTION Cancels (and deletes) the selected print job
'  ##BLOCK_REMARKS You should check the status of the print job is valid for cancelation.
Public Sub CancelJob()
Attribute CancelJob.VB_Description = "Cancels (and deletes) the selected print job"
Attribute CancelJob.VB_HelpID = 2684

Dim lRet As Long

If APIDispenser.System.WindowsOSVersion < ver_win_nt4 Then
    lRet = SetJob(mhPrinter, mJobId, 0, 0, JOB_CONTROL_CANCEL)
Else
    lRet = SetJob(mhPrinter, mJobId, 0, 0, JOB_CONTROL_DELETE)
End If
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CancelJob", GetLastSystemError
End If

End Sub

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get DataType() As String
Attribute DataType.VB_HelpID = 2685

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    DataType = mJOB_INFO_1.lpDataType
End If

End Property

'\\ --[DocumentName]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the document being printed.
'  ##BLOCK_REMARKS Typically this is the filename of the document, although it can be whatever the application that spools the job decides to call it.
'  ##BLOCK_REMARKS Therefore you should not assume this is an actual existing file's na
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get DocumentName() As String
Attribute DocumentName.VB_Description = "This is the name of the document being printed."
Attribute DocumentName.VB_HelpID = 2686

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    DocumentName = mJOB_INFO_1.lpDocumentName
End If

End Property


Friend Property Let JobId(ByVal newJobId As Long)

    mJobId = newJobId
    
End Property

Public Property Get JobId() As Long
Attribute JobId.VB_HelpID = 2687

    JobId = mJobId
    
End Property

'  ##BLOCK_DESCRIPTION This links the two jobs together as one job
'  ##BLOCK_REMARKS This requires NT4 or higher operating system version
'  ##BLOCK_REMARKS You must have administer level rights on both jobs concerned
Public Sub LinkToJob(ByVal newJob As ApiPrintJob)
Attribute LinkToJob.VB_Description = "This links the two jobs together as one job"
Attribute LinkToJob.VB_HelpID = 3108

If newJob.JobId <> JobId Then
    With mJOB_INFO_3
        .JobId = JobId
        .NextJobId = newJob.JobId
        .Reserved = 0
    End With
    Call SavePrintJobInfo(3)
End If

End Sub

'\\ --[MachineName]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the machine from which the job was added to the printer queue.
'  ##BLOCK_REMARKS If the job was spooled from this machine then the name will match the %computer name:EventVB~ApiSystem~MachineName%
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get MachineName() As String
Attribute MachineName.VB_Description = "This is the name of the machine from which the job was added to the printer queue."
Attribute MachineName.VB_HelpID = 2688

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    MachineName = mJOB_INFO_1.lpMachinename
End If

End Property

'  ##BLOCK_DESCRIPTION This is the user that will be notified if the job status changes
'  ##BLOCK_REMARKS You should set this to a user name known to the domain on which the printer resides
Public Property Get NotifyUser() As String
Attribute NotifyUser.VB_Description = "This is the user that will be notified if the job status changes"
Attribute NotifyUser.VB_HelpID = 3109

Call RefreshJobInfo(2)
NotifyUser = mJOB_INFO_2.lpNotifyName

End Property

'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the number of pages of the job that have been printed thus far.
'  #---------------------------------------------------------------------------------------------------------
Public Property Get PagesPrinted() As Long
Attribute PagesPrinted.VB_Description = "This is the number of pages of the job that have been printed thus far."
Attribute PagesPrinted.VB_HelpID = 2689

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    PagesPrinted = mJOB_INFO_1.PagesPrinted
End If

End Property

'  ##BLOCK_DESCRIPTION Pauses the selected print job
'  ##BLOCK_REMARKS You should check the status of the print job is valid for pausing.
Public Sub PauseJob()
Attribute PauseJob.VB_Description = "Pauses the selected print job"
Attribute PauseJob.VB_HelpID = 2690

Dim lRet As Long

lRet = SetJob(mhPrinter, mJobId, 0, 0, JOB_CONTROL_PAUSE)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":PauseJob", GetLastSystemError
End If

End Sub


'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the print jobs position in the print queue.
'  ##BLOCK_REMARKS This may be zero if the job is deleted.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Position() As Long
Attribute Position.VB_Description = "This is the print jobs position in the print queue."
Attribute Position.VB_HelpID = 2691

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    Position = mJOB_INFO_1.Position
End If

End Property

Friend Property Let PrinterHandle(ByVal newhPrinter As Long)

    mhPrinter = newhPrinter
    
End Property
'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the %printer:EventVB~ApiPrinter% that this job is queued against
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get PrinterName() As String
Attribute PrinterName.VB_Description = "This is the name of the printer:EventVB~ApiPrinter that this job is queued against"
Attribute PrinterName.VB_HelpID = 2692

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    PrinterName = mJOB_INFO_1.lpPrinterName
End If

End Property

'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the priority of the print job.  Jobs with higher priority get printed first.
'  ##BLOCK_REMARKS This will initially be set to the %printers default priority:EventVB~ApiPrinter~DefaultPriority%
'  ##BLOCK_REMARKS This value may be changed by a process or user with sufficient access rights.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Priority() As Long
Attribute Priority.VB_Description = "This is the priority of the print job.  Jobs with higher priority get printed first."
Attribute Priority.VB_HelpID = 2693

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    Priority = mJOB_INFO_1.Priority
End If

End Property

Private Sub RefreshJobInfo(ByVal Index As Integer)

Dim lRet As Long
Dim SizeNeeded As Long


'\\ Prevent unneccesary call to the API
If DateDiff("s", mLastApiCall(Index), Now) > 1 Then
    mLastApiCall(Index) = Now
        
    ReDim Preserve buffer(0 To 1) As Long
    lRet = GetJob(mhPrinter, mJobId, Index, buffer(0), UBound(buffer), SizeNeeded)
    
    If SizeNeeded > 0 Then
        ReDim Preserve buffer(0 To (SizeNeeded / 4) + 3) As Long
        lRet = GetJob(mhPrinter, mJobId, Index, buffer(0), UBound(buffer) * 4, SizeNeeded)
        
        If Index < 1 Or Index > 3 Then
            Debug.Print "Error in call to ApiPrintJob:RefreshJobInfo - invalid index"
        Else
            Select Case Index
            Case 1
                With mJOB_INFO_1
                    .JobId = buffer(0)
                    .lpPrinterName = StringFromPointer(buffer(1), 1024)
                    .lpMachinename = StringFromPointer(buffer(2), 1024)
                    .lpUserName = StringFromPointer(buffer(3), 1024)
                    .lpDocumentName = StringFromPointer(buffer(4), 1024)
                    .lpDataType = StringFromPointer(buffer(5), 1024)
                    .lpStatus = StringFromPointer(buffer(6), 1024)
                    .Status = buffer(7)
                    .Priority = buffer(8)
                    .Position = buffer(9)
                    .TotalPages = buffer(10)
                    .PagesPrinted = buffer(11)
                    'Submitted is also here...
                    With .Submitted
                        .wYear = LoWord(buffer(12))
                        .wMonth = HiWord(buffer(12))
                        .wDayOfWeek = LoWord(buffer(13))
                        .wDay = HiWord(buffer(13))
                        .wHour = LoWord(buffer(14))
                        .wMinute = HiWord(buffer(14))
                        .wSecond = LoWord(buffer(15))
                        .wMilliseconds = HiWord(buffer(15))
                    End With
                End With
            Case 2
                With mJOB_INFO_2
                    .JobId = buffer(0)
                    .lpPrinterName = StringFromPointer(buffer(1), 1024)
                    .lpMachinename = StringFromPointer(buffer(2), 1024)
                    .lpUserName = StringFromPointer(buffer(3), 1024)
                    .lpDocumentName = StringFromPointer(buffer(4), 1024)
                    .lpNotifyName = StringFromPointer(buffer(5), 1024)
                    .lpDataType = StringFromPointer(buffer(6), 1024)
                    .lpPrintProcessor = StringFromPointer(buffer(7), 1024)
                    .lpParameters = StringFromPointer(buffer(8), 1024)
                    .lpDriverName = StringFromPointer(buffer(9), 1024)
                    .lpDevMode = buffer(10) 'To do: Replace with ApiDevMode class when available
                    .lpStatus = StringFromPointer(buffer(11), 1024)
                    .lpSecurityDescriptor = buffer(12) 'To do: Replace...
                    .Status = buffer(13)
                    .Priority = buffer(14)
                    .Position = buffer(15)
                    .StartTime = buffer(16)
                    .UntilTime = buffer(17)
                    .TotalPages = buffer(18)
                    .JobSize = buffer(19)
                    'Submitted is here....
                    '4 long values, packed with 8 integers
                    With .Submitted
                        .wYear = LoWord(buffer(20))
                        .wMonth = HiWord(buffer(20))
                        .wDayOfWeek = LoWord(buffer(21))
                        .wDay = HiWord(buffer(21))
                        .wHour = LoWord(buffer(22))
                        .wMinute = HiWord(buffer(22))
                        .wSecond = LoWord(buffer(23))
                        .wMilliseconds = HiWord(buffer(23))
                    End With
                    .time = buffer(24)
                    .PagesPrinted = buffer(25)
                End With
            Case 3
                With mJOB_INFO_3
                    .JobId = buffer(0)
                    .NextJobId = buffer(1)
                    .Reserved = buffer(2)
                End With
            End Select
        End If
    End If
End If

End Sub

'  ##BLOCK_DESCRIPTION Resumes the selected print job if it was paused
'  ##BLOCK_REMARKS You should check the status of the print job is valid for resuming.
Public Sub ResumeJob()
Attribute ResumeJob.VB_Description = "Resumes the selected print job if it was paused"
Attribute ResumeJob.VB_HelpID = 2694

Dim lRet As Long

lRet = SetJob(mhPrinter, mJobId, 0, 0, JOB_CONTROL_RESUME)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":ResumeJob", GetLastSystemError
End If

End Sub

Private Sub SavePrintJobInfo(ByVal Level As Long)

Dim lRet As Long

Select Case Level
Case 1
    '\\ Overwrite existing buffer data with that held locally in JOB_INFO_1
    With mJOB_INFO_1
        buffer(0) = .JobId '\\ This cannot change and is ignored
        buffer(1) = LPCSTR(.lpPrinterName)  '\\ This is ignored and cannot be changed
        buffer(2) = LPCSTR(.lpMachinename)  '\\ also ignored
        buffer(3) = LPCSTR(.lpUserName)
        buffer(4) = LPCSTR(.lpDocumentName)
        buffer(5) = LPCSTR(.lpDataType)
        buffer(6) = LPCSTR(.lpStatus)
        buffer(7) = .Status
        buffer(8) = .Priority
        buffer(9) = .Position
        buffer(10) = .TotalPages
        buffer(11) = .PagesPrinted
        With .Submitted '\\ Also ignored
            buffer(12) = MakeLong(.wYear, .wMonth)
            buffer(13) = MakeLong(.wDayOfWeek, .wDay)
            buffer(14) = MakeLong(.wHour, .wMinute)
            buffer(15) = MakeLong(.wSecond, .wMilliseconds)
        End With
    End With

Case 2
    '\\ Overwrite existing buffer data with that held locally in JOB_INFO_2
    With mJOB_INFO_2
        buffer(0) = .JobId '\\ This cannot change and is ignored
        buffer(1) = LPCSTR(.lpPrinterName)  '\\ This is ignored and cannot be changed
        buffer(2) = LPCSTR(.lpMachinename)  '\\ also ignored
        buffer(3) = LPCSTR(.lpUserName)
        buffer(4) = LPCSTR(.lpDocumentName)
        buffer(5) = LPCSTR(.lpNotifyName)
        buffer(6) = LPCSTR(.lpDataType)
        buffer(7) = LPCSTR(.lpPrintProcessor)
        buffer(8) = LPCSTR(.lpParameters)
        buffer(9) = LPCSTR(.lpDriverName)
        buffer(10) = .lpDevMode
        buffer(11) = LPCSTR(.lpStatus)
        buffer(12) = .lpSecurityDescriptor
        buffer(13) = .Status
        buffer(14) = .Priority
        buffer(15) = .Position
        buffer(16) = .StartTime
        buffer(17) = .UntilTime
        buffer(18) = .TotalPages
        buffer(19) = .JobSize
        With .Submitted '\\ Also ignored
            buffer(20) = MakeLong(.wYear, .wMonth)
            buffer(21) = MakeLong(.wDayOfWeek, .wDay)
            buffer(22) = MakeLong(.wHour, .wMinute)
            buffer(23) = MakeLong(.wSecond, .wMilliseconds)
        End With
        buffer(24) = .time
        buffer(25) = .PagesPrinted
    End With
    
Case 3
    '\\ Overwrite existing buffer data with that held locally in JOB_INFO_3
    With mJOB_INFO_3
        buffer(0) = .JobId
        buffer(1) = .NextJobId
        buffer(2) = .Reserved
    End With
    
End Select

lRet = SetJob(mhPrinter, mJobId, Level, buffer(0), 0)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":SavePrintJobInfo", GetLastSystemError
Else
    '\\ Free up "long pointer to string" strings...
    Call ResetLPCSTRIndexes
End If

End Sub

'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This returns the current status of the print job.
'  ##BLOCK_REMARKS This is more reliable than using the %printer status:EventVB~ApiPrinter~Status%
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Status() As Long
Attribute Status.VB_Description = "This returns the current status of the print job."
Attribute Status.VB_HelpID = 2695

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    Status = mJOB_INFO_1.Status
End If

End Property


'  ##BLOCK_DESCRIPTION This job has been deleted
Public Property Get Deleted() As Boolean
        Deleted = ((Status And JOB_STATUS_DELETED) = JOB_STATUS_DELETED)
End Property

'  ##BLOCK_DESCRIPTION This job is being deleted
Public Property Get Deleting() As Boolean
        Deleting = ((Status And JOB_STATUS_DELETING) = JOB_STATUS_DELETING)
End Property

'  ##BLOCK_DESCRIPTION This job has been printed
Public Property Get Printed() As Boolean
    Printed = ((Status And JOB_STATUS_PRINTED) = JOB_STATUS_PRINTED)
End Property

'  ##BLOCK_DESCRIPTION This job is being printed
Public Property Get Printing() As Boolean
    
        Printing = ((Status And JOB_STATUS_PRINTING) = JOB_STATUS_PRINTING)

End Property

'  ##BLOCK_DESCRIPTION This job has an error
Public Property Get InError() As Boolean

    InError = ((Status And JOB_STATUS_ERROR) = JOB_STATUS_ERROR)

End Property

'  ##BLOCK_DESCRIPTION This job is not printing because the printer is off line
Public Property Get Offline() As Boolean
    
    Offline = ((Status And JOB_STATUS_OFFLINE) = JOB_STATUS_OFFLINE)
  
End Property

'  ##BLOCK_DESCRIPTION This job is not printing because the printer is out of paper
Public Property Get PaperOut() As Boolean
    
    PaperOut = ((Status And JOB_STATUS_PAPEROUT) = JOB_STATUS_PAPEROUT)

End Property

'  ##BLOCK_DESCRIPTION This job is not printing because it requires user intervention
Public Property Get UserInterventionRequired() As Boolean

    UserInterventionRequired = ((Status And JOB_STATUS_USER_INTERVENTION) = JOB_STATUS_USER_INTERVENTION)

End Property
'\\ --[Submitted]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This returns the time that the job was sent for printing.
'  ##BLOCK_DESCRIPTION The time is relative to the system time of the machine that issued the print job.
'  ##BLOCK_DESCRIPTION In wide area networks this may not be the same as the time on the server.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Submitted() As APISystemTime
Attribute Submitted.VB_Description = "This returns the time that the job was sent for printing."
Attribute Submitted.VB_HelpID = 2696

Dim oSystemTime As APISystemTime

Call RefreshJobInfo(1)

Set oSystemTime = New APISystemTime
oSystemTime.CreateFromPointer (VarPtr(mJOB_INFO_1.Submitted))

Set Submitted = oSystemTime

End Property

'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This returns the total number of pages in the print job.
'  ##BLOCK_REMARKS This setting will not be correct until the job has finished spooling as pages may be added until that point.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get TotalPages() As Long
Attribute TotalPages.VB_Description = "This returns the total number of pages in the print job."
Attribute TotalPages.VB_HelpID = 2697

If mhPrinter <> 0 And mJobId <> 0 Then
    Call RefreshJobInfo(1)
    TotalPages = mJOB_INFO_1.TotalPages
End If

End Property

'\\ --[StartTime]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This returns the time that the job started being processed.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get StartTime() As Long
Attribute StartTime.VB_Description = "This returns the time that the job started being processed."
Attribute StartTime.VB_HelpID = 2698

Call RefreshJobInfo(2)
StartTime = mJOB_INFO_2.StartTime

End Property

Public Property Get UntilTime() As Long
Attribute UntilTime.VB_HelpID = 2699

Call RefreshJobInfo(2)
UntilTime = mJOB_INFO_2.UntilTime

End Property


Public Property Let Username(ByVal newname As String)
Attribute Username.VB_Description = "This returns the name of the user whos account submitted the job to the printer.  "
Attribute Username.VB_HelpID = 3110

Call RefreshJobInfo(1)
mJOB_INFO_1.lpUserName = newname
mJOB_INFO_1.Position = JOB_POSITION_UNSPECIFIED
Call SavePrintJobInfo(1)

End Property

'  ##BLOCK_DESCRIPTION This returns the name of the user whos account submitted the job to the printer.
'  ##BLOCK_REMARKS The user name will be the login name for a Windows network.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Username() As String

    Call RefreshJobInfo(1)
    Username = mJOB_INFO_1.lpUserName

End Property



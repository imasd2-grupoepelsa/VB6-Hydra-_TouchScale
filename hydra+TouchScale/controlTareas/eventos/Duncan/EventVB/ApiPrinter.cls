VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2619
Attribute VB_Description = "This class provides the properties and methods for a getting information about and controlling a printer or print queue."
Option Explicit

Private Const CLSNAME = "ApiPrinter"

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for a getting information about and controlling a printer or print queue.

' ##MODULE_DESCRIPTION You can use this to get information about the status _
of any of the printers installed and (subject to having access rights) to _
control that printer.

' ##ENUMERATION_DESCRIPTION Printer_Control_Commands are sent to the printer to perform actions _
that control the current printer's job(s)
Public Enum Printer_Control_Commands
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_CONTROL_PAUSE Pause the printing of the currently active job
     PRINTER_CONTROL_PAUSE = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_CONTROL_RESUME Resume printing if paused
     PRINTER_CONTROL_RESUME = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_CONTROL_PURGE Terminate and delete the currently printing job
     PRINTER_CONTROL_PURGE = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_CONTROL_SET_STATUS Set the printer status for the current job
     PRINTER_CONTROL_SET_STATUS = 4
End Enum

' ##ENUMERATION_DESCRIPTION Printer_Status are the various possible states a printer can be in. Note _
that this information is provided by the printer driver and in many cases the status will always be _
"Ready" unless there is one or more jobs in error on that printer.
Public Enum Printer_Status
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_READY The printer is free and ready to print
    PRINTER_STATUS_READY = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PAUSED Printing is paused
    PRINTER_STATUS_PAUSED = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_ERROR Printing is suspended due to a general error
    PRINTER_STATUS_ERROR = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PENDING_DELETION Printing is suspended while a job is being deleted
    PRINTER_STATUS_PENDING_DELETION = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PAPER_JAM Printing is suspended because there has been a paper jam
    PRINTER_STATUS_PAPER_JAM = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PAPER_OUT Printing is suspended because the printer has run out of paper
    PRINTER_STATUS_PAPER_OUT = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_MANUAL_FEED Printing is suspended pending a manual paper feed
    PRINTER_STATUS_MANUAL_FEED = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PAPER_PROBLEM Printing is supended due to a paper problem
    PRINTER_STATUS_PAPER_PROBLEM = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_OFFLINE The printer is off line
    PRINTER_STATUS_OFFLINE = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_IO_ACTIVE The printer is reading data from the IO port
    PRINTER_STATUS_IO_ACTIVE = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_BUSY The printer is busy
    PRINTER_STATUS_BUSY = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PRINTING The printer is active and printing a job
    PRINTER_STATUS_PRINTING = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_OUTPUT_BIN_FULL The output tray is full and the printer has paused until it is emptied
    PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_NOT_AVAILABLE The printer status is not known
    PRINTER_STATUS_NOT_AVAILABLE = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_WAITING The printer is waiting for a job
    PRINTER_STATUS_WAITING = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PROCESSING The printer is active and processing a print job
    PRINTER_STATUS_PROCESSING = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_INITIALIZING The printer is initialising and not yet ready to print
    PRINTER_STATUS_INITIALIZING = &H8000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_WARMING_UP The printer is warming up and not yet ready to print
    PRINTER_STATUS_WARMING_UP = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_TONER_LOW Printing is suspended because the level of toner or ink is too low for a reasonable quality print
    PRINTER_STATUS_TONER_LOW = &H20000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_NO_TONER Printing has been suspended because the printer is out of toner or ink
    PRINTER_STATUS_NO_TONER = &H40000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_PAGE_PUNT (Win95 only) The printer is suspended while a page is deleted
    PRINTER_STATUS_PAGE_PUNT = &H80000 'win95
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_USER_INTERVENTION The printer has been paused by the user intervention
    PRINTER_STATUS_USER_INTERVENTION = &H100000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_OUT_OF_MEMORY Printing is suspended because the printer has run out of memory
    PRINTER_STATUS_OUT_OF_MEMORY = &H200000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_DOOR_OPEN Printing is supended because a door on the printer is open
    PRINTER_STATUS_DOOR_OPEN = &H400000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_SERVER_UNKNOWN Printing is suspended due to an unknown server error
    PRINTER_STATUS_SERVER_UNKNOWN = &H800000
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_STATUS_POWER_SAVE The printer is suspended in power saving mode
    PRINTER_STATUS_POWER_SAVE = &H1000000
End Enum

' ##ENUMERATION_DESCRIPTION
Public Enum Printer_Attributes
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_QUEUED
     PRINTER_ATTRIBUTE_QUEUED = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_DIRECT
     PRINTER_ATTRIBUTE_DIRECT = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_DEFAULT
     PRINTER_ATTRIBUTE_DEFAULT = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_SHARED
     PRINTER_ATTRIBUTE_SHARED = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_NETWORK
     PRINTER_ATTRIBUTE_NETWORK = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_HIDDEN
     PRINTER_ATTRIBUTE_HIDDEN = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_LOCAL
     PRINTER_ATTRIBUTE_LOCAL = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_ENABLE_DEVQ
     PRINTER_ATTRIBUTE_ENABLE_DEVQ = &H80
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS
     PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST
     PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST = &H200
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_WORK_OFFLINE
     PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_ENABLE_BIDI The printer can operate in bidirectional mode
     PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_RAW_ONLY The printer can only accept raw data
     PRINTER_ATTRIBUTE_RAW_ONLY = &H1000
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ATTRIBUTE_PUBLISHED The printer is shared as a network resource
     PRINTER_ATTRIBUTE_PUBLISHED = &H2000
End Enum

Private Type PRINTER_INFO_1
    flags As Long
    pDescription As String
    pName As String
    pComment As String
End Type

Private Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As Long 'LPDEVMODEA
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As Long 'PSECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    JobsCount As Long
    AveragePPM As Long
End Type


Private Type PRINTER_INFO_3
    pSecurityDescriptor As Long 'PSECURITY_DESCRIPTOR
End Type

Private Type PRINTER_INFO_4
    pPrinterName As String
    pServerName As String
    Attributes As Long
End Type

Private Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

Private Type PRINTER_INFO_6
    dwStatus As Long
End Type

Private Type PRINTER_INFO_7
    pszObjectGUID As String
    dwAction As Long
End Type

Private Type PRINTER_INFO_8
    pDevMode As Long 'DEVMODE
End Type

'api calls...

Private Declare Function SetPrinterApi Lib "winspool.drv" Alias _
                   "SetPrinterA" (ByVal hPrinter As Long, _
                       ByVal Level As Long, _
                       buffer As Long, _
                      ByVal Command As Long) As Long
                                                                                     
                                                                                     
    
'Aliased api calls....
Private Declare Function GetPrinterApi Lib "winspool.drv" Alias _
                           "GetPrinterA" (ByVal hPrinter As Long, _
                                            ByVal Level As Long, _
                                            buffer As Long, _
                                            ByVal pbSize As Long, _
                                            pbSizeNeeded As Long) As Long



Private Type PRINTER_DEFAULTS
  pDatatype As String
  pDevMode As Long
  DesiredAccess As Long
End Type

Private Declare Function OpenPrinter Lib "winspool.drv" _
        Alias "OpenPrinterA" (ByVal pPrinterName As String, _
        phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
        (ByVal hPrinter As Long) As Long


'\\ Printer handle....
Private mhPrinter As Long
Private msDeviceName As String

'\\ Private member variables for simplicity sake....
Private mPRINTER_INFO_1 As PRINTER_INFO_1
Private mPRINTER_INFO_2 As PRINTER_INFO_2
Private mPRINTER_INFO_3 As PRINTER_INFO_3
Private mPRINTER_INFO_4 As PRINTER_INFO_4
Private mPRINTER_INFO_5 As PRINTER_INFO_5
Private mPRINTER_INFO_6 As PRINTER_INFO_6
Private mPRINTER_INFO_7 As PRINTER_INFO_7
Private mPRINTER_INFO_8 As PRINTER_INFO_8
Private mPRINTER_INFO_9 As PRINTER_INFO_8 'currently this structure is identical 9 and 8

'\\ Declaration to enumerate the active jobs on this printer
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
                          (ByVal hPrinter As Long, _
                           ByVal FirstJob As Long, _
                           ByVal NoJobs As Long, _
                           ByVal Level As Long, _
                           pJob As Long, _
                           ByVal cdBuff As Long, _
                           pcbNeeded As Long, _
                           pcBytesReturned As Long) As Long
                           
'\\ To get a printer's device context....
Private Declare Function CreateDCByLong Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long

Private bChanged As Boolean

'\\ Global memory managemennt routines...
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As Long, ByVal Length As Long)
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cbSize As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvLocation As Long)

' ##ENUMERATION_DESCRIPTION PrinterAccessRights describe the possible settings for the desired access _
to the printer.  These are used to request a specific access to a printer.  This access _
will depend on the user or process access rights
Public Enum PrinterAccessRights
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ACCESS_ADMINISTER Allows administration of a printer
     PRINTER_ACCESS_ADMINISTER = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ACCESS_USE Allows printer general use (printing, querying)
     PRINTER_ACCESS_USE = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTER_ALL_ACCESS Allows use and administration.
     PRINTER_ALL_ACCESS = &HF000C
End Enum

Private pDef As PRINTER_DEFAULTS
Private buffer() As Long

' ##ENUMERATION_DESCRIPTION PrintDeviceCapabilitiesIndexes are passed to the DeviceCapabilities API to get information _
about aspects of a printer devices capabilities and setup
Public Enum PrintDeviceCapabilitiesIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_FIELDS Which fields of the device mode are used
    DC_FIELDS = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PAPERS Which Printer Paper Sizes the device supports
    DC_PAPERS = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PAPERSIZE The dimensions of the paper in 10ths of a millimeter
    DC_PAPERSIZE = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_MINEXTENT The minimum paper width and height the printer can support
    DC_MINEXTENT = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_MAXEXTENT The maximum paper width and height the printer can support
    DC_MAXEXTENT = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_BINS The standard paper bins supported by this printer
    DC_BINS = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_DUPLEX Whether the printer supports duplex printing
    DC_DUPLEX = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_SIZE
    DC_SIZE = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_EXTRA
    DC_EXTRA = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_VERSION
    DC_VERSION = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_DRIVER
    DC_DRIVER = 11
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_BINNAMES
    DC_BINNAMES = 12
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_ENUMRESOLUTIONS
    DC_ENUMRESOLUTIONS = 13
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_FILEDEPENDENCIES
    DC_FILEDEPENDENCIES = 14
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_TRUETYPE
    DC_TRUETYPE = 15
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PAPERNAMES
    DC_PAPERNAMES = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_ORIENTATION
    DC_ORIENTATION = 17
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_COPIES
    DC_COPIES = 18
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_BINADJUST
    DC_BINADJUST = 19
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_EMF_COMPLIANT
    DC_EMF_COMPLIANT = 20
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_DATATYPE_PRODUCED
    DC_DATATYPE_PRODUCED = 21
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_COLLATE
    DC_COLLATE = 22
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_MANUFACTURER
    DC_MANUFACTURER = 23
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_MODEL
    DC_MODEL = 24
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PERSONALITY
    DC_PERSONALITY = 25
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PRINTRATE
    DC_PRINTRATE = 26
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PRINTRATEUNIT
    DC_PRINTRATEUNIT = 27
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PRINTERMEM
    DC_PRINTERMEM = 28
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_MEDIAREADY
    DC_MEDIAREADY = 29
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_STAPLE
    DC_STAPLE = 30
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_PRINTRATEPPM
    DC_PRINTRATEPPM = 31
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_COLORDEVICE
    DC_COLORDEVICE = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION DC_NUP
    DC_NUP = 33
End Enum

' ##ENUMERATION_DESCRIPTION PrintRateUnits are the different ways of measuring the speed of a printer
Public Enum PrintRateUnits
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTRATEUNIT_PPM Speed in pages per minute
    PRINTRATEUNIT_PPM = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTRATEUNIT_CPS Speed in characters per second
    PRINTRATEUNIT_CPS = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTRATEUNIT_LPM Speed in lines per minute
    PRINTRATEUNIT_LPM = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRINTRATEUNIT_IPM Speed in inches per minute
    PRINTRATEUNIT_IPM = 4
End Enum


' ##ENUMERATION_DESCRIPTION PrinterBins are the standard printer paper sources
Public Enum PrinterBins
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_UPPER The upper (or only) tray
     DMBIN_UPPER = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_LOWER The lower paper tray
     DMBIN_LOWER = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_MIDDLE The middle paper tray
     DMBIN_MIDDLE = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_MANUAL The manual form feed
     DMBIN_MANUAL = 4
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_ENVELOPE The envelope tray
     DMBIN_ENVELOPE = 5
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_ENVMANUAL Manual envelope feed
     DMBIN_ENVMANUAL = 6
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_AUTO The printer automatically selects the source
     DMBIN_AUTO = 7
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_TRACTOR The tractor feed paper source
     DMBIN_TRACTOR = 8
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_SMALLFMT The smaller format paper tray
     DMBIN_SMALLFMT = 9
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_LARGEFMT The larger format paper tray
     DMBIN_LARGEFMT = 10
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_LARGECAPACITY The high capacity tray
     DMBIN_LARGECAPACITY = 11
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_CASSETTE The paper cassete source
     DMBIN_CASSETTE = 14
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_FORMSOURCE The preprinted from feed
     DMBIN_FORMSOURCE = 15
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMBIN_USER Paper source is device specific
     DMBIN_USER = 256
End Enum

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal DeviceName As String, ByVal Portname As String, ByVal Index As PrintDeviceCapabilitiesIndexes, ByVal lpBuf As Long, pDevMode As Long) As Long

Private mDeviceCapabilitiesRet As Long
Private mDevOutput() As Byte
Private WithEvents dmThis As ApiDeviceMode
Attribute dmThis.VB_VarHelpID = -1

'  ##BLOCK_DESCRIPTION The printer is ready
Public Property Get IsReady() As Boolean
        IsReady = (Status = 0)
End Property

'  ##BLOCK_DESCRIPTION This printer has a door (or tray) open
Public Property Get IsDoorOpen() As Boolean
    IsDoorOpen = ((Status And PRINTER_STATUS_DOOR_OPEN) = PRINTER_STATUS_DOOR_OPEN)
End Property

'  ##BLOCK_DESCRIPTION The printer has an error
Public Property Get IsInError() As Boolean
        IsInError = ((Status And PRINTER_STATUS_ERROR) = PRINTER_STATUS_ERROR)
End Property

'  ##BLOCK_DESCRIPTION This printer is initialising itself
Public Property Get IsInitialising() As Boolean
        IsInitialising = ((Status And PRINTER_STATUS_INITIALIZING) = PRINTER_STATUS_INITIALIZING)
End Property

'  ##BLOCK_DESCRIPTION This printer is paused waiting for a paper feed
Public Property Get IsAwaitingManualFeed() As Boolean
        IsAwaitingManualFeed = ((Status And PRINTER_STATUS_MANUAL_FEED) = PRINTER_STATUS_MANUAL_FEED)
End Property

'  ##BLOCK_DESCRIPTION This printer is out of toner (or ink)
Public Property Get IsOutOfToner() As Boolean
        IsOutOfToner = ((mPRINTER_INFO_2.Status And PRINTER_STATUS_NO_TONER) = PRINTER_STATUS_NO_TONER)
End Property

'  ##BLOCK_DESCRIPTION The printer is unavailable
Public Property Get IsUnavailable() As Boolean
        IsUnavailable = ((Status And PRINTER_STATUS_NOT_AVAILABLE) = PRINTER_STATUS_NOT_AVAILABLE)
End Property

'  ##BLOCK_DESCRIPTION This printer is currently offline
Public Property Get IsOffline() As Boolean
        IsOffline = ((Status And PRINTER_STATUS_OFFLINE) = PRINTER_STATUS_OFFLINE)
End Property

'  ##BLOCK_DESCRIPTION This printer is out of free memory
Public Property Get IsOutOfmemory() As Boolean
        IsOutOfmemory = ((Status And PRINTER_STATUS_OUT_OF_MEMORY) = PRINTER_STATUS_OUT_OF_MEMORY)
End Property

'  ##BLOCK_DESCRIPTION The printer output tray is full
Public Property Get IsOutputBinFull() As Boolean
        IsOutputBinFull = ((Status And PRINTER_STATUS_OUTPUT_BIN_FULL) = PRINTER_STATUS_OUTPUT_BIN_FULL)
End Property

'  ##BLOCK_DESCRIPTION The paper has a paper jam
Public Property Get IsPaperJammed() As Boolean
        IsPaperJammed = ((Status And PRINTER_STATUS_PAPER_JAM) = PRINTER_STATUS_PAPER_JAM)
End Property


'  ##BLOCK_DESCRIPTION This printer is out of paper
Public Property Get IsOutOfPaper() As Boolean
        IsOutOfPaper = ((Status And PRINTER_STATUS_PAPER_OUT) = PRINTER_STATUS_PAPER_OUT)
End Property


    Public Property Get Paused() As Boolean
            Paused = ((Status And PRINTER_STATUS_PAUSED) = PRINTER_STATUS_PAUSED)
    End Property



'  ##BLOCK_DESCRIPTION This printer is busy deleting a print job
Public Property Get IsDeletingJob() As Boolean
        IsDeletingJob = ((Status And PRINTER_STATUS_PENDING_DELETION) = PRINTER_STATUS_PENDING_DELETION)
End Property



'  ##BLOCK_DESCRIPTION This printer is in power save mode
Public Property Get IsInPowerSave() As Boolean
        IsInPowerSave = ((Status And PRINTER_STATUS_POWER_SAVE) = PRINTER_STATUS_POWER_SAVE)
End Property



'  ##BLOCK_DESCRIPTION The printer is printing
Public Property Get IsPrinting() As Boolean
        IsPrinting = ((Status And PRINTER_STATUS_PRINTING) = PRINTER_STATUS_PRINTING)
End Property



'  ##BLOCK_DESCRIPTION The printer is waiting for user intervention
Public Property Get IsWaitingOnUserIntervention() As Boolean
        IsWaitingOnUserIntervention = ((Status And PRINTER_STATUS_USER_INTERVENTION) = PRINTER_STATUS_USER_INTERVENTION)
End Property

'  ##BLOCK_DESCRIPTION The printer is warming up
Public Property Get IsWarmingUp() As Boolean
        IsWarmingUp = ((Status And PRINTER_STATUS_WARMING_UP) = PRINTER_STATUS_WARMING_UP)
End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Private Property Get Attributes() As Long

Call RefreshPrinterInfo(2)
Attributes = mPRINTER_INFO_2.Attributes

End Property

'\\ --[AveragePagesPerMinute]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION Average Pages Per Minute is an indication of the output speed of the selected printer
'  ##BLOCK_REMARKS This data is provided by the print spooler.
'  ##BLOCK_REMARKS It is not guaranteed to be supported by every printer driver.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get AveragePagesPerMinute() As Long
Attribute AveragePagesPerMinute.VB_Description = "Average Pages Per Minute is an indication of the output speed of the selected printer"
Attribute AveragePagesPerMinute.VB_HelpID = 2620

Call RefreshPrinterInfo(2)
AveragePagesPerMinute = mPRINTER_INFO_2.AveragePPM

End Property

Public Property Let Comment(ByVal newComment As String)
Attribute Comment.VB_Description = "The comment is an additional aid to identifying the printer  "
Attribute Comment.VB_HelpID = 2622

Dim lRet As Long

Call RefreshPrinterInfo(2)
mPRINTER_INFO_2.pComment = newComment
Call SavePrinterInfo(2)

End Property


'\\ --[Comment]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION The comment is an additional aid to identifying the printer
'  ##BLOCK_REMARKS This data is provided by the print spooler.
'  ##BLOCK_REMARKS The comment may be changed by a user with sufficient administratice rights on the print server.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Comment() As String

Call RefreshPrinterInfo(1)
Comment = mPRINTER_INFO_1.pComment

End Property

Private Property Get DataSize(ByVal Index As PrintDeviceCapabilitiesIndexes)

Select Case Index
    
End Select

End Property

Public Property Let DataType(ByVal newValue As String)
Attribute DataType.VB_HelpID = 2624

Call RefreshPrinterInfo(2)
mPRINTER_INFO_2.pDatatype = newValue
Call SavePrinterInfo(2)

End Property

Public Property Get DataType() As String

Call RefreshPrinterInfo(2)
DataType = mPRINTER_INFO_2.pDatatype

End Property


Public Property Let DefaultPriority(ByVal newSetting As Long)
Attribute DefaultPriority.VB_Description = "This is the default priority for new jobs spooled to this printer.  "
Attribute DefaultPriority.VB_HelpID = 2626

Call RefreshPrinterInfo(2)
mPRINTER_INFO_2.DefaultPriority = newSetting
Call SavePrinterInfo(2)

End Property

'\\ --[DefaultPriority]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the default priority for new jobs spooled to this printer.
'  ##BLOCK_REMARKS This data is provided by the print spooler.
'  ##BLOCK_REMARKS Individual jobs may have their priority changed from this default value.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get DefaultPriority() As Long

Call RefreshPrinterInfo(2)
DefaultPriority = mPRINTER_INFO_2.DefaultPriority

End Property



'\\ --[AveragePagesPerMinute]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is a longer description to help identify the selected printer
'  ##BLOCK_REMARKS This data is provided by the print spooler.
'  ##BLOCK_REMARKS This may be changed by any user with sufficient authority on the printer server.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Description() As String
Attribute Description.VB_Description = "This is a longer description to help identify the selected printer"
Attribute Description.VB_HelpID = 2627

Call RefreshPrinterInfo(1)
Description = mPRINTER_INFO_1.pDescription

End Property


'\\ --[DesiredAccess]-----------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION Sets the level of access desired to the selected printer
'  ##BLOCK_REMARKS Select PRINTER_ACCESS_ADMINISTER to adminster the printer
'  ##BLOCK_REMARKS Select PRINTER_ACCESS_USE to use the printer
'  ##BLOCK_REMARKS Select PRINTER_ALL_ACCESS to use and administer the printer
'\\ ------------------------------------------------------------------------------------------------------
Public Property Let DesiredAccess(ByVal newAccess As PrinterAccessRights)
Attribute DesiredAccess.VB_Description = "Sets the level of access desired to the selected printer"
Attribute DesiredAccess.VB_HelpID = 2628

    pDef.DesiredAccess = newAccess
    Dim lRet As Long
    
    If mhPrinter <> 0 Then
        Call ClosePrinter(mhPrinter)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":DeviceName", GetLastSystemError
        Else
            mhPrinter = 0
        End If
    End If
    If mhPrinter = 0 Then
        '\\ If the access hasn't been pre-specified, try all access
        If pDef.DesiredAccess = 0 Then
            pDef.DesiredAccess = PRINTER_ALL_ACCESS
        End If
        If msDeviceName <> "" Then
            '\\ Need to get the printer handle for the printer device named
            lRet = OpenPrinter(msDeviceName, mhPrinter, pDef)
        End If
    End If

End Property

'\\ --[DeviceContext]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the %device context:EventVB~ApiDeviceContext% upon which drawing, painting and text functions can be performed.
'  ##BLOCK_REMARKS Graphical operations on this device context take place on the current page.
'  ##BLOCK_REMARKS The %device context:EventVB~ApiDeviceContext% also supplies information about the graphical capabilities of the selected printer.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get DeviceContext() As ApiDeviceContext
Attribute DeviceContext.VB_Description = "This is the device context:EventVB~ApiDeviceContext upon which drawing, painting and text functions can be performed."
Attribute DeviceContext.VB_HelpID = 2629

Dim hPrinterDC As Long
Dim DCPrinter As New ApiDeviceContext

hPrinterDC = CreateDCByLong("WINSPOOL", DeviceName, vbNullString, ByVal 0&)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":DeviceContext", GetLastSystemError
End If

Set DeviceContext = DCPrinter

End Property

Public Property Get DeviceMode() As ApiDeviceMode
Attribute DeviceMode.VB_HelpID = 3099

Set dmThis = New ApiDeviceMode

Call RefreshPrinterInfo(2)
If dmThis.CreateFromPointer(mPRINTER_INFO_2.pDevMode) Then
    Set DeviceMode = dmThis
End If

End Property

'\\ --[DeviceName]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the unique name by which the printer is identified
'  ##BLOCK_REMARKS To change the physical printer that this ApiPrinter object points to,
'  ##BLOCK_REMARKS set this property to that physical printer's device na
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get DeviceName() As String
Attribute DeviceName.VB_Description = "This is the unique name by which the printer is identified  "
Attribute DeviceName.VB_HelpID = 2630

DeviceName = msDeviceName

End Property

Public Property Let DeviceName(ByVal newname As String)

Dim lRet As Long

If newname <> msDeviceName Then
    msDeviceName = newname
    If mhPrinter <> 0 Then
        Call ClosePrinter(mhPrinter)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":DeviceName", GetLastSystemError
        Else
            mhPrinter = 0
        End If
    End If
    If mhPrinter = 0 Then
        '\\ If the access hasn't been pre-specified, try all access
        If pDef.DesiredAccess = 0 Then
            pDef.DesiredAccess = PRINTER_ALL_ACCESS
        End If
        '\\ Need to get the printer handle for the printer device named
        lRet = OpenPrinter(newname, mhPrinter, pDef)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":DeviceName (Let)", GetLastSystemError
        End If
    End If
End If

End Property

'\\ --[Direct]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION If this property is set the printer is not served by a spooler.
'  ##BLOCK_REMARKS This data is provided by the printer driver/spooler driver.
'\\ ----------------------------------------------------------------------------------------
Public Property Get Direct() As Boolean
Attribute Direct.VB_Description = "If this property is set the printer is not served by a spooler."
Attribute Direct.VB_HelpID = 2631

Direct = Attributes And PRINTER_ATTRIBUTE_DIRECT

End Property

'\\ --[Driver]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the %printer driver:EventVB~ApiPrinterDriver% object that relates to the driver software installed for this printer.
'  ##BLOCK_REMARKS A printer driver is typically provided by the printer manufacturer.
'  ##BLOCK_REMARKS Different manufacturers may provide significantly different functionality in their respective drivers.
'  ##BLOCK_REMARKS This class wraps the most common (and useful) subset of this functionality.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Driver() As ApiPrinterDriver
Attribute Driver.VB_Description = "This is the printer driver:EventVB~ApiPrinterDriver object that relates to the driver software installed for this printer."
Attribute Driver.VB_HelpID = 2632

Dim drvThis As New ApiPrinterDriver

drvThis.Handle = mhPrinter
Set Driver = drvThis

End Property

'\\ --[DriverName]------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION The name of the driver software that is installed for this printer.
'  ##BLOCK_REMARKS Additional information can be retrieved by reference to the %printer driver:EventVB~ApiPrinterDriver%
'  ##BLOCK_REMARKS property of this ApiPrinter object.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get DriverName() As String
Attribute DriverName.VB_Description = "The name of the driver software that is installed for this printer."
Attribute DriverName.VB_HelpID = 2633

Call RefreshPrinterInfo(2)
DriverName = mPRINTER_INFO_2.pDriverName

End Property

'  ##BLOCK_DESCRIPTION Does the printer support duplex operation
Public Property Get Duplex() As Boolean
Attribute Duplex.VB_Description = "Does the printer support duplex operation"
Attribute Duplex.VB_HelpID = 3101

Call RefreshPrinterCapabilities(DC_DUPLEX)
Duplex = (mDeviceCapabilitiesRet > 0)

End Property



Public Property Let Location(ByVal newSetting As String)
Attribute Location.VB_Description = "This information is set to indicate the physical location of the printer.  "
Attribute Location.VB_HelpID = 2635

Call RefreshPrinterInfo(2)
mPRINTER_INFO_2.pLocation = newSetting
Call SavePrinterInfo(2)

End Property

'\\ --[Location]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This information is set to indicate the physical location of the printer.
'  ##BLOCK_REMARKS This information can be set by any user with sufficient access rights.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Location() As String

Call RefreshPrinterInfo(2)
Location = mPRINTER_INFO_2.pLocation

End Property

Public Property Get PaperBins() As Collection
Attribute PaperBins.VB_HelpID = 3102

Call RefreshPrinterCapabilities(DC_BINS)
Dim coThis As New Collection
Dim biThis As PrinterBins
Dim lItem As Long

For lItem = 0 To UBound(mDevOutput)
    With APIDispenser
        biThis = mDevOutput(lItem)
        coThis.Add biThis
    End With
Next lItem

Set PaperBins = coThis

End Property


Public Property Get Parameters() As String
Attribute Parameters.VB_HelpID = 2636

Call RefreshPrinterInfo(2)
Parameters = mPRINTER_INFO_2.pParameters

End Property

'\\ --[Pauseprinting]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION Suspends printing of any jobs queued on this printer.
'  ##BLOCK_REMARKS
'\\ ------------------------------------------------------------------------------------------------------
Public Sub PausePrinting()
Attribute PausePrinting.VB_Description = "Suspends printing of any jobs queued on this printer."
Attribute PausePrinting.VB_HelpID = 2637

Dim lRet As Long

lRet = SetPrinterApi(mhPrinter, 0, 0, PRINTER_CONTROL_PAUSE)

If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":PausePrinting", GetLastSystemError
End If

End Sub

'\\ --[Pauseprinting]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION Purges jobs queued on this printer.
'  ##BLOCK_REMARKS
'\\ ------------------------------------------------------------------------------------------------------
Public Sub PurgePrintJobs()
Attribute PurgePrintJobs.VB_Description = "Purges jobs queued on this printer."
Attribute PurgePrintJobs.VB_HelpID = 2886

Dim lRet As Long

lRet = SetPrinterApi(mhPrinter, 0, 0, PRINTER_CONTROL_PURGE)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":PurgePrintJobs", GetLastSystemError
End If

End Sub
'\\ --[PendingJobsCount]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the count of %print jobs:EventVB~ApiPrintJob% queued to print on the selected printer.
'\\ -------------------------------------------------------------------------------------------------------------------
Public Property Get PendingJobsCount() As Long
Attribute PendingJobsCount.VB_Description = "This is the count of print jobs:EventVB~ApiPrintJob queued to print on the selected printer."
Attribute PendingJobsCount.VB_HelpID = 2638

Call RefreshPrinterInfo(2)
PendingJobsCount = mPRINTER_INFO_2.JobsCount

End Property

'\\ --[PortName]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the port to which the printer is connected.
'  ##BLOCK_REMARKS For example "LPT1:"
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Portname() As String
Attribute Portname.VB_Description = "This is the name of the port to which the printer is connected."
Attribute Portname.VB_HelpID = 2639

Call RefreshPrinterInfo(2)
Portname = mPRINTER_INFO_2.pPortName

End Property

Friend Property Let PrinterHandle(ByVal newHandle As Long)

If newHandle <> mhPrinter Then
    mhPrinter = newHandle
End If

End Property


'\\ --[PrintJobs]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the collection of %print jobs:EventVB~ApiPrintJob% queued on the selected printer.
'  ##BLOCK_REMARKS This collection is subject to frequent change and should be read afresh whenever possible.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get PrintJobs() As Collection
Attribute PrintJobs.VB_Description = "This is the collection of print jobs:EventVB~ApiPrintJob queued on the selected printer."
Attribute PrintJobs.VB_HelpID = 2640

Dim colThis As Collection
Dim oPrintJob As ApiPrintJob

Dim lRet As Long
Dim pcbSizeRequired As Long, pcbBytesReturned As Long
Dim pJobId As Long

Dim buffer() As Long

'\\ Initialise the collection of print job objects
Set colThis = New Collection

ReDim Preserve buffer(0) As Long

lRet = EnumJobs(mhPrinter, 0, 255, 1, buffer(0), UBound(buffer), pcbSizeRequired, pcbBytesReturned)

If pcbSizeRequired > 0 Then
    
    '\\ Need to resize our array to cope with this data
    ReDim Preserve buffer(0 To (pcbSizeRequired / 4) + 3) As Long
    lRet = EnumJobs(mhPrinter, 0, 255, 1, buffer(0), UBound(buffer) * 4, pcbSizeRequired, pcbBytesReturned)
    
    '\\ Note any error if this has failed
    If Err.LastDllError <> 0 Then
        Call ReportError(Err.LastDllError, CLSNAME & ":PrintJobs", GetLastSystemError)
    End If
    
    '\\ At this stage buffer() contains an array of JOB_INFO_1 structures
    '\\ For each job...
    For pJobId = 0 To (pcbBytesReturned - 1) 'each record is 16 bytes long
        Set oPrintJob = New ApiPrintJob
        With oPrintJob
            .PrinterHandle = mhPrinter
            .JobId = buffer(pJobId * 16)
        End With
        colThis.Add oPrintJob
    Next pJobId
    
End If

'\\ pass the collection to the caller
Set PrintJobs = colThis

End Property

'\\ --[PrintJobs]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the print processor associated with the printer.
'  ##BLOCK_REMARKS This is often "WINSPOOL".
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get PrintProcessor() As String
Attribute PrintProcessor.VB_Description = "This is the name of the print processor associated with the printer."
Attribute PrintProcessor.VB_HelpID = 2641

Call RefreshPrinterInfo(2)
PrintProcessor = mPRINTER_INFO_2.pPrintProcessor

End Property

Public Property Let Priority(ByVal newSetting As Long)
Attribute Priority.VB_Description = "This is the priority assigned to jobs queued on the selected printer.  "
Attribute Priority.VB_HelpID = 2643

    Call RefreshPrinterInfo(2)
    mPRINTER_INFO_2.Priority = newSetting
    Call SavePrinterInfo(2)
    
End Property

'\\ --[PrintJobs]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the priority assigned to jobs queued on the selected printer.
'  ##BLOCK_REMARKS The priority of individual jobs may  be changed by users or processes with sufficient access rights.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get Priority() As Long

Call RefreshPrinterInfo(2)
Priority = mPRINTER_INFO_2.Priority

End Property

'\\ --[PrintJobs]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This attribute if the selected printer has a job queue associated with it.
'  ##BLOCK_REMARKS You should check this before accessing the %print jobs:EventVB~ApiPrinter~PrintJobs% property
'\\ -----------------------------------------------------------------------------------------------------------
Public Property Get Queued() As Boolean
Attribute Queued.VB_Description = "This attribute if the selected printer has a job queue associated with it."
Attribute Queued.VB_HelpID = 2644

Queued = Attributes And PRINTER_ATTRIBUTE_QUEUED

End Property


Private Sub RefreshPrinterCapabilities(ByVal mIndex As PrintDeviceCapabilitiesIndexes)

Dim lRet As Long
Dim pOutput As Long

If mPRINTER_INFO_2.pPortName = "" Then
    Call RefreshPrinterInfo(2)
End If

lRet = DeviceCapabilities(mPRINTER_INFO_1.pName, mPRINTER_INFO_2.pPortName, mIndex, ByVal 0&, ByVal 0&)
If lRet > 0 Then
    ReDim mDevOutput(lRet) As Byte
    mDeviceCapabilitiesRet = DeviceCapabilities(mPRINTER_INFO_1.pName, mPRINTER_INFO_2.pPortName, mIndex, VarPtr(mDevOutput(0)), ByVal 0&)
Else
    mDeviceCapabilitiesRet = 0
End If

If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Duplex", GetLastSystemError
End If

End Sub

Private Sub RefreshPrinterInfo(ByVal Index As Integer)

Dim lRet As Long
Dim SizeNeeded As Long


ReDim Preserve buffer(0 To 1) As Long
lRet = GetPrinterApi(mhPrinter, Index, buffer(0), UBound(buffer), SizeNeeded)
ReDim Preserve buffer(0 To (SizeNeeded / 4) + 3) As Long
lRet = GetPrinterApi(mhPrinter, Index, buffer(0), UBound(buffer) * 4, SizeNeeded)

bChanged = False

If Index < 1 Or Index > 9 Then
    Debug.Print "Error in call to ApiPrinter.RefreshPrinterInfo - invalid index"
Else
    Select Case Index
    Case 1
        With mPRINTER_INFO_1
            .flags = buffer(0)
            .pDescription = StringFromPointer(buffer(1), 1024)
            .pName = StringFromPointer(buffer(2), 1024)
            .pComment = StringFromPointer(buffer(3), 1024)
        End With
    Case 2
        With mPRINTER_INFO_2
            .pServerName = StringFromPointer(buffer(0), 1024)
            .pPrinterName = StringFromPointer(buffer(1), 1024)
            .pShareName = StringFromPointer(buffer(2), 1024)
            .pPortName = StringFromPointer(buffer(3), 1024)
            .pDriverName = StringFromPointer(buffer(4), 1024)
            .pComment = StringFromPointer(buffer(5), 1024)
            .pLocation = StringFromPointer(buffer(6), 1024)
            .pDevMode = buffer(7)
            .pSepFile = StringFromPointer(buffer(8), 1024)
            .pPrintProcessor = StringFromPointer(buffer(9), 1024)
            .pDatatype = StringFromPointer(buffer(10), 1024)
            .pParameters = StringFromPointer(buffer(11), 1024)
            .pSecurityDescriptor = buffer(12)
            .Attributes = buffer(13)
            .Priority = buffer(14)
            .DefaultPriority = buffer(15)
            .StartTime = buffer(16)
            .UntilTime = buffer(17)
            .Status = buffer(18)
            .JobsCount = buffer(19)
            .AveragePPM = buffer(20)
        End With
    Case 3
        With mPRINTER_INFO_3
            .pSecurityDescriptor = buffer(0)
        End With
    Case 4
        With mPRINTER_INFO_4
            .pPrinterName = StringFromPointer(buffer(0), 1024)
            .pServerName = StringFromPointer(buffer(1), 1024)
            .Attributes = buffer(2)
        End With
    Case 5
        With mPRINTER_INFO_5
            .pPrinterName = StringFromPointer(buffer(0), 1024)
            .pPortName = StringFromPointer(buffer(1), 1024)
            .Attributes = buffer(2)
            .DeviceNotSelectedTimeout = buffer(3)
            .TransmissionRetryTimeout = buffer(4)
        End With
    Case 6
        With mPRINTER_INFO_6
            .dwStatus = buffer(0)
        End With
    Case 7
        With mPRINTER_INFO_7
            .pszObjectGUID = StringFromPointer(buffer(0), 1024)
            .dwAction = buffer(1)
        End With
    Case 8
        With mPRINTER_INFO_8
            .pDevMode = buffer(0)
        End With
    Case 9
        '\\ Currently this structure is the same as 8...
        With mPRINTER_INFO_9
            .pDevMode = buffer(0)
        End With
    End Select
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, CLSNAME & ":RefreshPrinterInfo", GetLastSystemError
    End If
End If

End Sub

Public Sub ResumePrinting()
Attribute ResumePrinting.VB_HelpID = 2645

Dim lRet As Long

lRet = SetPrinterApi(mhPrinter, 0, 0, PRINTER_CONTROL_RESUME)

If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":ResumePrinting", GetLastSystemError
End If

End Sub

Private Sub SavePrinterInfo(ByVal Level As Long)

Dim lRet As Long

Select Case Level
Case 1
    ReportError -1, CLSNAME & ":SavePrinterInfo", "Cannot save level 1 printer info"
Case 2
    With mPRINTER_INFO_2
        buffer(0) = LPCSTR(.pServerName)
        buffer(1) = LPCSTR(.pPrinterName)
        buffer(2) = LPCSTR(.pShareName)
        buffer(3) = LPCSTR(.pPortName)
        buffer(4) = LPCSTR(.pDriverName)
        buffer(5) = LPCSTR(.pComment)
        buffer(6) = LPCSTR(.pLocation)
        buffer(7) = .pDevMode
        buffer(8) = LPCSTR(.pSepFile)
        buffer(9) = LPCSTR(.pPrintProcessor)
        buffer(10) = LPCSTR(.pDatatype)
        buffer(11) = LPCSTR(.pParameters)
        buffer(12) = .pSecurityDescriptor
        buffer(13) = .Attributes
        buffer(14) = .Priority
        buffer(15) = .DefaultPriority
        buffer(16) = .StartTime
        buffer(17) = .UntilTime
        buffer(18) = .Status
        buffer(19) = .JobsCount
        buffer(20) = .AveragePPM
    End With
Case 3
Case 4
Case 5
Case 6
Case 7
Case 8
Case 9
End Select

lRet = SetPrinterApi(mhPrinter, Level, buffer(0), 0)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":SavePrinterInfo", GetLastSystemError
Else
    Call ResetLPCSTRIndexes
End If

End Sub

Public Property Let SeparatorFile(ByVal newSetting As String)
Attribute SeparatorFile.VB_Description = "If this is not blank then the specified file is printed to seperate individual print jobs.  "
Attribute SeparatorFile.VB_HelpID = 2647

Call RefreshPrinterInfo(2)
mPRINTER_INFO_2.pSepFile = newSetting
Call SavePrinterInfo(2)

End Property

'\\ --[SeparatorFile]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION If this is not blank then the specified file is printed to seperate individual print jobs.
'\\ ----------------------------------------------------------------------------------------------------------------
Public Property Get SeparatorFile() As String

Call RefreshPrinterInfo(2)
SeparatorFile = mPRINTER_INFO_2.pSepFile

End Property

'\\ --[ServerName]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the name of the physical machine (often a dedicated server) to which the printer is attached.
'  ##BLOCK_REMARKS For a local printer this name will be the same as the %computer name:EventVB~ApiSystem~ComputerName%.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get ServerName() As String
Attribute ServerName.VB_Description = "This is the name of the physical machine (often a dedicated server) to which the printer is attached."
Attribute ServerName.VB_HelpID = 2648

Call RefreshPrinterInfo(4)
ServerName = mPRINTER_INFO_4.pServerName

End Property


'\\ --[ShareName]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This is the share name by which the printer can be seen from other machines connected to the network.
'  ##BLOCK_REMARKS This will be blank if the printer is not shared.
'\\ ------------------------------------------------------------------------------------------------------
Public Property Get ShareName() As String
Attribute ShareName.VB_Description = "This is the share name by which the printer can be seen from other machines connected to the network."
Attribute ShareName.VB_HelpID = 2649

   Call RefreshPrinterInfo(2)
   ShareName = mPRINTER_INFO_2.pShareName

End Property


'  ##BLOCK_DESCRIPTION Does the printer support job stapling
Public Property Get Stapler() As Boolean
Attribute Stapler.VB_Description = "Does the printer support job stapling"
Attribute Stapler.VB_HelpID = 3103

Call RefreshPrinterCapabilities(DC_STAPLE)
Stapler = (mDeviceCapabilitiesRet > 0)

End Property

'\\ --[Status]-----------------------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION This returns the current status of the selected printer.
'  ##BLOCK_REMARKS This setting may not be valid if there are no jobs queued for printing.
'  ##BLOCK_REMARKS Some printer drivers only update the status when a job returns a failure or success code.
'\\ ---------------------------------------------------------------------------------------------------------
Public Property Get Status() As Long
Attribute Status.VB_Description = "This returns the current status of the selected printer.  "
Attribute Status.VB_HelpID = 2650

    Call RefreshPrinterInfo(2)
    Status = mPRINTER_INFO_2.Status

End Property

Public Property Let Status(ByVal newSetting As Long)

Dim lRet As Long
lRet = SetPrinterApi(mhPrinter, 0, newSetting, PRINTER_CONTROL_SET_STATUS)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Status (Set)", GetLastSystemError
End If

End Property

Private Sub Class_Terminate()

If mhPrinter <> 0 Then
    Call ClosePrinter(mhPrinter)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Terminate", GetLastSystemError
    End If
    Set dmThis = Nothing
End If

End Sub


Private Sub dmThis_Changed()

Call RefreshPrinterInfo(2)
If dmThis.CopyToPointer(mPRINTER_INFO_2.pDevMode) Then
    Call SavePrinterInfo(2)
End If

End Sub



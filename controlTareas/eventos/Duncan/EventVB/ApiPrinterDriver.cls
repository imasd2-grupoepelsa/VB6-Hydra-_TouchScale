VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiPrinterDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2672
Attribute VB_Description="This is the printer driver:EventVB~ApiPrinterDriver object that relates to the driver software installed for this printer."
Option Explicit

'\\ --[ApiPrinterDriver]------------------------------------------------
'  ##MODULE_DESCRIPTION This is the %printer driver:EventVB~ApiPrinterDriver% object that relates to the driver software installed for this printer.
'  ##MODULE_REMARKS A printer driver is typically provided by the printer manufacturer.
'  ##MODULE_REMARKS Different manufacturers may provide significantly different functionality in their respective drivers.
'  ##MODULE_REMARKS This class wraps the most common (and useful) subset of this functionality.
'\\ ----------------------------------------------------------------------------------------

' ##ENUMERATION_DESCRIPTION PrinterDriverOSVersion is the version of the Windows operating system _
for which the driver was specifically written.
Public Enum PrinterDriverOSVersion
     ' ##ENUMERATION_MEMBER_DESCRIPTION DRIVER_OS_WIN9x The driver was written for Windows 95, 98 or Me
     DRIVER_OS_WIN9x = &H0
     ' ##ENUMERATION_MEMBER_DESCRIPTION DRIVER_OS_NT351 The driver was written for Windows NT 3.51
     DRIVER_OS_NT351 = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION DRIVER_OS_NT4 The driver was written for Windows NT4
     DRIVER_OS_NT4 = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION DRIVER_OS_WIN2K The driver was written for Windows 2000
     DRIVER_OS_WIN2K = &H3
End Enum

Private Type DRIVER_INFO_1
    lpDriverName As String
End Type

Private Type DRIVER_INFO_2
     lVersion As PrinterDriverOSVersion
     lpDriverName As String
     lpEnvironment As String
     lpDriverPath As String
     lpDataFile As String
     lpConfigFile As String
End Type

Private Type DRIVER_INFO_3
     lVersion As PrinterDriverOSVersion
     lpDriverName As String
     lpEnvironment As String
     lpDriverPath As String
     lpDataFile As String
     lpConfigFile As String
     lpHelpFile As String
     lpDependentFiles As String
     lpMonitorName As String
     lpDefaultDataType As String
End Type

'\\ The following types are only supported by drivers
'\\ written for Win2K and WinXP
Private Type DRIVER_INFO_4
     lVersion As PrinterDriverOSVersion
     lpDriverName As String
     lpEnvironment As String
     lpDriverPath As String
     lpDataFile As String
     lpConfigFile As String
     lpHelpFile As String
     lpDependentFiles As String
     lpMonitorName As String
     lpDefaultDataType As String
     lpCompatiblePreviousNames As String
End Type


'\\ API Functions used...
Private Declare Function GetPrinterDriver Lib "winspool.drv" Alias "GetPrinterDriverA" _
                  (ByVal hPrinter As Long, _
                   ByVal lpEnvironment As String, _
                   ByVal Level As Long, _
                   pDriverBuffer As Long, _
                   ByVal cBufferSize As Long, _
                   cBufferSizeRequired As Long) As Long
                   
Private Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" _
                  (ByVal lpServerName As String, _
                   ByVal lpEnvironment As String, _
                   ByVal Level As Long, _
                   lpDriverBuffer As Long, _
                   ByVal cBufferSize As Long, _
                   cBufferSizeRequired As Long) As Long
                   
'\\ Private member variables....
Private mhPrinter As Long
Private mEnvironment As String

Private mDRIVER_INFO_1 As DRIVER_INFO_1
Private mDRIVER_INFO_2 As DRIVER_INFO_2
Private mDRIVER_INFO_3 As DRIVER_INFO_3
Private mDRIVER_INFO_4 As DRIVER_INFO_4

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get CompatiblePreviousDriverNames() As String
Attribute CompatiblePreviousDriverNames.VB_HelpID = 2673

    Call RefreshDriverInfo(4)
    CompatiblePreviousDriverNames = mDRIVER_INFO_4.lpCompatiblePreviousNames
    
End Property

Public Property Get ConfigurationFile() As String
Attribute ConfigurationFile.VB_HelpID = 2674

    Call RefreshDriverInfo(2)
    ConfigurationFile = mDRIVER_INFO_2.lpConfigFile
    
End Property

Public Property Get DataFile() As String
Attribute DataFile.VB_HelpID = 2675

    Call RefreshDriverInfo(2)
    DataFile = mDRIVER_INFO_2.lpDataFile
    
End Property

Public Property Get DefaultDataType() As String
Attribute DefaultDataType.VB_HelpID = 2676

    Call RefreshDriverInfo(3)
    DefaultDataType = mDRIVER_INFO_3.lpDefaultDataType
    
End Property

Public Property Get Environment() As String
Attribute Environment.VB_HelpID = 2677
    
    Call RefreshDriverInfo(2)
    Environment = mDRIVER_INFO_2.lpEnvironment
    
End Property

Friend Property Let Handle(ByVal newhandle As Long)

    mhPrinter = newhandle
    
End Property

Public Property Get HelpFile() As String
Attribute HelpFile.VB_HelpID = 2678

    Call RefreshDriverInfo(3)
    HelpFile = mDRIVER_INFO_3.lpHelpFile
    
End Property

Public Property Get MonitorName() As String
Attribute MonitorName.VB_HelpID = 2679

    Call RefreshDriverInfo(3)
    MonitorName = mDRIVER_INFO_3.lpMonitorName
    
End Property


Public Property Get Name() As String
Attribute Name.VB_HelpID = 2680

Call RefreshDriverInfo(2)
Name = mDRIVER_INFO_2.lpDriverName

End Property

Public Property Get OperatingSystemVersion() As PrinterDriverOSVersion
Attribute OperatingSystemVersion.VB_HelpID = 2681

Call RefreshDriverInfo(2)
OperatingSystemVersion = mDRIVER_INFO_2.lVersion

End Property

Public Property Get Path() As String
Attribute Path.VB_HelpID = 2682

    Call RefreshDriverInfo(2)
    Path = mDRIVER_INFO_2.lpDriverPath
    
End Property

Private Sub RefreshDriverInfo(ByVal index As Integer)

Dim lRet As Long
Dim SizeNeeded As Long

Dim buffer() As Long

'\\ Find out how much buffer space we're going to need
ReDim Preserve buffer(0 To 1) As Long
lRet = GetPrinterDriver(mhPrinter, mEnvironment, index, buffer(0), UBound(buffer), SizeNeeded)
'\\ Resize the buffer to accomodate that amount
ReDim Preserve buffer(0 To (SizeNeeded / 4) + 3) As Long
lRet = GetPrinterDriver(mhPrinter, mEnvironment, index, buffer(0), UBound(buffer) * 4, SizeNeeded)

If index < 1 Or index > 4 Then
    '\\ Index is out of bounds - notify developer of error
    Debug.Print "Error in call to ApiPrinterDriver:RefreshPrinterDriverInfo - invalid index"
Else
    Select Case index
    Case 1
        With mDRIVER_INFO_1
            .lpDriverName = StringFromPointer(buffer(0), 1024)
        End With
    Case 2
        With mDRIVER_INFO_2
            .lVersion = buffer(0)
            .lpDriverName = StringFromPointer(buffer(1), 1024)
            .lpEnvironment = StringFromPointer(buffer(2), 1024)
            .lpDriverPath = StringFromPointer(buffer(3), 1024)
            .lpDataFile = StringFromPointer(buffer(4), 1024)
            .lpConfigFile = StringFromPointer(buffer(5), 1024)
        End With
    Case 3
        With mDRIVER_INFO_3
            .lVersion = buffer(0)
            .lpDriverName = StringFromPointer(buffer(1), 1024)
            .lpEnvironment = StringFromPointer(buffer(2), 1024)
            .lpDriverPath = StringFromPointer(buffer(3), 1024)
            .lpDataFile = StringFromPointer(buffer(4), 1024)
            .lpConfigFile = StringFromPointer(buffer(5), 1024)
            .lpHelpFile = StringFromPointer(buffer(6), 1024)
            .lpDependentFiles = StringFromPointer(buffer(7), 1024)
            .lpMonitorName = StringFromPointer(buffer(8), 1024)
            .lpDefaultDataType = StringFromPointer(buffer(9), 1024)
        End With
    Case 4
        With mDRIVER_INFO_4
            .lVersion = buffer(0)
            .lpDriverName = StringFromPointer(buffer(1), 1024)
            .lpEnvironment = StringFromPointer(buffer(2), 1024)
            .lpDriverPath = StringFromPointer(buffer(3), 1024)
            .lpDataFile = StringFromPointer(buffer(4), 1024)
            .lpConfigFile = StringFromPointer(buffer(5), 1024)
            .lpHelpFile = StringFromPointer(buffer(6), 1024)
            .lpDependentFiles = StringFromPointer(buffer(7), 1024)
            .lpMonitorName = StringFromPointer(buffer(8), 1024)
            .lpDefaultDataType = StringFromPointer(buffer(9), 1024)
            .lpCompatiblePreviousNames = StringFromPointer(buffer(10), 1024)
        End With
    End Select
End If
End Sub

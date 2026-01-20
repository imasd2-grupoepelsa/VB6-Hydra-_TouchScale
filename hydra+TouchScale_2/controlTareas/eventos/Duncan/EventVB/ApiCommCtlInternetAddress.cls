VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiCommCtlInternetAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2553
Attribute VB_Description = "This class provides methods and properties for manipulating instances of the "
Option Explicit

' ##MODULE_DESCRIPTION This class provides methods and properties for manipulating _
instances of the "Internet Address" windows custom control

Private mWND As ApiWindow


Private Const IPN_FIRST = -860
Public Enum enIPAddressNotificationMessages
    IPN_FIELDCHANGED = -860
End Enum

' ##ENUMERATION_DESCRIPTION enIPAddressMessages - Special windows messages pertaining to the internet address control
Public Enum enIPAddressMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_CLEARADDRESS - Clears the address in teh control
    IPM_CLEARADDRESS = 1124 ' ##  no parameters
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_SETADDRESS - Sets the address in the control
    IPM_SETADDRESS = 1125 ' ##  lparam = TCP/IP address
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_GETADDRESS - Returns the address held in the control
    IPM_GETADDRESS = 1126 ' ##  lresult = # of non black fields.  lparam = LPDWORD for TCP/IP address
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_SETRANGE - Sets the range of possible addrersses for the control
    IPM_SETRANGE = 1127 ' ##  wparam = field, lparam = range
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_SETFOCUS - Sets the focus to one of the address boxes
    IPM_SETFOCUS = 1128 ' ##  wparam = field
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPM_ISBLANK - Returns true if the control is blank
    IPM_ISBLANK = 1129 ' ##  no parameters
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type NMIPADDRESS
    NMHDR_hwndFrom As Long
    NMHDR_idfrom As Long
    NMHDR_code As Long
    iField As Long
    iValue As Long
End Type

' ##ENUMERATION_DESCRIPTION enIPAddressFields - The fields in an IP address control
Public Enum enIPAddressFields
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPAF_FIRST - The first box in the IP address
    IPAF_FIRST = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPAF_SECOND - The second box in the IP address
    IPAF_SECOND = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPAF_THIRD - The third box in the IP address
    IPAF_THIRD = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION IPAF_FOURTH - The fourth box in the IP address
    IPAF_FOURTH = 4
End Enum

Private bChanged As Boolean

Public Property Get Blank() As Boolean
Attribute Blank.VB_HelpID = 2554

Dim lret As Long

lret = SendMessage(m_hwnd, IPM_ISBLANK, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:Clear", GetLastSystemError
End If
Blank = CBool(lret)

End Property


Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get ClassName() As String
Attribute ClassName.VB_HelpID = 2555

    ClassName = "SysIPAddress32"
    
End Property

Public Sub Clear()
Attribute Clear.VB_HelpID = 2556

Dim lret As Long

lret = SendMessage(m_hwnd, IPM_CLEARADDRESS, 0, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:Clear", GetLastSystemError
End If

End Sub




Public Property Let IPAddress(ByVal newAddress As Long)
Attribute IPAddress.VB_HelpID = 2559

Dim lret As Long

lret = SendMessage(m_hwnd, IPM_SETADDRESS, 0, ByVal newAddress)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:IPAddress", GetLastSystemError
End If

End Property

Public Property Get IPAddress() As Long

Dim lret As Long
Dim lAddress As Long

lret = SendMessage(m_hwnd, IPM_GETADDRESS, 0, lAddress)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:IPAddress", GetLastSystemError
End If
IPAddress = lAddress

End Property

Private Property Get m_hwnd() As Long

m_hwnd = mWND.hwnd

End Property

Public Sub SetFieldFocus(ByVal WhichField As enIPAddressFields)
Attribute SetFieldFocus.VB_HelpID = 2560

Dim lret As Long

lret = SendMessage(m_hwnd, IPM_SETFOCUS, WhichField, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:SetFieldFocus", GetLastSystemError
End If

End Sub

Public Sub SetRange(ByVal WhichField As enIPAddressFields, ByVal Lower As Byte, ByVal Upper As Byte)
Attribute SetRange.VB_HelpID = 2561

Dim lret As Long
Dim lRange As Long

lRange = MakeInt(Lower, Upper)

lret = SendMessage(m_hwnd, IPM_SETFOCUS, WhichField, ByVal lRange)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiCommCtlInternetAddress:SetFieldFocus", GetLastSystemError
End If

End Sub


Public Property Get Window() As ApiWindow
Attribute Window.VB_HelpID = 2820

Set Window = mWND

End Property


Private Sub Class_Initialize()

'\\ make sure the common controls class is initialised
Call APIDispenser.InitCommonControls(ICC_INTERNET_CLASSES)

Dim lStyle As Long
Dim lExStyle As Long

If mWND Is Nothing Then
    Dim clsInternetAddress As New APIWndClass
    With clsInternetAddress
        .CreateFromClassname App.hInstance, ClassName
    End With
    lStyle = clsInternetAddress.Style
    Set mWND = APIDispenser.System.WindowDispenser.CreateWindow(lExStyle, clsInternetAddress, "", lStyle, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
End If

End Sub


Private Sub Class_Terminate()

Set mWND = Nothing

End Sub



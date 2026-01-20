Attribute VB_Name = "ICMP_PINGS"
Option Explicit
Private Const IP_SUCCESS As Long = 0
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const WS_VERSION_REQD As Long = &H101
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function gethostname Lib "wsock32" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" _
   (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
Private Declare Function inet_addr Lib "wsock32" _
   (ByVal s As String) As Long
Public Function GetStatusCode(status As Long) As String
   Dim msg As String
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
End Function
Public Function ICMP_Ping(sAddress As String, _
                     sDataToSend As String, _
                     Echo As ICMP_ECHO_REPLY) As Long
   Dim hPort As Long
   Dim dwAddress As Long
   Dim MyIP As String
   Dim Direccion As String
   Dim Bucle As Integer
   Dim MyStatus As Long
   
   Dim s As Variant
   
   sDataToSend = "12345678901234567890123456789012"
   
   'MsgBox sAddress
   s = sAddress
   
   Direccion = ConvertIP(sAddress)
   If Direccion <> "" Then
   Else
       Direccion = s
   End If

   
   If Direccion <> "" Then
        dwAddress = inet_addr(Direccion)
   Else
        MyIP = GetIPFromHostName(sAddress)
        If Trim(MyIP) <> "" Then dwAddress = inet_addr(MyIP)
   End If
   If dwAddress <> INADDR_NONE Then
        For Bucle = 1 To 2
            hPort = IcmpCreateFile()
            If hPort Then
                Call IcmpSendEcho(hPort, dwAddress, sDataToSend, _
                LenB(sDataToSend), 0, Echo, Len(Echo), PING_TIMEOUT)
                MyStatus = Echo.status
                Call IcmpCloseHandle(hPort)
                If MyStatus = 0 Then Exit For
            End If
            Do_Events
        Next Bucle
   Else
         ICMP_Ping = INADDR_NONE
   End If
   ICMP_Ping = MyStatus
End Function
Public Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       MsgBox "Error de la Pila WinsockOption Explicit"
    End If
End Sub
Public Function GetIPFromHostName(ByVal sHostName As String) As String
  'converts a host name to an IP address.
   Dim nbytes As Long
   Dim ptrHosent As Long
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String
   sAddress = Space$(4)
   ptrHosent = gethostbyname(sHostName & vbNullChar)
   If ptrHosent <> 0 Then
      ptrAddress = ptrHosent + 12
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
      GetIPFromHostName = IPToText(sAddress)
   End If
End Function
Private Function IPToText(ByVal IPAddress As String) As String
   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function
Public Function ConvertIP(Direccion As String) As String
    Dim Resp As String
    Dim Bytes(4) As String
    Dim Cont As Integer
    Dim Bucle As Integer
    Dim Error As Boolean
    Cont = 0
    For Bucle = 1 To Len(Direccion)
        If IsNumeric(Mid(Direccion, Bucle, 1)) Then
            Bytes(Cont) = Bytes(Cont) & Mid(Direccion, Bucle, 1)
        Else
            Cont = Cont + 1
            If Cont > 3 Or Mid(Direccion, Bucle, 1) <> "." Then
                Error = True
                Exit For
            End If
        End If
    Next Bucle
    If Error Then
        ConvertIP = ""
    Else
        ConvertIP = Val(Bytes(0)) & "." & Val(Bytes(1)) & "." & Val(Bytes(2)) & "." & Val(Bytes(3))
    End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiNamedPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2655
Option Explicit

'\\ -- [ ApiNamedPipe ]------------------------------------------------------
'\\ Class to manage a named pipe based communication
'\\ --------------------------------------------------------------------

Public Enum NamedPipeOpenMode
    PIPE_ACCESS_INBOUND = &H1
    PIPE_ACCESS_OUTBOUND = &H2
    PIPE_ACCESS_DUPLEX = &H3
End Enum

Public Enum NamedPipeEnd
    PIPE_CLIENT_END = &H0
    PIPE_SERVER_END = &H1
End Enum

'\\ These three enums are combined to NamedPipeMode in the documentation..
Public Enum NamedPipeWaitMode
    PIPE_WAIT = &H0
    PIPE_NOWAIT = &H1
End Enum

Public Enum NamedPipeReadMode
    PIPE_READMODE_BYTE = &H0
    PIPE_READMODE_MESSAGE = &H2
End Enum

Public Enum NamedPipeWriteMode
    PIPE_WRITEMODE_BYTE = &H0
    PIPE_WRITEMODE_MESSAGE = &H4
End Enum

Private Const PIPE_UNLIMITED_INSTANCES = 255

Private Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Long
End Type

Private Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" _
                           (ByVal lpName As String, _
                           ByVal dwOpenMode As Long, _
                           ByVal dwPipeMode As Long, _
                           ByVal nMaxInstances As Long, _
                           ByVal nOutBufferSize As Long, _
                           ByVal nInBufferSize As Long, _
                           ByVal nDefaultTimeOut As Long, _
                           lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" _
                           (ByVal lpNamedPipeName As String, _
                            ByVal nTimeOut As Long) As Long
                            
Private Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" _
                            (ByVal lpNamedPipeName As String, _
                             lpInBuffer As Any, _
                             ByVal lInBufferSize As Long, _
                             lpOutBuffer As Any, _
                             ByVal lOutBufferSize As Long, _
                             lBytesRead As Long, _
                             ByVal nTimeOut As Long) As Long
                             
Private Declare Function GetNamedPipeHandleState Lib "kernel32" Alias "GetNamedPipeHandleStateA" _
                            (ByVal hNamedPipeHandle As Long, _
                            lpState As Long, _
                            lpCurrInstances As Long, _
                            lpMaxCollectionCount As Long, _
                            lpCollectDataTimeOut As Long, _
                            ByVal lpUswerName As String, _
                            ByVal nMaxUsernameSize As Long) As Long
                            
Private Declare Function GetNamedPipeInfo Lib "kernel32" Alias "GetNamedPipeInfoA" _
                            (ByVal hNamedPipe As Long, _
                             lpFlags As Long, _
                             lpOutBufferSize As Long, _
                             lpInBufferSize As Long, _
                             lpMaxInstances As Long) As Long
                             

Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEventHandle As Long
End Type

Private Declare Function ConnectNamedPipe Lib "kernel32" _
                           (ByVal hNamedPipeHandle As Long, _
                           lpOverlapped As OVERLAPPED) As Long
                           
Private Declare Function DisconnectNamedPipe Lib "kernel32" _
                           (ByVal hNamedPipeHandle As Long) As Long
                           

Private mPipeHandle As Long

'\\ A pipe name must start with the standard prefix:
Private Const PIPE_NAME_PREFIX = "\\.\pipe\"
Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property
Public Property Get Handle() As Long
Attribute Handle.VB_HelpID = 2656

Handle = mPipeHandle

End Property



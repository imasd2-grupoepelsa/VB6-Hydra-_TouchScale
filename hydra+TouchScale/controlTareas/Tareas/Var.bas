Attribute VB_Name = "Var"
Option Explicit

' declaraciones api UDT y constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
    "CreateToolhelp32Snapshot" ( _
        ByVal lFlags As Long, _
        ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32" Alias _
    "Process32First" ( _
        ByVal hSnapShot As Long, _
        uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32" Alias _
    "Process32Next" ( _
        ByVal hSnapShot As Long, _
        uProcess As PROCESSENTRY32) As Long

Public Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' constante/mensaje para cerrar el proceso
Public Const PROCESS_TERMINATE As Long = (&H1)

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist _
                                + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Const MAX_PATH As Integer = 260

'estructura para las funciones ProcessNext, ProcessFirst
Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * MAX_PATH
End Type




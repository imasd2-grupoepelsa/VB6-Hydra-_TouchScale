Attribute VB_Name = "Var"
Option Explicit

'\\\\\\\\\\\\\\
'//////////////
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////////
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
''''''''''''''''''''''''''''''''''

Public PathSc10 As String
Public TimeAuto As Integer
Public TimeAutoSeg As Integer


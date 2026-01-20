VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIFileTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2096
Attribute VB_Description = "This class provides properties for the FILETIME structure which is used to describe the date and time a file was written or modified."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ##MODULE_DESCRIPTION This class provides properties for the FILETIME structure which _
is used to describe the date and time a file was written or modified.

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'\\ Createion successful?
Public CreatedOK As Boolean
Attribute CreatedOK.VB_VarHelpID = 2097
'\\ Member variables
Public dwLowDateTime As Long
Attribute dwLowDateTime.VB_VarHelpID = 2098
Public dwHighDateTime As Long
Attribute dwHighDateTime.VB_VarHelpID = 2099

'\\ Private memory handling functions
Private Declare Sub CopyMemoryFiletime Lib "kernel32" Alias "RtlMoveMemory" (Destination As FILETIME, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrFiletime Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrFiletime Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ Conversion to a system ti...
Private Declare Function FileTimeToSystemTimeApi Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
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

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this Filetime object from the location poiunted to by
'\\ lpFiletime
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpFileTime As Long) As Boolean

Dim ftThis As FILETIME

CreatedOK = False

If Not IsBadReadPtrFiletime(lpFileTime, Len(ftThis)) Then
    Call CopyMemoryFiletime(ftThis, lpFileTime, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            dwLowDateTime = .dwLowDateTime
            dwHighDateTime = .dwHighDateTime
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Function FileTimeToSystemTime() As APISystemTime
Attribute FileTimeToSystemTime.VB_HelpID = 2100

Dim ftThis As FILETIME
Dim stThis As SYSTEMTIME
Dim oSystemTime As APISystemTime

Dim lret As Long

With ftThis
    .dwHighDateTime = dwHighDateTime
    .dwLowDateTime = dwLowDateTime
End With

lret = FileTimeToSystemTimeApi(ftThis, stThis)
If Err.LastDllError = 0 Then
    Set oSystemTime = New APISystemTime
    If oSystemTime.CreateFromPointer(VarPtr(stThis)) Then
        FileTimeToSystemTime = oSystemTime
    End If
End If

End Function



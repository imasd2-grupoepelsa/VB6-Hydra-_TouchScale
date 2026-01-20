VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APISystemTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2101
Option Explicit

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

Public CreatedOK As Boolean
Public wYear As Integer
Public wMonth As Integer
Public wDayOfWeek As Integer
Public wDay As Integer
Public wHour As Integer
Public wMinute As Integer
Public wSecond As Integer
Public wMilliseconds As Integer


'\\ Private memory handling functions
Private Declare Sub CopyMemorySystemTime Lib "kernel32" Alias "RtlMoveMemory" (Destination As SYSTEMTIME, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrSystemTime Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrSystemTime Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private bChanged As Boolean
Attribute CreatedOK.VB_VarHelpID = 2102
Attribute wYear.VB_VarHelpID = 2103
Attribute wMonth.VB_VarHelpID = 2104
Attribute wDayOfWeek.VB_VarHelpID = 2105
Attribute wDay.VB_VarHelpID = 2106
Attribute wHour.VB_VarHelpID = 2107
Attribute wMinute.VB_VarHelpID = 2108
Attribute wSecond.VB_VarHelpID = 2109
Attribute wMilliseconds.VB_VarHelpID = 2110
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this SystemTime object from the location poiunted to by
'\\ lpSystemTime
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpSystemTime As Long) As Boolean

Dim ftThis As SYSTEMTIME

CreatedOK = False

If Not IsBadReadPtrSystemTime(lpSystemTime, Len(ftThis)) Then
    CopyMemory ftThis, ByVal lpSystemTime, ByVal Len(ftThis)
    If Err.LastDllError = 0 Then
        With ftThis
            wDay = .wDay
            wDayOfWeek = .wDayOfWeek
            wHour = .wHour
            wMilliseconds = .wMilliseconds
            wMinute = .wMinute
            wMonth = .wMonth
            wSecond = .wSecond
            wYear = .wYear
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function






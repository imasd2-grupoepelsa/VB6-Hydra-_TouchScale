VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiNotifyMessageHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2562
Option Explicit

'\\ --[ApiNotifyMessageHeader]--------------------
'\\ Wrapper for the NMHDR structurte passed in the
'\\ lParam of the WM_NOTIFY message
'\\ ----------------------------------------------
Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryNMHDR Lib "kernel32" Alias "RtlMoveMemory" (Destination As NMHDR, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrNMHDR Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrNMHDR Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long


Public hwndFrom As Long
Public idFrom As Long
Public Code As Long
Private bChanged As Boolean
Attribute CreatedOK.VB_VarHelpID = 2563
Attribute hwndFrom.VB_VarHelpID = 2564
Attribute idFrom.VB_VarHelpID = 2565
Attribute Code.VB_VarHelpID = 2566
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this NMHDR object from the location poiunted to by
'\\ lpNMHDR
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpNMHDR As Long) As Boolean

Dim ftThis As NMHDR

CreatedOK = False

If Not IsBadReadPtrNMHDR(lpNMHDR, Len(ftThis)) Then
    Call CopyMemoryNMHDR(ftThis, lpNMHDR, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            Code = .Code
            hwndFrom = .hwndFrom
            idFrom = .idFrom
            CreatedOK = True
        End With
    Else
        ReportError Err.LastDllError, "ApiNotifyMessageHeader:CreateFromPointer", GetLastSystemError
    End If
End If

CreateFromPointer = CreatedOK

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiNotifyMessageMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2567
Option Explicit

'\\ --[ApiNotifyMessageMouse]------------------------------------------------
'\\ The structure wrapper for an NMMOUSE type, passed as lParam
'\\ for certain WM_NOTIFY messages
'\\ NMCLICK duplicates this structure....
'\\ --------------------------------------------------------------


Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type NMMOUSE
    hdr As NMHDR
    dwItemSpec As Long
    dwItemData As Long
    pt As POINTAPI
    dwHitInfo As Long
End Type

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryNMMOUSE Lib "kernel32" Alias "RtlMoveMemory" (Destination As NMMOUSE, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrNMMOUSE Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrNMMOUSE Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Public ItemSpec As Long
Public ItemData As Long
Public HitInfo As Long
Public hdr As ApiNotifyMessageHeader
Public pt As APIPoint

Private bChanged As Boolean
Attribute CreatedOK.VB_VarHelpID = 2568
Attribute ItemSpec.VB_VarHelpID = 2569
Attribute ItemData.VB_VarHelpID = 2570
Attribute HitInfo.VB_VarHelpID = 2571
Attribute hdr.VB_VarHelpID = 2572
Attribute pt.VB_VarHelpID = 2573
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this NMMOUSE object from the location pointed to by
'\\ lpNMMOUSE
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpNMMOUSE As Long) As Boolean

Dim ftThis As NMMOUSE

CreatedOK = False

If Not IsBadReadPtrNMMOUSE(lpNMMOUSE, Len(ftThis)) Then
    Call CopyMemoryNMMOUSE(ftThis, lpNMMOUSE, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            Set hdr = New ApiNotifyMessageHeader
            hdr.CreateFromPointer (VarPtr(.hdr))
            Set pt = New APIPoint
            pt.CreateFromPointer (VarPtr(.pt))
            HitInfo = .dwHitInfo
            ItemSpec = .dwItemSpec
            ItemData = .dwItemData
        End With
    Else
        ReportError Err.LastDllError, "ApiNotifyMessageHeader:CreateFromPointer", GetLastSystemError
    End If
End If

CreateFromPointer = CreatedOK

End Function

Private Sub Class_Terminate()

Set hdr = Nothing
Set pt = Nothing

End Sub



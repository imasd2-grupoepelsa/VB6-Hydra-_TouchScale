VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2126
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public CreatedOK As Boolean
Public x As Long
Public y As Long

'\\ Private memory handling functions
Private Declare Sub CopyMemoryPointAPI Lib "kernel32" Alias "RtlMoveMemory" (Destination As POINTAPI, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrPointAPI Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrPointAPI Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Private bChanged As Boolean
Attribute CreatedOK.VB_VarHelpID = 2127
Attribute x.VB_VarHelpID = 2128
Attribute y.VB_VarHelpID = 2129
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this PointAPI object from the location poiunted to by
'\\ lpPointAPI
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corPointAPIness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpPointAPI As Long) As Boolean

Dim ftThis As POINTAPI

CreatedOK = False

If Not IsBadReadPtrPointAPI(lpPointAPI, Len(ftThis)) Then
    Call CopyMemoryPointAPI(ftThis, lpPointAPI, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            x = .x
            y = .y
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

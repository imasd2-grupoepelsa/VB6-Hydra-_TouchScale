VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiLogPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2756
Attribute VB_Description = "This class provides the properties and methods for a logical pen description."
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for a logical pen description.

' ##MODULE_DESCRIPTION A logical pen is a description of a pen that _
is used to perform any line drawing type operations on a %device context:EventVB~ApiDeviceContext%.

' ##MODULE_DESCRIPTION Not every device can produce the exact pen as defined here in which case _
the system will approximate the nearest possible alternative.

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

'\\ member variables
Private mLogPen As LOGPEN
Private mHPEN As Long
Private mStock As Boolean

Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long

Private Declare Function GetObjectLOGPEN Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As LOGPEN) As Long

Public Enum PenStyles
    PS_SOLID = 0
    PS_DASH = 1                  '/* -------  */
    PS_DOT = 2                   '/* .......  */
    PS_DASHDOT = 3               '/* _._._._  */
    PS_DASHDOTDOT = 4            '/* _.._.._  */
    PS_NULL = 5
    PS_INSIDEFRAME = 6
    PS_USERSTYLE = 7
    PS_ALTERNATE = 8
End Enum

Private Declare Sub CopyMemoryLogPen Lib "kernel32" Alias "RtlMoveMemory" (Destination As LOGPEN, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrLogPen Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrLogPen Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private CreatedOK As Boolean

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this LogPen object from the location poiunted to by
'\\ lpLogPen
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpLogPen As Long) As Boolean

Dim lbThis As LOGPEN

CreatedOK = False

If Not IsBadReadPtrLogPen(lpLogPen, Len(lbThis)) Then
    Call CopyMemoryLogPen(lbThis, lpLogPen, Len(lbThis))
    If Err.LastDllError = 0 Then
        With lbThis
            mLogPen.lopnColor = .lopnColor
            mLogPen.lopnStyle = .lopnStyle
            mLogPen.lopnWidth.x = .lopnWidth.x
            mLogPen.lopnWidth.y = .lopnWidth.y
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Property Set Colour(ByVal newColour As ApiColour)
Attribute Colour.VB_HelpID = 2758

    mLogPen.lopnColor = newColour.ColourRef
    
End Property

Public Property Get Colour() As ApiColour

    Dim colThis As ApiColour
    
    Set colThis = New ApiColour
    
    colThis.ColourRef = mLogPen.lopnColor
    
    Set Colour = colThis
    
End Property

Public Property Get Description() As String
Attribute Description.VB_HelpID = 2759

Dim sdesc As String

sdesc = "LOGPEN:Style="
Select Case mLogPen.lopnStyle
Case PS_SOLID
    sdesc = sdesc & "SOLID"
Case PS_DASH                 '/* -------  */
    sdesc = sdesc & "DASH"
Case PS_DOT                    '/* .......  */
    sdesc = sdesc & "DOT"
Case PS_DASHDOT                '/* _._._._  */
    sdesc = sdesc & "DASHDOT"
Case PS_DASHDOTDOT               '/* _.._.._  */
    sdesc = sdesc & "DASHDOTDOT"
Case PS_NULL
    sdesc = sdesc & "NULL"
Case PS_INSIDEFRAME
    sdesc = sdesc & "INSIDEFRAME"
Case PS_USERSTYLE
    sdesc = sdesc & "USER"
Case PS_ALTERNATE
    sdesc = sdesc & "ALTERNATE"
End Select

sdesc = sdesc & ",Width=" & mLogPen.lopnWidth.x & "." & mLogPen.lopnWidth.y
sdesc = sdesc & ",Colour=" & Colour.Description

Description = sdesc

End Property

Friend Property Get Handle() As Long

    If mHPEN = 0 Then
        mHPEN = CreatePenIndirect(mLogPen)
    End If
    Handle = mHPEN
    
End Property

Friend Property Let Handle(ByVal newHandle As Long)

Dim lret As Long

    If newHandle <> mHPEN Then
        mHPEN = newHandle
        If newHandle <> 0 Then
            lret = GetObjectLOGPEN(newHandle, Len(mLogPen), mLogPen)
            If Err.LastDllError <> 0 Then
                ReportError Err.LastDllError, "ApiLogPEN:Handle (Let)", GetLastSystemError
            End If
        End If
    End If
    
End Property

Friend Property Get IsStockObject() As Boolean

    IsStockObject = mStock
    '\\ Note:  This will need to be amended to read from the
    '\\ GDI object table and return True if the stock object's
    '\\ owner process id is zero...
    
End Property
Friend Property Let IsStockObject(ByVal bIs As Boolean)

    mStock = bIs

End Property
Public Property Let Style(ByVal NewStyle As PenStyles)
Attribute Style.VB_HelpID = 2761

    mLogPen.lopnStyle = NewStyle
    
End Property

Public Property Get Style() As PenStyles

    Style = mLogPen.lopnStyle
    
End Property



Public Property Let Width(ByVal newWidth As Long)
Attribute Width.VB_HelpID = 2763

    mLogPen.lopnWidth.x = newWidth
    
End Property

Public Property Get Width() As Long

    Width = mLogPen.lopnWidth.x
    
End Property



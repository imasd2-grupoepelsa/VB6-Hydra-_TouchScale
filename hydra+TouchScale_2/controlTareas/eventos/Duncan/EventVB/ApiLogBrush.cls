VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiLogBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2764
Attribute VB_Description="This class provides the properties and methods for a logical brush."
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for a logical brush.

' ##MODULE_DESCRIPTION A logical brush is a description of a brush that _
is used to perform any painting type operations on a %device context:EventVB~ApiDeviceContext%.

' ##MODULE_DESCRIPTION Not every device can produce the exact brush as defined here in which case _
the system will approximate the nearest possible alternative.


Private Type LogBrush
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private mHBRUSH As Long
Private mLogBrush As LogBrush

Private mhBitmap As Long

' ##ENUMERATION_DESCRIPTION The type of brush
Public Enum BrushStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_SOLID The brush is a solid colour
    BS_SOLID = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_NULL The brush is empty
    BS_NULL = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_HATCHED The brush is a hatch pattern
    BS_HATCHED = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_PATTERN The brush is made of a device dependent bitmap
    BS_PATTERN = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_INDEXED
    BS_INDEXED = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_DIBPATTERN The brush is made of a device independent bitmap
    BS_DIBPATTERN = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_DIBPATTERNPT
    BS_DIBPATTERNPT = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_PATTERN8X8 The brush is an 8x8 device dependent bitmap
    BS_PATTERN8X8 = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_DIBPATTERN8X8 The brush is an 8x8 device independent bitmap
    BS_DIBPATTERN8X8 = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION BS_MONOPATTERN The brush is a monochrome (black & white) bitmap
    BS_MONOPATTERN = 9
End Enum

' ##ENUMERATION_DESCRIPTION If the brush style is BS_HATCHED then this describes _
the type of hatching the brush uses
Public Enum HatchStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION HS_HORIZONTAL Horizontal hatching lines
     HS_HORIZONTAL = 0            '/* ----- */
     ' ##ENUMERATION_MEMBER_DESCRIPTION HS_VERTICAL Vertical hatching lines
     HS_VERTICAL = 1              '/* ||||| */
     ' ##ENUMERATION_MEMBER_DESCRIPTION HS_FDIAGONAL Forward slanting diagonal hatching
     HS_FDIAGONAL = 2             '/* \\\\\ */
     ' ##ENUMERATION_MEMBER_DESCRIPTION HS_BDIAGONAL Back slanting diagonal hatch
     HS_BDIAGONAL = 3             '/* ///// */
     ' ##ENUMERATION_MEMBER_DESCRIPTION HS_CROSS Vertical-Horizontal cross hatching
     HS_CROSS = 4                 '/* +++++ */
     ' ##ENUMERATION_MEMBER_DESCRIPTION HS_DIAGCROSS Diagonal cross hatching
     HS_DIAGCROSS = 5             '/* xxxxx */
End Enum

'\\ Getting a LOGBRUSH from its handle
Private Declare Function GetObjectLOGBRUSH Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As LogBrush) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long

'\\ Creating picture type brushes....
Private Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long


Private mStock As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryLogBrush Lib "kernel32" Alias "RtlMoveMemory" (Destination As LogBrush, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrLogBrush Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrLogBrush Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private CreatedOK As Boolean

Private bChanged As Boolean

Public Property Set Bitmap(ByVal bmThis As ApiBitmap)
Attribute Bitmap.VB_HelpID = 2890

If bmThis Is Nothing Then
    If mhBitmap <> 0 Then
        mhBitmap = 0
        mHBRUSH = 0
        Changed = True
    End If
Else
    If mhBitmap <> bmThis.hBitmap Then
        mhBitmap = bmThis.hBitmap
        mHBRUSH = 0
        Changed = True
    End If
End If

End Property


Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this LogBrush object from the location poiunted to by
'\\ lpLogBrush
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpLogBrush As Long) As Boolean

Dim lbThis As LogBrush

CreatedOK = False

If Not IsBadReadPtrLogBrush(lpLogBrush, Len(lbThis)) Then
    Call CopyMemoryLogBrush(lbThis, lpLogBrush, Len(lbThis))
    If Err.LastDllError = 0 Then
        With lbThis
            mLogBrush.lbColor = .lbColor
            mLogBrush.lbHatch = .lbHatch
            mLogBrush.lbStyle = .lbStyle
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
Public Property Set Colour(ByVal newColour As ApiColour)
Attribute Colour.VB_HelpID = 2766

    mLogBrush.lbColor = newColour.ColourRef
    bChanged = True
    
End Property

Public Property Get Colour() As ApiColour

    Dim colThis As ApiColour
    
    Set colThis = New ApiColour
    colThis.ColourRef = mLogBrush.lbColor
    
End Property

Public Property Let Hatch(ByVal newhatch As HatchStyles)
Attribute Hatch.VB_HelpID = 2768

    mLogBrush.lbHatch = newhatch
    bChanged = True
    
End Property

Public Property Get Hatch() As HatchStyles

    Hatch = mLogBrush.lbHatch
    
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

Friend Property Get Handle() As Long

    If mHBRUSH = 0 Then
        If mhBitmap = 0 Then
            mHBRUSH = CreateBrushIndirect(mLogBrush)
        Else
            mHBRUSH = CreatePatternBrush(mhBitmap)
        End If
        If Err.LastDllError Then
            ReportError Err.LastDllError, "ApiLogBrush:Handle (get)", GetLastSystemError
        End If
    End If
    Handle = mHBRUSH
    
End Property

Friend Property Let Handle(ByVal newHandle As Long)

Dim lret As Long

    If newHandle <> mHBRUSH Then
        mHBRUSH = newHandle
        If newHandle <> 0 Then
            lret = GetObjectLOGBRUSH(newHandle, Len(mLogBrush), mLogBrush)
            If Err.LastDllError <> 0 Then
                ReportError Err.LastDllError, "ApiLogBRUSH:Handle (Let)", GetLastSystemError
            End If
        End If
    End If
    
End Property

Public Property Let Style(ByVal newStyle As BrushStyles)
Attribute Style.VB_HelpID = 2770

    mLogBrush.lbStyle = newStyle
    bChanged = True
    
End Property


Public Property Get Style() As BrushStyles

    Style = mLogBrush.lbStyle
    
End Property



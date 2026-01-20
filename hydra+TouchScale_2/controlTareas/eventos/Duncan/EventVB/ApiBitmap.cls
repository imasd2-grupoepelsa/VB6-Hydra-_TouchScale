VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2409
Attribute VB_Description = "This class returns information about and allows the manipulation of bitmaps"
Option Explicit

' ##MODULE_DESCRIPTION This class returns information about and allows the manipulation of bitmaps


Private mhBitmap As Long
Private mHDC As Long

Private mStockBitmap As Boolean

Private Declare Function DeleteObjectApi Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleBitmapApi Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As BitmapCompressionIndexes
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type Bitmap '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0) As RGBQUAD
End Type

Private Type DIBSECTION
    BM As Bitmap
    BMIH As BITMAPINFOHEADER
    colourmask_Red As Long
    colourmask_Green As Long
    colourmask_Blue As Long
    hSection As Long
    Offset As Long
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Byte, lpBI As BITMAPINFO, ByVal wUsage As DeviveIndependentBitmapPalleteStyles) As Long

' ##ENUMERATION_DESCRIPTION DeviceIndependentBitmapPalleteStyles - How the device independent bitmap deals with palettes
Public Enum DeviveIndependentBitmapPalleteStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION DIB_RGB_COLORS The colour table in absolute Red/Green/Blue values
    DIB_RGB_COLORS = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DIB_PAL_COLORS The colour table in palette indices
    DIB_PAL_COLORS = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DIB_PAL_INDICES (also DIB_PAL_PHYSINDICES) There are no colour table indices into surf palette
    DIB_PAL_INDICES = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DIB_PAL_LOGINDICES There are no colour table indices into the device context palette
    DIB_PAL_LOGINDICES = 4
End Enum

' ##ENUMERATION_DESCRIPTION BitmapCompressionIndexes - The type of compression used in this bitmap
Public Enum BitmapCompressionIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_RGB - The bitmap is not compressed at all
    BI_RGB = 0&
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_RLE4 - The bitmap is run length encoded at 4 bits per pixel
    BI_RLE4 = 2&
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_RLE8 - The bitmap is run length encoded at 8 bits per pixel
    BI_RLE8 = 1&
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_BITFIELDS - Uncompressed with the colour table three DWORD color masks that specify the red, green, and blue components of each pixel
    BI_BITFIELDS = 3&
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_JPEG - The bitmap is in compressed JPEG format
    BI_JPEG = 4&
    ' ##ENUMERATION_MEMBER_DESCRIPTION BI_PNG - The bitmap is compressed in PNG format
    BI_PNG = 5&
End Enum

' ##ENUMERATION_DESCRIPTION StandardImageTypes - The set of image types that can be loaded
Public Enum StandardImageTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION IMAGE_BITMAP - The image is a bitmap picture
    IMAGE_BITMAP = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION IMAGE_ICON - The image is an icon
    IMAGE_ICON = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION IMAGE_CURSOR - The image is a cursor
    IMAGE_CURSOR = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION IMAGE_ENHMETAFILE - The image is an enhanced metafile image
    IMAGE_ENHMETAFILE = 3
End Enum

Private Declare Function GetObjectDibSection Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As DIBSECTION) As Long
Private Declare Function GetObjectBitmap Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Bitmap) As Long


Private mIsDibSection As Boolean
Private mDS As DIBSECTION

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Byte, lpvSource As Long, ByVal cbCopy As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property
Public Property Get Bits() As Collection
Attribute Bits.VB_HelpID = 2410

Dim colThis As New Collection
Dim lBit As Long
Dim lBytesRequired As Long
Dim buffer() As Byte

With mDS.BM
    If .bmBits <> 0 Then
        '\\ Dimension the temporary array to hold the bitmap data
        lBytesRequired = (.bmHeight * .bmWidthBytes)
        ReDim Preserve buffer(1 To lBytesRequired) As Byte
        '\\ Get the data from the bitmap
        Call CopyMemory(buffer(1), .bmBits, lBytesRequired)
        
    End If
End With

End Property


Public Property Get BitsPerPixel() As Long
Attribute BitsPerPixel.VB_HelpID = 2411

BitsPerPixel = mDS.BM.bmBitsPixel

End Property

Public Property Get BytesPerScanline() As Long
Attribute BytesPerScanline.VB_HelpID = 2412

    BytesPerScanline = mDS.BM.bmWidthBytes

End Property

Public Property Get ColourPlanes() As Long
Attribute ColourPlanes.VB_HelpID = 2413

    ColourPlanes = mDS.BM.bmPlanes

End Property


'------------------------------------------------------------------------------
'  ##BLOCK_DESCRIPTION ConvertToMaskRegion takes the bitmap data and converts it
'  ##BLOCK_DESCRIPTION to a region where every pixel that is of the same colour as
'  ##BLOCK_DESCRIPTION the Mask Colour is excluded from the region.
'  ##BLOCK_REMARKS This function can be used to make windows that clip to the
'  ##BLOCK_REMARKS picture they contain, for example.
'------------------------------------------------------------------------------
Public Function ConvertToMaskRegion(ByVal TargetWindow As ApiWindow, ByVal MaskColour As ApiColour) As ApiRegion
Attribute ConvertToMaskRegion.VB_Description = "ConvertToMaskRegion takes the bitmap data and converts it"
Attribute ConvertToMaskRegion.VB_HelpID = 2414



End Function


Friend Sub CreateCompatibleBitmap(ByVal hdc As Long, ByVal Width As Long, ByVal Height As Long)

Dim lret As Long

lret = CreateCompatibleBitmapApi(hdc, Width, Height)
If lret = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "CreateCompatibleBitmap", GetLastSystemError)
Else
    mhBitmap = lret
End If

End Sub


Public Property Get hBitmap() As Long
Attribute hBitmap.VB_HelpID = 2415

    hBitmap = mhBitmap
    
End Property

Public Property Let hBitmap(ByVal newHandle As Long)

    If mhBitmap <> newHandle Then
        mhBitmap = newHandle
        Call Refresh
    End If
    
End Property

Friend Property Let hdc(ByVal newDC As Long)

mHDC = newDC

End Property

'\\ --[Height]------------------------------------------------------------------------------
'\\ Returns the height of a bitmap in logical pixels
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Property Get Height() As Long
Attribute Height.VB_HelpID = 2416

    Height = mDS.BM.bmHeight

End Property

Private Function Refresh()

Dim lret As Long

If mIsDibSection Then
    lret = GetObjectDibSection(mhBitmap, LenB(mDS), mDS)
Else
    lret = GetObjectBitmap(mhBitmap, Len(mDS.BM), mDS.BM)
End If
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiBitmap:Refresh", GetLastSystemError
End If

End Function

'\\ --[Width]------------------------------------------------------------------------------
'\\ Returns the width of a bitmap in logical pixels
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Property Get Width() As Long
Attribute Width.VB_HelpID = 2417
    
    Width = mDS.BM.bmWidth

End Property

Private Sub Class_Terminate()

Dim lret As Long

If mhBitmap > 0 And Not (mStockBitmap) Then
    lret = DeleteObjectApi(mhBitmap)
    If lret = 0 Or Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiBitmap:Terminate", GetLastSystemError)
    End If
End If

End Sub



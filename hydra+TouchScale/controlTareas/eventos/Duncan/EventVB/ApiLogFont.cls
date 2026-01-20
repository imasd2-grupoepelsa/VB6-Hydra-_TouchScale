VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiLogFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2709
Attribute VB_Description="This class provides the properties and methods for a logical font description."
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for a logical font description.

' ##MODULE_DESCRIPTION A logical font is a description of a font that _
is used to perform any text printing operations on a %device context:EventVB~ApiDeviceContext%.

' ##MODULE_DESCRIPTION Not every device can produce the exact font as defined here in which case _
the system will approximate the nearest possible alternative.


Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
            "CreateFontIndirectA" (lpLOGFONT As LOGFONT) As Long
            

Public Enum FontOutputPrecisions
    OUT_DEFAULT_PRECIS = 0
    OUT_STRING_PRECIS = 1
    OUT_CHARACTER_PRECIS = 2
    OUT_STROKE_PRECIS = 3
    OUT_TT_PRECIS = 4
    OUT_DEVICE_PRECIS = 5
    OUT_RASTER_PRECIS = 6
    OUT_TT_ONLY_PRECIS = 7
    OUT_OUTLINE_PRECIS = 8
    OUT_SCREEN_OUTLINE_PRECIS = 9
    OUT_PS_ONLY_PRECIS = 10
End Enum

Public Enum FontClipPrecisions
    CLIP_DEFAULT_PRECIS = 0
    CLIP_CHARACTER_PRECIS = 1
    CLIP_STROKE_PRECIS = 2
End Enum

Public Enum FontQualitySettings
    DEFAULT_QUALITY = 0
    DRAFT_QUALITY = 1
    PROOF_QUALITY = 2
    NONANTIALIASED_QUALITY = 3
    ANTIALIASED_QUALITY = 4
End Enum

Public Enum FontFamilies
    FF_DONTCARE = 0    '  Don't care or don't know.
    FF_ROMAN = 16      '  Variable stroke width, serifed.
    FF_SWISS = 32      '  Variable stroke width, sans-serifed.
    FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.
    FF_SCRIPT = 64     '  Cursive, etc.
    FF_DECORATIVE = 80 '  Old English, etc.
End Enum

Public Enum FontPitchSettings
    DEFAULT_PITCH = 0
    FIXED_PITCH = 1
    VARIABLE_PITCH = 2
    MONO_FONT = 8
End Enum

Public Enum FontWeights
     FW_DONTCARE = 0
     FW_THIN = 100
     FW_EXTRALIGHT = 200
     FW_LIGHT = 300
     FW_NORMAL = 400
     FW_MEDIUM = 500
     FW_SEMIBOLD = 600
     FW_BOLD = 700
     FW_EXTRABOLD = 800
     FW_HEAVY = 900
End Enum

Public Enum FontCharsets
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1
    SYMBOL_CHARSET = 2
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    HANGUL_CHARSET = 129
    GB2312_CHARSET = 134
    CHINESEBIG5_CHARSET = 136
    OEM_CHARSET = 255
    JOHAB_CHARSET = 130
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    VIETNAMESE_CHARSET = 163
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    RUSSIAN_CHARSET = 204
    MAC_CHARSET = 77
    BALTIC_CHARSET = 186
End Enum


'\\ -- Private member variables....
Private mLogFont As LOGFONT
Private mHFONT As Long

'\\ Private memory handling functions
Private Declare Sub CopyMemoryLOGFONT Lib "kernel32" Alias "RtlMoveMemory" (Destination As LOGFONT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrLOGFONT Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrLOGFONT Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Public CreatedOK As Boolean

'\\ Getting a LOGFONT from its handle
Private Declare Function GetObjectLOGFONT Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As LOGFONT) As Long

Private mStock As Boolean

Private bChanged As Boolean
Attribute CreatedOK.VB_VarHelpID = 2710
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Let Bold(ByVal newBold As Boolean)
Attribute Bold.VB_HelpID = 2712

    If newBold Then
        If mLogFont.lfWeight < FW_BOLD Then
            mLogFont.lfWeight = FW_BOLD
        End If
    Else
        If mLogFont.lfWeight >= FW_BOLD Then
            mLogFont.lfWeight = FW_NORMAL
        End If
    End If
    
End Property

Public Property Get Bold() As Boolean

    Bold = (mLogFont.lfWeight >= FW_BOLD)
    
End Property

Public Property Let Charset(ByVal newCharset As FontCharsets)
Attribute Charset.VB_HelpID = 2714

    mLogFont.lfCharSet = newCharset
    
End Property

Public Property Get Charset() As FontCharsets

    Charset = mLogFont.lfCharSet
    
End Property

Public Property Let ClipPrecision(ByVal newPrecision As FontClipPrecisions)
Attribute ClipPrecision.VB_HelpID = 2716

    mLogFont.lfClipPrecision = newPrecision
    
End Property

Public Property Get ClipPrecision() As FontClipPrecisions

    ClipPrecision = mLogFont.lfClipPrecision
    
End Property

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this LOGFONT object from the location pointed to by
'\\ lpLOGFONT
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpLOGFONT As Long) As Boolean

Dim ftThis As LOGFONT
Dim nIndex As Long

CreatedOK = False

If Not IsBadReadPtrLOGFONT(lpLOGFONT, Len(ftThis)) Then
    Call CopyMemoryLOGFONT(ftThis, lpLOGFONT, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            mLogFont.lfCharSet = .lfCharSet
            mLogFont.lfCharSet = .lfClipPrecision
            mLogFont.lfEscapement = .lfEscapement
            For nIndex = 0 To LF_FACESIZE
                mLogFont.lfFaceName(nIndex) = .lfFaceName(nIndex)
            Next nIndex
            mLogFont.lfHeight = .lfHeight
            mLogFont.lfItalic = .lfItalic
            mLogFont.lfOrientation = .lfOrientation
            mLogFont.lfOutPrecision = .lfOutPrecision
            mLogFont.lfPitchAndFamily = .lfPitchAndFamily
            mLogFont.lfQuality = .lfQuality
            mLogFont.lfStrikeOut = .lfStrikeOut
            mLogFont.lfUnderline = .lfUnderline
            mLogFont.lfWeight = .lfWeight
            mLogFont.lfWidth = .lfWidth
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
Public Property Let Escapement(ByVal newEscapement As Long)
Attribute Escapement.VB_HelpID = 2718

    mLogFont.lfEscapement = newEscapement
    
End Property

Public Property Get Escapement() As Long

    Escapement = mLogFont.lfEscapement
    
End Property


Public Property Let FaceName(ByVal newname As String)
Attribute FaceName.VB_HelpID = 2720

    Dim nIndex
    
    For nIndex = 0 To LF_FACESIZE
        If Len(newname) > nIndex Then
            mLogFont.lfFaceName(nIndex) = Asc(Mid$(newname, nIndex + 1, 1))
        Else
            '\\ Blank out any existing characters from a bigger font name
            mLogFont.lfFaceName(nIndex) = 0
        End If
    Next nIndex
    
End Property

Public Property Get FaceName() As String

    Dim nIndex As Long
    Dim sRet As String
    
    For nIndex = 0 To LF_FACESIZE
        sRet = sRet & Chr$(mLogFont.lfFaceName(nIndex))
    Next nIndex
    
    If InStr(sRet, Chr$(0)) Then
        sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
    
    FaceName = sRet
    
End Property

Public Property Let Family(ByVal newFamily As FontFamilies)
Attribute Family.VB_HelpID = 2722

    mLogFont.lfPitchAndFamily = Pitch Or newFamily
    
End Property

Public Property Get Family() As FontFamilies

    Family = (mLogFont.lfPitchAndFamily And Not 7)
    
End Property

Friend Property Let Handle(ByVal newhandle As Long)

Dim lRet As Long

    If newhandle <> mHFONT Then
        mHFONT = newhandle
        If newhandle <> 0 Then
            lRet = GetObjectLOGFONT(newhandle, Len(mLogFont), mLogFont)
            If Err.LastDllError <> 0 Then
                ReportError Err.LastDllError, "ApiLogFont:Handle (Let)", GetLastSystemError
            End If
        End If
    End If
    
End Property

Friend Property Get Handle() As Long

    If mHFONT = 0 Then
        mHFONT = CreateFontIndirect(mLogFont)
    End If
    Handle = mHFONT
    
End Property

Public Property Let Height(ByVal newHeight As Long)
Attribute Height.VB_HelpID = 2724

    mLogFont.lfHeight = newHeight
    
End Property


Public Property Get Height() As Long

    Height = mLogFont.lfHeight
    
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

Public Property Let Italic(ByVal bItalic As Boolean)
Attribute Italic.VB_HelpID = 2726

    mLogFont.lfItalic = bItalic
    
End Property

Public Property Get Italic() As Boolean

    Italic = CBool(mLogFont.lfItalic)
    
End Property

Public Property Let Orientation(ByVal newOrientation As Long)
Attribute Orientation.VB_HelpID = 2728

If mLogFont.lfOrientation <> newOrientation Then
    mLogFont.lfOrientation = newOrientation
    mHFONT = 0
End If

End Property

Public Property Get Orientation() As Long

    Orientation = mLogFont.lfOrientation
    
End Property

Public Property Let OutputPrecision(ByVal newPrecision As FontOutputPrecisions)
Attribute OutputPrecision.VB_HelpID = 2730

    mLogFont.lfOutPrecision = newPrecision
    
End Property

Public Property Get OutputPrecision() As FontOutputPrecisions

    OutputPrecision = mLogFont.lfOutPrecision
    
End Property

Public Property Let Pitch(ByVal newPitch As FontPitchSettings)
Attribute Pitch.VB_HelpID = 2732

    mLogFont.lfPitchAndFamily = newPitch Or Family
    
End Property

Public Property Get Pitch() As FontPitchSettings

    '\\ Pitch is held in bits 0,1,2 of the LOGFONT Pitch and Family
    '\\ member, thus and it with 00000111 or 9 in decimals
    Pitch = (mLogFont.lfPitchAndFamily And 7)
    
End Property

Public Property Let Quality(ByVal newQuality As FontQualitySettings)
Attribute Quality.VB_HelpID = 2734

    mLogFont.lfQuality = newQuality
    
End Property

Public Property Get Quality() As FontQualitySettings

    Quality = mLogFont.lfQuality
    
End Property

Public Property Let StrikeOut(ByVal newStrikeout As Boolean)
Attribute StrikeOut.VB_HelpID = 2736

    mLogFont.lfStrikeOut = newStrikeout
    
End Property

Public Property Get StrikeOut() As Boolean

    StrikeOut = mLogFont.lfStrikeOut
    
End Property

Public Property Let Underline(ByVal newUnderline As Boolean)
Attribute Underline.VB_HelpID = 2738

    mLogFont.lfUnderline = newUnderline
    
End Property

Public Property Get Underline() As Boolean

    Underline = CBool(mLogFont.lfUnderline)
    
End Property

Public Property Let Weight(ByVal newWeight As Long)
Attribute Weight.VB_HelpID = 2740

    mLogFont.lfWeight = newWeight
    
End Property

Public Property Get Weight() As Long

    Weight = mLogFont.lfWeight
    
End Property

Public Property Let Width(ByVal newWidth As Long)
Attribute Width.VB_HelpID = 2742

    mLogFont.lfWidth = newWidth
    
End Property

Public Property Get Width() As Long

    Width = mLogFont.lfWidth
    
End Property



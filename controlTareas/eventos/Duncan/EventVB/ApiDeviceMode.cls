VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiDeviceMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 3114
Attribute VB_Description="This class encapsulates the functionality to get and set the device mode (or initialisation and environment) of a device - typically a printer"
Option Explicit

Private Const CLSNAME = "ApiDeviceMode"

' ##MODULE_DESCRIPTION This class encapsulates the functionality to get and set the _
device mode (or initialisation and environment) of a device - typically a printer

' ##ENUMERATION_DESCRIPTION DeviceOrientations describe the orientation of the device
Public Enum DeviceOrientations
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMORIENT_PORTRAIT The device is in portait (short edge on top) mode
    DMORIENT_PORTRAIT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMORIENT_LANDSCAPE The device is in landscape (long edge on top) mode
    DMORIENT_LANDSCAPE = 2
End Enum

' ##ENUMERATION_DESCRIPTION PrinterPaperSizes are the predefined paper types.  Typically a printer will only support a very small subset of these.
Public Enum PrinterPaperSizes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER Letter 8 1/2 x 11 in
    DMPAPER_LETTER = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTERSMALL Letter Small 8 1/2 x 11 in
    DMPAPER_LETTERSMALL = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_TABLOID Tabloid 11 x 17 in
    DMPAPER_TABLOID = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LEDGER Ledger 17 x 11 in
    DMPAPER_LEDGER = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LEGAL Legal 8 1/2 x 14 in
    DMPAPER_LEGAL = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_STATEMENT Statement 5 1/2 x 8 1/2 in
    DMPAPER_STATEMENT = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_EXECUTIVE Executive 7 1/4 x 10 1/2 in
    DMPAPER_EXECUTIVE = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A3 A3 297 x 420 mm
    DMPAPER_A3 = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4 A4 210 x 297 mm
    DMPAPER_A4 = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4SMALL A4 Small 210 x 297 mm
    DMPAPER_A4SMALL = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A5 A5 148 x 210 mm
    DMPAPER_A5 = 11
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B4 B4 (JIS) 250 x 354
    DMPAPER_B4 = 12
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B5 B5 (JIS) 182 x 257 mm
    DMPAPER_B5 = 13
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_FOLIO Folio 8 1/2 x 13 in
    DMPAPER_FOLIO = 14
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_QUARTO Quarto 215 x 275 mm
    DMPAPER_QUARTO = 15
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_10X14 10x14 in
    DMPAPER_10X14 = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_11X17 11x17 in
    DMPAPER_11X17 = 17
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_NOTE Note 8 1/2 x 11 in
    DMPAPER_NOTE = 18
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_9 Envelope #9 3 7/8 x 8 7/8
    DMPAPER_ENV_9 = 19
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_10 Envelope #10 4 1/8 x 9 1/2
    DMPAPER_ENV_10 = 20
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_11 Envelope #11 4 1/2 x 10 3/8
    DMPAPER_ENV_11 = 21
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_12 Envelope #12 4 \276 x 11
    DMPAPER_ENV_12 = 22
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_14 Envelope #14 5 x 11 1/2
    DMPAPER_ENV_14 = 23
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_CSHEET C size sheet
    DMPAPER_CSHEET = 24
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_DSHEET D size sheet
    DMPAPER_DSHEET = 25
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ESHEET E size sheet
    DMPAPER_ESHEET = 26
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_DL Envelope DL 110 x 220mm
    DMPAPER_ENV_DL = 27
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_C5 Envelope C5 162 x 229 mm
    DMPAPER_ENV_C5 = 28
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_C3 Envelope C3  324 x 458 mm
    DMPAPER_ENV_C3 = 29
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_C4 Envelope C4  229 x 324 mm
    DMPAPER_ENV_C4 = 30
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_C6 Envelope C6  114 x 162 mm
    DMPAPER_ENV_C6 = 31
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_C65 Envelope C65 114 x 229 mm
    DMPAPER_ENV_C65 = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_B4 Envelope B4  250 x 353 mm
    DMPAPER_ENV_B4 = 33
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_B5 Envelope B5  176 x 250 mm
    DMPAPER_ENV_B5 = 34
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_B6 Envelope B6  176 x 125 mm
    DMPAPER_ENV_B6 = 35
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_ITALY Envelope 110 x 230 mm
    DMPAPER_ENV_ITALY = 36
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_MONARCH Envelope Monarch 3.875 x 7.5 in
    DMPAPER_ENV_MONARCH = 37
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_PERSONAL 6 3/4 Envelope 3 5/8 x 6 1/2 in
    DMPAPER_ENV_PERSONAL = 38
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_FANFOLD_US US Std Fanfold 14 7/8 x 11 in
    DMPAPER_FANFOLD_US = 39
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_FANFOLD_STD_GERMAN German Std Fanfold 8 1/2 x 12 in
    DMPAPER_FANFOLD_STD_GERMAN = 40
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_FANFOLD_LGL_GERMAN German Legal Fanfold 8 1/2 x 13 in
    DMPAPER_FANFOLD_LGL_GERMAN = 41
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ISO_B4 B4 (ISO) 250 x 353 mm
    DMPAPER_ISO_B4 = 42
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JAPANESE_POSTCARD Japanese Postcard 100 x 148 mm
    DMPAPER_JAPANESE_POSTCARD = 43
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_9X11 9 x 11 in
    DMPAPER_9X11 = 44
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_10X11 10 x 11 in
    DMPAPER_10X11 = 45
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_15X11 15 x 11 in
    DMPAPER_15X11 = 46
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_ENV_INVITE Envelope Invite 220 x 220 mm
    DMPAPER_ENV_INVITE = 47
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER_EXTRA Letter Extra 9 \275 x 12 in
    DMPAPER_LETTER_EXTRA = 50
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LEGAL_EXTRA Legal Extra 9 \275 x 15 in
    DMPAPER_LEGAL_EXTRA = 51
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_TABLOID_EXTRA Tabloid Extra 11.69 x 18 in
    DMPAPER_TABLOID_EXTRA = 52
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4_EXTRA A4 Extra 9.27 x 12.69 in
    DMPAPER_A4_EXTRA = 53
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER_TRANSVERSE Letter Transverse 8 \275 x 11 in
    DMPAPER_LETTER_TRANSVERSE = 54
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4_TRANSVERSE A4 Transverse 210 x 297 mm
    DMPAPER_A4_TRANSVERSE = 55
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER_EXTRA_TRANSVERSE Letter Extra Transverse 9\275 x 12 in
    DMPAPER_LETTER_EXTRA_TRANSVERSE = 56
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A_PLUS SuperA/SuperA/A4 227 x 356 mm
    DMPAPER_A_PLUS = 57
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B_PLUS SuperB/SuperB/A3 305 x 487 mm
    DMPAPER_B_PLUS = 58
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER_PLUS Letter Plus 8.5 x 12.69 in
    DMPAPER_LETTER_PLUS = 59
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4_PLUS A4 Plus 210 x 330 mm
    DMPAPER_A4_PLUS = 60
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A5_TRANSVERSE A5 Transverse 148 x 210 mm
    DMPAPER_A5_TRANSVERSE = 61
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B5_TRANSVERSE B5 (JIS) Transverse 182 x 257 mm
    DMPAPER_B5_TRANSVERSE = 62
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A3_EXTRA A3 Extra 322 x 445 mm
    DMPAPER_A3_EXTRA = 63
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A5_EXTRA A5 Extra 174 x 235 mm
    DMPAPER_A5_EXTRA = 64
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B5_EXTRA B5 (ISO) Extra 201 x 276 mm
    DMPAPER_B5_EXTRA = 65
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A2 A2 420 x 594 mm
    DMPAPER_A2 = 66
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A3_TRANSVERSE A3 Transverse 297 x 420 mm
    DMPAPER_A3_TRANSVERSE = 67
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A3_EXTRA_TRANSVERSE A3 Extra Transverse 322 x 445 mm
    DMPAPER_A3_EXTRA_TRANSVERSE = 68
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_DBL_JAPANESE_POSTCARD Japanese Double Postcard 200 x 148 mm
    DMPAPER_DBL_JAPANESE_POSTCARD = 69
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A6 A6 105 x 148 mm
    DMPAPER_A6 = 70
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_KAKU2 Japanese Envelope Kaku #2
    DMPAPER_JENV_KAKU2 = 71
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_KAKU3 Japanese Envelope Kaku #3
    DMPAPER_JENV_KAKU3 = 72
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_CHOU3 Japanese Envelope Chou #3
    DMPAPER_JENV_CHOU3 = 73
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_CHOU4 Japanese Envelope Chou #4
    DMPAPER_JENV_CHOU4 = 74
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_LETTER_ROTATED Letter Rotated 11 x 8 1/2 11 in
    DMPAPER_LETTER_ROTATED = 75
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A3_ROTATED A3 Rotated 420 x 297 mm
    DMPAPER_A3_ROTATED = 76
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A4_ROTATED A4 Rotated 297 x 210 mm
    DMPAPER_A4_ROTATED = 77
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A5_ROTATED A5 Rotated 210 x 148 mm
    DMPAPER_A5_ROTATED = 78
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B4_JIS_ROTATED B4 (JIS) Rotated 364 x 257 mm
    DMPAPER_B4_JIS_ROTATED = 79
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B5_JIS_ROTATED B5 (JIS) Rotated 257 x 182 mm
    DMPAPER_B5_JIS_ROTATED = 80
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JAPANESE_POSTCARD_ROTATED Japanese Postcard Rotated 148 x 100 mm
    DMPAPER_JAPANESE_POSTCARD_ROTATED = 81
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED Double Japanese Postcard Rotated 148 x 200 mm
    DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_A6_ROTATED A6 Rotated 148 x 105 mm
    DMPAPER_A6_ROTATED = 83
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_KAKU2_ROTATED Japanese Envelope Kaku #2 Rotated
    DMPAPER_JENV_KAKU2_ROTATED = 84
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_KAKU3_ROTATED Japanese Envelope Kaku #3 Rotated
    DMPAPER_JENV_KAKU3_ROTATED = 85
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_CHOU3_ROTATED Japanese Envelope Chou #3 Rotated
    DMPAPER_JENV_CHOU3_ROTATED = 86
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_CHOU4_ROTATED Japanese Envelope Chou #4 Rotated
    DMPAPER_JENV_CHOU4_ROTATED = 87
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B6_JIS B6 (JIS) 128 x 182 mm
    DMPAPER_B6_JIS = 88
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_B6_JIS_ROTATED B6 (JIS) Rotated 182 x 128 mm
    DMPAPER_B6_JIS_ROTATED = 89
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_12X11 12 x 11 in
    DMPAPER_12X11 = 90
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_YOU4 Japanese Envelope You #4
    DMPAPER_JENV_YOU4 = 91
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_JENV_YOU4_ROTATED Japanese Envelope You #4 Rotated
    DMPAPER_JENV_YOU4_ROTATED = 92
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P16K PRC 16K 146 x 215 mm
    DMPAPER_P16K = 93
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P32K PRC 32K 97 x 151 mm
    DMPAPER_P32K = 94
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P32KBIG PRC 32K(Big) 97 x 151 mm
    DMPAPER_P32KBIG = 95
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_1 PRC Envelope #1 102 x 165 mm
    DMPAPER_PENV_1 = 96
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_2 PRC Envelope #2 102 x 176 mm
    DMPAPER_PENV_2 = 97
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_3 PRC Envelope #3 125 x 176 mm
    DMPAPER_PENV_3 = 98
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_4 PRC Envelope #4 110 x 208 mm
    DMPAPER_PENV_4 = 99
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_5 PRC Envelope #5 110 x 220 mm
    DMPAPER_PENV_5 = 100
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_6 PRC Envelope #6 120 x 230 mm
    DMPAPER_PENV_6 = 101
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_7 PRC Envelope #7 160 x 230 mm
    DMPAPER_PENV_7 = 102
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_8 PRC Envelope #8 120 x 309 mm
    DMPAPER_PENV_8 = 103
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_9 PRC Envelope #9 229 x 324 mm
    DMPAPER_PENV_9 = 104
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_10 PRC Envelope #10 324 x 458 mm
    DMPAPER_PENV_10 = 105
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P16K_ROTATED PRC 16K Rotated
    DMPAPER_P16K_ROTATED = 106
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P32K_ROTATED PRC 32K Rotated
    DMPAPER_P32K_ROTATED = 107
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_P32KBIG_ROTATED PRC 32K(Big) Rotated
    DMPAPER_P32KBIG_ROTATED = 108
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_1_ROTATED PRC Envelope #1 Rotated 165 x 102 mm
    DMPAPER_PENV_1_ROTATED = 109
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_2_ROTATED PRC Envelope #2 Rotated 176 x 102 mm
    DMPAPER_PENV_2_ROTATED = 110
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_3_ROTATED PRC Envelope #3 Rotated 176 x 125 mm
    DMPAPER_PENV_3_ROTATED = 111
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_4_ROTATED PRC Envelope #4 Rotated 208 x 110 mm
    DMPAPER_PENV_4_ROTATED = 112
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_5_ROTATED PRC Envelope #5 Rotated 220 x 110 mm
    DMPAPER_PENV_5_ROTATED = 113
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_6_ROTATED PRC Envelope #6 Rotated 230 x 120 mm
    DMPAPER_PENV_6_ROTATED = 114
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_7_ROTATED PRC Envelope #7 Rotated 230 x 160 mm
    DMPAPER_PENV_7_ROTATED = 115
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_8_ROTATED PRC Envelope #8 Rotated 309 x 120 mm
    DMPAPER_PENV_8_ROTATED = 116
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_9_ROTATED PRC Envelope #9 Rotated 324 x 229 mm
    DMPAPER_PENV_9_ROTATED = 117
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMPAPER_PENV_10_ROTATED PRC Envelope #10 Rotated 458 x 324 mm
    DMPAPER_PENV_10_ROTATED = 118
End Enum

' ##ENUMERATION_DESCRIPTION DeviceModeResolutions are the different resolution settings for devices that _
support multiple resolutions
Public Enum DeviceModeResolutions
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMRES_DRAFT Device is set to draft quality
    DMRES_DRAFT = -1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMRES_LOW Device is set to low resolution quality
    DMRES_LOW = -2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMRES_MEDIUM Device is set to an intermediate resolution quality
    DMRES_MEDIUM = -3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMRES_HIGH Devide is set to high resolution quality
    DMRES_HIGH = -4
End Enum

' ##ENUMERATION_DESCRIPTION DeviceColourModes are the possible colour use settings for colour devices
Public Enum DeviceColourModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMCOLOR_MONOCHROME Do not use colour
    DMCOLOR_MONOCHROME = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMCOLOR_COLOR Do use colour
    DMCOLOR_COLOR = 2
End Enum

' ##ENUMERATION_DESCRIPTION DeviceDuplexSettings is how a duplex device is set for duplex printing
Public Enum DeviceDuplexSettings
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMDUP_SIMPLEX Device is not using double sided printing
     DMDUP_SIMPLEX = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMDUP_VERTICAL Device is performing double sided printing on the short edge
     DMDUP_VERTICAL = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMDUP_HORIZONTAL Device is performing double sided printing on the long edge
     DMDUP_HORIZONTAL = 3
End Enum

' ##ENUMERATION_DESCRIPTION DeviceTrueTypeFontSettings specifies how the device handles TrueType fonts
Public Enum DeviceTrueTypeFontSettings
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMTT_BITMAP Device treats TrueType fonts as bitmaps
     DMTT_BITMAP = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMTT_DOWNLOAD Device downloads TrueType fonts as soft fonts
     DMTT_DOWNLOAD = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMTT_SUBDEV Device substitutes TrueType fonts for the nearest printer font
     DMTT_SUBDEV = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMTT_DOWNLOAD_OUTLINE Device downloads TrueType fonts as outline soft fonts
     DMTT_DOWNLOAD_OUTLINE = 4
End Enum

' ##ENUMERATION_DESCRIPTION DeviceCollateSettings set whether or not multi-copy print jobs should be collated
Public Enum DeviceCollateSettings
    ' ##ENUMERATION_MEMBER_DESCRIPTION DMCOLLATE_FALSE Do not collate the jobs
     DMCOLLATE_FALSE = 0
     ' ##ENUMERATION_MEMBER_DESCRIPTION DMCOLLATE_TRUE Do collate the jobs
     DMCOLLATE_TRUE = 1
End Enum

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'\\ For windows version 4 and higher
Private Type DEVMODE_4
    dmBase As DEVMODE
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type

'\\ For windows Version 5 and higher
Private Type DEVMODE_5
    dmBase4 As DEVMODE_4
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

'\\ Private member variables
Private mChanged As Boolean
Private mlpDevMode As Long
Private mDevMode As DEVMODE_5

'\\ API decls to copy this structure to/from memory
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long

Private Declare Sub CopyMemoryDMBase Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEVMODE, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromDMBase Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As DEVMODE, ByVal Length As Long)
Private Declare Sub CopyMemoryDM4 Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEVMODE_4, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromDM4 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As DEVMODE_4, ByVal Length As Long)
Private Declare Sub CopyMemoryDM5 Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEVMODE_5, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromDM5 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As DEVMODE_5, ByVal Length As Long)

Public Event Changed()
Attribute Changed.VB_HelpID = 3115

Friend Property Let Changed(ByVal bnewVal As Boolean)

    mChanged = bnewVal
    If mChanged Then
        If mlpDevMode <> 0 Then
            '\\ Save the devmode structure change...
            If CopyToPointer(mlpDevMode) Then
                mChanged = False
            End If
        End If
        RaiseEvent Changed
    End If
    
End Property


Friend Property Get Changed() As Boolean

    Changed = mChanged
    
End Property

Public Property Let ColourMode(ByVal newval As DeviceColourModes)
Attribute ColourMode.VB_HelpID = 3117
Attribute ColourMode.VB_Description = "The mode that the colour printer is set to  "

    If mDevMode.dmBase4.dmBase.dmColor <> newval Then
        mDevMode.dmBase4.dmBase.dmColor = newval
        Changed = True
    End If
    
End Property

' ##BLOCK_DESCRIPTION The mode that the colour printer is set to
' ##BLOCK_REMARKS This member has no meaning for monochrome devices
Public Property Get ColourMode() As DeviceColourModes

    ColourMode = mDevMode.dmBase4.dmBase.dmColor
    
End Property


Public Property Let Copies(ByVal newValue As Integer)
Attribute Copies.VB_HelpID = 3119
Attribute Copies.VB_Description = "The number of copies to print  "

If mDevMode.dmBase4.dmBase.dmCopies <> newValue Then
    mDevMode.dmBase4.dmBase.dmCopies = newValue
    Changed = True
End If

End Property

' ##BLOCK_DESCRIPTION The number of copies to print
' ##BLOCK_REMARKS Not all printers support multiple copy printing, in which case you will have to spool the same job a number of times
Public Property Get Copies() As Integer

    Copies = mDevMode.dmBase4.dmBase.dmCopies
    
End Property

Friend Function CopyToPointer(lpDevMode As Long) As Boolean

If Not IsBadWritePtr(lpDevMode, mDevMode.dmBase4.dmBase.dmSize) Then
    If mDevMode.dmBase4.dmBase.dmSize = LenB(mDevMode) Then
        Call CopyMemoryFromDM5(lpDevMode, mDevMode, LenB(mDevMode))
    ElseIf mDevMode.dmBase4.dmBase.dmSize = LenB(mDevMode.dmBase4) Then
        Call CopyMemoryFromDM4(lpDevMode, mDevMode.dmBase4, LenB(mDevMode.dmBase4))
    Else
        Call CopyMemoryFromDMBase(lpDevMode, mDevMode.dmBase4.dmBase, LenB(mDevMode.dmBase4.dmBase))
    End If
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":CopyToPointer", GetLastSystemError
    Else
        CopyToPointer = True
    End If
End If

End Function

' ##BLOCK_DESCRIPTION The "friendly" name of the device
' ##BLOCK_REMARKS This member may be truncated if more that 32 chars long
Public Property Get DeviceName() As String
Attribute DeviceName.VB_HelpID = 3120
Attribute DeviceName.VB_Description = "The "friendly" name of the device"

    DeviceName = Trim$(mDevMode.dmBase4.dmBase.dmDeviceName)

End Property

Public Property Let Duplex(ByVal newval As DeviceDuplexSettings)
Attribute Duplex.VB_HelpID = 3122
Attribute Duplex.VB_Description = "The duplex setting of the device  "

If mDevMode.dmBase4.dmBase.dmDuplex <> newval Then
    mDevMode.dmBase4.dmBase.dmDuplex = newval
    Changed = True
End If

End Property

' ##BLOCK_DESCRIPTION The duplex setting of the device
' ##BLOCK_REMARKS DMDUP_SIMPLEX means the device is not using double sided printing
' ##BLOCK_REMARKS DMDUP_VERTICAL means the device is performing double sided printing flipped on the short edge
' ##BLOCK_REMARKS DMDUP_HORIZONTAL means the device is performing double sided printing flipped on the long edge
Public Property Get Duplex() As DeviceDuplexSettings

    Duplex = mDevMode.dmBase4.dmBase.dmDuplex
    
End Property

' ##BLOCK_DESCRIPTION The name of the form used to print - e.g. "Letter"
' ##BLOCK_REMARKS This is only valid for printers running on Windows NT and above
Public Property Get FormName() As String
Attribute FormName.VB_HelpID = 3123
Attribute FormName.VB_Description = "The name of the form used to print - e.g. "Letter""

    FormName = Trim$(mDevMode.dmBase4.dmBase.dmFormName)
    
End Property

Public Property Let Orientation(ByVal newValue As DeviceOrientations)
Attribute Orientation.VB_HelpID = 3125
Attribute Orientation.VB_Description = "The orientation of the device  "

    If newValue <> mDevMode.dmBase4.dmBase.dmOrientation Then
        mDevMode.dmBase4.dmBase.dmOrientation = newValue
        Changed = True
    End If
    
End Property

' ##BLOCK_DESCRIPTION The orientation of the device
' ##BLOCK_REMARKS Not all devices support portrait orientation
Public Property Get Orientation() As DeviceOrientations

    Orientation = mDevMode.dmBase4.dmBase.dmOrientation
    
End Property


Public Property Let PaperLength(ByVal newValue As Integer)
Attribute PaperLength.VB_HelpID = 3127
Attribute PaperLength.VB_Description = "The length of the paper that the device is set to print  "

If newValue <> mDevMode.dmBase4.dmBase.dmPaperLength Then
    mDevMode.dmBase4.dmBase.dmPaperLength = newValue
    Changed = True
End If

End Property

' ##BLOCK_DESCRIPTION The length of the paper that the device is set to print
' ##BLOCK_REMARKS Measured in tenths of a millimeter
Public Property Get PaperLength() As Integer

    PaperLength = mDevMode.dmBase4.dmBase.dmPaperLength
    
End Property

Public Property Let PaperSize(ByVal newValue As PrinterPaperSizes)
Attribute PaperSize.VB_HelpID = 3129

If newValue <> mDevMode.dmBase4.dmBase.dmPaperSize Then
    mDevMode.dmBase4.dmBase.dmPaperSize = newValue
    Changed = True
End If

End Property

Friend Function CreateFromPointer(lpDevMode As Long) As Boolean

If Not IsBadReadPtr(lpDevMode, Len(mDevMode.dmBase4.dmBase)) Then
    '\\ Keep a copy of the pointer for saving...
    mlpDevMode = lpDevMode
    Call CopyMemoryDMBase(mDevMode.dmBase4.dmBase, lpDevMode, Len(mDevMode.dmBase4.dmBase))
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":CreateFromPointer", GetLastSystemError
    Else
        If mDevMode.dmBase4.dmBase.dmSize = LenB(mDevMode.dmBase4) Then
            Call CopyMemoryDM4(mDevMode.dmBase4, lpDevMode, Len(mDevMode.dmBase4))
        ElseIf mDevMode.dmBase4.dmBase.dmSize = LenB(mDevMode) Then
            Call CopyMemoryDM5(mDevMode, lpDevMode, Len(mDevMode))
        End If
        CreateFromPointer = True
    End If
End If

End Function

Public Property Get PaperSize() As PrinterPaperSizes

    PaperSize = mDevMode.dmBase4.dmBase.dmPaperSize
    
End Property


Public Property Let PaperWidth(ByVal newValue As Integer)
Attribute PaperWidth.VB_HelpID = 3131
Attribute PaperWidth.VB_Description = "The width of the paper that the device is set to print  "

If newValue <> mDevMode.dmBase4.dmBase.dmPaperWidth Then
    mDevMode.dmBase4.dmBase.dmPaperWidth = newValue
    Changed = True
End If

End Property

' ##BLOCK_DESCRIPTION The width of the paper that the device is set to print
' ##BLOCK_REMARKS Measured in tenths of a millimeter
Public Property Get PaperWidth() As Integer

    PaperWidth = mDevMode.dmBase4.dmBase.dmPaperWidth
    
End Property


Public Property Let ScaleFactor(ByVal newValue As Integer)
Attribute ScaleFactor.VB_HelpID = 3133
Attribute ScaleFactor.VB_Description = "The scaling factor between the logical and physical pages  "

If mDevMode.dmBase4.dmBase.dmScale <> newValue Then
    mDevMode.dmBase4.dmBase.dmScale = newValue
    Changed = True
End If

End Property

' ##BLOCK_DESCRIPTION The scaling factor between the logical and physical pages
' ##BLOCK_REMARKS This is expressed as a percent - i.e. if ScaleFactor is 100 the scale is 1 to 1
Public Property Get ScaleFactor() As Integer

    ScaleFactor = mDevMode.dmBase4.dmBase.dmScale

End Property


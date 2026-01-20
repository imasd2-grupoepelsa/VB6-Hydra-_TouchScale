VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiColourAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2418
Attribute VB_Description="This class is used to manipulate the "Colour Adjustment" settings of a device context:EventVB~ApiDeviceContext."
Option Explicit

' ##MODULE_DESCRIPTION This class is used to manipulate the "Colour Adjustment" _
settings of a %device context:EventVB~ApiDeviceContext%.

' ##MODULE_DESCRIPTION These settings affect the way that bitmaps colours are _
changed when the bitmaps are put on that %device context:EventVB~ApiDeviceContext%.

' ##MODULE_DESCRIPTION You can use this to alter the colour balance settings _
as per in a graphics package.

Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Private Type COLORADJUSTMENT
    caSize As Integer
    caFlags As Integer
    caIlluminantIndex As Integer
    caRedGamma As Integer
    caGreenGamma As Integer
    caBlueGamma As Integer
    caReferenceBlack As Integer
    caReferenceWhite As Integer
    caContrast As Integer
    caBrightness As Integer
    caColorfulness As Integer
    caRedGreenTint As Integer
End Type
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long

Private mParentDC As Long

Private mColourAdjustment As COLORADJUSTMENT

Public Enum enColourAdjustFlags
     CA_NEGATIVE = &H1
     CA_LOG_FILTER = &H2
End Enum

Public Enum enIlluninantIndexes
    ILLUMINANT_DEVICE_DEFAULT = 0
    ILLUMINANT_TUNGSTEN = 1
    ILLUMINANT_NOON_SUNLIGHT = 2
    ILLUMINANT_NTSC_DAYLIGHT = 3
    ILLUMINANT_NORMAL_PRINT = 4
    ILLUMINANT_BOND_PRINT = 5
    ILLUMINANT_STANDARD_DAYLIGHT = 6
    ILLUMINANT_NORTHEN_DAYLIGHT = 7
    ILLUMINANT_FLOURESCENT_LIGHT = 8
End Enum

Private Const RGB_GAMMA_MIN    As Long = 2500
Private Const RGB_GAMMA_MAX    As Long = 65000

'/* Min and max for ReferenceBlack and ReferenceWhite */
Private Const REFERENCE_WHITE_MIN  As Long = 6000
Private Const REFERENCE_WHITE_MAX  As Long = 10000
Private Const REFERENCE_BLACK_MIN  As Long = 0
Private Const REFERENCE_BLACK_MAX   As Long = 4000
Private Const COLOR_ADJ_MIN    As Long = -100
Private Const COLOR_ADJ_MAX    As Long = 100

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get BlueGamma() As Integer
Attribute BlueGamma.VB_HelpID = 2419

    BlueGamma = mColourAdjustment.caBlueGamma
    
End Property

Public Property Let BlueGamma(ByVal newGamma As Integer)

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caBlueGamma <> newGamma Then
       mColourAdjustment.caBlueGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Let Brightness(ByVal NewValue As Integer)
Attribute Brightness.VB_HelpID = 2421

    If NewValue < COLOR_ADJ_MIN Then
        NewValue = COLOR_ADJ_MIN
    ElseIf NewValue > COLOR_ADJ_MAX Then
        NewValue = COLOR_ADJ_MAX
    End If
    
    If NewValue <> mColourAdjustment.caBrightness Then
        mColourAdjustment.caBrightness = NewValue
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get Brightness() As Integer

    Brightness = mColourAdjustment.caBrightness
    
End Property

Public Property Let Colourfulness(ByVal NewValue As Integer)
Attribute Colourfulness.VB_HelpID = 2423

    If NewValue < COLOR_ADJ_MIN Then
        NewValue = COLOR_ADJ_MIN
    ElseIf NewValue > COLOR_ADJ_MAX Then
        NewValue = COLOR_ADJ_MAX
    End If
    
    If NewValue <> mColourAdjustment.caColorfulness Then
        mColourAdjustment.caColorfulness = NewValue
        Call RefreshColourAdjustment
    End If

End Property

Public Property Get Colourfulness() As Integer

Colourfulness = mColourAdjustment.caColorfulness

End Property

Public Property Let Contrast(ByVal newContrast As Integer)
Attribute Contrast.VB_HelpID = 2425

    If newContrast < COLOR_ADJ_MIN Then
        newContrast = COLOR_ADJ_MIN
    ElseIf newContrast > COLOR_ADJ_MAX Then
        newContrast = COLOR_ADJ_MAX
    End If
    
    If newContrast <> mColourAdjustment.caContrast Then
        mColourAdjustment.caContrast = newContrast
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get Contrast() As Integer

    Contrast = mColourAdjustment.caContrast
    
End Property

Public Property Let GreenGamma(ByVal newGamma As Integer)
Attribute GreenGamma.VB_HelpID = 2427

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caGreenGamma <> newGamma Then
       mColourAdjustment.caGreenGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get GreenGamma() As Integer

    GreenGamma = mColourAdjustment.caGreenGamma
    
End Property


Public Property Let IlluminantIndex(ByVal newIndex As enIlluninantIndexes)
Attribute IlluminantIndex.VB_HelpID = 2429

    If mColourAdjustment.caIlluminantIndex <> newIndex Then
        mColourAdjustment.caIlluminantIndex = newIndex
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get IlluminantIndex() As enIlluninantIndexes

IlluminantIndex = mColourAdjustment.caIlluminantIndex

End Property

Public Property Set ParentDC(ByVal newDC As ApiDeviceContext)
Attribute ParentDC.VB_HelpID = 2430

Dim lRet As Long

If newDC.hdc <> mParentDC Then
    mParentDC = newDC.hdc
    lRet = GetColorAdjustment(mParentDC, mColourAdjustment)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiColourAdjustment:ParentDC", GetLastSystemError)
    End If
End If

End Property
Public Property Get RedGamma() As Integer
Attribute RedGamma.VB_HelpID = 2431

    RedGamma = mColourAdjustment.caRedGamma
    
End Property

Public Property Let RedGamma(ByVal newGamma As Integer)

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caRedGamma <> newGamma Then
       mColourAdjustment.caRedGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Let RedGreenTint(ByVal newTint As Integer)
Attribute RedGreenTint.VB_HelpID = 2433

    If newTint < COLOR_ADJ_MIN Then
        newTint = COLOR_ADJ_MIN
    ElseIf newTint > COLOR_ADJ_MAX Then
        newTint = COLOR_ADJ_MAX
    End If
    
    If newTint <> mColourAdjustment.caRedGreenTint Then
        mColourAdjustment.caRedGreenTint = newTint
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get RedGreenTint() As Integer

    RedGreenTint = mColourAdjustment.caRedGreenTint
    
End Property

Private Function RefreshColourAdjustment()

Dim lRet As Long

If mParentDC <> 0 Then
    lRet = SetColorAdjustment(mParentDC, mColourAdjustment)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiColourAdjustment:RefreshColourAdjustment", GetLastSystemError)
    End If
End If

End Function


Private Sub Class_Initialize()

mColourAdjustment.caSize = LenB(mColourAdjustment)

End Sub



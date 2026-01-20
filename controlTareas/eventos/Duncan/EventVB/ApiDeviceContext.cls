VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiDeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2340
Attribute VB_Description = "This class provides methods and properties pertaining to a device context."
Option Explicit

' ##MODULE_DESCRIPTION This class provides methods and properties pertaining to a device context.

' ##MODULE_DESCRIPTION A device context is akin to the canvas upon which all graphical operations _
take place.  It provides information as to what types of operations a graphical device is capable of and _
a set of standard methods to perform operations on that canvas.

Private Const CLSNAME = "ApiDeviceContext"

'\\ API Declarations ---------------------------------------------------
Private Declare Function GetDeviceCapsApi Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long

' ##ENUMERATION_DESCRIPTION Get the capabilities of the device that created this device context
Public Enum enDeviceCapsIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION cDRIVERVERSION  Device driver version
    cDRIVERVERSION = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION cTECHNOLOGY Device classification
    cTECHNOLOGY = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION cHORZSIZE Horizontal size in millimeters
    cHORZSIZE = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION cVERTSIZE Vertical size in millimeters
    cVERTSIZE = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION cHORZRES Horizontal width in pixels
    cHORZRES = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION cVERTRES Vertical width in pixels
    cVERTRES = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION cLOGPIXELSX Logical pixels/inch in X
    cLOGPIXELSX = 88
    ' ##ENUMERATION_MEMBER_DESCRIPTION cLOGPIXELSY Logical pixels/inch in Y
    cLOGPIXELSY = 90
    ' ##ENUMERATION_MEMBER_DESCRIPTION cBITSPIXEL Number of bits per pixel
    cBITSPIXEL = 12
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPLANES Number of planes
    cPLANES = 14
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMBRUSHES Number of brushes the device has
    cNUMBRUSHES = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMCOLORS Number of colors the device supports
    cNUMCOLORS = 24
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMFONTS Number of fonts the device has
    cNUMFONTS = 22
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMMARKERS Number of markers the device has
    cNUMMARKERS = 20
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMPENS Number of pens the device has
    cNUMPENS = 18
    ' ##ENUMERATION_MEMBER_DESCRIPTION cASPECTX Length of the X leg
    cASPECTX = 40
    ' ##ENUMERATION_MEMBER_DESCRIPTION cASPECTXY Length of the hypotenuse
    cASPECTXY = 44
    ' ##ENUMERATION_MEMBER_DESCRIPTION cASPECTY Length of the Y leg
    cASPECTY = 42
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPDEVICESIZE Size required for device descriptor
    cPDEVICESIZE = 26
    ' ##ENUMERATION_MEMBER_DESCRIPTION cCLIPCAPS Clipping capabilities
    cCLIPCAPS = 36
    ' ##ENUMERATION_MEMBER_DESCRIPTION cSIZEPALETTE Number of entries in physical palette
    cSIZEPALETTE = 104
    ' ##ENUMERATION_MEMBER_DESCRIPTION cNUMRESERVED Number of reserved entries in palette
    cNUMRESERVED = 106
    ' ##ENUMERATION_MEMBER_DESCRIPTION cCOLORRES Actual color resolution
    cCOLORRES = 108
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPHYSICALOFFSETX Physical Printable Area x margin
    cPHYSICALOFFSETX = 112
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPHYSICALOFFSETY Physical Printable Area y margin
    cPHYSICALOFFSETY = 113
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPHYSICALHEIGHT Physical Height in device units
    cPHYSICALHEIGHT = 111
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPHYSICALWIDTH Physical Width in device units
    cPHYSICALWIDTH = 110
    ' ##ENUMERATION_MEMBER_DESCRIPTION cSCALINGFACTORX Scaling factor x
    cSCALINGFACTORX = 114
    ' ##ENUMERATION_MEMBER_DESCRIPTION cSCALINGFACTORY Scaling factor y
    cSCALINGFACTORY = 115
    ' ##ENUMERATION_MEMBER_DESCRIPTION cRASTERCAPS Bitblt capabilities
    cRASTERCAPS = 38
    ' ##ENUMERATION_MEMBER_DESCRIPTION cCURVECAPS Curve capabilities
    cCURVECAPS = 28
    ' ##ENUMERATION_MEMBER_DESCRIPTION cLINECAPS Line capabilities
    cLINECAPS = 30
    ' ##ENUMERATION_MEMBER_DESCRIPTION cPOLYGONALCAPS Polygonal capabilities
    cPOLYGONALCAPS = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION cTEXTCAPS Text capabilities
    cTEXTCAPS = 34
End Enum

' ##ENUMERATION_DESCRIPTION These are the classifications of device which have _
%device context:ApiDeviceContext% support
Public Enum DeviceContextTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_PLOTTER Vector plotter (Forexample an x,y pen plotter as used in technical drawing)
    DT_PLOTTER = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_RASCAMERA Raster camera devices
    DT_RASCAMERA = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_RASDISPLAY Raster display (VDU or plasma screen for example)
    DT_RASDISPLAY = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_RASPRINTER Raster printer (Dot matrix, bubblejet or laser printer)
    DT_RASPRINTER = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_CHARSTREAM Character-stream, PLP
    DT_CHARSTREAM = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_METAFILE Metafile, VDM
    DT_METAFILE = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION DT_DISPFILE Display-file
    DT_DISPFILE = 6
End Enum

Public Enum enDeviceMesaurementScale
    DMS_Millimeters = 1
    DMS_Pixels = 2
End Enum

' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% output clipping capabilities
Public Enum ClippingCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION CP_NONE No clipping of output
    CP_NONE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION CP_RECTANGLE Output can be clipped to rectangles
    CP_RECTANGLE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CP_REGION Output can be clipped to regions
    CP_REGION = 2              '
End Enum

' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% raster capabilities
Public Enum RasterCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_BANDING Device requires banding support
    RC_BANDING = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_BIGFONT supports >64K fonts
    RC_BIGFONT = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_BITBLT Can do standard BLT.
    RC_BITBLT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_BITMAP64 Device can support >64K bitmap
    RC_BITMAP64 = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_DEVBITS
    RC_DEVBITS = &H8000
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_DI_BITMAP supports DIB to memory
    RC_DI_BITMAP = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_DIBTODEV supports DIBitsToDevice
    RC_DIBTODEV = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_FLOODFILL supports FloodFill
    RC_FLOODFILL = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_NONE Device does not support raster operations
    RC_NONE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_PALETTE supports a palette
    RC_PALETTE = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_SAVEBITMAP
    RC_SAVEBITMAP = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_SCALING Device requires scaling support
    RC_SCALING = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_STRETCHBLT supports StretchBlt
    RC_STRETCHBLT = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION RC_STRETCHDIB supports StretchDIBits
    RC_STRETCHDIB = &H2000
End Enum


' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% curve drawing capabilities
Public Enum CurveCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_CIRCLES Can do circles
    CC_CIRCLES = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_CHORD Can do chord arcs
    CC_CHORD = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_ELLIPSES Can do ellipses
    CC_ELLIPSES = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_PIE Can do pie wedges
    CC_PIE = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_NONE Curves not supported
    CC_NONE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_WIDE Can do wide lines
    CC_WIDE = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_STYLED Can do styled lines
    CC_STYLED = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_WIDESTYLED Can do wide styled lines
    CC_WIDESTYLED = 64
    ' ##ENUMERATION_MEMBER_DESCRIPTION CC_INTERIORS Can do interiors
    CC_INTERIORS = 128
End Enum

' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% line drawing capabilities
Public Enum LineCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_NONE Line drawing not supported
    LC_NONE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_POLYLINE Can do polylines
    LC_POLYLINE = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_MARKER Can do markers
    LC_MARKER = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_POLYMARKER Can do polymarkers
    LC_POLYMARKER = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_WIDE Can do wide lines
    LC_WIDE = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_STYLED Can do styled lines
    LC_STYLED = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_WIDESTYLED Can do wide styled lines
    LC_WIDESTYLED = 64
    ' ##ENUMERATION_MEMBER_DESCRIPTION LC_INTERIORS Can do interiors
    LC_INTERIORS = 128
End Enum


' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% polygon drawing capabilities
Public Enum PolygonCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_NONE Polygonals not supported
    PC_NONE = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_POLYGON Can do polygons
    PC_POLYGON = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_RECTANGLE Can do rectangles
    PC_RECTANGLE = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_TRAPEZOID Can do trapezoids
    PC_TRAPEZOID = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_SCANLINE Can do scanlines
    PC_SCANLINE = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_WIDE Can do wide borders
    PC_WIDE = 16
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_STYLED Can do styled borders
    PC_STYLED = 32
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_WIDESTYLED Can do wide styled borders
    PC_WIDESTYLED = 64
    ' ##ENUMERATION_MEMBER_DESCRIPTION PC_INTERIORS Can do interiors
    PC_INTERIORS = 128
End Enum

' ##ENUMERATION_DESCRIPTION The %device context's:ApiDeviceContext% text drawing capabilities
Public Enum TextCapabilities
    ' ##ENUMERATION_MEMBER_DESCRIPTION Text output to character level precision
    TC_OP_CHARACTER = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION Text output to stroke level precision
    TC_OP_STROKE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_CP_STROKE Can clip to stroke level precision
    TC_CP_STROKE = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_CR_90 Can rotate text 90 degrees
    TC_CR_90 = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_CR_ANY Can rotate text to any angle
    TC_CR_ANY = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_SF_X_YINDEP Can do X/Y independent font scaling
    TC_SF_X_YINDEP = &H20              '  Can do ScaleFreedom      X_YINDEPENDENT
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_SA_DOUBLE Can do double precision font scaling
    TC_SA_DOUBLE = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_SA_INTEGER Can do integer precision font scaling
    TC_SA_INTEGER = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_SA_CONTIN Can do continuously scalable fonts
    TC_SA_CONTIN = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_IA_ABLE Can do italic text
    TC_IA_ABLE = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_UA_ABLE Can do underlined text
    TC_UA_ABLE = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_RA_ABLE Can do raster fonts
    TC_RA_ABLE = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_SCROLLBLT Can do scrolling text
    TC_SCROLLBLT = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION TC_VA_ABLE Can do vector fonts
    TC_VA_ABLE = &H4000
End Enum

' ##ENUMERATION_DESCRIPTION This describes the mapping mode between logical coordinates and physical coordinates on the device context. _
It is broadly analagous to the VB ScaleMode property
Public Enum MappingModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_ANISOTROPIC Arbitrary x and y scales, like ScaleMode = vbScalemodeUser
    MM_ANISOTROPIC = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_HIENGLISH 1 unit is 0.001 inch
    MM_HIENGLISH = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_HIMETRIC 1 unit is 0.01 millimeters
    MM_HIMETRIC = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_ISOTROPIC x and y are arbitrary but same as each other
    MM_ISOTROPIC = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_LOENGLISH 1 unit is 0.01 inch
    MM_LOENGLISH = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_LOMETRIC 1 unit is 0.1mm
    MM_LOMETRIC = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_TEXT 1 unit is one device pixel
    MM_TEXT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION MM_TWIPS 1 unit is one twip (approx. 1/1440 of an inch)
    MM_TWIPS = 6
End Enum


'\\ DrawEdge Related
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

' ##ENUMERATION_DESCRIPTION Specifies the style of border to draw in the draw edge call
Public Enum EdgeBorderStyles
    ' ##ENUMERATION_MEMBER_DESCRIPTION BDR_RAISEDINNER The inner edge of teh border is raised
    BDR_RAISEDINNER = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BDR_RAISEDOUTER The outer edge of the border is raised
    BDR_RAISEDOUTER = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BDR_SUNKENINNER The inner edge of the border is sunken
    BDR_SUNKENINNER = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION BDR_SUNKENOUTER The outer edge of the border is sunken
    BDR_SUNKENOUTER = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION EDGE_BUMP The outer is raised and the inner sunken
    EDGE_BUMP = &H9
    ' ##ENUMERATION_MEMBER_DESCRIPTION EDGE_ETCHED The outer is sunken and the inner raised
    EDGE_ETCHED = &H6
    ' ##ENUMERATION_MEMBER_DESCRIPTION EDGE_RAISED Both inner and outer border are raised
    EDGE_RAISED = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION EDGE_SUNKEN Both inner and outer border are sunken
    EDGE_SUNKEN = &H10
End Enum

' ##ENUMERATION_DESCRIPTION Specifies the type of border to draw in the draw edge function
Public Enum BorderFlags
   ' ##ENUMERATION_MEMBER_DESCRIPTION BF_LEFT Draw the left hand border
    BF_LEFT = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_TOP Draw the top border
    BF_TOP = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_RIGHT Draw the right hand border
    BF_RIGHT = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_BOTTOM Draw the bottom border
    BF_BOTTOM = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_TOPLEFT Draw the top and left borders
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_TOPRIGHT Draw the top and right borders
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_BOTTOMLEFT Draw the bottom and left borders
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_BOTTOMRIGHT Draw the bottom and right borders
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_RECT Draw all four rectangle borders
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_DIAGONAL Draw a diagonal border
    BF_DIAGONAL = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_DIAGONAL_ENDTOPRIGHT Draw a diagonal border from bottom left to top right
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_DIAGONAL_ENDTOPLEFT Draw a diagonal border from bottom right to top left
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_DIAGONAL_ENDBOTTOMLEFT Draw a diagonal border from top right to bottom left
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_DIAGONAL_ENDBOTTOMRIGHT Draw a diagonal border from top left to bottom right
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_MIDDLE Fill the middle of the border
    BF_MIDDLE = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_SOFT Make a soft edge border
    BF_SOFT = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_ADJUST Calculate the space left over
    BF_ADJUST = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_FLAT For flat rather than 3D borders
    BF_FLAT = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION BF_MONO For monochrome borders
    BF_MONO = &H8000
End Enum

'\\ Drawing related...
Private Declare Function DrawIconApi Lib "user32" Alias "DrawIcon" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawEdgeApi Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long

' ##ENUMERATION_DESCRIPTION The type of control drawn by DrawFrameControl
Private Enum DrawFrameControlTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFC_CAPTION Draw a form caption bar
    DFC_CAPTION = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFC_MENU Draw a top level menu
    DFC_MENU = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFC_SCROLL Draw a scroll bar
    DFC_SCROLL = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFC_BUTTON Draw a button
    DFC_BUTTON = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFC_POPUPMENU Draw a popup menu control
    DFC_POPUPMENU = 5
End Enum

' ##ENUMERATION_DESCRIPTION The part of the caption type control to draw
Public Enum DrawFrameCaptionComponents
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CAPTIONCLOSE Draw the caption close button
    DFCS_CAPTIONCLOSE = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CAPTIONMIN Draw the caption minimise button
    DFCS_CAPTIONMIN = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CAPTIONMAX Draw the caption maximise button
    DFCS_CAPTIONMAX = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CAPTIONRESTORE Draw the caption restore button
    DFCS_CAPTIONRESTORE = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CAPTIONHELP Draw a caption help button
    DFCS_CAPTIONHELP = &H4
End Enum

' ##ENUMERATION_DESCRIPTION The style of menu control to draw
Public Enum DrawFrameMenuComponents
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_MENUARROW Draw a submenu arrow
    DFCS_MENUARROW = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_MENUCHECK Draw a menu checkbox
    DFCS_MENUCHECK = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_MENUBULLET Draw a menu bullet mark
    DFCS_MENUBULLET = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_MENUARROWRIGHT Draw a right hand submenu arrow
    DFCS_MENUARROWRIGHT = &H4
End Enum

' ##ENUMERATION_DESCRIPTION The part of the scrollbar type control to draw
Public Enum DrawFrameScrollComponents
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLUP Draw the up arrow of the scrollbar
    DFCS_SCROLLUP = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLDOWN Draw the down arrow of the scrollbar
    DFCS_SCROLLDOWN = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLLEFT Draw the left arrow of the scrollbar
    DFCS_SCROLLLEFT = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLRIGHT Draw the right arrow of the scroll bar
    DFCS_SCROLLRIGHT = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLCOMBOBOX Draw a combo box scrollbar
    DFCS_SCROLLCOMBOBOX = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLSIZEGRIP Draw the left hand size grip
    DFCS_SCROLLSIZEGRIP = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_SCROLLSIZEGRIPRIGHT Draw the right hand size grip
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
End Enum

' ##ENUMERATION_DESCRIPTION The style of button control to draw
Public Enum DrawFrameButtonComponents
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTONCHECK Button is a checkbox type button
    DFCS_BUTTONCHECK = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTONRADIOIMAGE Button is the image for a radio button
    DFCS_BUTTONRADIOIMAGE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTONRADIOMASK Button is the mask for a radio button
    DFCS_BUTTONRADIOMASK = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTONRADIO Button is a radio button
    DFCS_BUTTONRADIO = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTON3STATE Button is a tri-state button
    DFCS_BUTTON3STATE = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_BUTTONPUSH Button is a standard push button
    DFCS_BUTTONPUSH = &H10
End Enum
    
' ##ENUMERATION_DESCRIPTION The state of the button to draw
Public Enum DrawFrameButtonComponentsState
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_NORMAL Draw the button in normal (up) state
    DFCS_NORMAL = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_INACTIVE Draw the button disabled
    DFCS_INACTIVE = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_PUSHED Draw the button pushed down
    DFCS_PUSHED = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_CHECKED Draw the checkbox style button checked
    DFCS_CHECKED = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_HOT Draw a button in hot-tracking state
    DFCS_HOT = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_MONO Draw a mono border type button
    DFCS_MONO = &H8000
End Enum

' ##ENUMERATION_DESCRIPTION The style of popup menu control to draw
Public Enum DrawFrameMenuComponentsPopupMenu
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_TRANSPARENT
    DFCS_TRANSPARENT = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_ADJUSTRECT
    DFCS_ADJUSTRECT = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION DFCS_FLAT
    DFCS_FLAT = &H4000
End Enum

Private Declare Function DrawFrameControlApi Lib "user32" Alias "DrawFrameControl" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

'\\ The graphics mode...
Private Declare Function SetGraphicsModeApi Lib "gdi32" Alias "SetGraphicsMode" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function GetGraphicsModeApi Lib "gdi32" Alias "GetGraphicsMode" (ByVal hdc As Long) As Long
' ##ENUMERATION_DESCRIPTION The graphics mode of the %device context:ApiDeviceContext%
Public Enum GraphicsMode
    ' ##ENUMERATION_MEMBER_DESCRIPTION GM_COMPATIBLE The graphics mode is compatible with 16 bit windows and does not support world transforms
    GM_COMPATIBLE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION GM_ADVANCED The graphics mode supports world transforms
    GM_ADVANCED = 2
End Enum

' ##ENUMERATION_DESCRIPTION Mode of operation of the stretchblt command to this device context
Public Enum StretchBltModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_BLACKONWHITE - Performs an AND operation between source and target of the stretchblt operation
    SBM_BLACKONWHITE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_WHITEONBLACK - Performs an OR operation between source and target of the stretchblt operation
    SBM_WHITEONBLACK = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_COLOURONCOLOUR - The target is overwritten with the source
    SBM_COLOURONCOLOUR = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION SBM_HALFTONE - The target and source are combined using halftone averaging
    SBM_HALFTONE = 4
End Enum

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long

'\\ Brush origin
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

' ##ENUMERATION_DESCRIPTION The horizontal alignment of text operations on this device context
Public Enum HorizontalTextAlignments
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_LEFT - Text is left aligned horizontally
    TA_LEFT = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_RIGHT - Text is right aligned horizontally
    TA_RIGHT = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_CENTRE - Text is centred horizontally
    TA_CENTRE = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_RTLREADING - The text is written right to left
    TA_RTLREADING = 256
End Enum

' ##ENUMERATION_DESCRIPTION The vertical alignment of text operations on this device context
Public Enum VerticalTextAlignments
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_TOP - Text is vertically aligned to the top
    TA_TOP = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_BOTTOM - Text is aligned to the bottom
    TA_BOTTOM = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_BASELINE - Text is aligned on the baseline of the text
    TA_BASELINE = 24
End Enum

' ##ENUMERATION_DESCRIPTION Sets whether text operations update the current position of the cursor
Private Enum UpdateTextPos
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_NOUPDATECP - Text operations do not update the current cursor position
    TA_NOUPDATECP = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION TA_UPDATECP - Text operations do update the current cursor position
    TA_UPDATECP = 1
End Enum

Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long

'\\ Text colour
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'\\ Background colour
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'\\ Blt functions...
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' ##ENUMERATION_DESCRIPTION These describe how two raster images are combined in a BitBlt function
Public Enum RasterOperationCodes
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_SRCCOPY Destination is overwritten by source
    ROP_SRCCOPY = &HCC0020
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_SRCPAINT Destination is (source OR destination)
    ROP_SRCPAINT = &HEE0086
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_SRCAND Destination = (source AND destination)
    ROP_SRCAND = &H8800C6
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_SRCINVERT Destination = (source XOR destination)
    ROP_SRCINVERT = &H660046
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_SRCERASE Destination = (source AND (NOT destination))
    ROP_SRCERASE = &H440328
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_NOTSRCCOPY Destination = (NOT source)
    ROP_NOTSRCCOPY = &H330008
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_NOTSRCERASE Destination = (NOT src) AND (NOT dest)
    ROP_NOTSRCERASE = &H1100A6
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_MERGECOPY Destination = (source AND pattern)
    ROP_MERGECOPY = &HC000CA
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_MERGEPAINT Destination = (NOT source) OR dest
    ROP_MERGEPAINT = &HBB0226
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_PATCOPY Destination = pattern
    ROP_PATCOPY = &HF00021
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_PATPAINT Destination = DPSnoo
    ROP_PATPAINT = &HFB0A09
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_PATINVERT Destination = pattern XOR dest
    ROP_PATINVERT = &H5A0049
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_DSTINVERT Destination = (NOT dest)
    ROP_DSTINVERT = &H550009
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_BLACKNESS Destination made black
    ROP_BLACKNESS = &H42
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_WHITENESS Destination made white
    ROP_WHITENESS = &HFF0062
    ' ##ENUMERATION_MEMBER_DESCRIPTION ROP_NOMIRRORBITMAP Do not Mirror the bitmap in this call
    ROP_NOMIRRORBITMAP = &H80000000
End Enum

' ##ENUMERATION_DESCRIPTION Sets the options for extended text operations on this device context
Public Enum ExtendedTextOutOptions
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_GRAYED - Text is displayed as if it were disabled
    ETO_GRAYED = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_OPAQUE - The background rectangle of the text is filled with the device context background colour
    ETO_OPAQUE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_CLIPPED - The text is clipped to the given rectangle
    ETO_CLIPPED = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_GLYPH_INDEX
    ETO_GLYPH_INDEX = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_RTLREADING - Text is written right to left
    ETO_RTLREADING = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_NUMERICSLOCAL - Numbers are written with the local format (if this is different from arabic numerals)
    ETO_NUMERICSLOCAL = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_NUMERICSLATIN - Numbers are always written in arabic numerals (0..9)
    ETO_NUMERICSLATIN = &H800
    ' ##ENUMERATION_MEMBER_DESCRIPTION ETO_IGNORELANGUAGE - Local language is ignored and text is output in ASCII (not advisable)
    ETO_IGNORELANGUAGE = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION
    ETO_PDY = &H2000
End Enum

'\\ Simple text drawing
Private Declare Function TextOutApi Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'\\ More advanced text, with bounding box etc.
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'\\ Font manipulation....
Private mPrevFont As Long

'\\ GDI Objects....
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private mInitialFont As ApiLogFont
Private mInitialPen As ApiLogPen
Private mInitialBrush As ApiLogBrush

'\\ Region drawing functions....
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'\\ Compatible device context creation...
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

'\\ Background mode
Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

' ##ENUMERATION_DESCRIPTION Describes how drawing and text operations affect the background
Public Enum BackgroundModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION TRANSPARENT - The area behind the drawing (or text) operation is not affected
    TRANSPARENT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION OPAQUE - The background is filled with the background brush before the drawing (or text) operation takes place
    OPAQUE = 2
End Enum

'\\ Getting and setting the CurrentX and CurrentY properties...
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long

'\\ 'Standard' GDI drawing operations...
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function EllipseApi Lib "gdi32" Alias "Ellipse" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RectangleApi Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'\\ Arc drawing modifier
Private Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long

' ##ENUMERATION_DESCRIPTION Describes the direction of arc drawing
Public Enum ArcDrawingDirections
    ' ##ENUMERATION_MEMBER_DESCRIPTION AD_CLOCKWISE Draw arcs clockwise
    AD_CLOCKWISE = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION AD_COUNTERCLOCKWISE Draw arcs anti-clockwise
    AD_COUNTERCLOCKWISE = 1
End Enum

'\\ Painting operations
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

' ##ENUMERATION_DESCRIPTION Modifies how the flood fill routine operates
Public Enum FloodFillModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION FLOODFILLBORDER Fill the area bounded by the given test colour
    FLOODFILLBORDER = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION FLOODFILLSURFACE Fill all conjoined area of the given test colour
    FLOODFILLSURFACE = 1
End Enum

Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern

'\\ Member variables...
Private mHDC As Long
Private bChanged As Boolean

Private Type Size
    cx As Long
    cy As Long
End Type
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Size) As Long
Private Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Private Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As Size) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Size) As Long

Private Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Size) As Long
Private Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Size) As Long

' ##BLOCK_DESCRIPTION Draws an arc defined by the input parameters in the currently selected %pen:EventVB~ApiLogPen%
' ##BLOCK_REMARKS BoundingRectangle - The rectangle that completely contains the arc
' ##BLOCK_REMARKS StartRadial - The starting point of the arc
' ##BLOCK_REMARKS EndRadial - The end point of the arc
Public Sub Arc(ByVal BoundingRectangle As APIRect, ByVal StartRadial As APIPoint, ByVal EndRadial As APIPoint)
Attribute Arc.VB_Description = "Draws an arc defined by the input parameters in the currently selected pen:EventVB~ApiLogPen"
Attribute Arc.VB_HelpID = 2857

Dim lRet As Long

If Not BoundingRectangle Is Nothing Then
    With BoundingRectangle
        lRet = ArcTo(mHDC, .Left, .Top, .Right, .Bottom, StartRadial.x, StartRadial.y, EndRadial.x, EndRadial.y)
    End With
End If
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Arc", GetLastSystemError
End If

End Sub


Public Property Let ArcDrawingDirection(ByVal newDirection As ArcDrawingDirections)
Attribute ArcDrawingDirection.VB_Description = "ArcDrawingDirection - Which direction an Arc is drawn in on this device context  "
Attribute ArcDrawingDirection.VB_HelpID = 2859

Dim lRet As Long
lRet = SetArcDirection(mHDC, newDirection)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":ArcDrawingDirection", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION ArcDrawingDirection - Which direction an Arc is drawn in on this device context
Public Property Get ArcDrawingDirection() As ArcDrawingDirections

ArcDrawingDirection = GetArcDirection(mHDC)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":ArcDrawingDirection", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION BrushFill - Fills the given rectangle with the given brush
' ##BLOCK_REMARKS FillRectangle - The rectangle (relative to the window) to fill
' ##BLOCK_REMARKS Brush - The brush to use to do the filling
Public Sub BrushFill(ByVal FillRectangle As APIRect, ByVal Brush As ApiLogBrush)
Attribute BrushFill.VB_Description = "BrushFill - Fills the given rectangle with the given brush"
Attribute BrushFill.VB_HelpID = 3057

Dim lRet As Long
Dim prevBrush As ApiLogBrush

Set prevBrush = SelectedBrush
Set SelectedBrush = Brush
With FillRectangle
    lRet = PatBlt(mHDC, .Left, .Top, (.Right - .Left), (.Bottom - .Top), PATCOPY)
End With
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":BrushFill", GetLastSystemError
End If

End Sub

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Set BackgroundColour(ByVal newCol As ApiColour)
Attribute BackgroundColour.VB_Description = "The current background colour of this device context  "
Attribute BackgroundColour.VB_HelpID = 2342

Call SetBkColor(mHDC, newCol.ColourRef)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":BackgroundColour", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION The current background colour of this device context
' ##BLOCK_REMARKS This colour is the backround on text and drawing operations
Public Property Get BackgroundColour() As ApiColour

Dim colThis As ApiColour
Dim lColRef As Long

Set colThis = New ApiColour
lColRef = GetBkColor(mHDC)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":BackgroundColour", GetLastSystemError
End If

Set BackgroundColour = colThis

End Property

Public Property Let BackStyle(ByVal NewStyle As BackgroundModes)
Attribute BackStyle.VB_Description = "The current background style of this device context  "
Attribute BackStyle.VB_HelpID = 2344

Dim lRet As Long
lRet = SetBkMode(mHDC, NewStyle)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":BackStyle (Let)", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION The current background style of this device context
' ##BLOCK_REMARKS The background can be clear or opaque
Public Property Get BackStyle() As BackgroundModes

BackStyle = GetBkMode(mHDC)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":BackStyle (Get)", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION The bits per pixel value describes how many bits of memory are required per pixel
' ##BLOCK_REMARKS The higher the bits per pixel value the more colours the device context supports
Public Property Get BitsPerPixel() As Long
Attribute BitsPerPixel.VB_Description = "The bits per pixel value describes how many bits of memory are required per pixel"
Attribute BitsPerPixel.VB_HelpID = 2345

    BitsPerPixel = GetDeviceCaps(cBITSPIXEL)
    
End Property


'\\ --[BlockTransferBitData]-------------------------------------------------------------------------
' ##BLOCK_DESCRIPTION Transfers the data from a rectangle section of a device context to a rectangle section of this device context
'\\ -------------------------------------------------------------------------------------------------
Public Sub BlockTransferBitData(ByVal Source As ApiDeviceContext, ByVal SourceSection As APIRect, ByVal TargetSection As APIRect, ByVal Operation As RasterOperationCodes)
Attribute BlockTransferBitData.VB_Description = "Transfers the data from a rectangle section of a device context to a rectangle section of this device context"
Attribute BlockTransferBitData.VB_HelpID = 2346

Dim lRet As Long

lRet = StretchBlt(mHDC, TargetSection.Left, TargetSection.Top, _
                  TargetSection.Right - TargetSection.Left, _
                  TargetSection.Bottom - TargetSection.Top, _
                  Source.hdc, SourceSection.Left, _
                  SourceSection.Top, _
                  SourceSection.Right - SourceSection.Left, _
                  SourceSection.Bottom - SourceSection.Top, Operation)
                  
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":BlockTransferBitData", GetLastSystemError
End If

End Sub

' ##BLOCK_DESCRIPTION The collection of %brushes:EventVB~ApiLogBrush% in this device context
' ##BLOCK_REMARKS Additional brushes can be used by setting the selected brush property
Public Property Get Brushes() As Collection
Attribute Brushes.VB_Description = "The collection of brushes:EventVB~ApiLogBrush in this device context"
Attribute Brushes.VB_HelpID = 2347

Dim colBrushes As Collection

Set colBrushes = GetDCBrushCollection(mHDC)

If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Brushes", GetLastSystemError
End If

Set Brushes = colBrushes

End Property

Public Property Set BrushOrigin(ByVal neworigin As APIPoint)
Attribute BrushOrigin.VB_HelpID = 2349

Dim ptTemp As POINTAPI

Call SetBrushOrgEx(mHDC, neworigin.x, neworigin.y, ptTemp)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":BrushOrigin", GetLastSystemError
End If

End Property

Public Property Get BrushOrigin() As APIPoint

Dim ptThis As APIPoint
Dim ptTemp As POINTAPI

Set ptThis = New APIPoint

Call GetBrushOrgEx(mHDC, ptTemp)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":BrushOrigin (Get)", GetLastSystemError
Else
    ptThis.CreateFromPointer VarPtr(ptTemp)
End If

Set BrushOrigin = ptThis

End Property

Public Property Get ClipingCapabilities() As ClippingCapabilities
Attribute ClipingCapabilities.VB_HelpID = 2350

    ClipingCapabilities = GetDeviceCaps(cCLIPCAPS)

End Property


Public Property Get ColourAdjustment() As ApiColourAdjustment
Attribute ColourAdjustment.VB_HelpID = 2351

Dim mCol As ApiColourAdjustment

Set mCol = New ApiColourAdjustment
Set mCol.ParentDC = Me

Set ColourAdjustment = mCol

End Property

Public Property Get ColourPlanes() As Long
Attribute ColourPlanes.VB_HelpID = 2352

    ColourPlanes = GetDeviceCaps(cPLANES)
    
End Property

Public Property Get ColourResolution() As Long
Attribute ColourResolution.VB_HelpID = 2353

    ColourResolution = GetDeviceCaps(cCOLORRES)

End Property

' ##BLOCK_DESCRIPTION Creates a bitmap with the same properties as this device context
Public Function CreateCompatibleBitmap(ByVal bmWidth As Long, ByVal bmHeight As Long) As ApiBitmap
Attribute CreateCompatibleBitmap.VB_Description = "Creates a bitmap with the same properties as this device context"
Attribute CreateCompatibleBitmap.VB_HelpID = 2354

Dim bitmapThis As ApiBitmap

Set bitmapThis = New ApiBitmap

bitmapThis.CreateCompatibleBitmap mHDC, bmWidth, bmHeight

Set CreateCompatibleBitmap = bitmapThis

End Function

Public Property Let CurrentX(ByVal newval As Long)
Attribute CurrentX.VB_Description = "The current x position for drawing operations  "
Attribute CurrentX.VB_HelpID = 2861

Dim ptCurrent As POINTAPI
Dim lRet As Long

lRet = MoveToEx(mHDC, newval, CurrentY, ptCurrent)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CurrentX", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION The current x position for drawing operations
Public Property Get CurrentX() As Long

Dim ptCurrent As POINTAPI

Dim lRet As Long
lRet = GetCurrentPositionEx(mHDC, ptCurrent)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CurrentX", GetLastSystemError
Else
    CurrentX = ptCurrent.x
End If

End Property

Public Property Let CurrentY(ByVal newval As Long)
Attribute CurrentY.VB_Description = "The current y position for drawing operations  "
Attribute CurrentY.VB_HelpID = 2863

Dim ptCurrent As POINTAPI
Dim lRet As Long

lRet = MoveToEx(mHDC, CurrentX, newval, ptCurrent)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CurrentY", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION The current y position for drawing operations
Public Property Get CurrentY() As Long

Dim ptCurrent As POINTAPI

Dim lRet As Long
lRet = GetCurrentPositionEx(mHDC, ptCurrent)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":CurrentY", GetLastSystemError
Else
    CurrentY = ptCurrent.y
End If


End Property


Public Property Get CurveCapability(ByVal Capability As CurveCapabilities) As Boolean
Attribute CurveCapability.VB_HelpID = 2355

Dim lRet As Long

lRet = GetDeviceCaps(cCURVECAPS)
CurveCapability = (lRet And Capability)

End Property

Public Property Get DeviceType() As DeviceContextTypes
Attribute DeviceType.VB_HelpID = 2356

DeviceType = GetDeviceCaps(cTECHNOLOGY)

End Property


' ##BLOCK_DESCRIPTION Draws the frame of a button in the rectangle RectIn
' ##BLOCK_REMARKS ButtonStyle - The type of button to draw (push button, radio etc.)
' ##BLOCK_REMARKS ButtonState - The state to draw the button in (up, down, disabled etc.)
Public Sub DrawFrameButton(ByVal RectIn As APIRect, ByVal ButtonStyle As DrawFrameButtonComponents, ByVal ButtonState As DrawFrameButtonComponentsState)
Attribute DrawFrameButton.VB_Description = "Draws the frame of a button in the rectangle RectIn"
Attribute DrawFrameButton.VB_HelpID = 2357

Dim lStyle As Long

lStyle = (ButtonStyle Or ButtonState)

Call DrawFrameControl(RectIn, DFC_BUTTON, lStyle)

End Sub

'\\ --[DrawFrameCaption]------------------------------------------------------------------
' ##BLOCK_DESCRIPTION Draws a caption rectangle in the given RECT
' ##BLOCK_REMARKS CaptionStyle - The part of the standard caption to draw (close box, minimise box etc.)
'\\ ---------------------------------------------------------------------------------
Public Sub DrawFrameCaption(ByVal RectIn As APIRect, ByVal CaptionStyle As DrawFrameCaptionComponents)
Attribute DrawFrameCaption.VB_Description = "Draws a caption rectangle in the given RECT"
Attribute DrawFrameCaption.VB_HelpID = 2358

Call DrawFrameControl(RectIn, DFC_CAPTION, CaptionStyle)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":DrawFrameCaption", GetLastSystemError
End If

End Sub


'\\ --[DrawEdge]----------------------------------------------------------
' ##BLOCK_DESCRIPTION Draws the edge of the rectangle in RectIn as per the parameters passed in.
' ##BLOCK_REMARKS Edge - The type of edge to draw (sunken,raised etc)
' ##BLOCK_REMARKS Flags - A combination of BorderFlags to set which edges to draw
'\\ ----------------------------------------------------------------------
Public Sub DrawEdge(ByVal RectIn As APIRect, ByVal Edge As EdgeBorderStyles, ByVal flags As Long)
Attribute DrawEdge.VB_Description = "Draws the edge of the rectangle in RectIn as per the parameters passed in."
Attribute DrawEdge.VB_HelpID = 2359

Dim lRet As Long
Dim rcThis As RECT

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Right = RectIn.Right
    .Top = RectIn.Top
End With

lRet = DrawEdgeApi(mHDC, rcThis, Edge, flags)
If (Err.LastDllError <> 0) Or (lRet = 0) Then
    '\\ An error occured
    Call ReportError(Err.LastDllError, CLSNAME & ":DrawEdge", GetLastSystemError)
End If

End Sub

Private Sub DrawFrameControl(ByVal RectIn As APIRect, ByVal FrameType As DrawFrameControlTypes, ByVal Style As Long)

Dim rcThis As RECT
Dim lRet  As Long

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Right = RectIn.Right
    .Top = RectIn.Top
End With

lRet = DrawFrameControlApi(mHDC, rcThis, FrameType, Style)
If (Err.LastDllError > 0) Or (lRet = 0) Then
    Call ReportError(Err.LastDllError, CLSNAME & ":DrawFrameControl", GetLastSystemError)
End If

End Sub

' ##BLOCK_DESCRIPTION Draws the given menu style frame in the rectangle RectIn
' ##BLOCK_REMARKS MenuStyle - The type of menu frame to draw
Public Sub DrawFrameMenu(ByVal RectIn As APIRect, ByVal MenuStyle As DrawFrameMenuComponents)
Attribute DrawFrameMenu.VB_Description = "Draws the given menu style frame in the rectangle RectIn"
Attribute DrawFrameMenu.VB_HelpID = 2360

Call DrawFrameControl(RectIn, DFC_MENU, MenuStyle)

End Sub

' ##BLOCK_DESCRIPTION Draws the frame of a popup menu in the given rectangle
' ##BLOCK_REMARKS Style - The type of popup
Public Sub DrawFramePopup(ByVal RectIn As APIRect, ByVal Style As DrawFrameMenuComponentsPopupMenu)
Attribute DrawFramePopup.VB_Description = "Draws the frame of a popup menu in the given rectangle"
Attribute DrawFramePopup.VB_HelpID = 2361

Call DrawFrameControl(RectIn, DFC_POPUPMENU, Style)

End Sub

' ##BLOCK_DESCRIPTION Draws a scrollbar control in the given rectangle
' ##BLOCK_REMARKS ScrollStyle - The part of the scrollbar to draw
Public Sub DrawFrameScroll(ByVal RectIn As APIRect, ByVal ScrollStyle As DrawFrameScrollComponents)
Attribute DrawFrameScroll.VB_Description = "Draws a scrollbar control in the given rectangle"
Attribute DrawFrameScroll.VB_HelpID = 2362

Call DrawFrameControl(RectIn, DFC_SCROLL, ScrollStyle)
If (Err.LastDllError <> 0) Then
    '\\ An error occured
    Call ReportError(Err.LastDllError, CLSNAME & ":DrawFrameScroll", GetLastSystemError)
End If

End Sub

' ##BLOCK_DESCRIPTION Draws the given icon at the x,y position specified
Public Sub DrawIcon(ByVal xPos As Long, ByVal yPos As Long, ByVal Icon As ApiIcon)
Attribute DrawIcon.VB_Description = "Draws the given icon at the x,y position specified"
Attribute DrawIcon.VB_HelpID = 2363

Dim lRet As Long

lRet = DrawIconApi(mHDC, xPos, yPos, Icon.hIcon)
If (Err.LastDllError > 0) Or (lRet = 0) Then
    Call ReportError(Err.LastDllError, CLSNAME & ":DrawIcon", GetLastSystemError)
End If

End Sub

'\\ --[DrawRegion]-----------------------------------------------------------
' ##BLOCK_DESCRIPTION Draws the region passed in on the current device context, according to the brush and width/height or filled properties
' ##BLOCK_REMARKS rgnIn - The %region:EventVB~ApiRegion% to draw
' ##BLOCK_REMARKS brshIn - The %brush:EventVB~ApiLogBrush% to use
' ##BLOCK_REMARKS LineWidth - The x width of the line around the region
' ##BLOCK_REMARKS LineHeight - The y width of the line around the region
' ##BLOCK_REMARKS Fill - True to fill in the region
'\\ -------------------------------------------------------------------------
Public Sub DrawRegion(ByVal rgnIn As ApiRegion, ByVal brshIn As ApiLogBrush, ByVal LineWidth As Long, ByVal LineHeight As Long, ByVal Fill As Boolean)
Attribute DrawRegion.VB_Description = "Draws the region passed in on the current device context, according to the brush and width/height or filled properties"
Attribute DrawRegion.VB_HelpID = 2364

Dim lRet As Long

If Fill Then
    lRet = FillRgn(mHDC, rgnIn.hRgn, brshIn.Handle)
Else
    lRet = FrameRgn(mHDC, rgnIn.hRgn, brshIn.Handle, LineWidth, LineHeight)
End If
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":DrawRegion", GetLastSystemError
End If

End Sub

Public Property Get DriverVersion() As Long
Attribute DriverVersion.VB_HelpID = 2365

    DriverVersion = GetDeviceCaps(cDRIVERVERSION)
    
End Property


' ##BLOCK_DESCRIPTION Draws an ellipse defined by the input parameters in the currently selected %pen:EventVB~ApiLogPen%
' ##BLOCK_REMARKS BoundingRectangle - The rectangle that completely contains the ellipse
Public Sub Ellipse(ByVal BoundingRectangle As APIRect)
Attribute Ellipse.VB_Description = "Draws an ellipse defined by the input parameters in the currently selected pen:EventVB~ApiLogPen"
Attribute Ellipse.VB_HelpID = 2864

Dim lRet As Long

With BoundingRectangle
    lRet = EllipseApi(mHDC, .Left, .Top, .Right, .Bottom)
End With
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Ellipse", GetLastSystemError
End If

End Sub

' ##BLOCK_DESCRIPTION Fills out from the start point to an area bounded by, or made up of, a given colour
' ##BLOCK_REMARKS StartFrom - The point to fill out from
' ##BLOCK_REMARKS TestColour - The colour defining the fill area
' ##BLOCK_REMARKS FillMode - How to control the fill
Public Sub FloodFill(ByVal StartFrom As APIPoint, ByVal TestColour As ApiColour, ByVal Mode As FloodFillModes)
Attribute FloodFill.VB_Description = "Fills out from the start point to an area bounded by, or made up of, a given colour"
Attribute FloodFill.VB_HelpID = 2865

Dim lRet As Long

lRet = ExtFloodFill(mHDC, StartFrom.x, StartFrom.y, TestColour.ColourRef, Mode)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":FloodFill", GetLastSystemError
End If

End Sub

Public Function PaintWithDesktop()
Attribute PaintWithDesktop.VB_HelpID = 3058

Dim lRet As Long

lRet = PaintDesktop(mHDC)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":PaintWithDesktop", GetLastSystemError
End If

End Function

' ##BLOCK_DESCRIPTION Draws an rectangle defined by the input parameters in the currently selected %pen:EventVB~ApiLogPen%
' ##BLOCK_REMARKS BoundingRectangle - The rectangle to draw
Public Sub Rectangle(ByVal BoundingRectangle As APIRect)
Attribute Rectangle.VB_Description = "Draws an rectangle defined by the input parameters in the currently selected pen:EventVB~ApiLogPen"
Attribute Rectangle.VB_HelpID = 2866

Dim lRet As Long

With BoundingRectangle
    lRet = RectangleApi(mHDC, .Left, .Top, .Right, .Bottom)
End With
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Ellipse", GetLastSystemError
End If

End Sub
Private Function GetDeviceCaps(ByVal dcIndex As enDeviceCapsIndexes) As Long

GetDeviceCaps = GetDeviceCapsApi(mHDC, dcIndex)

End Function






Public Property Get GraphicsMode() As GraphicsMode
Attribute GraphicsMode.VB_HelpID = 2366

Dim lRet As Long

lRet = GetGraphicsModeApi(mHDC)
If Err.LastDllError > 0 Or lRet = 0 Then
   Call ReportError(Err.LastDllError, CLSNAME & ":GraphicsMode (Get)", GetLastSystemError)
End If
GraphicsMode = lRet

End Property

Public Property Let GraphicsMode(ByVal gmNew As GraphicsMode)

Dim lRet As Long

lRet = SetGraphicsModeApi(mHDC, gmNew)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":GraphicsMode (Set)", GetLastSystemError
End If

End Property

Public Property Get hdc() As Long
Attribute hdc.VB_HelpID = 2367

    hdc = mHDC
    
End Property

Public Property Let hdc(ByVal newHDC As Long)

    mHDC = newHDC
    
End Property
Public Property Get Height(ByVal MeasurementScale As enDeviceMesaurementScale) As Long
Attribute Height.VB_HelpID = 2368

If MeasurementScale = DMS_Millimeters Then
    Height = GetDeviceCaps(cVERTSIZE)
Else
    Height = GetDeviceCaps(cVERTRES)
End If

End Property

Public Property Get HorizontalTextAlignment() As HorizontalTextAlignments
Attribute HorizontalTextAlignment.VB_HelpID = 2369

Dim lAlign As Long

lAlign = mTextAlign

Select Case True
Case lAlign And TA_RTLREADING
    HorizontalTextAlignment = TA_RTLREADING
Case lAlign And TA_CENTRE
    HorizontalTextAlignment = TA_CENTRE
Case lAlign And TA_RIGHT
    HorizontalTextAlignment = TA_RIGHT
Case Else
    HorizontalTextAlignment = TA_LEFT
End Select

End Property


' ##BLOCK_DESCRIPTION Draws a line between the two points in the current %pen:EventVB~ApiLogPen%
' ##BLOCK_REMARKS The CurrentX and CurrentY properties are set to the end point after this call
Public Sub Line(ByVal StartPoint As APIPoint, ByVal EndPoint As APIPoint)
Attribute Line.VB_Description = "Draws a line between the two points in the current pen:EventVB~ApiLogPen"
Attribute Line.VB_HelpID = 2869

Dim lRet As Long
Dim ptThis As POINTAPI

lRet = MoveToEx(mHDC, StartPoint.x, StartPoint.y, ptThis)
If Err.LastDllError Then
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, CLSNAME & ":Line", GetLastSystemError
    End If
Else
    lRet = LineTo(mHDC, EndPoint.x, EndPoint.y)
    If Err.LastDllError Then
        If Err.LastDllError > 0 Then
            ReportError Err.LastDllError, CLSNAME & ":Line", GetLastSystemError
        End If
    End If
End If

End Sub

Public Property Get LineCapability(ByVal Capability As LineCapabilities) As Boolean
Attribute LineCapability.VB_HelpID = 2370

Dim lRet As Long

lRet = GetDeviceCaps(cLINECAPS)
LineCapability = (lRet And Capability)

End Property

Public Property Let MappingMode(ByVal newMode As MappingModes)
Attribute MappingMode.VB_HelpID = 2372

Static OldMappingMode As MappingModes

If newMode <> OldMappingMode Then
    OldMappingMode = SetMapMode(mHDC, newMode)
End If

End Property

Public Property Get MappingMode() As MappingModes

MappingMode = GetMapMode(mHDC)

End Property


Private Property Let mTextAlign(ByVal newAlign As Long)

    Call SetTextAlign(mHDC, newAlign)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, CLSNAME & ":SetTextAlign", GetLastSystemError
    End If
    
End Property

Private Property Get mTextAlign() As Long

mTextAlign = GetTextAlign(mHDC)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":GetTextAlign", GetLastSystemError
End If

End Property

Public Property Get Pens() As Collection
Attribute Pens.VB_HelpID = 2373

Dim colPens As Collection

Set colPens = GetDCPenCollection(mHDC)

If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Pens", GetLastSystemError
End If

Set Pens = colPens

End Property

Public Property Get PixelAspectDiagonal() As Long
Attribute PixelAspectDiagonal.VB_HelpID = 2374

    PixelAspectDiagonal = GetDeviceCaps(cASPECTXY)

End Property

Public Property Get PixelAspectHeight() As Long
Attribute PixelAspectHeight.VB_HelpID = 2375

    PixelAspectHeight = GetDeviceCaps(cASPECTY)
    
End Property

Public Property Get PixelAspectWidth() As Long
Attribute PixelAspectWidth.VB_HelpID = 2376

    PixelAspectWidth = GetDeviceCaps(cASPECTX)
    
End Property

Public Property Get PixelsPerInch_Horizontal() As Long
Attribute PixelsPerInch_Horizontal.VB_HelpID = 2377

PixelsPerInch_Horizontal = GetDeviceCaps(cLOGPIXELSX)

End Property
Public Property Get PixelsPerInch_Vertical() As Long
Attribute PixelsPerInch_Vertical.VB_HelpID = 2378

PixelsPerInch_Vertical = GetDeviceCaps(cLOGPIXELSY)

End Property



Public Property Get PolygonCapability(ByVal Capability As PolygonCapabilities) As Boolean
Attribute PolygonCapability.VB_HelpID = 2379

Dim lRet As Long

lRet = GetDeviceCaps(cPOLYGONALCAPS)
PolygonCapability = (lRet And Capability)

End Property

Public Property Get RasterCapability(ByVal Capability As RasterCapabilities) As Boolean
Attribute RasterCapability.VB_HelpID = 2380

Dim lRet As Long

lRet = GetDeviceCaps(cRASTERCAPS)
RasterCapability = (lRet And Capability)


End Property

Public Sub ScrollViewport(ByVal dx As Long, ByVal dy As Long)
Attribute ScrollViewport.VB_HelpID = 3161

Dim szThis As Size
Call OffsetViewportOrgEx(mHDC, dx, dy, szThis)

End Sub

Public Sub ScrollWindow(ByVal dx As Long, ByVal dy As Long)
Attribute ScrollWindow.VB_HelpID = 3162

Dim szThis As Size
Call OffsetWindowOrgEx(mHDC, dx, dy, szThis)

End Sub

Public Property Set SelectedBrush(ByVal brushNew As ApiLogBrush)
Attribute SelectedBrush.VB_HelpID = 2382

Dim lRet As Long

If brushNew Is Nothing Then
    If mInitialBrush Is Nothing Then
        '\\ Nothing has ever been selected to unselect
    Else
        lRet = SelectObject(mHDC, mInitialBrush.Handle)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":SelectedBrush (Set)", GetLastSystemError
        End If
    End If
Else
    lRet = SelectObject(mHDC, brushNew.Handle)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelectedBrush (Set)", GetLastSystemError
    End If
    '\\ Initial fon is not ours to unselect
    If mInitialBrush Is Nothing Then
        Set mInitialBrush = New ApiLogBrush
        mInitialBrush.Handle = lRet
    End If
End If

End Property

Public Property Get SelectedBrush() As ApiLogBrush

    Dim lOldBrush As Long
    Dim lNewBrush As Long
    Dim lBrush As ApiLogBrush
    
    lOldBrush = SelectObject(mHDC, APIDispenser.System.GraphicalDeviceInterface.StockBrush(BLACK_BRUSH).Handle)
    lNewBrush = SelectObject(mHDC, lOldBrush)
    
    Set lBrush = New ApiLogBrush
    lBrush.Handle = lOldBrush
    
    Set SelectedBrush = lBrush
    
End Property

Public Property Get SelectedFont() As ApiLogFont
Attribute SelectedFont.VB_HelpID = 2383

    Dim lOldFont As Long
    Dim lNewFont As Long
    Dim lFont As ApiLogFont
    
    lOldFont = SelectObject(mHDC, APIDispenser.System.GraphicalDeviceInterface.StockFont(ANSI_FIXED_FONT).Handle)
    lNewFont = SelectObject(mHDC, lOldFont)
    
    Set lFont = New ApiLogFont
    lFont.Handle = lOldFont
    
    Set SelectedFont = lFont
    
End Property


Public Property Set SelectedFont(ByVal fntNew As ApiLogFont)

Dim lRet As Long

If fntNew Is Nothing Then
    If mInitialFont Is Nothing Then
        '\\ Nothing has ever been selected to unselect
    Else
        lRet = SelectObject(mHDC, mInitialFont.Handle)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":SelectedFont (Set)", GetLastSystemError
        End If
    End If
Else
    lRet = SelectObject(mHDC, fntNew.Handle)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelectedFont (Set)", GetLastSystemError
    End If
    '\\ Initial fon is not ours to unselect
    If mInitialFont Is Nothing Then
        Set mInitialFont = New ApiLogFont
        mInitialFont.Handle = lRet
    End If
End If

End Property

Public Property Set SelectedPen(ByVal newpen As ApiLogPen)
Attribute SelectedPen.VB_HelpID = 2385

Dim lRet As Long

If newpen Is Nothing Then
    If mInitialPen Is Nothing Then
        '\\ Nothing has ever been selected to unselect
    Else
        lRet = SelectObject(mHDC, mInitialPen.Handle)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":SelectedPen (Set)", GetLastSystemError
        End If
    End If
Else
    lRet = SelectObject(mHDC, newpen.Handle)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SelectedPen (Set)", GetLastSystemError
    End If
    '\\ Initial fon is not ours to unselect
    If mInitialPen Is Nothing Then
        Set mInitialPen = New ApiLogPen
        mInitialPen.Handle = lRet
    End If
End If


End Property

Public Property Get SelectedPen() As ApiLogPen

    Dim lOldpen As Long
    Dim lNewpen As Long
    Dim lPen As ApiLogPen
    
    lOldpen = SelectObject(mHDC, APIDispenser.System.GraphicalDeviceInterface.StockPen(NULL_PEN).Handle)
    lNewpen = SelectObject(mHDC, lOldpen)
    
    Set lPen = New ApiLogPen
    lPen.Handle = lOldpen
    
    Set SelectedPen = lPen
    
End Property

Public Sub SetExtents(ByVal WindowWidth As Long, ByVal WindowHeight As Long, ByVal ViewportWidth As Long, ByVal ViewportHeight As Long)
Attribute SetExtents.VB_HelpID = 3164

Dim gx As Long
Dim gy As Long

gx = GreatestCommonDenominator(Abs(WindowWidth), Abs(ViewportWidth))
gy = GreatestCommonDenominator(Abs(WindowHeight), Abs(ViewportHeight))

Dim szThis As Size

Call SetWindowExtEx(mHDC, WindowWidth / gx, WindowHeight / gy, szThis)
Call SetViewportExtEx(mHDC, ViewportWidth / gx, ViewportHeight / gy, szThis)

End Sub

Public Property Let StretchBltMode(ByVal smNew As StretchBltModes)
Attribute StretchBltMode.VB_HelpID = 2387

Call SetStretchBltMode(mHDC, smNew)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":StretchBltMode (Let)", GetLastSystemError
End If
'\\ Note that if the mode is halftone you should reset the brush origin...
If smNew = SBM_HALFTONE Then
    Dim ptThis As APIPoint
    Set ptThis = New APIPoint
    ptThis.x = 0
    ptThis.y = 0
    Set BrushOrigin = ptThis
End If

End Property

Public Property Get StretchBltMode() As StretchBltModes

    StretchBltMode = GetStretchBltMode(mHDC)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, "ApiDeviceContect:StretchBltMode (Get)", GetLastSystemError
    End If
    
End Property


Public Property Get TextCapability(ByVal Capability As TextCapabilities) As Boolean
Attribute TextCapability.VB_HelpID = 2388

Dim lRet As Long

lRet = GetDeviceCaps(cTEXTCAPS)
TextCapability = (lRet And Capability)

End Property

Public Property Set TextColour(ByVal newColour As ApiColour)
Attribute TextColour.VB_HelpID = 2390

Call SetTextColor(mHDC, newColour.ColourRef)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":TextColour", GetLastSystemError
End If

End Property

Public Property Get TextColour() As ApiColour

Dim colourThis As ApiColour
Dim lcolourRef As Long

Set colourThis = New ApiColour
lcolourRef = GetTextColor(mHDC)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":TextColour", GetLastSystemError
End If
colourThis.ColourRef = lcolourRef
Set TextColour = colourThis

End Property

Public Sub TextOut(ByVal xPos As Long, ByVal yPos As Long, ByVal Text As String)
Attribute TextOut.VB_HelpID = 2391

Dim lRet As Long

lRet = TextOutApi(mHDC, xPos, yPos, Text, Len(Text))
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":TextOut", GetLastSystemError
End If

End Sub

Public Property Let UpdateTextCurrentPosition(ByVal bIn As Boolean)
Attribute UpdateTextCurrentPosition.VB_HelpID = 2393
 
 Dim lAlign As Long
 
 If bIn Then
    '\\ Add TA_UPDATECP to mTextAlign
    mTextAlign = mTextAlign Or TA_UPDATECP
 Else
    '\\ Remove TA_UPDATECP from mTextAlign
    mTextAlign = mTextAlign Xor TA_UPDATECP
 End If
 
End Property

Public Property Get UpdateTextCurrentPosition() As Boolean

Dim lAlign As Long

lAlign = mTextAlign
UpdateTextCurrentPosition = (lAlign And TA_UPDATECP)

End Property

Public Property Get VerticalTextAlignment() As VerticalTextAlignments
Attribute VerticalTextAlignment.VB_HelpID = 2394

Dim lAlign As Long

lAlign = mTextAlign

Select Case True
Case lAlign And TA_BASELINE
    VerticalTextAlignment = TA_BASELINE
Case lAlign And TA_BOTTOM
    VerticalTextAlignment = TA_BOTTOM
Case Else
    VerticalTextAlignment = TA_TOP
End Select

End Property

Public Property Get Width(ByVal MeasurementScale As enDeviceMesaurementScale) As Long
Attribute Width.VB_HelpID = 2395

If MeasurementScale = DMS_Millimeters Then
    Width = GetDeviceCaps(cHORZSIZE)
Else
    Width = GetDeviceCaps(cHORZRES)
End If

End Property


Public Property Let ViewportHeight(ByVal newval As Long)
Attribute ViewportHeight.VB_Description = "ViewportHeight - The height of the viewport that corresponds to the mapping of the window onto the device  "
Attribute ViewportHeight.VB_HelpID = 3063

Dim szThis As Size
Call GetViewportExtEx(mHDC, szThis)
Call SetViewportExtEx(mHDC, szThis.cx, newval, szThis)

End Property

' ##BLOCK_DESCRIPTION ViewportHeight - The height of the viewport that corresponds to the mapping of the window onto the device
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get ViewportHeight() As Long

Dim szThis As Size
Call GetViewportExtEx(mHDC, szThis)
ViewportHeight = szThis.cy

End Property

Public Property Let ViewportLeft(ByVal newval As Long)
Attribute ViewportLeft.VB_Description = "ViewportLeft - The left extent of the viewport that corresponds to the mapping of the window onto the device  "
Attribute ViewportLeft.VB_HelpID = 3065


Dim szThis As Size
Call GetViewportOrgEx(mHDC, szThis)
Call SetViewportOrgEx(mHDC, newval, szThis.cy, szThis)

End Property

' ##BLOCK_DESCRIPTION ViewportLeft - The left extent of the viewport that corresponds to the mapping of the window onto the device
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get ViewportLeft() As Long

Dim szThis As Size
Call GetViewportOrgEx(mHDC, szThis)
ViewportLeft = szThis.cx

End Property

Public Property Let ViewportTop(ByVal newval As Long)
Attribute ViewportTop.VB_Description = "ViewportTop - The top extent of the viewport that corresponds to the mapping of the window onto the device  "
Attribute ViewportTop.VB_HelpID = 3067

Dim szThis As Size
Call GetViewportOrgEx(mHDC, szThis)
Call SetViewportOrgEx(mHDC, szThis.cx, newval, szThis)

End Property

' ##BLOCK_DESCRIPTION ViewportTop - The top extent of the viewport that corresponds to the mapping of the window onto the device
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get ViewportTop() As Long

Dim szThis As Size
Call GetViewportOrgEx(mHDC, szThis)
ViewportTop = szThis.cy

End Property

Public Property Let ViewportWidth(ByVal newval As Long)
Attribute ViewportWidth.VB_Description = "ViewportHeight - The width of the viewport that corresponds to the mapping of the window onto the device  "
Attribute ViewportWidth.VB_HelpID = 3069

Dim szThis As Size
Call GetViewportExtEx(mHDC, szThis)
Call SetViewportExtEx(mHDC, newval, szThis.cy, szThis)

End Property

' ##BLOCK_DESCRIPTION ViewportHeight - The width of the viewport that corresponds to the mapping of the window onto the device
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get ViewportWidth() As Long

Dim szThis As Size
Call GetViewportExtEx(mHDC, szThis)
ViewportWidth = szThis.cx

End Property


Public Property Let WindowHeight(ByVal newval As Long)
Attribute WindowHeight.VB_Description = "ViewportHeight - The height of the logical window  "
Attribute WindowHeight.VB_HelpID = 3071

Dim szThis As Size
Call GetWindowExtEx(mHDC, szThis)
Call SetWindowExtEx(mHDC, szThis.cx, newval, szThis)

End Property

' ##BLOCK_DESCRIPTION ViewportHeight - The height of the logical window
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get WindowHeight() As Long

Dim szThis As Size
Call GetWindowExtEx(mHDC, szThis)
WindowHeight = szThis.cy

End Property

Public Property Let WindowLeft(ByVal newval As Long)
Attribute WindowLeft.VB_Description = "WindowLeft - The left extent of the logical window  "
Attribute WindowLeft.VB_HelpID = 3073


Dim szThis As Size
Call GetWindowOrgEx(mHDC, szThis)
Call SetWindowOrgEx(mHDC, newval, szThis.cy, szThis)

End Property

' ##BLOCK_DESCRIPTION WindowLeft - The left extent of the logical window
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get WindowLeft() As Long

Dim szThis As Size
Call GetWindowOrgEx(mHDC, szThis)
WindowLeft = szThis.cx

End Property

Public Property Let WindowTop(ByVal newval As Long)
Attribute WindowTop.VB_Description = "WindowTop - The top extent of the logical window  "
Attribute WindowTop.VB_HelpID = 3075

Dim szThis As Size
Call GetWindowOrgEx(mHDC, szThis)
Call SetWindowOrgEx(mHDC, szThis.cx, newval, szThis)

End Property

' ##BLOCK_DESCRIPTION WindowTop - The top extent of the logical window
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get WindowTop() As Long

Dim szThis As Size
Call GetWindowOrgEx(mHDC, szThis)
WindowTop = szThis.cy

End Property

Public Property Let WindowWidth(ByVal newval As Long)
Attribute WindowWidth.VB_Description = "WindowTop - The width of the logical window  "
Attribute WindowWidth.VB_HelpID = 3077

Dim szThis As Size
Call GetWindowExtEx(mHDC, szThis)
Call SetWindowExtEx(mHDC, newval, szThis.cy, szThis)

End Property

' ##BLOCK_DESCRIPTION WindowTop - The width of the logical window
' ##BLOCK_REMARKS This is only valid if the mapping mode is MM_ISOTROPIC or MM_ANISOTROPIC
Public Property Get WindowWidth() As Long

Dim szThis As Size
Call GetWindowExtEx(mHDC, szThis)
WindowWidth = szThis.cx

End Property

Private Sub Class_Initialize()

#If PRODUCTION Then
    On Error Resume Next
#End If

End Sub

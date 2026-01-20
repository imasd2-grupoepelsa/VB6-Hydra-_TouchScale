VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiGraphicalDeviceInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2745
Attribute VB_Description = "This class provides properties and methods for interacting with the Graphical Device Interface.  This is an interface to the graphical hardware (screens, printers, plotters etc.) which allows a wide range of physical devices to display graphics in a consistent manner."
Option Explicit

' ##MODULE_DESCRIPTION This class provides properties and methods for interacting _
with the Graphical Device Interface.  This is an interface to the graphical _
hardware (screens, printers, plotters etc.) which allows a wide range of physical _
devices to display graphics in a consistent manner.

' ##MODULE_DESCRIPTION The GDI provides a number of stock objects such as _
%brushes:EventVB~ApiLogBrush%,%pens:EventVB~ApiLogPen% and %fonts:EventVB~ApiLogFont% _
which can be used on every %device context:EventVB~ApiDeviceContext% without _
having to be recreated each ti

' ##MODULE_DESCRIPTION The GDI also provides information about the system colours that the _
user has chosen  for the various elements of their windows interface.

Private Const CLSNAME = "ApiGraphicalDeviceInterface"


'\\ 1 - GDI stock objects....
Public Enum GDIStockBrushes
    BLACK_BRUSH = 4
    DKGRAY_BRUSH = 3
    GRAY_BRUSH = 2
    NULL_BRUSH = 5
    LTGRAY_BRUSH = 1
    WHITE_BRUSH = 0
    DC_BRUSH = 18
End Enum

Public Enum GDIStockPens
    BLACK_PEN = 7
    NULL_PEN = 8
    WHITE_PEN = 6
    DC_PEN = 19
End Enum

Public Enum GDIStockFonts
    OEM_FIXED_FONT = 10
    ANSI_FIXED_FONT = 11
    ANSI_VAR_FONT = 12
    SYSTEM_FONT = 13
    DEVICE_DEFAULT_FONT = 14
    SYSTEM_FIXED_FONT = 16
    DEFAULT_GUI_FONT = 17
End Enum

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'\\ System defined colours
'(This matches the vb defined SystemColorConstants enum type)
Public Enum SystemColourIndexes
     COLOR_SCROLLBAR = 0
     COLOR_BACKGROUND = 1
     COLOR_ACTIVECAPTION = 2
     COLOR_INACTIVECAPTION = 3
     COLOR_MENU = 4
     COLOR_WINDOW = 5
     COLOR_WINDOWFRAME = 6
     COLOR_MENUTEXT = 7
     COLOR_WINDOWTEXT = 8
     COLOR_CAPTIONTEXT = 9
     COLOR_ACTIVEBORDER = 10
     COLOR_INACTIVEBORDER = 11
     COLOR_APPWORKSPACE = 12
     COLOR_HIGHLIGHT = 13
     COLOR_HIGHLIGHTTEXT = 14
     COLOR_BTNFACE = 15
     COLOR_BTNSHADOW = 16
     COLOR_GRAYTEXT = 17
     COLOR_BTNTEXT = 18
     COLOR_INACTIVECAPTIONTEXT = 19
     COLOR_BTNHIGHLIGHT = 20
     COLOR_3DDKSHADOW = 21
     COLOR_3DLIGHT = 22
     COLOR_INFOTEXT = 23
     COLOR_INFOBK = 24
     COLOR_HOTLIGHT = 26
     COLOR_GRADIENTACTIVECAPTION = 27
     COLOR_GRADIENTINACTIVECAPTION = 28
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long


'\\ Stock bitmaps
Public Enum GDIStockBitmaps
     OBM_CLOSE = 32754
     OBM_UPARROW = 32753
     OBM_DNARROW = 32752
     OBM_RGARROW = 32751
     OBM_LFARROW = 32750
     OBM_REDUCE = 32749
     OBM_ZOOM = 32748
     OBM_RESTORE = 32747
     OBM_REDUCED = 32746
     OBM_ZOOMD = 32745
     OBM_RESTORED = 32744
     OBM_UPARROWD = 32743
     OBM_DNARROWD = 32742
     OBM_RGARROWD = 32741
     OBM_LFARROWD = 32740
     OBM_MNARROW = 32739
     OBM_COMBO = 32738
     OBM_UPARROWI = 32737
     OBM_DNARROWI = 32736
     OBM_RGARROWI = 32735
     OBM_LFARROWI = 32734
     OBM_OLD_CLOSE = 32767
     OBM_SIZE = 32766
     OBM_OLD_UPARROW = 32765
     OBM_OLD_DNARROW = 32764
     OBM_OLD_RGARROW = 32763
     OBM_OLD_LFARROW = 32762
     OBM_BTSIZE = 32761
     OBM_CHECK = 32760
     OBM_CHECKBOXES = 32759
     OBM_BTNCORNERS = 32758
     OBM_OLD_REDUCE = 32757
     OBM_OLD_ZOOM = 32756
     OBM_OLD_RESTORE = 32755
End Enum

Private Declare Function LoadStockBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long

Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GdiGetBatchLimit Lib "gdi32" () As Long
Private Declare Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Let BatchLimit(ByVal newLimit As Long)
Attribute BatchLimit.VB_HelpID = 2747

Dim lret As Long

lret = GdiSetBatchLimit(newLimit)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiGraphicalDeviceContext:BatchLimit", GetLastSystemError
End If

End Property

Public Property Get BatchLimit() As Long

BatchLimit = GdiGetBatchLimit()
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":BatchLimit", GetLastSystemError
End If

End Property


Public Sub Flush()
Attribute Flush.VB_HelpID = 2748

Dim lret As Long
lret = GdiFlush
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiGrahicalDeviceInterface:Flush", GetLastSystemError
End If

End Sub

Public Property Get Regions() As ApiRgnDispenser
Attribute Regions.VB_HelpID = 2749

Dim rgnThis As ApiRgnDispenser

Set rgnThis = New ApiRgnDispenser

Set Regions = rgnThis

End Property


Public Property Get StockBitmap(ByVal Index As GDIStockBitmaps) As ApiBitmap
Attribute StockBitmap.VB_HelpID = 2750

Dim hBitmap As Long
Dim bmpThis As ApiBitmap

Set bmpThis = New ApiBitmap

hBitmap = LoadStockBitmap(0, Index)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":StockBitmap", GetLastSystemError
End If
bmpThis.hBitmap = hBitmap

Set StockBitmap = bmpThis

End Property

Public Property Get StockBrush(ByVal Index As GDIStockBrushes) As ApiLogBrush
Attribute StockBrush.VB_HelpID = 2751

    Dim lBrush As ApiLogBrush
    
    Set lBrush = New ApiLogBrush
    
    With lBrush
        .IsStockObject = True
        .Handle = GetStockObject(Index)
    End With
    
    Set StockBrush = lBrush
    
End Property


Public Property Get StockFont(ByVal Index As GDIStockFonts) As ApiLogFont
Attribute StockFont.VB_HelpID = 2752

Dim lFont As ApiLogFont

Set lFont = New ApiLogFont

With lFont
    .IsStockObject = True
    .Handle = GetStockObject(Index)
End With

Set StockFont = lFont

End Property

Public Property Get StockPen(ByVal Index As GDIStockPens) As ApiLogPen
Attribute StockPen.VB_HelpID = 2753

Dim lPen As ApiLogPen

Set lPen = New ApiLogPen

With lPen
    .IsStockObject = True
    .Handle = GetStockObject(Index)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":StockPen", GetLastSystemError
    End If
End With

Set StockPen = lPen

End Property


Public Property Set SystemColour(ByVal Index As SystemColourIndexes, ByVal newColour As ApiColour)
Attribute SystemColour.VB_HelpID = 2755

    Dim lret As Long
    
    lret = SetSysColors(1, Index, newColour.ColourRef)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":SystemColour (set)", GetLastSystemError
    End If
    
End Property

Public Property Get SystemColour(ByVal Index As SystemColourIndexes) As ApiColour

Dim colThis As New ApiColour

colThis.ColourRef = GetSysColor(Index)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":SystemColour (get)", GetLastSystemError
End If

Set SystemColour = colThis

End Property



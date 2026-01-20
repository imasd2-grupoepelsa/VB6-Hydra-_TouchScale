VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2400
Attribute VB_Description = "This class provides properties and methods for working with Icons.  An icon is a small picture which is used to represent an application and can be displayed on the windows desktop and also in the title bar of a window."
Option Explicit

' ##MODULE_DESCRIPTION This class provides properties and methods for _
working with Icons.  An icon is a small picture which is used to represent _
an application and can be displayed on the windows desktop and also in the title _
bar of a window.


Private Declare Function LoadIconApi Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function GetIconInfoApi Lib "user32" Alias "GetIconInfo" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private mhIcon As Long

Private mStockIcon As Boolean

'\\ Stock icons
Public Enum enStockIcons
    IDI_WINLOGO = 32517&
    IDI_APPLICATION = 32512&
    IDI_ASTERISK = 32516&
    IDI_EXCLAMATION = 32515&
    IDI_HAND = 32513&
    IDI_QUESTION = 32514&
    IDI_UNKNOWN = 32518&
End Enum

' ##ENUMERATION_DESCRIPTION IconTypes - The different types of icons associated with a window
Public Enum IconTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION ICON_SMALL - The small [16 x 16] icon
    ICON_SMALL = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION ICON_BIG - The larger [32 x 32] icon
    ICON_BIG = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION ICON_SMALL2 - The small icon.  If none use a system defined small icon (Windows 2000/XP only)
    ICON_SMALL2 = 2
End Enum

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property



Public Property Get ColourBitmap() As ApiBitmap
Attribute ColourBitmap.VB_HelpID = 2401

Dim info As ICONINFO

Dim bmThis As ApiBitmap
Dim lret As Long

Set bmThis = New ApiBitmap

lret = GetIconInfoApi(mhIcon, info)
If lret = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:ColourBitmap", GetLastSystemError)
End If

With info
    bmThis.hBitmap = .hbmColor
End With

Set ColourBitmap = bmThis

End Property

Public Property Let hIcon(ByVal newHandle As Long)
Attribute hIcon.VB_HelpID = 2403

If newHandle <> mhIcon Then
    mhIcon = newHandle
End If

End Property

Public Property Get hIcon() As Long
    
    hIcon = mhIcon
    
End Property

Public Property Get Hotspot() As APIPoint
Attribute Hotspot.VB_HelpID = 2404

Dim info As ICONINFO

Dim pointThis As APIPoint
Dim lret As Long

Set pointThis = New APIPoint

lret = GetIconInfoApi(mhIcon, info)
If lret = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:HotSpot", GetLastSystemError)
End If
With info
    pointThis.x = .xHotspot
    pointThis.y = .yHotspot
End With

Set Hotspot = pointThis

End Property

Public Sub LoadSystemIcon(ByVal StockIcon As enStockIcons)
Attribute LoadSystemIcon.VB_HelpID = 2405

Dim lret As Long

lret = LoadIconApi(0, StockIcon)
If Err.LastDllError = 0 Then
    hIcon = lret
    '\\ Flag that this icon was loded from a stock icon...
    mStockIcon = True
End If

End Sub


Public Property Get MaskBitmap() As ApiBitmap
Attribute MaskBitmap.VB_HelpID = 2406

Dim info As ICONINFO

Dim bmThis As ApiBitmap
Dim lret As Long

Set bmThis = New ApiBitmap

lret = GetIconInfoApi(mhIcon, info)
If lret = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:MaskBitmap", GetLastSystemError)
End If

With info
    bmThis.hBitmap = .hbmMask
End With

Set MaskBitmap = bmThis

End Property


Private Sub Class_Terminate()
Dim lret As Long

If mhIcon > 0 And Not (mStockIcon) Then
    lret = DestroyIcon(mhIcon)
    If lret = 0 Or Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiIcon:Terminate", GetLastSystemError)
    End If
End If

End Sub



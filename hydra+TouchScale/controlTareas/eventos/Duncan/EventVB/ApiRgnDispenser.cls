VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiRgnDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2328
Option Explicit

Private Declare Function CreateRectRgnApi Lib "gdi32" Alias "CreateRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgnApi Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgnApi Lib "gdi32" Alias "CreatePolygonRgn" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgnApi Lib "gdi32" Alias "CreateEllipticRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgnApi Lib "gdi32" Alias "CombineRgn" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long


Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Enum enCombineRegionModes
    RGN_AND = 1
    RGN_COPY = 5
    RGN_DIFF = 4
    RGN_OR = 2
    RGN_XOR = 3
End Enum

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Function CombineRegions(ByVal rgnSource1 As ApiRegion, ByVal rgnSource2 As ApiRegion, ByVal CombineMode As enCombineRegionModes) As ApiRegion
Attribute CombineRegions.VB_HelpID = 2329

Dim lRet As Long
Dim rgnThis As ApiRegion
Dim hRgnDest As Long

Dim rcThis As New APIRect

Set rgnThis = APIDispenser.System.GraphicalDeviceInterface.Regions.NewRectRgn(rcThis)

lRet = CombineRgnApi(rgnThis.hRgn, rgnSource1.hRgn, rgnSource2.hRgn, CombineMode)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiRgnDispenser:CombineRegions", GetLastSystemError
End If
Set CombineRegions = rgnThis

End Function

Public Function NewEllipticRgn(ByVal RectIn As APIRect) As ApiRegion
Attribute NewEllipticRgn.VB_HelpID = 2330

Dim rgnOut As ApiRegion
Dim lRet As Long

With RectIn
    lRet = CreateEllipticRgnApi(.Left, .Top, .Right, .Bottom)
End With
Set rgnOut = New ApiRegion
rgnOut.hRgn = lRet

Set NewEllipticRgn = rgnOut

End Function

Public Function NewPolygonRgn(ByVal Points As colApiPoint) As ApiRegion
Attribute NewPolygonRgn.VB_HelpID = 2331

Dim arPoints() As POINTAPI
Dim lIndex As Long
Dim lRet As Long
Dim ptThis As APIPoint
Dim rgnOut As ApiRegion

For Each ptThis In Points
    ReDim Preserve arPoints(0 To lIndex) As POINTAPI
    arPoints(lIndex).x = ptThis.x
    arPoints(lIndex).y = ptThis.y
    lIndex = lIndex + 1
Next ptThis
If lIndex > 0 Then
    lRet = CreatePolygonRgnApi(arPoints(0), (lIndex - 1), 1)
End If
Set rgnOut = New ApiRegion
rgnOut.hRgn = lRet
Set NewPolygonRgn = rgnOut

End Function

Public Function NewRectRgn(ByVal RectIn As APIRect) As ApiRegion
Attribute NewRectRgn.VB_HelpID = 2332

Dim rgnOut As ApiRegion
Dim lRet As Long

With RectIn
    lRet = CreateRectRgnApi(.Left, .Top, .Right, .Bottom)
End With
Set rgnOut = New ApiRegion
rgnOut.hRgn = lRet

Set NewRectRgn = rgnOut

End Function

Public Function NewRoundRectRgn(ByVal RectIn As APIRect, ByVal CornerWidth As Long, ByVal Cornerheight As Long) As ApiRegion
Attribute NewRoundRectRgn.VB_HelpID = 2333

Dim rgnOut As ApiRegion
Dim lRet As Long

With RectIn
    lRet = CreateRoundRectRgnApi(.Left, .Top, .Right, .Bottom, CornerWidth, Cornerheight)
End With
Set rgnOut = New ApiRegion
rgnOut.hRgn = lRet

Set NewRoundRectRgn = rgnOut

End Function



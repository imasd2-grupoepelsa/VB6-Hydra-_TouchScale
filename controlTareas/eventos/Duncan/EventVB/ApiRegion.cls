VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_HelpID = 2324
Attribute VB_Description="This class provides the properties and methods for manipulating regions."
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for manipulating regions.

' ##MODULE_DESCRIPTION A region is a bounded shape that can be as simple _
or as complex as you like.  More complex regions are typically made by _
combining a number of simple (rectangular, rounded rectangular, elliptic or polygonal) _
regions using boolean operators.

'\\ Offset this region
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long


'\\ Getting more details about the region...
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RGNDATA) As Long
Private Declare Function GetRegionDataLong Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Long) As Long

'\\ Freeing the object
Private Declare Function DeleteObjectApi Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type RGNDATAHEADER
    dwSize As Long
    iType As Long
    nCount As Long
    nRgnSize As Long
    rcBound As RECT
End Type

Private Type RGNDATA
    rdh As RGNDATAHEADER
    buffer() As Byte
End Type

'\\ Member variables
Private mhRgn As Long
Private mRgnData As RGNDATA
Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Let hRgn(ByVal newHRgn As Long)
Attribute hRgn.VB_HelpID = 2326

If newHRgn <> mhRgn Then
    mhRgn = newHRgn
End If

End Property

Public Property Get hRgn() As Long

    hRgn = mhRgn
    
End Property

Public Sub OffsetRegion(ByVal dX As Long, ByVal dY As Long)
Attribute OffsetRegion.VB_HelpID = 2327

Dim lRet As Long

lRet = OffsetRgn(mhRgn, dX, dY)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiRegion:OffsetRegion", GetLastSystemError
End If

End Sub

Private Sub RefreshRegionData()

Dim lRet As Long
Dim lSize As Long
Dim buffer() As Byte

lRet = GetRegionDataLong(mhRgn, lSize, ByVal 0&)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiRegion:GetData", GetLastSystemError
Else
    '\\ Lret is the size of the buffer needed for the region
    ReDim mRgnData.buffer(lRet) As Byte
    lRet = GetRegionData(mhRgn, lRet, mRgnData)
    If Err.LastDllError Then
        ReportError Err.LastDllError, "ApiRegion:GetData", GetLastSystemError
    End If
End If


End Sub


Private Sub Class_Terminate()

Call DeleteObjectApi(mhRgn)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiRegion:Terminate", GetLastSystemError
End If

End Sub



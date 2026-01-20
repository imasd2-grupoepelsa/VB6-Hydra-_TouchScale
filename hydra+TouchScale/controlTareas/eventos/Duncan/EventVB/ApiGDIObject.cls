VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiGDIObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2743
Attribute VB_Description="This class is a wrapper for the "Object" functions that are used to manipulate the GDI object table exported by GDI32.dll"
Option Explicit

'\\ --[ApiGDIObject]----------------------------------------------
' ##MODULE_DESCRIPTION This class is a wrapper for the "Object" functions that _
are used to manipulate the GDI object table exported by GDI32.dll

' ##ENUMERATION_DESCRIPTION
Public Enum GDIObjectTypes
    OBJ_BRUSH = 2
    OBJ_DC = 3
    OBJ_BITMAP = 7
    OBJ_ENHMETADC = 12
    OBJ_ENHMETAFILE = 13
    OBJ_EXTPEN = 11
    OBJ_FONT = 6   '\\ApiLogFont
    OBJ_MEMDC = 10
    OBJ_METADC = 4
    OBJ_METAFILE = 9
    OBJ_PAL = 5
    OBJ_PEN = 1
    OBJ_REGION = 8
End Enum


Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetGdiObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private mObjectHandle As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Friend Property Get ObjectHandle() As Long

    ObjectHandle = mObjectHandle

End Property

Friend Property Let ObjectHandle(ByVal newhandle As Long)

    mObjectHandle = newhandle
    
End Property

Public Property Get ObjectType() As GDIObjectTypes
Attribute ObjectType.VB_HelpID = 2744

Dim lRet As Long

If mObjectHandle <> 0 Then
    lRet = GetObjectType(mObjectHandle)
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, "ApiGDIObject:ObjectType", GetLastSystemError
    End If
    ObjectType = lRet
End If

End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2111
Attribute VB_Description = "This class provides the properties and methods for working with rectangles."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties and methods _
for working with rectangles.

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public CreatedOK As Boolean
Attribute CreatedOK.VB_VarHelpID = 2112

Private mLeft As Long
Private mTop As Long
Private mRight As Long
Private mBottom As Long

'\\ Private memory handling functions
Private Declare Sub CopyMemoryRect Lib "kernel32" Alias "RtlMoveMemory" (Destination As RECT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromRect Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As RECT, ByVal Length As Long)
Private Declare Function IsBadReadPtrRect Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrRect Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ RECT manipulation....
Private Declare Function InvalidateRectAPI Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetMenuItemRectAPI Lib "user32" Alias "GetMenuItemRect" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private bChanged As Boolean
Public Property Let Bottom(ByVal newBottom As Long)
Attribute Bottom.VB_HelpID = 2114

    If newBottom <> mBottom Then
        bChanged = True
        mBottom = newBottom
    End If
    
End Property

Public Property Get Bottom() As Long

    Bottom = mBottom
    
End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Function ContainsPoint(ByVal x As Long, ByVal y As Long) As Boolean
Attribute ContainsPoint.VB_HelpID = 2115

Dim rcTest As RECT

With rcTest
    .Left = Left
    .Top = Top
    .Right = Right
    .Bottom = Bottom
End With

ContainsPoint = PtInRect(rcTest, x, y)
If Err.LastDllError Then
    Call ReportError(Err.LastDllError, "ApiRect:ContainsPoint", GetLastSystemError)
End If

End Function


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this Rect object from the location poiunted to by
'\\ lpRect
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Friend Function CreateFromPointer(lpRect As Long) As Boolean

Dim ftThis As RECT

CreatedOK = False

If Not IsBadReadPtrRect(lpRect, Len(ftThis)) Then
    Call CopyMemoryRect(ftThis, lpRect, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            mLeft = .Left
            mRight = .Right
            mTop = .Top
            mBottom = .Bottom
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Function CreateFromWindow(ByVal hwnd As Long) As Boolean
Attribute CreateFromWindow.VB_HelpID = 2116

Dim lret As Long
Dim lpRect As RECT

CreatedOK = False
lret = GetWindowRect(hwnd, lpRect)
If Err.LastDllError = 0 Then
    With lpRect
        mLeft = .Left
        mRight = .Right
        mTop = .Top
        mBottom = .Bottom
    End With
    If Err.LastDllError = 0 Then
        CreatedOK = True
    End If
End If

CreateFromWindow = CreatedOK

End Function


Public Function GetMenuItemRect(ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long) As APIRect
Attribute GetMenuItemRect.VB_HelpID = 2117

Dim rcThis As RECT
Dim rcRet As APIRect
Dim lret As Long

lret = GetMenuItemRectAPI(hwnd, hMenu, uItem, rcThis)
If Err.LastDllError = 0 Then
    Set rcRet = New APIRect
    If rcRet.CreateFromPointer(VarPtr(rcThis)) Then
        Set GetMenuItemRect = rcRet
    End If
End If

End Function

Public Function InvalidateRect(ByVal hwnd As Long, ByVal uErase As Boolean) As Long
Attribute InvalidateRect.VB_HelpID = 2118

Dim lret As Long
Dim rcThis As RECT

With rcThis
    .Left = mLeft
    .Right = mRight
    .Top = mTop
    .Bottom = mBottom
End With

lret = InvalidateRectAPI(hwnd, rcThis, uErase)
If Err.LastDllError = 0 Then
    InvalidateRect = lret
End If

End Function


Public Property Let Left(ByVal newleft As Long)
Attribute Left.VB_HelpID = 2120

    If newleft <> mLeft Then
        bChanged = True
        mLeft = newleft
    End If
    
End Property

Public Property Get Left() As Long
    Left = mLeft
End Property


Public Property Let Right(ByVal newRight As Long)
Attribute Right.VB_HelpID = 2122

    If newRight <> mRight Then
        bChanged = True
        mRight = newRight
    End If
    
End Property

Public Property Get Right() As Long

    Right = mRight
    
End Property

Public Sub SaveToPointer(lpRect As Long)
Attribute SaveToPointer.VB_HelpID = 2123

Dim ftThis As RECT

With ftThis
    .Left = mLeft
    .Right = mRight
    .Top = mTop
    .Bottom = mBottom
End With

If Not IsBadReadPtrRect(lpRect, Len(ftThis)) Then
    Call CopyMemoryFromRect(lpRect, ftThis, Len(ftThis))
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, "ApiRect:SaveToPointer", GetLastSystemError
    End If
End If


End Sub


Public Property Let Top(ByVal newTop As Long)
Attribute Top.VB_HelpID = 2125

    If newTop <> mTop Then
        bChanged = True
        mTop = newTop
    End If
    
End Property

Public Property Get Top() As Long

    Top = mTop
    
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiGlobalmemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2460
Option Explicit

'\\ Global memory management functions
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private mMyData() As Byte
Private mMyDataSize As Long
Private mHmem As Long


Public Enum enGlobalmemoryAllocationConstants
    GMEM_FIXED = &H0
    GMEM_DISCARDABLE = &H100
    GMEM_MOVEABLE = &H2
    GMEM_NOCOMPACT = &H10
    GMEM_NODISCARD = &H20
    GMEM_ZEROINIT = &H40
End Enum

Private mAllocationType As enGlobalmemoryAllocationConstants

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Let AllocationType(ByVal newType As enGlobalmemoryAllocationConstants)
Attribute AllocationType.VB_HelpID = 2462

mAllocationType = newType

End Property

Public Property Get AllocationType() As enGlobalmemoryAllocationConstants

    AllocationType = mAllocationType
    
End Property


Private Sub CopyDataToGlobal()

Dim lret As Long

If mHmem > 0 Then
    lret = GlobalLock(mHmem)
    If lret > 0 Then
        Call CopyMemory(ByVal mHmem, mMyData(0), mMyDataSize)
        Call GlobalUnlock(mHmem)
    End If
End If

End Sub

Public Sub CopyFromHandle(ByVal hMemHandle As Long)
Attribute CopyFromHandle.VB_HelpID = 2463

Dim lret As Long
Dim lPtr As Long

lret = GlobalSize(hMemHandle)
If lret > 0 Then
    mMyDataSize = lret
    lPtr = GlobalLock(hMemHandle)
    If lPtr > 0 Then
        ReDim mMyData(0 To mMyDataSize - 1) As Byte
        CopyMemory mMyData(0), ByVal lPtr, mMyDataSize
        Call GlobalUnlock(hMemHandle)
    End If
End If

End Sub

Public Sub CopyToHandle(ByVal hMemHandle As Long)
Attribute CopyToHandle.VB_HelpID = 2464

Dim lSize As Long
Dim lPtr As Long

'\\ Don't copy if its empty
If Not IsEmpty Then
    lSize = GlobalSize(hMemHandle)
    '\\ Don't attempt to copy if zero size...
    If lSize > 0 Then
        If lPtr > 0 Then
            CopyMemory ByVal lPtr, mMyData(0), lSize
            Call GlobalUnlock(hMemHandle)
        End If
    End If
End If

End Sub


'\\ --[Handle]------------------------------------------------------
'\\ Returns a Global Memroy handle that is valid and filled with the
'\\ info held in this object's private byte array
'\\ ----------------------------------------------------------------
Public Property Get Handle() As Long
Attribute Handle.VB_HelpID = 2465

If mHmem = 0 Then
    If mMyDataSize > 0 Then
        mHmem = GlobalAlloc(AllocationType, mMyDataSize)
    End If
End If
Call CopyDataToGlobal

Handle = mHmem

End Property

Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_HelpID = 2466

    IsEmpty = (mMyDataSize = 0)

End Property

Public Sub Free()
Attribute Free.VB_HelpID = 2467

If mHmem > 0 Then
    Call GlobalFree(mHmem)
    mHmem = 0
    mMyDataSize = 0
    ReDim mMyData(0) As Byte
End If

End Sub

Private Sub Class_Terminate()

If mHmem > 0 Then
    Call GlobalFree(mHmem)
End If

End Sub



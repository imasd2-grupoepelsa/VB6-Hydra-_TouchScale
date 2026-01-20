Attribute VB_Name = "MemoryAPIFunctions"
'\\ -----[MemoryAPIFunctions]----------------------------------------------------
'\\ For copying structures (and strings) to/from pointers
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------

Option Explicit

'\\ Memory manipulation routines
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'\\ Pointer validation in StringFromPointer
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

Dim sTemp() As String
Dim lIndex As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Private Declare Function lstrcpyToBuffer Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Declare Sub CopyMemoryByte Lib "kernel32" Alias "RtlMoveMemory" (Destination As Byte, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryWord Lib "kernel32" Alias "RtlMoveMemory" (Destination As Integer, ByVal Source As Long, ByVal Length As Long)

Public Function MakeInt(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer

MakeInt = Val("&H" & PadString(Hex(HiByte), 2, "0") & PadString(Hex(LoByte), 2, "0"))

End Function

Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long

MakeLong = Val("&H" & PadString(Hex(HiWord), 4, "0") & PadString(Hex(LoWord), 4, "0"))

End Function
Public Function HiByte(ByVal w As Integer) As Byte

    Call CopyMemoryByte(HiByte, VarPtr(w) + 1, 1)
    
End Function

Public Function LoByte(w As Integer) As Byte

    Call CopyMemoryByte(LoByte, VarPtr(w), 1)
    
End Function

Public Function LoWord(dw As Long) As Integer

Call CopyMemoryWord(LoWord, VarPtr(dw), 2)

End Function

Public Function HiWord(dw As Long) As Integer

Call CopyMemoryWord(HiWord, VarPtr(dw) + 2, 2)

End Function


Public Function LPCSTR(ByVal sIn As String) As Long

Dim lChar As Long

If sIn = "" Then
    LPCSTR = 0
Else
    lIndex = lIndex + 1
    ReDim Preserve sTemp(1 To lIndex) As String
    sTemp(lIndex) = StrConv(sIn, vbFromUnicode)
    LPCSTR = StrPtr(sTemp(lIndex))
End If

End Function


Public Function LPFNfromLong(ByVal lIn As Long) As Long

    '\\ This is to get around the restriction whereby the AddressOf function
    '\\ can only be used in a parameter construct.
    LPFNfromLong = lIn
    
End Function

Public Function PadString(ByVal sIn As String, ByVal Length As Long, ByVal char As String)

If Len(sIn) >= Length Then
    PadString = Left$(sIn, Length)
Else
    PadString = String$(Length - Len(sIn), char) & sIn
End If

End Function

Public Sub ResetLPCSTRIndexes()

lIndex = 0
Erase sTemp

End Sub

'\\ --[StringFromPointer]-------------------------------------------------------------------
'\\ Returns a VB string from an API returned string pointer
'\\ Parameters:
'\\   lpString - The long pointer to the string
'\\   lMaxlength - the size of empty buffer to allow
'\\ HISTORY:
'\\  DEJ 28/02/2001 Check pointer is a valid string pointer...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String

Dim sRet As String
Dim lret As Long

If lpString = 0 Then
    StringFromPointer = ""
    Exit Function
End If

lret = lstrlen(lpString)
If lret < lMaxLength Then
    lMaxLength = lret
End If

If IsBadStringPtrByLong(lpString, lMaxLength) Then
    '\\ An error has occured - do not attempt to use this pointer
    Call ReportError(Err.LastDllError, "StringFromPointer", "Attempt to read bad string pointer: " & lpString)
    StringFromPointer = ""
    Exit Function
End If

'\\ Pre-initialise the return string...
sRet = Space$(lMaxLength)
Call lstrcpyToBuffer(sRet, lpString)
'CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
If Err.LastDllError = 0 Then
    If InStr(sRet, Chr$(0)) > 0 Then
        sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
End If

StringFromPointer = sRet

End Function


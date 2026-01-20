Attribute VB_Name = "test"
Option Explicit

Private Declare Function UnicodeToAnsi Lib "ntdll.dll" Alias "RtlUnicodeStringToAnsiString" (ByRef DestinationString As ANSI_STRING, ByVal SourceString As Long, Optional ByVal AllocateDestinationString As Byte) As Long
Private Declare Function AnsiToUnicode Lib "ntdll.dll" Alias "RtlAnsiStringToUnicodeString" (ByVal DestinationString As Long, ByRef SourceString As ANSI_STRING, Optional ByVal AllocateDestinationString As Byte) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type UNICODE_STRING
    Len As Integer
    MaxLen As Integer
    Buffer As String
End Type

Private Type ANSI_STRING
    Len As Integer
    MaxLen As Integer
    Buffer As Long
End Type
Public Function ByteToStr(bArray() As Byte) As String
    Dim lPntr As Long
    Dim bTmp() As Byte
    ReDim bTmp(UBound(bArray) * 2 + 1)
    For lPntr = 0 To UBound(bArray)
        bTmp(lPntr * 2) = bArray(lPntr)
    Next lPntr
    Let ByteToStr = bTmp
End Function

Public Function StrToByte(strInput As String) As Byte()
    Dim lPntr As Long
    Dim bTmp() As Byte
    Dim bArray() As Byte
    If Len(strInput) = 0 Then Exit Function
    ReDim bTmp(LenB(strInput) - 1) 'Memory length
    ReDim bArray(Len(strInput) - 1) 'String length
    CopyMemory bTmp(0), ByVal StrPtr(strInput), LenB(strInput)
    'Examine every second byte
    For lPntr = 0 To UBound(bArray)
        If bTmp(lPntr * 2 + 1) > 0 Then
            bArray(lPntr) = Asc(Mid$(strInput, lPntr + 1, 1))
        Else
            bArray(lPntr) = bTmp(lPntr * 2)
        End If
    Next lPntr
    StrToByte = bArray
End Function

Public Function UniToAnsi(sUnicode As String) As Byte()
    Dim UniString As UNICODE_STRING
    Dim AnsiString As ANSI_STRING
    Dim Buffer() As Byte
    If Len(sUnicode) = 0 Then Exit Function
    UniString.Buffer = sUnicode
    UniString.Len = LenB(UniString.Buffer)
    UniString.MaxLen = UniString.Len + 2
    AnsiString.Len = Len(UniString.Buffer)
    AnsiString.MaxLen = AnsiString.Len + 1
    ReDim Buffer(AnsiString.Len) As Byte
    AnsiString.Buffer = VarPtr(Buffer(0))
    If UnicodeToAnsi(AnsiString, VarPtr(UniString)) = 0 Then
        UniToAnsi = Buffer
        ReDim Preserve UniToAnsi(UBound(Buffer) - 1)
        sUnicode = ByteToStr(UniToAnsi)
    End If
End Function

'Public Function AnsiToUni(sAnsi As String) As Byte()
'    Dim UniString As UNICODE_STRING
'    Dim AnsiString As ANSI_STRING
'    Dim Buffer() As Byte
'    If Len(sAnsi) = 0 Then Exit Function
'    AnsiString.Buffer = sAnsi
'    AnsiString.Len = LenB(AnsiString.Buffer)
'    AnsiString.MaxLen = AnsiString.Len + 1
'    UniString.Len = Len(AnsiString.Buffer)
'    UniString.MaxLen = UniString.Len + 1
'    ReDim Buffer(UniString.Len) As Byte
'    UniString.Buffer = VarPtr(Buffer(0))
'    If AnsiToUnicode(VarPtr(UniString), AnsiString) = 0 Then
'        AnsiToUni = Buffer
'        ReDim Preserve AnsiToUni(UBound(Buffer) - 1)
'        sAnsi = ByteToStr(AnsiToUni)
'    End If
'End Function
Public Function AnsiToUni(sAnsi As String) As Byte()
    Dim Buffer() As Byte, ANSI_STR As ANSI_STRING, UNI_STR As UNICODE_STRING '

    ANSI_STR.Buffer = sAnsi '"How Quickly Daft Jumping Zebras Vex"
    ANSI_STR.Len = LenB(ANSI_STR.Buffer)
    ANSI_STR.MaxLen = ANSI_STR.Len + 1       '+ LenB(vbNullChar)

    UNI_STR.Len = Len(ANSI_STR.Buffer)
    ReDim Buffer(0& To UNI_STR.Len) As Byte
    UNI_STR.Buffer = VarPtr(Buffer(0&))
    UNI_STR.MaxLen = UNI_STR.Len + 2     '+ Len(vbNullChar)

    'If RtlUnicodeStringToAnsiString(ANSI_STR, VarPtr(UNI_STR)) = STATUS_SUCCESS Then
        UNI_STR.Buffer = Space$(ANSI_STR.Len) '
        Call AnsiToUnicode(VarPtr(UNI_STR), ANSI_STR)
        ReDim Preserve AnsiToUni(UBound(Buffer) - 1)
        sAnsi = ByteToStr(AnsiToUni)
    'End If
End Function


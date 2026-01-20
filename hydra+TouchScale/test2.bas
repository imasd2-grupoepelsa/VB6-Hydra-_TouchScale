Attribute VB_Name = "test2"
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
'Private Declare Function GetACP Lib "kernel32" () As Long

'Public Function DecodeString(ByVal strSource As String, Optional FromCodePage As Long = -1) As String
'    Dim strTemp As String'
'
'    If strSource = vbNullString Then Exit Function
'    strTemp = VUnicodeToAnsi(strSource, 0)
'    DecodeString = VAnsiToUnicode(strTemp, FromCodePage)
'End Function

'Public Function VAnsiToUnicode(ByVal strSource As String, Optional ByVal codePage As Long = -1, Optional lFlags As Long = 0) As String
'    Dim strBuffer As String
'    Dim cwch As Long
'    Dim pwz As Long
'    Dim pwzBuffer As Long

'    If codePage = -1 Then codePage = GetACP()
'    pwz = StrPtr(strSource)
'    cwch = MultiByteToWideChar(codePage, lFlags, pwz, -1, 0&, 0&)
'    strBuffer = String$(cwch + 1, vbNullChar)
'    pwzBuffer = StrPtr(strBuffer)
'    cwch = MultiByteToWideChar(codePage, lFlags, pwz, -1, pwzBuffer, Len(strBuffer))
'    AnsiToUnicode = left(strBuffer, cwch - 1)
'End Function

'Public Function VUnicodeToAnsi(ByVal strSource As String, Optional ByVal codePage As Long = -1, Optional lFlags As Long = 0) As String
'    Dim strBuffer As String
'    Dim cwch As Long
'    Dim pwz As Long
'    Dim pwzBuffer As Long''

'    If codePage = -1 Then codePage = GetACP()
'    pwz = StrPtr(strSource)
'    cwch = WideCharToMultiByte(codePage, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
'    strBuffer = String$(cwch + 1, vbNullChar)
'    pwzBuffer = StrPtr(strBuffer)
'    cwch = WideCharToMultiByte(codePage, lFlags, pwz, -1, pwzBuffer, Len(strBuffer), ByVal 0&, ByVal 0&)
'    UnicodeToAnsi = left(strBuffer, cwch - 1)
'End Function
Option Explicit

'Private Const CP_UTF8 As Long = 65001 ' UTF-8 Code Page

''Sys call to convert multiple byte chars to a char
'Private Declare Function MultiByteToWideChar Lib "KERNEL32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cchMultiByte As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long) As Long

''------------------------------------------------------------------
'' NAME:         DecodeURI (PUBLIC)
'' DESCRIPTION:  Decodes a UTF8 encoded string
'' CALLED BY:    HandleNavigate
'' PARAMETERS:
''  EncodedURL (I,REQ) - the UTF-8 encoded string to decode
'' RETURNS:      the the decoded UTF-8 string
''------------------------------------------------------------------
'Public Function DecodeURI(ByVal EncodedURI As String) As String
'    Dim bANSI() As Byte
'    Dim bUTF8() As Byte
'    Dim lIndex As Long
'    Dim lUTFIndex As Long'
'
'    If Len(EncodedURI) = 0 Then
'        Exit Function
'    End If''
'
'EncodedURI = Replace$(EncodedURI, "+", " ")         ' In case encoding isn't used.
'    bANSI = StrConv(EncodedURI, vbFromUnicode)          ' Convert from unicode text to ANSI values
'    ReDim bUTF8(UBound(bANSI))                          ' Declare dynamic array, get length
'    For lIndex = 0 To UBound(bANSI)                     ' from 0 to length of ANSI
'        If bANSI(lIndex) = &H25 Then                    ' If we have ASCII 37, %, then
'            bUTF8(lUTFIndex) = Val("&H" & Mid$(EncodedURI, lIndex + 2, 2)) ' convert hex to ANSI
'            lIndex = lIndex + 2                         ' this character was encoded into two bytes
'        Else
'            bUTF8(lUTFIndex) = bANSI(lIndex)            ' otherwise don't need to do anything special
'        End If
'        lUTFIndex = lUTFIndex + 1                       ' advance utf index
'    Next
'    DecodeURI = FromUTF8(bUTF8, lUTFIndex)              ' convert to string
'End Function

''------------------------------------------------------------------
'' NAME:         FromUTF8 (Private)
'' DESCRIPTION:  Use the system call MultiByteToWideChar to
''               get chars using more than one byte and return
''               return the whole string
'' CALLED BY:    DecodeURI
'' PARAMETERS:
''  UTF8 (I,REQ)   - the ID of the element to return
''  Length (I,REQ) - length of the string
'' RETURNS:      the full raw data of this field
''------------------------------------------------------------------
'Private Function FromUTF8(ByRef UTF8() As Byte, ByVal Length As Long) As String
'    Dim lDataLength As Long'
'
'    lDataLength = MultiByteToWideChar(CP_UTF8, 0, VarPtr(UTF8(0)), Length, 0, 0)  ' Get the length of the data.
'    FromUTF8 = String$(lDataLength, 0)                                         ' Create array big enough
'    MultiByteToWideChar CP_UTF8, 0, VarPtr(UTF8(0)), _
'                        Length, StrPtr(FromUTF8), lDataLength                  '
'End Function

'Private Declare Function GetStringTypeW Lib "kernel32" (ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, lpCharType As Integer) As Long
'
'Private Const CT_CTYPE1 As Long = &H1
'Private Const C1_SPACE  As Long = &H8
'Private Const C1_ALPHA  As Long = &H100
'
''Private Sub Form_Load()
''    Dim s As String
''
''    ' // Check greek string
''    s = ToProperCase(StrConv(StrConv("?? ????? ??? ??????????!", vbFromUnicode), vbUnicode, &H408))
''    TextOut Me.hdc, 0, 0, StrPtr(s), Len(s)
''
''End Sub
'
'Public Function ToProperCase(strIn As String) As String
'    Dim i   As Long
'    Dim t() As Integer
'    Dim b   As Boolean
'
'    If Len(strIn) = 0 Then Exit Function
'
'    ToProperCase = LCase(strIn)
'
'    ReDim t(Len(strIn) - 1)
'
'    If GetStringTypeW(CT_CTYPE1, StrPtr(ToProperCase), Len(ToProperCase), t(0)) = 0 Then
'
'        Err.Raise 51
'        ToProperCase = vbNullString
'        Exit Function
'
'    End If
'
'    b = True
'
'    For i = 0 To UBound(t)
'
'        If t(i) And C1_SPACE Then
'            b = True
'        Else
'
'            If (t(i) And C1_ALPHA) And b Then
'                Mid$(ToProperCase, i + 1, 1) = UCase(Mid$(ToProperCase, i + 1, 1))
'            End If
'
'            b = False
'
'        End If
'
'    Next
'
'End Function

'--------------------------------
'   WToA
'
'   UNICODE to ANSI conversion, via a given codepage
'--------------------------------
Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    Dim lpUsedDefaultChar As Long
    
    If cpg = -1 Then cpg = 1253 'GetACP()
    pwz = StrPtr(st)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    WToA = left$(stBuffer, cwch - 1)
End Function

'--------------------------------
' AToW
'
' ANSI to UNICODE conversion, via a given codepage.
'--------------------------------
Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
Dim stBuffer As String
Dim cwch As Long
Dim pwz As Long
Dim pwzBuffer As Long
Dim sResp As String

sResp = ""

'CadenadeLog "st:" & st

    If cpg = -1 Then cpg = 1253 'GetACP()
    pwz = StrPtr(st)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, Len(st), 0&, 0&)
'CadenadeLog "Respuesta MultiByteToWideChar:" & CStr(cwch)
'CadenadeLog "como va st:" & st
    stBuffer = String$(cwch * 2, Chr(2))
'CadenadeLog "a"

    pwzBuffer = StrPtr(stBuffer)
'CadenadeLog "b"

    cwch = MultiByteToWideChar(cpg, lFlags, pwz, Len(st) * 2, pwzBuffer, Len(stBuffer))
'CadenadeLog "2ºcomo va st:" & st
'CadenadeLog "2ºRespuesta MultiByteToWideChar:" & CStr(cwch)
'CadenadeLog "stbuffer:" & stBuffer
    For pwz = 1 To Len(stBuffer) Step 2
    '    CadenadeLog "asc(" & CStr(pwz) & "):" & Asc(Mid(stBuffer, pwz, 1))
        sResp = sResp & Mid(stBuffer, pwz, 1)
    Next pwz
'CadenadeLog sResp

    AToW = sResp 'Replace(left$(stBuffer, Len(stBuffer)), Chr(2), "")
End Function

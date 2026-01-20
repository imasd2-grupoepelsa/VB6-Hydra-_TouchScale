Attribute VB_Name = "utftostring"
Option Explicit

Public Function DecodeBase64(ByVal strData As String) As Byte()
Dim s() As Byte
Dim nL As Integer
Dim nC As Integer
    If IsNull(strData) Then GoTo salDecode
    nL = Len(strData)
    ReDim s(nL)
    For nC = 1 To nL
        s(nC) = Asc(Mid(strData, nC, 1))
    Next nC
salDecode:
    DecodeBase64 = s
End Function

' accept a byte array containing utf-8 data
' and convert it to a string
Public Function ConvertUtf8BytesToString(ByRef data() As Byte) As String

    Dim objStream As ADODB.Stream
    Dim strTmp As String
    
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.type = adTypeBinary
    objStream.Open
    
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.type = adTypeText
    strTmp = objStream.ReadText
    
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = Mid(strTmp, 2)

End Function

' accept a byte array containing utf-8 data and convert it to a string
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As String

    Dim objStream As ADODB.Stream
    Dim data() As Byte
    Dim nC, nL As Long
    Dim s As String
    If strText = "" Then
        s = ""
        GoTo salconver
    End If
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.type = adTypeText
    objStream.Open
    
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.type = adTypeBinary
    objStream.Read 3 ' skip first 3 bytes as this is the utf8 detection header
    data = objStream.Read()
    nL = objStream.Size
    ' close up and return
    objStream.Close
    On Error GoTo salconver
    s = ""
    For nC = 0 To nL - 1
        s = s & Chr(data(nC))
    Next nC
salconver:
    On Error GoTo 0
    ConvertStringToUtf8Bytes = s

End Function


Attribute VB_Name = "utftostring"
Option Explicit

' Convert UTF8 byte arrays into string and vice versa
' Article at: http://www.nonhostile.com/howto-convert-byte-array-utf8-string-vb6.asp?src=code188
' (c) Copyright 2006 Tim Hastings - please feel free to use in anyway you wish
'
' http://www.nonhostile.com/

'Public Sub Main()

'    Dim strB64 As String
'    Dim data() As Byte
'    Dim strTmp As String
'
'    ' define test data as base64 and decode to array of bytes
'    strB64 = "R3JlZXRpbmdzIGFuZCBTYWx1dGF0aW9uISAo4oKsKSBhbmQgc29"
'    strB64 = strB64 & "tZSBVcmR1OiDaqdix2KfahtuMINm+2Kfaqdiz2KrYp9mG24w="
'    data = DecodeBase64(strB64)'
'
'    ' convert from utf-8 to string
'    strTmp = ConvertUtf8BytesToString(data)
'
'    ' convert back to bytes
'    data = ConvertStringToUtf8Bytes(strTmp)
'
'End Sub


' accept a byte array containing utf-8 data
' and convert it to a string
Public Function ConvertUtf8BytesToString(ByRef data() As Byte) As String

    Dim objStream As ADODB.Stream
    Dim strTmp As String
    Dim nL As Integer
    Dim nC As Integer

    ' init stream
    Set objStream = New ADODB.Stream
    If Dir(App.Path & "\linux.txt") <> "" Then    '5.4.3
        objStream.Charset = "utf-8"
    Else
        If id = 6000 Then
            objStream.Charset = "iso-8859-7"
        Else
            objStream.Charset = "iso-8859-9"
        End If
    End If
    objStream.mode = adModeReadWrite
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
    Set objStream = Nothing
    ConvertUtf8BytesToString = Mid(strTmp, 2)

End Function


Public Function ConvertUtf8BytesToStringnW(ByRef data() As Byte) As String

    Dim objStream As ADODB.Stream
    Dim strTmp As String
    Dim nL As Integer
    Dim nC As Integer

    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "iso-8859-7"
    objStream.mode = adModeReadWrite
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
    Set objStream = Nothing
    ConvertUtf8BytesToStringnW = Mid(strTmp, 2)

End Function
' accept a byte array containing utf-8 data
' and convert it to a string
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As String

    Dim objStream As ADODB.Stream
    Dim data() As Byte
    Dim nC, nL As Long
    Dim s As String

    If id = 6000 Then
        s = UniToAnsi(strText)
    End If
    '''''''''''''''
    '    If Dir(App.Path & "\linux.txt") <> "" Then '5.4.3
    '
    '    If strText = "" Or IsNull(strText) Then
    '        s = ""
    '        GoTo salconver
    '    End If
    '    ' init stream
    '    Set objStream = New ADODB.Stream
    '    objStream.Charset = "utf-8"
    '    objStream.mode = adModeReadWrite
    '    objStream.type = adTypeText
    '    objStream.Open
    '
    '    ' write bytes into stream
    '    objStream.WriteText strText
    '    objStream.Flush
    '
    '    ' rewind stream and read text
    '    objStream.Position = 0
    '    objStream.type = adTypeBinary
    '    objStream.Read 3 ' skip first 3 bytes as this is the utf8 detection header
    '    data = objStream.Read()
    '    nL = objStream.size
    '    ' close up and return
    '    objStream.Close
    '    On Error GoTo salconver
    '    s = ""
    '    For nC = 0 To nL - 1
    '        s = s & Chr(data(nC))
    '    Next nC
    'salconver:
    '    On Error GoTo 0
    '    Else
    s = strText
    '    End If
    ConvertStringToUtf8Bytes = s

End Function

' decode base64 data into a byte array
' defined in the article - http://www.nonhostile.com/page000007.asp?src=code188
'Public Function DecodeBase64(ByVal strData As String) As Byte()
'
'    Dim objXML As MSXML2.DOMDocument
'    Dim objNode As MSXML2.IXMLDOMElement
'
'    ' help from MSXML
'    Set objXML = New MSXML2.DOMDocument
'    Set objNode = objXML.createElement("b64")
'    objNode.dataType = "bin.base64"
'    objNode.TexT = strData
'    DecodeBase64 = objNode.nodeTypedValue
'
'    ' thanks, bye
'    Set objNode = Nothing
'    Set objXML = Nothing
'
'End Function
'se sustituye por:
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







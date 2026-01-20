Attribute VB_Name = "modMain"
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
    
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeBinary
    objStream.Open
    
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeText
    strTmp = objStream.ReadText
    
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = strTmp

End Function


' accept a byte array containing utf-8 data
' and convert it to a string
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As Byte()

    Dim objStream As ADODB.Stream
    Dim data() As Byte
    
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
    
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3 ' skip first 3 bytes as this is the utf8 detection header
    data = objStream.Read()
    
    ' close up and return
    objStream.Close
    ConvertStringToUtf8Bytes = data

End Function


' decode base64 data into a byte array
' defined in the article - http://www.nonhostile.com/page000007.asp?src=code188
Public Function DecodeBase64(ByVal strData As String) As Byte()

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
   
    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
   
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

End Function








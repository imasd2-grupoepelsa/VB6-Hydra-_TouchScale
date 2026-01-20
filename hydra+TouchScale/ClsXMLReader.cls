VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXMLReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sPath As String
Private Type TXmlLabel
    CloseLabel As Boolean
    nID As Long
    nIDParent As Long
    sName As String
    nAttributes As Long
    Attributes() As String
    Values() As String
    sString As String
End Type
Private Buffer() As TXmlLabel
Private nBuffer As Long
Private Ready As Boolean
Private CurrentPos As Long
Private CurrentParent As Long
Public Property Let document(sFile As String)
    sPath = sFile
End Property
Public Property Get document() As String
    document = sPath
End Property


Public Property Get LabelID() As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    If CurrentParent = -1 Then
        LabelID = -1
    Else
        LabelID = Buffer(CurrentPos).nID
    End If
    
End Property
Public Property Get LabelData() As String
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    If CurrentParent = -1 Then
        LabelData = ""
    Else
        LabelData = Buffer(CurrentPos).sString
    End If
    
End Property
Public Property Get LabelName() As String
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    If CurrentParent = -1 Then
        LabelName = ""
    Else
        LabelName = Buffer(CurrentPos).sName
    End If
End Property
Public Sub GetAttributes(nAttrib As Long, sAttribName() As String, sAttribValue() As String)
    
    Dim Bucle As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Sub
    End If
    
    nAttrib = 0
    If CurrentParent <> -1 Then
        nAttrib = Buffer(CurrentPos).nAttributes
        ReDim sAttribName(nAttrib)
        ReDim sAttribValue(nAttrib)
        For Bucle = 0 To nAttrib - 1
            sAttribName(Bucle) = Buffer(CurrentPos).Attributes(Bucle)
            sAttribValue(Bucle) = Buffer(CurrentPos).Values(Bucle)
        Next Bucle
        
    End If
    
End Sub

Public Function GetChildren(Ids() As Long, sNames() As String) As Long
    Dim Bucle As Long
    Dim Resp As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    Resp = 0
    For Bucle = 0 To nBuffer - 1
        If Buffer(Bucle).nIDParent = Buffer(CurrentPos).nID Then
            If Buffer(Bucle).CloseLabel = False Then
                Resp = Resp + 1
                ReDim Preserve Ids(Resp)
                ReDim Preserve sNames(Resp)
                Ids(Resp - 1) = Buffer(Bucle).nID
                sNames(Resp - 1) = Buffer(Bucle).sName
            End If
        End If
    Next Bucle
    GetChildren = Resp
End Function
'****************************************************************************
' Movimiento por el fichero
'****************************************************************************
Public Sub Reset()
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Sub
    End If
    
    CurrentParent = -1
    CurrentPos = 0
End Sub
Public Function GoBack()
    Dim Bucle As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
       
    CurrentParent = Buffer(CurrentPos).nIDParent
    For Bucle = 0 To nBuffer - 1
        If Buffer(Bucle).nID = CurrentParent Then
            CurrentPos = Bucle
            Exit For
        End If
    Next Bucle
    
End Function
Public Function SearchChildID(ByVal nNum As Long) As Boolean
    Dim MyOK As Boolean
    Dim Bucle As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    MyOK = False
    For Bucle = 0 To nBuffer - 1
        If Buffer(Bucle).nIDParent = CurrentParent Then
            If nNum = Buffer(Bucle).nID Then
                MyOK = True
                CurrentParent = Buffer(Bucle).nID
                CurrentPos = Bucle
                Exit For
            End If
        End If
        
    Next Bucle
    
    SearchChildID = MyOK
End Function
Public Function SearchString(ByVal scad As String) As Boolean
    Dim MyOK As Boolean
    Dim Bucle As Long
    
    If Not Ready Then
        On Error GoTo 0
        Err.Raise 10003, , "No document Opened"
        Exit Function
    End If
    
    MyOK = False
    For Bucle = 0 To nBuffer - 1
        If Buffer(Bucle).nIDParent = CurrentParent Then
            If Trim(UCase(scad)) = Trim(UCase(Buffer(Bucle).sName)) Then
                MyOK = True
                CurrentParent = Buffer(Bucle).nID
                CurrentPos = Bucle
                Exit For
            End If
        End If
        
    Next Bucle
    SearchString = MyOK
End Function
Public Sub OpenDoc()
    Dim Arch As Integer
    Dim sBuf As String
    Dim BufTmp() As String
    Dim nBuftmp As String
    If Trim(sPath) = "" Then
        On Error GoTo 0
        Err.Raise 10000, , "No document to open"
    Else
        Arch = FreeFile()
        On Error Resume Next
        Open sPath For Input As #Arch
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise 10001, , "Error opening document"
        Else
            On Error GoTo 0
            nBuftmp = 0
            Do Until EOF(Arch)
                Line Input #Arch, sBuf
                sBuf = Trim(sBuf)
                If sBuf <> "" Then
                    nBuftmp = nBuftmp + 1
                    ReDim Preserve BufTmp(nBuftmp)
                    BufTmp(nBuftmp - 1) = sBuf
                    
                End If
            Loop
            Close #Arch
            If nBuftmp = 0 Then
                Err.Raise 10002, "No valid document"
            Else
                If Parse_Buffer(nBuftmp, BufTmp) Then
                    If Parse_Consistence() Then
                        CurrentPos = -1
                        CurrentParent = -1
                        Ready = True
                    Else
                        Err.Raise 10002, "No valid document"
                    End If
                Else
                    Err.Raise 10002, "No valid document"
                End If
            End If
        End If
    End If
    
End Sub
Private Function Parse_Consistence() As Boolean
    Dim Bucle As Long
    Dim BRet As Long
    Dim nID As Long
    Dim Encontrado As Boolean
    Dim CurParent() As Long
    Dim NCur As Long
    Dim MyOK As Boolean
    MyOK = True
    NCur = 0
    nID = -1
    For Bucle = 0 To nBuffer - 1
        If Left(Buffer(Bucle).sName, 1) = "/" Then
            If NCur = 0 Then
                '!!
                MsgBox "ERR"
                '!!
                MyOK = False
                Exit For
            Else
                Encontrado = False
                
                For BRet = Bucle - 1 To 0 Step -1
                    If NCur >= 2 Then
                        If CurParent(NCur - 2) = Buffer(BRet).nIDParent Then
                            If Mid(Buffer(Bucle).sName, 2) = Buffer(BRet).sName Then
                                Encontrado = True
                                Exit For
                            End If
                        End If
                    Else
                        If Buffer(BRet).nIDParent = -1 Then
                            If Mid(Buffer(Bucle).sName, 2) = Buffer(BRet).sName Then
                                Encontrado = True
                                Exit For
                            End If
                        End If
                    End If
                Next BRet
                NCur = NCur - 1
                If Not Encontrado Then
                    MyOK = False
                    '!!
                MsgBox "ERR"
                '!!
                    Exit For
                Else
                    Buffer(Bucle).CloseLabel = True
                    Buffer(Bucle).nID = Buffer(BRet).nID
                    Buffer(Bucle).nIDParent = Buffer(BRet).nIDParent
                End If
            End If
        Else
            nID = nID + 1
            NCur = NCur + 1
            ReDim Preserve CurParent(NCur)
            CurParent(NCur - 1) = nID
            Buffer(Bucle).nID = nID
            Buffer(Bucle).CloseLabel = False
            If NCur > 1 Then
                Buffer(Bucle).nIDParent = CurParent(NCur - 2)
            Else
                Buffer(Bucle).nIDParent = -1
            End If
        End If
    Next Bucle
    Parse_Consistence = MyOK
End Function
Private Function Parse_Buffer(ByVal nBuftmp As Long, BufTmp() As String) As Boolean
    Dim Bucle As Long
    Dim BigBuf As String
    Dim Zone As Long
    Dim MyOK As Boolean
    Dim scad As String
    Dim ZoneName As String
    Dim AttrName As String
    Dim MyXML As TXmlLabel
    Dim BLite As Long
    Dim AutoCierra As Boolean
    Zone = 0
    BigBuf = ""
    MyOK = True
    nBuffer = 0
    '
    'For Bucle = 0 To nBuftmp - 1
    '    BigBuf = BigBuf & BufTmp(Bucle) & Chr(13)
    'Next Bucle
    BigBuf = Join(BufTmp, Chr(13))
    '
    If InStr(1, BigBuf, Chr(9)) > 0 Then
        BigBuf = Replace(BigBuf, Chr(9), " ")
    End If
    '
    For Bucle = 1 To Len(BigBuf)
        scad = Mid(BigBuf, Bucle, 1)
        Select Case Zone
            Case 0
                ZoneName = ""
                Zone = 1
                If scad <> "<" Then
                    MyOK = False
                    Exit For
                End If
            Case 1
                If scad = ">" Then
                    If parse_name(ZoneName, MyXML) Then
                        If Left(ZoneName, 1) <> "?" Then
                            AutoCierra = False
                            If Right(MyXML.sName, 1) = "/" Then
                                AutoCierra = True
                                MyXML.sName = Left(MyXML.sName, Len(MyXML.sName) - 1)
                            End If
                            
                            nBuffer = nBuffer + 1
                            ReDim Preserve Buffer(nBuffer)
                            
                            Buffer(nBuffer - 1).sName = MyXML.sName
                            Buffer(nBuffer - 1).nAttributes = MyXML.nAttributes
                            If MyXML.nAttributes > 0 Then
                                ReDim Buffer(nBuffer - 1).Attributes(MyXML.nAttributes)
                                ReDim Buffer(nBuffer - 1).Values(MyXML.nAttributes)
                                For BLite = 0 To MyXML.nAttributes - 1
                                    Buffer(nBuffer - 1).Attributes(BLite) = MyXML.Attributes(BLite)
                                    Buffer(nBuffer - 1).Values(BLite) = MyXML.Values(BLite)
                                    Buffer(nBuffer - 1).sString = ""
                                Next BLite
                            End If
                            
                            If AutoCierra Then
                                nBuffer = nBuffer + 1
                                ReDim Preserve Buffer(nBuffer)
                                Buffer(nBuffer - 1).sName = "/" & MyXML.sName
                                Buffer(nBuffer - 1).nAttributes = 0
                            End If
                            
                        End If
                        Zone = 2
                    Else
                        MyOK = False
                        Exit For
                    End If
                Else
                    ZoneName = ZoneName & scad
                End If
            Case 2
                If scad = "<" Then
                    AttrName = Trim(AttrName)
                    If nBuffer > 0 Then Buffer(nBuffer - 1).sString = AttrName
                    AttrName = ""
                    Bucle = Bucle - 1
                    Zone = 0
                Else
                    If scad <> Chr(13) Then
                        AttrName = AttrName & scad
                    Else
                        If Bucle = Len(BigBuf) Then Zone = 0
                    End If
                End If
        End Select
    Next Bucle
    If Zone <> 0 Then MyOK = False
    Parse_Buffer = MyOK
End Function
Private Function parse_name(ByVal scad As String, MyXML As TXmlLabel) As Boolean

    Dim MyOK As Boolean
    Dim Bucle As Long
    Dim Buf As String
    Dim Zone As Long
    Dim AttrOpen As Boolean
    Dim tmpAttr As String
    MyXML.sName = ""
    MyXML.nAttributes = 0
    Zone = 0
    Buf = ""
    AttrOpen = False
    MyOK = True

    If Len(scad) = 0 Then
        MyOK = False
    Else
        For Bucle = 1 To Len(scad)
            
            Select Case Zone
                Case 0
                    If Mid(scad, Bucle, 1) = Chr(13) Then MyOK = False
                    If Mid(scad, Bucle, 1) = " " And Buf = "" Then MyOK = False
                    If MyOK Then
                        If Mid(scad, Bucle, 1) = " " Then
                            AttrOpen = True
                            Zone = 1
                            Bucle = Bucle - 1
                            MyXML.sName = Buf  ' nombre
                            Buf = ""
                        Else
                            Buf = Buf & Mid(scad, Bucle, 1)
                            If Bucle = Len(scad) Then MyXML.sName = Buf
                        End If
                    Else
                        Exit For
                    End If
                    
                Case 1
                    If Mid(scad, Bucle, 1) <> " " Then
                        '************************
                        ' nombre
                        Buf = ""
                        Bucle = Bucle - 1
                        Zone = 2
                    Else
                        AttrOpen = True
                        If Bucle = Len(scad) Then
                            MyOK = False
                            Exit For
                        End If
                    End If
                Case 2
                    If Mid(scad, Bucle, 1) = "=" Then
                        '******************
                        ' nombre attrib
                        tmpAttr = Buf
                        If Buf = "" Then
                            MyOK = False
                            Exit For
                        End If
                        Zone = 3
                    Else
                        If Mid(scad, Bucle, 1) <> " " Then
                            Buf = Buf & Mid(scad, Bucle, 1)
                            If Buf = "?" Then
                                AttrOpen = False
                                Zone = 1
                            End If
                        End If
                    End If
                Case 3
                    If Mid(scad, Bucle, 1) <> " " Then
                        If Mid(scad, Bucle, 1) <> Chr(13) Then
                            If Mid(scad, Bucle, 1) = Chr(34) Then
                                Zone = 4
                                Buf = ""
                            Else
                                MyOK = False
                                Exit For
                            End If
                        End If
                    End If
                Case 4
                    If Mid(scad, Bucle, 1) <> Chr(34) Then
                        Buf = Buf & Mid(scad, Bucle, 1)
                    Else
                        ' attr
                        MyXML.nAttributes = MyXML.nAttributes + 1
                        ReDim Preserve MyXML.Attributes(MyXML.nAttributes)
                        ReDim Preserve MyXML.Values(MyXML.nAttributes)
                        MyXML.Attributes(MyXML.nAttributes - 1) = tmpAttr
                        MyXML.Values(MyXML.nAttributes - 1) = Buf
                        Buf = ""
                        AttrOpen = False
                        Zone = 1
                    End If
            End Select
                
         
        Next Bucle
    End If
    If MyOK Then
        If Left(MyXML.sName, 1) = "/" And MyXML.nAttributes > 0 Then
            MyOK = False
        End If
       
        
    End If
    If AttrOpen Then MyOK = False
    parse_name = MyOK
End Function

Private Sub Class_Initialize()
    Ready = False
End Sub

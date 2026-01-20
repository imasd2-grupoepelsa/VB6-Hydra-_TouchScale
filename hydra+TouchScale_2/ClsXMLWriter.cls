VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXMLWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sPath As String
Private Arch As Integer
Private CurZone() As String
Private NCur As Long
Private Indentation As Long
Public Property Let document(sFile As String)
    sPath = sFile
End Property
Public Property Get document() As String
    document = sPath
End Property
Public Sub OpenDoc(ByVal sDocName As String)
    If sPath = "" Then
        Err.Raise 10000, , "No document to open"
    Else
        Arch = FreeFile()
        On Error Resume Next
        Open sPath For Output As #Arch
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise 10001, , "Error opening document"
        Else
            Indentation = 0
            NCur = 0
            PrintData "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & " ?>"
            PrintData "<" & sDocName & ">"
            NCur = 1
            ReDim Preserve CurZone(NCur)
            CurZone(0) = sDocName
            On Error GoTo 0
        End If
    End If
End Sub
Public Sub CloseDoc()
    Do While NCur > 0
        EndDataBlock
    Loop
    On Error Resume Next
    If Arch <> -1 Then Close #Arch
    On Error GoTo 0
End Sub
Public Sub SetSimpleBlock(sKey As String)
    Dim AttribName() As String
    Dim AttribValue() As String
    SetDataBlock sKey, 0, AttribName, AttribValue
End Sub
Public Sub SetDataBlock(sKey As String, nAttrib As Long, AttribName() As String, AttribValue() As String)
    Dim sBuf As String
    Dim Bucle As Long
    Indentation = Indentation + 1
    NCur = NCur + 1
    ReDim Preserve CurZone(NCur)
    CurZone(NCur - 1) = sKey
    sBuf = "<" & sKey
    If nAttrib > 0 Then
        For Bucle = 0 To nAttrib - 1
            sBuf = sBuf & " " & AttribName(Bucle) & "=" & Chr(34) & AttribValue(Bucle) & Chr(34)
        Next Bucle
    End If
    sBuf = sBuf & ">"
    PrintData sBuf
    Indentation = Indentation + 1
End Sub
Public Sub SetData(sKey As String, sValue As String)
    Dim sBuf As String
    sBuf = "<" & sKey & "> " & sValue & " </" & sKey & ">"
    PrintData sBuf
End Sub

Public Sub EndDataBlock()
    If NCur > 0 Then
        NCur = NCur - 1
        Indentation = Indentation - 1
        PrintData "</" & CurZone(NCur) & ">"
        Print #Arch, ""
        Indentation = Indentation - 1
    End If
End Sub
Private Sub Class_Initialize()
    sPath = ""
    Arch = -1
End Sub

Private Sub Class_Terminate()
    Me.CloseDoc
End Sub
Private Sub PrintData(sData As String)
    Dim MyLoop As Long
    Dim sBuf As String
    For MyLoop = 0 To Indentation - 1
        sBuf = sBuf & Space(4)
    Next MyLoop
    sBuf = sBuf & sData
    Print #Arch, sBuf
End Sub

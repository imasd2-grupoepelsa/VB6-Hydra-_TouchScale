VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2453
Attribute VB_Description = "This class is used to get information about and perform operations on the system clipboard"
Option Explicit

' ##MODULE_DESCRIPTION This class is used to get information about and perform _
operations on the system clipboard

' ##SEE_ALSO The clipboard is associated with the %ApiWindow:EventVB~ApiWindow%

Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetClipboardOwner Lib "user32" () As Long

Private Declare Function GetClipboardViewer Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long

' ##ENUMERATION_DESCRIPTION ClipboardFormats - The standard formats for clipboard data
Public Enum ClipboardFormats
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_BITMAP - Data is a bitmap handle
    CF_BITMAP = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_DIB - Data is a device independent bitmap
    CF_DIB = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_DIF - Data is in Software Arts' Data Interchange format
    CF_DIF = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_ENHMETAFILE - Data is in an enhanced metafile format
    CF_ENHMETAFILE = 14
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_METAFILEPICT - Metafile picture format
    CF_METAFILEPICT = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_OEMTEXT - Text in the OEM character set
    CF_OEMTEXT = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_PALETTE - Data is a palette description
    CF_PALETTE = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_PENDATA - Data for the extensions for pen computing
    CF_PENDATA = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_RIFF - Complex audio data
    CF_RIFF = 11
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_SYLK - MS Symbolic Link
    CF_SYLK = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_TEXT - Standardised ANSI text
    CF_TEXT = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_TIFF - Tagged Image file format data
    CF_TIFF = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_UNICODETEXT - Text in unicode format
    CF_UNICODETEXT = 13
    ' ##ENUMERATION_MEMBER_DESCRIPTION CF_WAVE - Simple wave form audio data
    CF_WAVE = 12
End Enum

Public ParenthWnd As Long
Attribute ParenthWnd.VB_VarHelpID = 2454

Private myMemory As ApiGlobalmemory
Private mLastFormat As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get BackedUp() As Boolean
Attribute BackedUp.VB_HelpID = 2455

    BackedUp = Not (myMemory Is Nothing)
    
End Property

'\\ --[Backup]------------------------------------------------------
'\\ Makes an in-memory copy of the clipboard's contents so that they
'\\ can be restored easily
'\\ ----------------------------------------------------------------
Public Sub Backup()
Attribute Backup.VB_HelpID = 2456

Dim lret As Long
Dim AllFormats As Collection
Dim lFormat As Long

'\\ Need to get all the formats first...
Set AllFormats = ClipboardFormats

lret = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
End If
If lret Then
    If AllFormats.Count > 0 Then
        '\\ Get the first format that holds any data
        For lFormat = 0 To AllFormats.Count - 1
            lret = GetClipboardData(lFormat)
            If lret > 0 Then
                Set myMemory = New ApiGlobalmemory
                Call myMemory.CopyFromHandle(lret)
                '\\ Keep a note of this format
                mLastFormat = lFormat
                Exit For
            End If
            'clipboard
        Next lFormat
    End If
    lret = CloseClipboard()
End If


End Sub

Public Property Get ClipboardFormats() As Collection
Attribute ClipboardFormats.VB_HelpID = 2457

Dim lret As Long
Dim colFormats As Collection

lret = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
End If
If lret > 0 Then
    Set colFormats = New Collection
    '\\ Get the first available format
    lret = EnumClipboardFormats(0)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
    End If
    While lret > 0
        colFormats.Add lret
        '\\ Get the next available format
        lret = EnumClipboardFormats(lret)
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
        End If
    Wend
    '\\ Close the clipboard object to make it available to other apps.
    lret = CloseClipboard()
End If

Set ClipboardFormats = colFormats

End Property

'\\ --[Restore]-----------------------------------------------------
'\\ Takes the in-memory copy of the clipboard object and restores it
'\\ to the clipboard.
'\\ ----------------------------------------------------------------
Public Sub Restore()
Attribute Restore.VB_HelpID = 2458

Dim lret As Long

If BackedUp Then
    lret = OpenClipboard(ParenthWnd)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiClipboard:Restore", GetLastSystemError)
    End If
    If lret Then
        myMemory.AllocationType = GMEM_FIXED
        lret = SetClipboardData(mLastFormat, myMemory.Handle)
        myMemory.Free
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
        End If
        lret = CloseClipboard()
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", GetLastSystemError)
        End If
    End If
End If

End Sub

Public Property Get Text() As String
Attribute Text.VB_HelpID = 2459

Dim sRet As String
If Clipboard.GetFormat(vbCFText) Then
    sRet = Clipboard.GetText()
End If

End Property

Private Sub Class_Terminate()

Set myMemory = Nothing

End Sub



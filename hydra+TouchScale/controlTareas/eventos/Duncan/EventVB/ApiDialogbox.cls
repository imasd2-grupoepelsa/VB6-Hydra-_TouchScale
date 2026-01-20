VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiDialogbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2396
Attribute VB_Description="This class provides properties and methods for the creation and manipulation of dialog boxes such as message boxes and the like."
Option Explicit

' ##MODULE_DESCRIPTION This class provides properties and methods for the creation and manipulation of _
dialog boxes such as message boxes and the like.

Private Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Private Type DLGTEMPLATE
    Style As Long
    dwExtendedStyle As Long
    cdit As Integer
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
End Type
Private Type DLGITEMTEMPLATE
    Style As Long
    dwExtendedStyle As Long
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    id As Integer
End Type

Public Enum enDialogStyleBits
    DS_ABSALIGN = &H1&
    DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION
    DS_NOIDLEMSG = &H100         '  WM_ENTERIDLE message will not be sent
    DS_SETFONT = &H40            '  User specified font for Dlg controls
    DS_SETFOREGROUND = &H200     '  not in win3.1
    DS_SYSMODAL = &H2&
End Enum


'\\ Private member functions
Private mCaption As String
Private mMessageText As String
Private mCountDownCaption As String
Private mCountdownSeconds As Long

'\\ Dialog main template
Private mDialogTemplate As DLGTEMPLATE

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Get Caption() As String
Attribute Caption.VB_HelpID = 2397
    Caption = mCaption
End Property

Public Property Let Caption(ByVal sNewCaption As String)

If sNewCaption <> mCaption Then
    mCaption = sNewCaption
End If

End Property

Public Property Let MessageText(ByVal sNewmessageText As String)
Attribute MessageText.VB_HelpID = 2399

If sNewmessageText <> mMessageText Then
    mMessageText = sNewmessageText
End If

End Property

Public Property Get MessageText() As String

    MessageText = mMessageText
    
End Property



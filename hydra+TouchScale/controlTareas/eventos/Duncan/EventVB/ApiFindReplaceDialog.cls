VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiFindReplaceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2657
Attribute VB_Description="This class provides a wrapper for the Find/Replace standard dialog box.  This allows you to add interactive Find/Replace functionality to any window in your project."
Option Explicit

' ##MODULE_DESCRIPTION This class provides a wrapper for the Find/Replace _
standard dialog box.  This allows you to add interactive Find/Replace functionality to _
any window in your project.

Private Enum FindReplaceDialogueFlags
    FR_DOWN = &H1
    FR_WHOLEWORD = &H2
    FR_MATCHCASE = &H4
    FR_FINDNEXT = &H8
    FR_REPLACE = &H10
    FR_REPLACEALL = &H20
    FR_DIALOGTERM = &H40
    FR_SHOWHELP = &H80
    FR_ENABLEHOOK = &H100
    FR_ENABLETEMPLATE = &H200
    FR_NOUPDOWN = &H400
    FR_NOMATCHCASE = &H800
    FR_NOWHOLEWORD = &H1000
    FR_ENABLETEMPLATEHANDLE = &H2000
    FR_HIDEUPDOWN = &H4000
    FR_HIDEMATCHCASE = &H8000
    FR_HIDEWHOLEWORD = &H10000
End Enum

Private Type FINDREPLACE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    flags As Long 'Combination of FR_ flags
    lpFindWhat As String 'Pointer to a String
    lpReplaceWhat As String 'Pointer to a String
    nFindStrLen As Integer
    nReplaceStrLen As Integer
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long 'Pointer to string
End Type

Private Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (lpFindReplace As FINDREPLACE) As Long
Private Declare Function ReplaceText Lib "comdlg32.dll" Alias "FindTextA" (lpFindReplace As FINDREPLACE) As Long

Public Enum FindReplaceDialogMode
   FRDM_FIND = &H0
   FRDM_REPLACE = &H1
End Enum

Private mMode As FindReplaceDialogMode
Private mFindReplace As FINDREPLACE

'\\ String used for messages from the Find/Replace dialog using RegisterWindowsmessage:
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private mDialogHwnd As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get DialogWindowHandle() As Long
Attribute DialogWindowHandle.VB_HelpID = 2658

    DialogWindowHandle = mDialogHwnd
    
End Property

Public Property Get FindReplaceWindowMessage() As Long
Attribute FindReplaceWindowMessage.VB_HelpID = 2659

    FindReplaceWindowMessage = RegisterWindowMessage(FINDMSGSTRING)

End Property


Public Property Let FindString(ByVal strToFind As String)
Attribute FindString.VB_HelpID = 2661

    mFindReplace.lpFindWhat = strToFind
        
End Property

Public Property Get FindString() As String

    FindString = mFindReplace.lpFindWhat
    
End Property

Public Property Get Mode() As FindReplaceDialogMode
Attribute Mode.VB_HelpID = 2662

Mode = mMode

End Property

Public Property Let Mode(ByVal newMode As FindReplaceDialogMode)

mMode = newMode

End Property

Public Property Set ParentWindow(ByVal theWindow As ApiWindow)
Attribute ParentWindow.VB_HelpID = 2663

    mFindReplace.hwndOwner = theWindow.hwnd
    
End Property

Public Property Get ReplaceString() As String
Attribute ReplaceString.VB_HelpID = 2664

    ReplaceString = mFindReplace.lpReplaceWhat
    
End Property

Public Property Let ReplaceString(ByVal strToReplace As String)

    mFindReplace.lpReplaceWhat = strToReplace
    
End Property

'\\ --[Show]---------------------------------------------------
'\\ Causes the dialog box to be shown
'\\ -----------------------------------------------------------
Public Sub Show()
Attribute Show.VB_HelpID = 2665

If mFindReplace.hwndOwner = 0 Then
    ReportError -1, "ApiFindReplaceDialog:Show", "Parent form must be set before the Find/Replace dialog can be shown"
    Exit Sub
End If

Select Case mMode
Case FRDM_FIND
   'Call the FindText API
   mDialogHwnd = FindText(mFindReplace)
Case FRDM_REPLACE
   'Call the ReplaceText API
   mDialogHwnd = ReplaceText(mFindReplace)
End Select
If Err.LastDllError <> 0 Then
   Call ReportError(Err.LastDllError, "ApiFindReplaceDialog:Show", GetLastSystemError)
End If

End Sub



Public Property Let ShowHelp(ByVal bShow As Boolean)
Attribute ShowHelp.VB_HelpID = 2667

If bShow Then
    mFindReplace.flags = mFindReplace.flags Or FR_SHOWHELP
Else
    mFindReplace.flags = mFindReplace.flags And (Not FR_SHOWHELP)
End If

End Property

Public Property Get ShowHelp() As Boolean

ShowHelp = (mFindReplace.flags = (mFindReplace.flags And FR_SHOWHELP))

End Property


Public Property Let ShowMatchCase(ByVal bShow As Boolean)
Attribute ShowMatchCase.VB_HelpID = 2669

If Not bShow Then
    mFindReplace.flags = mFindReplace.flags Or FR_HIDEMATCHCASE
Else
    mFindReplace.flags = mFindReplace.flags And (Not FR_HIDEMATCHCASE)
End If

End Property

Public Property Get ShowMatchCase() As Boolean

ShowMatchCase = Not (mFindReplace.flags = (mFindReplace.flags And FR_HIDEMATCHCASE))

End Property


Public Property Let ShowMatchWholeWord(ByVal bShow As Boolean)
Attribute ShowMatchWholeWord.VB_HelpID = 2671

If Not bShow Then
    mFindReplace.flags = mFindReplace.flags Or FR_HIDEWHOLEWORD
Else
    mFindReplace.flags = mFindReplace.flags And (Not FR_HIDEWHOLEWORD)
End If

End Property

Public Property Get ShowMatchWholeWord() As Boolean

ShowMatchWholeWord = Not (mFindReplace.flags = (mFindReplace.flags And FR_HIDEWHOLEWORD))

End Property


Private Sub Class_Initialize()

'\\ Initialise the strings....
With mFindReplace
    .lpFindWhat = String$(80, 0)
    .lpReplaceWhat = String$(80, 0)
    .nFindStrLen = 80
    .nReplaceStrLen = 80
    .lStructSize = Len(mFindReplace)
End With

End Sub

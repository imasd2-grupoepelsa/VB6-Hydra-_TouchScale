VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiIconDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2434
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long


' ##BLOCK_DESCRIPTION IconsFromFile - Returns the collection of icons in a file
' ##BLOCK_REMARKS The file can be an executable, dll or icon file
Public Property Get IconsFromFile(ByVal Filename As String) As Collection
Attribute IconsFromFile.VB_HelpID = 2435
Attribute IconsFromFile.VB_Description = "IconsFromFile - Returns the collection of icons in a file"

Dim lIndex As Long
Dim lIconCount As Long
Dim lret As Long

Dim colIcons As Collection
Dim thisIcon As ApiIcon

'\\ Initialise the collection
Set colIcons = New Collection

'\\ Get the number of items
lIconCount = ExtractIcon(App.hInstance, Filename, -1)
If lIconCount > 0 Then
    For lIndex = 0 To lIconCount - 1
        lret = ExtractIcon(App.hInstance, Filename, lIndex)
        If lret > 0 Then
            Set thisIcon = New ApiIcon
            thisIcon.hIcon = lret
            colIcons.Add thisIcon
        End If
    Next lIndex
End If

'\\ Return the collection
Set IconsFromFile = colIcons

End Property

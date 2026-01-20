VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2190
   ClientTop       =   1545
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   Begin RichTextLib.RichTextBox hpl 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   -2147483634
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"form1.frx":0077
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) _
   As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpData As Long) As Long

Private Const m_sDescription = _
                  "Version:1.0" & vbCrLf & _
                  "StartHTML:aaaaaaaaaa" & vbCrLf & _
                  "EndHTML:bbbbbbbbbb" & vbCrLf & _
                  "StartFragment:cccccccccc" & vbCrLf & _
                  "EndFragment:dddddddddd" & vbCrLf
                  
Private m_cfHTMLClipFormat As Long

Function RegisterCF() As Long


   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format") 'RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
   
End Function

Public Sub PutHTMLClipboard(sHtmlFragment As String, _
   Optional sContextStart As String = "<HTML><BODY>", _
   Optional sContextEnd As String = "</BODY></HTML>")
   
   Dim sData As String
   
   If RegisterCF = 0 Then Exit Sub
   
   'Add the starting and ending tags for the HTML fragment
   sContextStart = sContextStart & "<!--StartFragment -->"
   sContextEnd = "<!--EndFragment -->" & sContextEnd
   
   'Build the HTML given the description, the fragment and the context.
   'And, replace the offset place holders in the description with values
   'for the offsets of StartHMTL, EndHTML, StartFragment and EndFragment.
   sData = m_sDescription & sContextStart & sHtmlFragment & sContextEnd
   sData = Replace(sData, "aaaaaaaaaa", _
                   Format(Len(m_sDescription), "0000000000"))
   sData = Replace(sData, "bbbbbbbbbb", Format(Len(sData), "0000000000"))
   sData = Replace(sData, "cccccccccc", Format(Len(m_sDescription & _
                   sContextStart), "0000000000"))
   sData = Replace(sData, "dddddddddd", Format(Len(m_sDescription & _
                   sContextStart & sHtmlFragment), "0000000000"))

   'Add the HTML code to the clipboard
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      
      hMemHandle = GlobalAlloc(0, Len(sData) + 10)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            
            CopyMemory ByVal lpData, ByVal sData, Len(sData)
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData m_cfHTMLClipFormat, hMemHandle
                        
         End If
      
      End If
   
      Call CloseClipboard
   End If

End Sub

Public Function GetHTMLClipboard() As String

   Dim sData As String
   
   If RegisterCF = 0 Then Exit Function
   
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      Dim nClipSize As Long
      
      GlobalUnlock hMemHandle

      'Retrieve the data from the clipboard
      hMemHandle = GetClipboardData(m_cfHTMLClipFormat)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            nClipSize = lstrlen(lpData)
            sData = String(nClipSize + 10, 0)
            

            Call CopyMemory(ByVal sData, ByVal lpData, nClipSize)
            
            Dim nStartFrag As Long, nEndFrag As Long
            Dim nIndx As Long
            
            'If StartFragment appears in the data's description,
            'then retrieve the offset specified in the description
            'for the start of the fragment. Likewise, if EndFragment
            'appears in the description, then retrieve the
            'corresponding offset.
            nIndx = InStr(sData, "StartFragment:")
            If nIndx Then
               nStartFrag = CLng(Mid(sData, _
                                 nIndx + Len("StartFragment:"), 10))

            End If
            nIndx = InStr(sData, "EndFragment:")
            If nIndx Then
               nEndFrag = CLng(Mid(sData, nIndx + Len("EndFragment:"), 10))
            End If
            
            'Return the fragment given the starting and ending
            'offsets
            If (nStartFrag > 0 And nEndFrag > 0) Then
               GetHTMLClipboard = Mid(sData, nStartFrag + 1, _
                                 (nEndFrag - nStartFrag))
            End If
                        
         End If
      
      End If

   
      Call CloseClipboard
   End If


End Function

'Private Sub Command1_Click()
'Dim nf As Integer
'Dim stexto() As String
'Dim sResul() As String
'Dim sresul1() As String
'Dim sS As String
'Dim nC, nC1 As Integer
'Dim nL, nL1 As Integer
'Dim nPos As Integer
'
'    stexto = Split(GetHTMLClipboard(), "</p>")
'
'    hpl.Caption = ""
'    nC = UBound(stexto)
'    If nC > 0 Then
'        nf = FreeFile()
'        Open App.Path & "\texto.txt" For Output As #nf
'        For nL = 0 To nC - 1
'            If stexto(nL) <> "" Then
'                sS = Replace(stexto(nL), vbCrLf, "") & "</p>"
'                Print #nf, Format(nL, "00") & " - " & sS
'                sResul = Split(sS, ">")
'                nC1 = UBound(sResul)
'                For nL1 = 0 To nC1 - 1
'                    sResul(nL1) = Replace(sResul(nL1), "</p", "")
'                    sResul(nL1) = Replace(sResul(nL1), "<o:p", "")
'                    sResul(nL1) = Replace(sResul(nL1), "</o:p", "")
'                    If sResul(nL1) <> "" Then
'                        If Mid(sResul(nL1), 1, 1) <> "<" Then
'                            If Mid(sResul(nL1 - 1), 1, 2) <> "<p" Then
'                                hpl.Caption = hpl.Caption & "~" & sResul(nL1) & "~" & "^"
'                            Else
'                                hpl.Caption = hpl.Caption & sResul(nL1) & "^"
'                            End If
'                        End If
'                    End If
'                Next nL1
'            End If
'        Next nL
'        Close #nf
'    End If
'End Sub


'-----------------------------------------------
'http://www.vb-helper.com/howto_rtf_to_html.html
'Thanks to Brady Hegberg.
'-----------------------------------------------
Private Sub Command1_Click()
Dim old_start As Integer
Dim old_length As Integer

    old_start = RTB.SelStart
    old_length = RTB.SelLength
    RTB.SelStart = 0
    RTB.SelLength = Len(RTB.Text)

    hpl.Caption = RTF2HTML(RTB.SelRTF)
    
    RTB.SelStart = old_start
    RTB.SelLength = old_length

    ' Save the result into an HTML file.
    Dim file_name As String
    Dim fnum As Integer

    file_name = App.Path
    If Left$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "texto.txt"
    fnum = FreeFile
    Open file_name For Output As fnum
    Print #fnum, hpl.Caption
    Close #fnum

End Sub

Function RTF2HTML(strRTF As String) As String
    'Version 2.7
    
    'The current version of this function is available at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip
    
    'More information can be found at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html
    
    'Converts Rich Text encoded text to HTML format
    'if you find some text that this function doesn't
    'convert properly please email the text to
    'bradyh@bitstream.net
    Dim strHTML As String
    Dim l As Long
    Dim lTmp As Long
    Dim lTmp2 As Long
    Dim lTmp3 As Long
    Dim lRTFLen As Long
    Dim lBOS As Long                 'beginning of section
    Dim lEOS As Long                 'end of section
    Dim strTmp As String
    Dim strTmp2 As String
    Dim strEOS As String             'string to be added to end of section
    Dim strBOS As String             'string to be added to beginning of section
    Dim strEOP As String             'string to be added to end of paragraph
    Dim strBOL As String             'string to be added to the begining of each new line
    Dim strEOL As String             'string to be added to the end of each new line
    Dim strEOCL As String            'string to be added to the end of current line
    Dim strEOLL As String            'string to be added to the end of previous line
    Dim strCurFont As String         'current font code eg: "f3"
    Dim strCurFontSize As String     'current font size eg: "fs20"
    Dim strCurColor As String        'current font color eg: "cf2"
    Dim strFontFace As String        'Font face for current font
    Dim strFontColor As String       'Font color for current font
    Dim lFontSize As Integer         'Font size for current font
    Const gHellFrozenOver = False    'always false
    Dim gSkip As Boolean             'skip to next word/command
    Dim strCodes As String           'codes for ascii to HTML char conversion
    Dim strCurLine As String         'temp storage for text for current line before being added to strHTML
    Dim strColorTable() As String    'table of colors
    Dim lColors As Long              '# of colors
    Dim strFontTable() As String     'table of fonts
    Dim lFonts As Long               '# of fonts
    Dim strFontCodes As String       'list of font code modifiers
    Dim gSeekingText As Boolean      'True if we have to hit text before inserting a </FONT>
    Dim gText As Boolean             'true if there is text (as opposed to a control code) in strTmp
    Dim strAlign As String           '"center" or "right"
    Dim gAlign As Boolean            'if current text is aligned

    'setup HTML codes
    strCodes = "&nbsp;  {00}&copy;  {a9}&acute; {b4}&laquo; {ab}&raquo; {bb}&iexcl; {a1}&iquest;{bf}&Agrave;{c0}&agrave;{e0}&Aacute;{c1}"
    strCodes = strCodes & "&aacute;{e1}&Acirc; {c2}&acirc; {e2}&Atilde;{c3}&atilde;{e3}&Auml;  {c4}&auml;  {e4}&Aring; {c5}&aring; {e5}&AElig; {c6}"
    strCodes = strCodes & "&aelig; {e6}&Ccedil;{c7}&ccedil;{e7}&ETH;   {d0}&eth;   {f0}&Egrave;{c8}&egrave;{e8}&Eacute;{c9}&eacute;{e9}&Ecirc; {ca}"
    strCodes = strCodes & "&ecirc; {ea}&Euml;  {cb}&euml;  {eb}&Igrave;{cc}&igrave;{ec}&Iacute;{cd}&iacute;{ed}&Icirc; {ce}&icirc; {ee}&Iuml;  {cf}"
    strCodes = strCodes & "&iuml;  {ef}&Ntilde;{d1}&ntilde;{f1}&Ograve;{d2}&ograve;{f2}&Oacute;{d3}&oacute;{f3}&Ocirc; {d4}&ocirc; {f4}&Otilde;{d5}"
    strCodes = strCodes & "&otilde;{f5}&Ouml;  {d6}&ouml;  {f6}&Oslash;{d8}&oslash;{f8}&Ugrave;{d9}&ugrave;{f9}&Uacute;{da}&uacute;{fa}&Ucirc; {db}"
    strCodes = strCodes & "&ucirc; {fb}&Uuml;  {dc}&uuml;  {fc}&Yacute;{dd}&yacute;{fd}&yuml;  {ff}&THORN; {de}&thorn; {fe}&szlig; {df}&sect;  {a7}"
    strCodes = strCodes & "&para;  {b6}&micro; {b5}&brvbar;{a6}&plusmn;{b1}&middot;{b7}&uml;   {a8}&cedil; {b8}&ordf;  {aa}&ordm;  {ba}&not;   {ac}"
    strCodes = strCodes & "&shy;   {ad}&macr;  {af}&deg;   {b0}&sup1;  {b9}&sup2;  {b2}&sup3;  {b3}&frac14;{bc}&frac12;{bd}&frac34;{be}&times; {d7}"
    strCodes = strCodes & "&divide;{f7}&cent;  {a2}&pound; {a3}&curren;{a4}&yen;   {a5}...     {85}"
    
    'setup color table
    lColors = 0
    ReDim strColorTable(0)
    lBOS = InStr(strRTF, "\colortbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\red")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strColorTable(lColors)
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 4, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 5, 1)), Mid(strRTF, lBOS + 5, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\green")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 6, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 8, 1)), Mid(strRTF, lBOS + 8, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\blue")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 5, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\red")
                lColors = lColors + 1
            Wend
        End If
    End If
    
    'setup font table
    lFonts = 0
    ReDim strFontTable(0)
    lBOS = InStr(strRTF, "\fonttbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\f0")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strFontTable(lFonts)
                While ((Mid(strRTF, lBOS, 1) <> " ") And (lBOS <= lEOS))
                    lBOS = lBOS + 1
                Wend
                lBOS = lBOS + 1
                strTmp = Mid(strRTF, lBOS, InStr(lBOS, strRTF, ";") - lBOS)
                strFontTable(lFonts) = strFontTable(lFonts) & strTmp
                lBOS = InStr(lBOS, strRTF, "\f" & (lFonts + 1))
                lFonts = lFonts + 1
            Wend
        End If
    End If
    
    strHTML = ""
    lRTFLen = Len(strRTF)
    'seek first line with text on it
    
    lBOS = InStr(strRTF, "\deflang")
    If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 4
    
    lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally

    While Not gHellFrozenOver
        strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
        l = lBOS
        While l <= lEOS
            strTmp = Mid(strRTF, l, 1)
            Select Case strTmp
            Case "{"
                l = l + 1
            Case "}"
                'strCurLine = strCurLine & strEOS
                strCurLine = strEOS
                strEOS = ""
                l = l + 1
            Case "\"    'special code
                l = l + 1
                strTmp = Mid(strRTF, l, 1)
                Select Case strTmp
                Case "b"
                    If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                        'b = bold
                        strCurLine = strCurLine & "<B>"
                        strEOS = "</B>" & strEOS
                        If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                    ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
                        strTmp = "•"     'bullet
                        l = l + 6
                        gText = True
                    Else
                        gSkip = True
                    End If
                Case "c"
                    If ((Mid(strRTF, l, 2) = "cf") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                        'cf = color font
                        lTmp = Val(Mid(strRTF, l + 2, 5))
                        If lTmp <= UBound(strColorTable) Then
                            strCurColor = "cf" & lTmp
                            strFontColor = "#" & strColorTable(lTmp)
                            gSeekingText = True
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                            lTmp = lTmp + 1
                        Wend
                        If (Mid(strRTF, lTmp, 1) = " ") Then
                            l = lTmp
                        Else
                            l = lTmp - 1
                        End If
                    Else
                        gSkip = True
                    End If
                Case "e"
                    If (Mid(strRTF, l, 7) = "emdash ") Then
                        strTmp = "—"
                        l = l + 6
                        gText = True
                    Else
                        gSkip = True
                    End If
                Case "f"
                    If IsNumeric(Mid(strRTF, l + 1, 1)) Then
                        'f# = font
                        'first get font number
                        lTmp = l + 2
                        strTmp2 = Mid(strRTF, l + 1, 1)
                        While IsNumeric(Mid(strRTF, lTmp, 1))
                            strTmp2 = strTmp2 & Mid(strRTF, lTmp2, 1)
                            lTmp = lTmp + 1
                        Wend
                        lTmp = Val(strTmp2)
                        strCurFont = "f" & lTmp
                        If ((lTmp <= UBound(strFontTable)) And (strFontTable(lTmp) <> strFontTable(0))) Then
                            'insert codes if lTmp is a valid font # AND the font is not the default font
                            strFontFace = strFontTable(lTmp)
                            gSeekingText = True
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                            lTmp = lTmp + 1
                        Wend
                        If (Mid(strRTF, lTmp, 1) = " ") Then
                            l = lTmp
                        Else
                            l = lTmp - 1
                        End If
                    ElseIf ((Mid(strRTF, l + 1, 1) = "s") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                        'fs# = font size
                        'first get font size
                        lTmp = l + 3
                        strTmp2 = Mid(strRTF, l + 2, 1)
                        While IsNumeric(Mid(strRTF, lTmp, 1))
                            strTmp2 = strTmp2 & Mid(strRTF, lTmp, 1)
                            lTmp = lTmp + 1
                        Wend
                        lTmp = Val(strTmp2)
                        strCurFontSize = "fs" & lTmp
                        lFontSize = Int((lTmp / 5) - 2)
                        If lFontSize = 2 Then
                            strCurFontSize = ""
                            lFontSize = 0
                        Else
                            gSeekingText = True
                            If lFontSize > 8 Then lFontSize = 8
                            If lFontSize < 1 Then lFontSize = 1
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                            lTmp = lTmp + 1
                        Wend
                        If (Mid(strRTF, lTmp, 1) = " ") Then
                            l = lTmp
                        Else
                            l = lTmp - 1
                        End If
                    Else
                        gSkip = True
                    End If
                Case "i"
                    If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                        strCurLine = strCurLine & "<I>"
                        strEOS = "</I>" & strEOS
                        If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                    Else
                        gSkip = True
                    End If
                Case "l"
                    If (Mid(strRTF, l, 10) = "ldblquote ") Then
                        strTmp = "“"
                        l = l + 9
                        gText = True
                    ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
                        strTmp = "‘"
                        l = l + 6
                        gText = True
                    Else
                        gSkip = True
                    End If
                Case "p"
                    If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
                        If (Len(strFontColor & strFontFace) > 0) Then
                            If Not gSeekingText Then strCurLine = strCurLine & "</FONT>"
                            strFontColor = ""
                            strFontFace = ""
                        End If
                        If gAlign Then
                            strCurLine = strCurLine & "</TD></TR></TABLE><BR>"
                            gAlign = False
                        End If
                        strCurLine = strCurLine & strEOS
                        strEOS = ""
                        If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5    'catch next \ but skip a space
                    ElseIf (Mid(strRTF, l, 9) = "pnlvlblt\") Then
                        'bulleted list
                        strEOS = ""
                        strBOS = "<UL>"
                        strBOL = "<LI>"
                        strEOL = "</LI>"
                        strEOP = "</UL>"
                        l = l + 7    'catch next \
                    ElseIf (Mid(strRTF, l, 7) = "pntext\") Then
                        l = InStr(l, strRTF, "}")   'skip to end of braces
                    ElseIf (Mid(strRTF, l, 6) = "pntxtb") Then
                        l = InStr(l, strRTF, "}")   'skip to end of braces
                    ElseIf (Mid(strRTF, l, 10) = "pard\plain") Then
                        strCurLine = strCurLine & strEOS & strEOP
                        strEOS = ""
                        strEOP = ""
                        strBOL = ""
                        strEOL = "<BR>"
                        l = l + 3    'catch next \
                    Else
                        gSkip = True
                    End If
                Case "q"
                    If ((Mid(strRTF, l, 3) = "qc\") Or (Mid(strRTF, l, 3) = "qc ")) Then
                        'qc = centered
                        strAlign = "center"
                        'move "cursor" position to next rtf code
                        If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                        l = l + 1
                    ElseIf ((Mid(strRTF, l, 3) = "qr\") Or (Mid(strRTF, l, 3) = "qr ")) Then
                        'qr = right justified
                        strAlign = "right"
                        'move "cursor" position to next rtf code
                        If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                        l = l + 1
                    Else
                        gSkip = True
                    End If
                Case "r"
                    If (Mid(strRTF, l, 7) = "rquote ") Then
                        strTmp = "’"
                        l = l + 6
                        gText = True
                    ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
                        strTmp = "”"
                        l = l + 9
                        gText = True
                    Else
                        gSkip = True
                    End If
                Case "t"
                    If (Mid(strRTF, l, 4) = "tab ") Then
                        strTmp = Chr$(9)   'tab
                        l = l + 3
                        gText = True
                    Else
                        gSkip = True
                    End If
                Case "u"
                    If ((Mid(strRTF, l, 3) = "ul ") Or (Mid(strRTF, l, 3) = "ul\")) Then
                        strCurLine = strCurLine & "<U>"
                        strEOS = "</U>" & strEOS
                        If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1 Else l = l + 2
                    Else
                        gSkip = True
                    End If
                Case "'"
                    strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
                    lTmp = InStr(strCodes, strTmp2)
                    If lTmp = 0 Then
                        strTmp = Chr("&H" & Mid(strTmp2, 2, 2))
                    Else
                        strTmp = Trim(Mid(strCodes, lTmp - 8, 8))
                    End If
                    l = l + 1
                    gText = True
                Case "~"
                    strTmp = " "
                    gText = True
                Case "{", "}", "\"
                    gText = True
                Case vbLf, vbCr, vbCrLf    'always use vbCrLf
                    strCurLine = strCurLine & vbCrLf
                Case Else
                    gSkip = True
                End Select
                If gSkip = True Then
                    'skip everything up until the next space or "\" or "}"
                    While InStr(" \}", Mid(strRTF, l, 1)) = 0
                        l = l + 1
                    Wend
                    gSkip = False
                    If (Mid(strRTF, l, 1) = "\") Then l = l - 1
                End If
                l = l + 1
            Case vbLf, vbCr, vbCrLf
                l = l + 1
            Case Else
                gText = True
            End Select
            If gText Then
                If ((Len(strFontColor & strFontFace) > 0) And gSeekingText) Then
                    If Len(strAlign) > 0 Then
                        gAlign = True
                        If strAlign = "center" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""center""><TD>"
                        ElseIf strAlign = "right" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""right""><TD>"
                        End If
                        strAlign = ""
                    End If
                    If Len(strFontFace) > 0 Then
                        strFontCodes = strFontCodes & " FACE=""" & strFontFace & """"
                    End If
                    If Len(strFontColor) > 0 Then
                        strFontCodes = strFontCodes & " COLOR=""" & strFontColor & """"
                    End If
                    If Len(strCurFontSize) > 0 Then
                        strFontCodes = strFontCodes & " SIZE = """ & lFontSize & """"
                    End If
                    strCurLine = strCurLine & "<FONT" & strFontCodes & ">"
                    strFontCodes = ""
                End If
                strCurLine = strCurLine & strTmp
                l = l + 1
                gSeekingText = False
                gText = False
            End If
        Wend

        lBOS = lEOS + 2
        lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
        strHTML = strHTML & strEOLL & strBOS & strBOL & strCurLine
        strEOLL = strEOL
        If Len(strEOL) = 0 Then strEOL = "<BR>"
        
        If lEOS = 0 Then GoTo finally
        strBOS = ""
        strCurLine = ""
    Wend

finally:
    'strHTML = strHTML & strEOS
    'clear up any hanging fonts
    If (Len(strFontColor & strFontFace) > 0) Then strHTML = strHTML & "</FONT>"
    
    RTF2HTML = strHTML
End Function



Private Sub Command2_Click()
Dim old_start As Integer
Dim old_length As Integer
Dim nC As Integer
Dim sT As String
Dim lSel As Boolean
    'old_start = RTB.SelStart
    'old_length = RTB.SelLength
    'RTB.SelStart = 0
    'RTB.SelLength = Len(RTB.Text)

    'hpl.Caption = RTF2HTML(RTB.SelRTF)
    
    sT = ""
    lSel = False
    For nC = 1 To Len(RTB.Text)
        RTB.SelStart = nC
        RTB.SelLength = 1
        If (RTB.SelBold Or RTB.SelItalic Or RTB.SelUnderline) And lSel = False Then
            sT = sT & "~" & Mid(RTB.Text, nC, 1)
            lSel = True
        Else
            If (RTB.SelBold Or RTB.SelItalic Or RTB.SelUnderline) And lSel = True Then
                sT = sT & Mid(RTB.Text, nC, 1)
            Else
                If lSel = False Then
                    sT = sT & Mid(RTB.Text, nC, 1)
                Else
                    sT = sT & Mid(RTB.Text, nC, 1) & "~"
                    lSel = False
                End If
            
            End If
            
        End If
    Next nC
    
    sT = sT
    sT = Replace(sT, "~~", "")
    sT = Replace(sT, vbCrLf, "^")
    
    
    hpl.Text = sT
    
    'RTB.SelStart = old_start
    'RTB.SelLength = old_length

End Sub

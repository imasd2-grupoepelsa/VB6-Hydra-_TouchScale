Attribute VB_Name = "modLocationLocationLocation"
' modLocationLocationLocation
' 1998/09/01 Copyright 1998-2004, Larry Rebich
' 2000/10/01 Used in BrandingModel
' 2000/10/01 Used in Branding
' 2001/01/31 Used in PMModel Version 4
' 2001/01/31 Used in PMTips Version 4
' 2001/04/02 Add support for Version 4.
' 2003/12/03 Use in Infinos Version 5
' 2004/10/27 Used in PMTest
' 2005/10/19 Update in Tip-Of-The-Month

    Option Explicit
    DefLng A-Z

    Const mcsCompany        As String = "The Bridge"               '<----- Change for Your Company
    Const mcsProductName    As String = "LocationLocationLocation" '<------Change to Your Product's Name

    Public gsCompany        As String   '<--------Store your company name here
    Public gsProductName    As String   '<--------Store you product name here

    ' Used with Saving a location
    Const mcsComma          As String = ","
    Const mcsState          As String = "State"            'Normal, Minimized, Maximized
    Const mcsLocation       As String = "Location"         'Left, Top, Width, Height
    Const mcsScreenWidth    As String = "Screen Width"
    Const mcsScreenHeight   As String = "Screen Height"
    Public Const gcshWnd    As String = "hWnd"
    
    Private Enum eDock      '2000/08/06 Added for docking
        eLeft = 1
        eRight = 2
        eTop = 3
        eBottom = 4
    End Enum
    Const mcieDockLeft = 1
    Const mcieDockRight = 2
    Const mcieDockTop = 3
    Const mcieDockBottom = 4
    
    Public Enum eCorner     '2002/08/11 Added for corner
        eTopLeft = 1
        eTopRight = 2
        eBottomLeft = 3
        eBottomRight = 4
    End Enum
    
' 1994/11/12 Center the Message within the Zone defined by these boundaries
    Public TB_ZoneLeft          As Long
    Public TB_ZoneTop           As Long
    Public TB_ZoneWidth         As Long
    Public TB_ZoneHeight        As Long
    
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Const SPI_GETWORKAREA = 48
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

' Used to set Tips Always on Top-------------------------------------
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_SHOWWINDOW = &H40
'--------------------------------------------------------------------
    
    ' Window Arrange
    Public Const gciWinCenter          As Integer = 1
    Public Const gciWinTileHorizontal  As Integer = gciWinCenter + 1
    Public Const gciWinTileVertical    As Integer = gciWinTileHorizontal + 1
    Public Const gciWinCascade         As Integer = gciWinTileHorizontal + 2
    Public Const gciWinArrangeIcons    As Integer = gciWinTileHorizontal + 3
    Const mcsWinCenter          As String = "Ce&nter"
    Const mcsWinTileHorizontal  As String = "Tile &Horizontally"
    Const mcsWinTileVertical    As String = "Tile &Vertically"
    Const mcsWinCascade         As String = "&Cascade"
    Const mcsWinArrangeIcons    As String = "&Arrange Icons"
    
    ' Window Size Menu Indexes
    Const mciSepSize     As Integer = gciWinArrangeIcons + 1
    Const mci640x480     As Integer = mciSepSize + 1
    Public Const gci640x480 = mci640x480
    Const mci800x600     As Integer = mciSepSize + 2
    Public Const gci800x600 = mci800x600
    Const mci1024x768    As Integer = mciSepSize + 3
    Public Const gci1024x768 = mci1024x768
    Const mci1152x864    As Integer = mciSepSize + 4
    Public Const gci1152x864 = mci1152x864
    Const mci1280x768    As Integer = mciSepSize + 5        '2003/01/17 Sony SDM-V72W
    Public Const gci1280x768 = mci1280x768
    Const mci1280x1024   As Integer = mciSepSize + 6
    Public Const gci1280x1024 = mci1280x1024
    Const mci1400x1050   As Integer = mciSepSize + 7
    Public Const gci1400x1050 = mci1400x1050
    Const mci1600x1200   As Integer = mciSepSize + 8
    Public Const gci1600x1200 = mci1600x1200
    Const mciFullScreen  As Integer = mciSepSize + 9
    Public Const gciFullScreen = mciFullScreen
    
    Const mcs640x480     As String = "&640x480 [VGA]"       '1
    Const mcs800x600     As String = "&800x600 [SVGA]"      '2
    Const mcs1024x768    As String = "1&024x768 [XVGA]"     '3
    Const mcs1152x864    As String = "1&152x864"            '4
    Const mcs1280x768    As String = "1280x768 [&Wide]"     '5 2003/01/17
    Const mcs1280x1024   As String = "1&280x1024"           '6
    Const mcs1400x1050   As String = "1&400x1050"           '7
    Const mcs1600x1200   As String = "1&600x1200"           '8
    Const mcsFullScreen  As String = "&Full Screen"         '9
    
    ' Dock
    Const mciSepDock     As Integer = mciFullScreen + 1
    Const mciDockLeft    As Integer = mciSepDock + 1
    Const mciDockRight   As Integer = mciSepDock + 2
    Const mciDockTop     As Integer = mciSepDock + 3
    Const mciDockBottom  As Integer = mciSepDock + 4
    Const mcsDockLeft    As String = "Fill &Left"
    Const mcsDockRight   As String = "Fill &Right"
    Const mcsDockTop     As String = "Fill &Top"
    Const mcsDockBottom  As String = "Fill &Bottom"
    
    ' 2002/08/11 Corner
    Const mciSepCorner          As Integer = mciDockBottom + 1
    Const mciCornerUpLeft       As Integer = mciSepCorner + 1
    Const mciCornerUpRight      As Integer = mciSepCorner + 2
    Const mciCornerLowerLeft    As Integer = mciSepCorner + 3
    Const mciCornerLowerRight   As Integer = mciSepCorner + 4
    Const mcsCornerUpLeft       As String = "Top Left"
    Const mcsCornerUpRight      As String = "Top Right"
    Const mcsCornerLowerLeft    As String = "Bottom Left"
    Const mcsCornerLowerRight   As String = "Bottom Right"
    
    ' Window Sizes
    Dim mlVGAWidth          As Long
    Dim mlVGAHeight         As Long
    Public glVGAWidth       As Long '2000/09/22 Added
    Public glVGAHeight      As Long '2000/09/22 Added
    Public ml800x600Width   As Long '2000/12/30 Made public
    Public ml800x600Height  As Long '2001/01/01 Made public
    Public ml1024x768Width  As Long '2000/12/20 Made public
    Public ml1024x768Height As Long '2000/12/20 Made public
    Public ml1152x864Width      As Long     '2000/12/30 Made public
    Public ml1152x864Height     As Long     '2001/05/09 Made Public
    Public gl1280x768Width      As Long     '2003/01/17 Added to support Sony SDM-V72M
    Public gl1280x768Height     As Long     '2003/01/17 Added to support Sony SDM-V72M
    Public ml1280x1024Width     As Long     '2001/05/09 Made Public
    Public ml1280x1024Height    As Long     '2001/05/09 Made Public
    Public ml1400x1050Width     As Long     '2001/01/28 Added
    Public ml1400x1050Height    As Long     '2001/01/28 Added
    Public ml1600x1200Width     As Long     '2001/01/28 Added
    Public ml1600x1200Height    As Long     '2001/05/09 Made Public
    '
    

Public Sub TB_CenterForm(frm As Form, Optional vntOffsetLeft As Variant, Optional vntOffsetTop As Variant, Optional vntConsiderTaskbar As Variant)
' Center the form within the zone or in the center of the screen. Larry.
    Dim lLeft As Long, lTop As Long
    Dim lOffsetLeft As Long
    Dim lOffsetTop  As Long
    Dim bConsiderTaskbar As Boolean
    
    If Not IsMissing(vntOffsetLeft) Then
        lOffsetLeft = vntOffsetLeft
    End If
    If Not IsMissing(vntOffsetTop) Then
        lOffsetTop = vntOffsetTop
    End If
    If Not IsMissing(vntConsiderTaskbar) Then
        bConsiderTaskbar = vntConsiderTaskbar
    End If
    With frm
        If .WindowState = vbNormal Then                 'do only if Normal, not minimized or maximized
            If bConsiderTaskbar Then                    'use center within work area
                TB_CenterForm32 frm, lOffsetLeft, lOffsetTop
                Exit Sub
            ElseIf TB_ZoneWidth + TB_ZoneHeight > 0 Then    'any values here then use boundaries
                lLeft = (TB_ZoneLeft + (TB_ZoneWidth \ 2)) - (.Width \ 2) + lOffsetLeft
                lTop = (TB_ZoneTop + (TB_ZoneHeight \ 2)) - (.Height \ 2) + lOffsetTop
                If lLeft + .Width > Screen.Width Then       'make sure it is on the screen
                    lLeft = Screen.Width - .Width
                ElseIf lLeft < 0 Then
                    lLeft = 0
                End If
                If lTop + .Height > Screen.Height Then
                    lTop = Screen.Height - .Height
                ElseIf lTop < 0 Then
                    lTop = 0
                End If
            Else                                    'center in middle of the screen
                lLeft = ((Screen.Width - .Width) \ 2) + lOffsetLeft     'Center form horizontally.
                lTop = (((Screen.Height - .Height) \ 2) + lOffsetTop)   'vertically
            End If
            If .Left = lLeft And .Top = lTop Then   'no move needed
            Else
                .Move lLeft, lTop                   'move to centered location
            End If
        End If
    End With
End Sub

Public Sub TB_ZoneFullScreen()
    TB_ZoneWidth = 0
    TB_ZoneHeight = 0
End Sub

Public Sub TB_Zone(frm As Form, Optional vntUseScaleValues As Variant)
' 94/11/12 The form will be shown within the boundry of the form
    'Pass the name of the form to act as the boundry
    Dim bUseScaleValues As Boolean
    If Not IsMissing(vntUseScaleValues) Then
        bUseScaleValues = vntUseScaleValues
    End If
    With frm
        If bUseScaleValues Then
            TB_ZoneLeft = .ScaleLeft + .Left
            TB_ZoneTop = .ScaleTop + .Top
            TB_ZoneWidth = .ScaleWidth
            TB_ZoneHeight = .ScaleHeight
        Else
            TB_ZoneLeft = .Left
            TB_ZoneTop = .Top
            TB_ZoneWidth = .Width
            TB_ZoneHeight = .Height
        End If
    End With
End Sub

Public Sub TB_CenterFormInMDI(mdi As MDIForm, frm As Form)
' Use 'old' method to center in MDI work area
    Dim l As Long, t As Long
    With frm
        If .WindowState = vbNormal Then
            l = (mdi.ScaleWidth - .Width) \ 2
            t = (mdi.ScaleHeight - .Height) \ 2
            If l < 0 Then       'can't have caption off the screen
                l = 0
            End If
            If t < 0 Then
                t = 0
            End If
            .Move l, t
        End If
    End With
End Sub

Public Sub TB_FillFormInMDI(mdi As MDIForm, frm As Form)
' Fill the mdi work area with the form
    If frm.WindowState = vbNormal Then
        frm.Move 0, 0, mdi.ScaleWidth, mdi.ScaleHeight
    End If
End Sub

Public Function TB_GetFormInformation(frm As Form, iState As Integer, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, Optional vntFormName As Variant, _
    Optional vntFullScreen As Variant) As Integer
' Return True if this is successful, otherwise false
' Return the values found in the registry
    Dim lL As Long, lT As Long, lW As Long, lH As Long
    Dim sState      As String
    Dim sLocation   As String
    Dim sKey        As String
    
    If IsMissing(vntFormName) Then
        sKey = SetKeyForm(frm)
    Else
        sKey = SetKeyForm(frm, vntFormName)
    End If
    sLocation = RegRead(sKey, mcsLocation)              'read the setting
    If sLocation <> "" Then                             'none found
        If ConvertLocationToIntegers(sLocation, lL, lT, lW, lH) Then
            TB_MakeSureOnScreen lL, lT, lW, lH          'make sure it is on the screen
            lLeft = lL
            lTop = lT
            lWidth = lW
            lHeight = lH
            TB_GetFormInformation = True
        End If
    End If
    sState = RegRead(sKey, mcsState)
    iState = ConvertStateToInteger(sState)
    If iState = vbMinimized Then
        iState = vbNormal
    End If
    If Not IsMissing(vntFullScreen) Then
        vntFullScreen = TB_SaveGetFullScreen(frm, sKey, False)
    End If
End Function

Public Function TB_SaveFormInformation(frm As Form, Optional vntFormName As Variant) As Integer
' Save information about a form in the registry
    Dim iState      As Integer
    Dim sFormName   As String
    Dim sKey        As String
    With frm
        sFormName = .Name
        iState = .WindowState
    End With
    If IsMissing(vntFormName) Then
        sKey = SetKeyForm(frm)
    Else
        sKey = SetKeyForm(frm, vntFormName)
    End If
    RegWrite sKey, mcsState, CStr(ConvertStateToString(iState))
    If iState = vbNormal Then   'don't change the location if not normal state
        RegWrite sKey, mcsLocation, CStr(ConvertLocationToString(frm))
    End If
    TB_SaveGetFullScreen frm, sKey, True
    TB_SaveFormInformation = True
End Function

Public Function TB_SaveGetFullScreen(frm As Form, sKey As String, bSave As Boolean) As Boolean
    Dim l, t, w, h
    Dim sValue As String
    Const csFillDesktop = "FillDesktop"
    TB_GetDesktopWorkArea l, t, w, h
    If bSave Then
        If frm.Width = w And frm.Height = h Then
            RegWrite sKey, csFillDesktop, CStr(True)
        Else
            RegWrite sKey, csFillDesktop, ""
        End If
    Else
        sValue = RegRead(sKey, csFillDesktop)
        If sValue <> "" Then
            TB_SaveGetFullScreen = CBool(sValue)
        End If
    End If
End Function

Public Sub TB_SaveScreenSize(frm As Form)
' 96/07/03 Save the screen size so that if the user changes resolution we show MMI in its default location. Larry.
    Dim sKey As String
    sKey = SetKeyForm(frm)
    RegWrite sKey, mcsScreenWidth, CStr(Screen.Width)
    RegWrite sKey, mcsScreenHeight, CStr(Screen.Height)
End Sub
Public Function TB_GetScreenSize(frm As Form, lWidth As Long, lHeight As Long) As Boolean
' 96/07/03 Get the screen size saved by subroutine SaveScreenSize
    Dim sKey As String
    Dim bBad As Boolean
    sKey = SetKeyForm(frm)
    lWidth = Val(RegRead(sKey, mcsScreenWidth))
    If lWidth = 0 Then
        bBad = True
    End If
    lHeight = Val(RegRead(sKey, mcsScreenHeight))
    If lHeight = 0 Then
        bBad = True
    End If
    If Not bBad Then
        TB_GetScreenSize = True
    End If
End Function

Public Function ConvertStateToString(iState As Integer) As String
' return the WindowState as as string
    Select Case iState
        Case vbNormal
            ConvertStateToString = "Normal"
        Case vbMinimized
            ConvertStateToString = "Minimized"
        Case vbMaximized
            ConvertStateToString = "Maximized"
    End Select
End Function

Private Function ConvertStateToInteger(sState As String) As Integer
' return the WindowState as an integer
' 95/08/30 Bug - select statement used sState not s - Larry
    Dim iTemp As Integer        'temp integer
    Dim sTemp As String         'temp string
    sTemp = LCase$(sState)      'convert to all lower
    Select Case sTemp           'what state - 95/08/30 Use s not sState - Larry
        Case "normal"
            iTemp = vbNormal
        Case "minimized"
            iTemp = vbMinimized
        Case "maximized"
            iTemp = vbMaximized
    End Select
    ConvertStateToInteger = iTemp   'return the value
End Function
Private Function ConvertLocationToString(frm As Form) As String
' Return the forms location in a string, for example: "20,40,4600,3200" for left,top,width,height
    Dim sTemp As String
    With frm
        sTemp = CStr(.Left) & mcsComma
        sTemp = sTemp & CStr(.Top) & mcsComma
        sTemp = sTemp & CStr(.Width) & mcsComma
        sTemp = sTemp & CStr(.Height)
    End With
    ConvertLocationToString = sTemp
End Function

Private Function ConvertLocationToIntegers(sTheString As String, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long) As Integer
' Return the four numbers found in a string variable as long integers, for example:
' input:  sTheString = "20,40,4600,3200"
' output: lLeft = 20
'         lTop = 40
'         lWidth = 4600
'         lHeight = 3200
' Return True if successful, False if can't process or all values are zero
    Dim i As Integer                'place counter
    Dim l As Long                   'temp long integer
    Dim iCnt As Integer             'counter
    Dim sTemp As String             'temporary string
    sTemp = "@" & sTheString & "@"  '95/08/30 Make sure we have 3 commas
    ' check, make sure we have 3 commas
    While InStr(sTemp, mcsComma)
        iCnt = iCnt + 1
        i = InStr(sTemp, mcsComma)
        sTemp = Mid$(sTemp, i + 1)
    Wend
    If iCnt <> 3 Then Exit Function 'error, none so exit
    
    iCnt = 0
    sTemp = sTheString & mcsComma & " "  'add an extra comma and space
    On Error GoTo ConvertLocationToIntegersError    'bad string or something
    While InStr(sTemp, mcsComma)
        i = InStr(sTemp, mcsComma)       'where is it
        l = CLng(Mid$(sTemp, 1, i - 1)) 'convert it to an integer
        iCnt = iCnt + 1
        Select Case iCnt
            Case 1
                lLeft = l
            Case 2
                lTop = l
            Case 3
                lWidth = l
            Case 4
                lHeight = l
            Case Else
                Exit Function       'should not
        End Select
        sTemp = Mid$(sTemp, i + 1)
    Wend
    If iCnt = 4 Then
        If lLeft + lTop + lWidth + lHeight > 0 Then
            ConvertLocationToIntegers = True    'success
        End If
    End If
    Exit Function
ConvertLocationToIntegersError:                 'bad, no real harm
End Function

Public Function SetKeyCompany(Optional vntCompany As Variant) As String
    Dim sCompany As String
    If Not IsMissing(vntCompany) Then
        sCompany = vntCompany
    Else
        If gsCompany <> "" Then
            sCompany = gsCompany
        Else
            sCompany = mcsCompany
        End If
    End If
    SetKeyCompany = "Software\" & sCompany
End Function

Public Function SetKeyProduct(Optional vntProductName As Variant) As String
' 1998/02/25 Add optional Product Name
    Dim sTemp As String
    sTemp = SetKeyCompany
    If IsMissing(vntProductName) Then
        If gsProductName <> "" Then
            sTemp = sTemp & "\" & gsProductName
        Else
            sTemp = sTemp & "\" & mcsProductName
        End If
    Else
        If vntProductName <> "" Then
            sTemp = sTemp & "\" & vntProductName
        End If
    End If
    SetKeyProduct = sTemp
End Function

Public Function SetKeyForm(frm As Form, Optional vntFormName As Variant) As String
    Dim sTemp As String
    sTemp = SetKeyProduct
    If IsMissing(vntFormName) Then
        sTemp = sTemp & "\" & frm.Name
    Else
        If vntFormName <> "" Then
            sTemp = sTemp & "\" & Trim$(vntFormName)
        End If
    End If
    SetKeyForm = sTemp
End Function

Public Sub TB_WindowItem(frm As Form, Index As Integer, Optional vntSkipCenter As Variant)
' Process the window menu item
    Dim bSkipCenter As Boolean
    
    If Not IsMissing(vntSkipCenter) Then
        bSkipCenter = vntSkipCenter
    End If
    
    Select Case Index
        Case gciWinCenter
            TB_CenterForm32 frm
        Case gciWinTileHorizontal To gciWinArrangeIcons
            TB_WindowArrange frm, Index
'        Case mci640x480 To mci1280x1024
        Case mci640x480 To mci1600x1200     '2001/10/08 Add mci1600x1200 support
            TB_WindowItemSize frm, Index, bSkipCenter
        Case mciFullScreen
            TB_FillDesktop32 frm
        Case mciDockLeft To mciDockBottom
            TB_DoDock frm, Index
        Case mciCornerUpLeft To mciCornerLowerRight
            TB_Corner frm, Index
    End Select
End Sub

Private Sub TB_WindowArrange(mdifrm As MDIForm, Index As Integer)
    Dim iTemp As Integer
    Select Case Index
        Case gciWinTileHorizontal
            iTemp = vbTileHorizontal
        Case gciWinTileVertical
            iTemp = vbTileVertical
        Case gciWinCascade
            iTemp = vbCascade
        Case gciWinArrangeIcons
            iTemp = vbArrangeIcons
    End Select
    mdifrm.Arrange iTemp
End Sub

Public Sub TB_WindowItemSize(frm As Form, Index As Integer, Optional vntSkipCenter As Variant)
' Set window size based upon standard sizes
' 2001/07/19 Add vntSkipCenter
' 2003/01/17 Add support for Sony SDM-V72W wide monitor

    Dim l, t, w, h
    Dim bSkipCenter As Boolean
    
    If Not IsMissing(vntSkipCenter) Then
        bSkipCenter = vntSkipCenter
    End If
    
    TB_CalcWindowSizes
    Select Case Index
        Case mci640x480
            w = mlVGAWidth
            h = mlVGAHeight
        Case mci800x600
            w = ml800x600Width
            h = ml800x600Height
        Case mci1024x768
            w = ml1024x768Width
            h = ml1024x768Height
        Case mci1152x864
            w = ml1152x864Width
            h = ml1152x864Height
        Case mci1280x768                '2003/01/17 Added
            w = gl1280x768Width         '2003/01/17
            h = gl1280x768Height        '2003/01/17
        Case mci1280x1024
            w = ml1280x1024Width
            h = ml1280x1024Height
        Case mci1400x1050               '2001/10/08 Added
            w = ml1400x1050Width
            h = ml1400x1050Height
        Case mci1600x1200               '2001/10/08 Added
            w = ml1600x1200Width
            h = ml1600x1200Height
    End Select
    TB_WindowSizeOnWindow l, t, w, h    'make sure it is on the screen
    With frm
        If .WindowState = vbMaximized Then
            .WindowState = vbNormal
        End If
        .Move .Left, .Top, w, h
    End With
    If Not bSkipCenter Then             '2001/07/19 Added
        TB_CenterForm32 frm
    Else
        With frm                        '2001/07/19 Added
            l = .Left
            t = .Top
            w = .Width
            h = .Height
            If Not TB_MakeSureOnScreen(l, t, w, h) Then
                .Move l, t, w, h
            End If
        End With
    End If
    ' The next commands fires a FormResize event in the Child form if any.
    On Error Resume Next
    frm.ActiveForm.FormResize
End Sub

Public Sub TB_CalcWindowSizes()
' 1996/07/19 Calculate the various window sizes. Larry.
    Static bHerebefore As Boolean
    Dim lTWX As Long
    Dim lTWY As Long
    
    If bHerebefore Then Exit Sub        'only need to do once
    bHerebefore = True                  'only once

    lTWX = Screen.TwipsPerPixelX        'local variable, faster
    lTWY = Screen.TwipsPerPixelY

    ' Convert to Twips

    mlVGAWidth = 640 * lTWX
    mlVGAHeight = 480 * lTWY
    glVGAWidth = mlVGAWidth             '2000/09/22 Added
    glVGAHeight = mlVGAHeight           '2000/09/22 Added

    ml800x600Width = 800 * lTWX
    ml800x600Height = 600 * lTWY

    ml1024x768Width = 1024 * lTWX
    ml1024x768Height = 768 * lTWY

    ml1152x864Width = 1152 * lTWX
    ml1152x864Height = 864 * lTWY
    
    gl1280x768Width = 1280 * lTWX       '2003/01/17 Added
    gl1280x768Height = 768 * lTWY       '2003/01/17 Added

    ml1280x1024Width = 1280 * lTWX
    ml1280x1024Height = 1024 * lTWY

    ml1400x1050Width = 1400 * lTWX      '2001/01/28 Added
    ml1400x1050Height = 1050 * lTWY

    ml1600x1200Width = 1600 * lTWX      '2001/01/28 Added
    ml1600x1200Height = 1200 * lTWY

End Sub

Public Sub TB_Window(frm As Form)
' 2001/02/05 Added, easier to remember than TB_SetWindowMenu
    TB_SetWindowMenu frm
End Sub

Private Sub TB_SetWindowMenu(frm As Form)
' 1996/07/19 Disable Menu Item if size does not apply. Larry
    TB_SetWindowArrangeMenus frm
    TB_SetDockMenu frm
    TB_SetCornerMenu frm
    With frm
        ' Center item
        .mnuWindowItem(gciWinCenter).Enabled = Not TB_WindowIsItAlreadyCentered(frm)
        
        ' Full Screen Item
        .mnuWindowItem(mciFullScreen).Enabled = Not TB_WindowIsItAlreadyFullScreen(frm)
        
        ' Size items
        If .Width = mlVGAWidth And .Height = mlVGAHeight Then
            .mnuWindowItem(mci640x480).Enabled = False
        Else
            .mnuWindowItem(mci640x480).Enabled = True
        End If
        If .Width = ml800x600Width And .Height = ml800x600Height Then
            .mnuWindowItem(mci800x600).Enabled = False
        Else
            .mnuWindowItem(mci800x600).Enabled = True
        End If
        If .Width = ml1024x768Width And .Height = ml1024x768Height Then
            .mnuWindowItem(mci1024x768).Enabled = False
        Else
            .mnuWindowItem(mci1024x768).Enabled = True
        End If
        
        If .Width = ml1152x864Width And .Height = ml1152x864Height Then
            .mnuWindowItem(mci1152x864).Enabled = False
        Else
            .mnuWindowItem(mci1152x864).Enabled = True
        End If
        
        If .Width = gl1280x768Width And .Height = gl1280x768Height Then     '2003/01/17 Added
            .mnuWindowItem(mci1280x768).Enabled = False
        Else
            .mnuWindowItem(mci1280x768).Enabled = True
        End If
    
        If .Width = ml1280x1024Width And .Height = ml1280x1024Height Then
            .mnuWindowItem(mci1280x1024).Enabled = False
        Else
            .mnuWindowItem(mci1280x1024).Enabled = True
        End If
    
        If .Width = ml1400x1050Width And .Height = ml1400x1050Height Then   '2001/10/08 Added
            .mnuWindowItem(mci1400x1050).Enabled = False
        Else
            .mnuWindowItem(mci1400x1050).Enabled = True
        End If
        
        If .Width = ml1600x1200Width And .Height = ml1600x1200Height Then   '2001/10/08 Added
            .mnuWindowItem(mci1600x1200).Enabled = False
        Else
            .mnuWindowItem(mci1600x1200).Enabled = True
        End If
    
    End With
End Sub

Private Sub TB_SetWindowArrangeMenus(frm As Form)
' Set the Arrange menu items enabled/disabled based upon if there are any child forms
    Dim i               As Integer
    Dim bValueChild     As Boolean
    Dim bValueMin       As Boolean
    
    For i = 1 To Forms.Count - 1
        On Error GoTo TB_SetWindowArrangeMenusResume    'in case form does not have property MDIChild
        If Forms(i).MDIChild Then
            If Forms(i).WindowState = vbMinimized Then  'is it vbMinimized
                bValueMin = True
            Else
                bValueChild = True
            End If
        End If
TB_SetWindowArrangeMenusResume:
    Next
    With frm
        For i = gciWinTileHorizontal To gciWinArrangeIcons
            If i = gciWinArrangeIcons Then
                .mnuWindowItem(i).Enabled = bValueMin
            Else
                .mnuWindowItem(i).Enabled = bValueChild
            End If
        Next
    End With
End Sub
Public Sub TB_BuildWindowMenu(frm As Form, Optional vntNoBuildArrangeMenu As Variant, Optional vntDockMenu As Variant, _
   Optional vntNoBuildSizeMenu As Variant, Optional vntCornerMenu As Variant)
    Dim i As Integer
    Dim bNoBuildArrangeMenu As Boolean
    Dim bDockMenu           As Boolean
    Dim bNoBuildSizeMenu    As Boolean
    Dim bCornerMenu         As Boolean      '2002/08/11
    If Not IsMissing(vntNoBuildArrangeMenu) Then
        bNoBuildArrangeMenu = vntNoBuildArrangeMenu
    End If
    If Not IsMissing(vntDockMenu) Then
        bDockMenu = vntDockMenu
    End If
    If Not IsMissing(vntNoBuildSizeMenu) Then
        bNoBuildSizeMenu = vntNoBuildSizeMenu
    End If
    If Not IsMissing(vntCornerMenu) Then    '2002/08/11
        bCornerMenu = vntCornerMenu
    End If
    RemoveAllMenuItems frm
    With frm
        For i = gciWinCenter To mciFullScreen
            Load .mnuWindowItem(i)
            With .mnuWindowItem(i)
                .Visible = True
            End With
        Next
        ' Arrange
        .mnuWindowItem(gciWinCenter).Caption = mcsWinCenter
        .mnuWindowItem(gciWinTileHorizontal).Caption = mcsWinTileHorizontal
        .mnuWindowItem(gciWinTileVertical).Caption = mcsWinTileVertical
        .mnuWindowItem(gciWinCascade).Caption = mcsWinCascade
        .mnuWindowItem(gciWinArrangeIcons).Caption = mcsWinArrangeIcons
        ' Size
        .mnuWindowItem(mciSepSize).Caption = "-"
        .mnuWindowItem(mci640x480).Caption = mcs640x480
        .mnuWindowItem(mci800x600).Caption = mcs800x600
        .mnuWindowItem(mci1024x768).Caption = mcs1024x768
        .mnuWindowItem(mci1152x864).Caption = mcs1152x864
        .mnuWindowItem(mci1280x768).Caption = mcs1280x768       '2003/01/17
        .mnuWindowItem(mci1280x1024).Caption = mcs1280x1024
        .mnuWindowItem(mci1400x1050).Caption = mcs1400x1050
        .mnuWindowItem(mci1600x1200).Caption = mcs1600x1200
        .mnuWindowItem(mciFullScreen).Caption = mcsFullScreen
        .mnuWindowItem(0).Visible = False
    End With
    TB_CalcWindowSizes
    With frm
        ' Make sure it will fit in Desktop Area
        If Not TB_WindowSizeCanItFit(ml1600x1200Width, ml1600x1200Height) Then
            .mnuWindowItem(mci1600x1200).Visible = False
            If Not TB_WindowSizeCanItFit(ml1400x1050Width, ml1400x1050Height) Then
                .mnuWindowItem(mci1400x1050).Visible = False
                If Not TB_WindowSizeCanItFit(ml1280x1024Width, ml1280x1024Height) Then
                    .mnuWindowItem(mci1280x1024).Visible = False
                    If Not TB_WindowSizeCanItFit(ml1152x864Width, ml1152x864Height) Then
                        .mnuWindowItem(mci1152x864).Visible = False
                        If Not TB_WindowSizeCanItFit(gl1280x768Width, gl1280x768Height) Then    '2003/01/17
                            .mnuWindowItem(mci1280x768).Visible = False                         '2003/01/17
                            If Not TB_WindowSizeCanItFit(ml1024x768Width, ml1024x768Height) Then
                                .mnuWindowItem(mci1024x768).Visible = False
                                If Not TB_WindowSizeCanItFit(ml800x600Width, ml800x600Height) Then
                                    .mnuWindowItem(mci800x600).Visible = False
                                    If Not TB_WindowSizeCanItFit(mlVGAWidth, mlVGAHeight) Then
                                        .mnuWindowItem(mci640x480).Enabled = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    If bNoBuildArrangeMenu Then
        With frm
            For i = gciWinTileHorizontal To gciWinArrangeIcons
                .mnuWindowItem(i).Visible = False
            Next
        End With
    End If
    If bDockMenu Then
        With frm
            For i = mciFullScreen + 1 To mciFullScreen + 1 + 4
                Load .mnuWindowItem(i)
                With .mnuWindowItem(i)
                    .Visible = True
                End With
            Next
            .mnuWindowItem(mciSepDock).Caption = "-"
            .mnuWindowItem(mciDockLeft).Caption = mcsDockLeft
            .mnuWindowItem(mciDockRight).Caption = mcsDockRight
            .mnuWindowItem(mciDockTop).Caption = mcsDockTop
            .mnuWindowItem(mciDockBottom).Caption = mcsDockBottom
        End With
    End If
    If bCornerMenu Then         '2002/08/11 Add corner menu
        With frm
            For i = mciSepCorner To mciCornerLowerRight
                Load .mnuWindowItem(i)
                With .mnuWindowItem(i)
                    .Visible = True
                End With
            Next
            .mnuWindowItem(mciSepCorner).Caption = "-"
            .mnuWindowItem(mciCornerUpLeft).Caption = mcsCornerUpLeft
            .mnuWindowItem(mciCornerUpRight).Caption = mcsCornerUpRight
            .mnuWindowItem(mciCornerLowerLeft).Caption = mcsCornerLowerLeft
            .mnuWindowItem(mciCornerLowerRight).Caption = mcsCornerLowerRight
        End With
    End If
    If bNoBuildSizeMenu Then
        With frm
            For i = mciSepSize To mciFullScreen
                .mnuWindowItem(i).Visible = False
            Next
        End With
    End If
End Sub

Private Sub RemoveAllMenuItems(frm As Form)
    Dim i As Integer
    
    With frm
        .mnuWindowItem(0).Visible = True
        For i = .mnuWindowItem.Count - 1 To 1 Step -1
            Unload .mnuWindowItem(i)
        Next
    End With
End Sub

Public Function TB_CenterForm32(frm As Form, Optional vntOffsetLeft As Variant, Optional vntOffsetTop As Variant, Optional vntAlreadyCentered As Variant) As Boolean
' Center within the desktop work area
    Dim tDesktopArea    As RECT
    Dim lOffsetLeft     As Long
    Dim lOffsetTop      As Long
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim lLeft   As Long
    Dim lTop    As Long
    Dim lWidth  As Long
    Dim lHeight As Long
    
    If Not IsMissing(vntOffsetLeft) Then                'any offsets?
        lOffsetLeft = vntOffsetLeft
    End If
    If Not IsMissing(vntOffsetTop) Then
        lOffsetTop = vntOffsetTop
    End If
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    
    lLeft = (lScreenWidth - frm.Width) \ 2 + lScreenLeft + lOffsetLeft
    lTop = (lScreenHeight - frm.Height) \ 2 + lScreenTop + lOffsetTop
    
    With frm
        If .WindowState = vbNormal Then
            If .Left <> lLeft Or .Top <> lTop Then          'move it?
                If IsMissing(vntAlreadyCentered) Then       'just testing to see if centered, don't move it
                    .Move lLeft, lTop                       'do move
                    TB_CenterForm32 = True                  'report moved
                End If
            Else
                If Not IsMissing(vntAlreadyCentered) Then
                    vntAlreadyCentered = True
                End If
            End If
        End If
    End With
End Function


Public Function TB_FillDesktop32(frm As Form, Optional vntAlreadyFull As Variant) As Boolean
' Move the form to fill the desktop work area
    Dim tDesktopArea    As RECT
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    With frm
        If lScreenLeft <> .Left Or lScreenTop <> .Top Or lScreenWidth <> .Width Or lScreenHeight <> .Height Then
            If IsMissing(vntAlreadyFull) Then
                If .WindowState <> vbNormal Then
                    .WindowState = vbNormal
                End If
                .Move lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
                TB_FillDesktop32 = True
            End If
        Else
            If Not IsMissing(vntAlreadyFull) Then
                vntAlreadyFull = True
            End If
        End If
    End With
End Function

Public Function TB_GetDesktopWorkArea(lScreenLeft As Long, lScreenTop As Long, lScreenWidth As Long, lScreenHeight As Long) As Boolean
' Get the desktop work area using API SystemParametersInfo
    Const SPI_GETWORKAREA = 48
    Dim tDesktopArea As RECT
    
    SystemParametersInfo SPI_GETWORKAREA, 0, tDesktopArea, 0    'issue the API
    lScreenLeft = tDesktopArea.Left * Screen.TwipsPerPixelX
    lScreenTop = tDesktopArea.Top * Screen.TwipsPerPixelY
    lScreenWidth = (tDesktopArea.Right - tDesktopArea.Left) * Screen.TwipsPerPixelX
    lScreenHeight = (tDesktopArea.Bottom - tDesktopArea.Top) * Screen.TwipsPerPixelY
End Function

Public Function TB_WindowSizeOnWindow(l As Long, t As Long, w As Long, h As Long) As Boolean
' Make sure it is in the Work area
    Dim tDesktopArea    As RECT
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    If w > lScreenWidth Or h > lScreenHeight Then
        l = lScreenLeft
        t = lScreenTop
        w = lScreenWidth
        h = lScreenHeight
        TB_WindowSizeOnWindow = True
    End If
End Function

Public Function TB_WindowSizeCanItFit(w As Long, h As Long) As Boolean
' Make sure it will fit
    Dim tDesktopArea    As RECT
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    If w <= lScreenWidth And h <= lScreenHeight Then
        TB_WindowSizeCanItFit = True
    End If
End Function

Public Function TB_WindowIsItAlreadyFullScreen(mdifrm As Form) As Boolean
' Is the form already filling the desktop work area
    Dim bValue As Boolean
    TB_FillDesktop32 mdifrm, bValue
    TB_WindowIsItAlreadyFullScreen = bValue
End Function
Public Function TB_WindowIsItAlreadyCentered(mdifrm As Form) As Boolean
' Is it already centered in the desktop work area
    Dim bValue  As Boolean
    TB_CenterForm32 mdifrm, , , bValue      'get value, is it already centered
    TB_WindowIsItAlreadyCentered = bValue
End Function

Public Function TB_MakeSureOnScreen(lL As Long, lT As Long, lW As Long, lH As Long) As Boolean
' Make sure the form is on the screen, used at form load
    Dim tDesktopArea    As RECT
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim bValue          As Boolean
    
    bValue = True
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    If lW > lScreenWidth Then
        lW = lScreenWidth
        bValue = False                  'was not on screen
    End If
    If lH > lScreenHeight Then
        lH = lScreenHeight
        bValue = False                  'was not on screen
    End If
    If lL < 0 Then
        lL = 0
        bValue = False                  'was not on screen
    ElseIf lL + lW > lScreenWidth Then
        lL = lScreenWidth - lW
        bValue = False                  'was not on screen
    End If
    If lT < 0 Then
        lT = 0
        bValue = False                  'was not on screen
    ElseIf lT + lH > lScreenHeight Then
        lT = lScreenHeight - lH
        bValue = False                  'was not on screen
    End If
    TB_MakeSureOnScreen = bValue
End Function

Private Sub TB_Dock(frm As Form, iWhere As eDock)
' 2000/08/06 Sub created by Larry Rebich while in Weed, California
' 2001/05/10 Modified to support right and top 'start' bars.

    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    Select Case iWhere
        Case mcieDockLeft
            lScreenWidth = frm.Width
        Case mcieDockRight
            lScreenLeft = lScreenLeft + lScreenWidth - frm.Width
            lScreenWidth = frm.Width
        Case mcieDockTop
            lScreenHeight = frm.Height
        Case mcieDockBottom
            lScreenTop = lScreenTop + lScreenHeight - frm.Height
            lScreenHeight = frm.Height
    End Select
    With frm
        .Move lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    End With
End Sub

Public Sub TB_DoDock(frm As Form, iIndex As Integer)
' 2000/08/06 Sub created by Larry Rebich while in Weed, California
    Select Case iIndex
        Case mciDockLeft
            TB_Dock frm, eLeft
        Case mciDockRight
            TB_Dock frm, eRight
        Case mciDockTop
            TB_Dock frm, eTop
        Case mciDockBottom
            TB_Dock frm, eBottom
    End Select
End Sub

Public Sub TB_SetWindowMenuUseMinMax(frm As Form, lMinWidth As Long, lMaxWidth As Long, lMinHeight As Long, lMaxHeight As Long)
' 2001/05/10 Sub created by Larry Rebich while in La Quinta, California.
' Set the Window Size menu items based upon min and max window settings.
' Works in conjunction with clsMinWindow
    Dim l, t, w, h
    TB_GetDesktopWorkArea l, t, w, h
    With frm
        If lMaxWidth > 0 And lMaxHeight > 0 Then
            If lMaxWidth < w Or lMaxHeight < h Then     'no full screen unless max is larger than full screen
                .mnuWindowItem(mciFullScreen).Enabled = False
            End If
            If lMaxWidth < ml1600x1200Width Or lMaxHeight < ml1600x1200Height Then
                .mnuWindowItem(mci1600x1200).Enabled = False
            End If
            If lMaxWidth < ml1400x1050Width Or lMaxHeight < ml1400x1050Height Then
                .mnuWindowItem(mci1400x1050).Enabled = False
            End If
            If lMaxWidth < ml1280x1024Width Or lMaxHeight < ml1280x1024Height Then
                .mnuWindowItem(mci1280x1024).Enabled = False
            End If
            If lMaxWidth < gl1280x768Width Or lMaxHeight < gl1280x768Height Then    '2003/01/17
                .mnuWindowItem(mci1280x768).Enabled = False
            End If
            If lMaxWidth < ml1152x864Width Or lMaxHeight < ml1152x864Height Then
                .mnuWindowItem(mci1152x864).Enabled = False
            End If
            If lMaxWidth < ml1024x768Width Or lMaxHeight < ml1024x768Height Then
                .mnuWindowItem(mci1024x768).Enabled = False
            End If
            If lMaxWidth < ml800x600Width Or lMaxHeight < ml800x600Height Then
                .mnuWindowItem(mci800x600).Enabled = False
            End If
            If lMaxWidth < mlVGAWidth Or lMaxHeight < mlVGAHeight Then
                .mnuWindowItem(mci640x480).Enabled = False
            End If
        End If
        If lMinWidth > 0 And lMinHeight > 0 Then
            If lMinWidth > mlVGAWidth Or lMinHeight > mlVGAHeight Then
                .mnuWindowItem(mci640x480).Enabled = False
            End If
        End If
    End With
End Sub

Public Sub TB_DoCorner(frm As Form, iWhere As eCorner)
' 2002/08/11 Sub created by Larry Rebich while in Kelowna, BC, Canada.
    TB_Corner frm, iWhere + mciCornerUpLeft - 1
End Sub

Private Sub TB_Corner(frm As Form, iWhere As Integer)
' 2002/08/11 Sub created by Larry Rebich while in Kelowna, BC, Canada.
'    Dim iIndex As Integer
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim l As Long, t As Long, w As Long, h As Long
    
    If frm.WindowState = vbNormal Then
'        iIndex = iWhere + mciCornerUpLeft - 1
        TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    
        Select Case iWhere
            Case mciCornerUpLeft
                l = lScreenLeft
                t = lScreenTop
            Case mciCornerUpRight
                l = lScreenLeft + lScreenWidth - frm.Width
                t = lScreenTop
            Case mciCornerLowerLeft
                l = lScreenLeft
                t = lScreenTop + lScreenHeight - frm.Height
            Case mciCornerLowerRight
                l = lScreenLeft + lScreenWidth - frm.Width
                t = lScreenTop + lScreenHeight - frm.Height
        End Select
        frm.Move l, t
    End If
End Sub

Private Sub TB_SetDockMenu(frm As Form)
' 2002/08/11 Sub created by Larry Rebich while in Kelowna, BC, Canada.
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim l As Long, t As Long, w As Long, h As Long
    Dim i As Integer
    
    On Error GoTo TB_SetDockMenuEH
    
    If frm.WindowState = vbNormal Then
        TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight

        With frm        'easier to work with
            l = .Left
            t = .Top
            w = .Width
            h = .Height
        End With

        If l = lScreenLeft And t = lScreenTop And h = lScreenHeight Then
            frm.mnuWindowItem(mciDockLeft).Enabled = False
        Else
            frm.mnuWindowItem(mciDockLeft).Enabled = True
        End If
    
        If (l = lScreenLeft + lScreenWidth - w) And t = lScreenTop And h = lScreenHeight Then
            frm.mnuWindowItem(mciDockRight).Enabled = False
        Else
            frm.mnuWindowItem(mciDockRight).Enabled = True
        End If

        If l = lScreenLeft And t = lScreenTop And w = lScreenWidth Then
            frm.mnuWindowItem(mciDockTop).Enabled = False
        Else
            frm.mnuWindowItem(mciDockTop).Enabled = True
        End If
    
        If l = lScreenLeft And (t = lScreenTop + lScreenHeight - h) And w = lScreenWidth Then
            frm.mnuWindowItem(mciDockBottom).Enabled = False
        Else
            frm.mnuWindowItem(mciDockBottom).Enabled = True
        End If
    Else
        For i = mciDockLeft To mciDockBottom
            frm.mnuWindowItem(i).Enabled = False
        Next
    End If

TB_SetDockMenuEH:

End Sub
Private Sub TB_SetCornerMenu(frm As Form)
' 2002/08/11 Sub created by Larry Rebich while in Kelowna, BC, Canada.
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim l As Long, t As Long, w As Long, h As Long
    Dim i As Integer
    
    On Error GoTo TB_SetCornerMenuEH    'in case does not exist
    
    If frm.WindowState = vbNormal Then
        TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight

        With frm        'easier to work with
            l = .Left
            t = .Top
            w = .Width
            h = .Height
        End With
    
        If l = lScreenLeft And t = lScreenTop Then
            frm.mnuWindowItem(mciCornerUpLeft).Enabled = False
        Else
            frm.mnuWindowItem(mciCornerUpLeft).Enabled = True
        End If

        If (l = (lScreenLeft + lScreenWidth - w)) And (t = lScreenTop) Then
            frm.mnuWindowItem(mciCornerUpRight).Enabled = False
        Else
            frm.mnuWindowItem(mciCornerUpRight).Enabled = True
        End If
    
        If (l = lScreenLeft) And (t = lScreenTop + lScreenHeight - h) Then
            frm.mnuWindowItem(mciCornerLowerLeft).Enabled = False
        Else
            frm.mnuWindowItem(mciCornerLowerLeft).Enabled = True
        End If
        If (l = (lScreenLeft + lScreenWidth - w)) And (t = lScreenTop + lScreenHeight - h) Then
            frm.mnuWindowItem(mciCornerLowerRight).Enabled = False
        Else
            frm.mnuWindowItem(mciCornerLowerRight).Enabled = True
        End If
    Else
        For i = mciCornerUpLeft To mciCornerLowerRight
            frm.mnuWindowItem(i).Enabled = False
        Next
    End If
TB_SetCornerMenuEH:
End Sub


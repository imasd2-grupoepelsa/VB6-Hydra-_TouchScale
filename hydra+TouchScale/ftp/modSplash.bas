Attribute VB_Name = "modSplash"
Const GWL_WNDPROC = (-4)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long

Const PROP_PREVPROC = "PrevProc"
Const PROP_FORM = "FormObject"

Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
   ByVal hWnd As Long, _
   ByVal lpString As String, _
   ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
   ByVal hWnd As Long, _
   ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" ( _
   ByVal hWnd As Long, _
   ByVal lpString As String) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   Dest As Any, _
   Src As Any, _
   ByVal DestL As Long)

Const WM_PRINTCLIENT = &H318

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32" ( _
   ByVal hWnd As Long, _
   lpRect As RECT) As Long

Private Declare Function apiOleTranslateColor Lib "oleaut32" Alias "OleTranslateColor" ( _
   ByVal lOleColor As Long, _
   ByVal lHPalette As Long, _
   lColorRef As Long) As Long

Enum AnimateWindowFlags
   AW_HOR_POSITIVE = &H1
   AW_HOR_NEGATIVE = &H2
   AW_VER_POSITIVE = &H4
   AW_VER_NEGATIVE = &H8
   AW_CENTER = &H10
   AW_HIDE = &H10000
   AW_ACTIVATE = &H20000
   AW_SLIDE = &H40000
   AW_BLEND = &H80000
End Enum

Private Declare Function apiAnimateWindow Lib "user32" Alias "AnimateWindow" ( _
   ByVal hWnd As Long, _
   ByVal dwTime As Long, _
   ByVal dwFlags As Long) As Long
Private Declare Function MulDiv Lib "kernel32" ( _
   ByVal Mul As Long, _
   ByVal Nom As Long, _
   ByVal Den As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
   ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hDestDC As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, _
   ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" ( _
   ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
   ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
   ByVal hWnd As Long, _
   ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" ( _
   ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
   ByVal hDC As Long, _
   lpRect As RECT, _
   ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
   ByVal hDC As Long, _
   ByVal hObject As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
   ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

'
' AnimateWindow
'
' Wrapper for AnimateWindow api
'
Sub AnimateWindow( _
   ByVal Form As Form, _
   ByVal dwTime As Long, _
   ByVal dwFlags As AnimateWindowFlags)

   ' Set the properties
   SetProp Form.hWnd, PROP_PREVPROC, GetWindowLong(Form.hWnd, GWL_WNDPROC)
   SetProp Form.hWnd, PROP_FORM, ObjPtr(Form)

   ' Subclass the window
   SetWindowLong Form.hWnd, GWL_WNDPROC, AddressOf AnimateWinProc
   
   ' Call AnimateWindow API
   apiAnimateWindow Form.hWnd, dwTime, dwFlags
   
   ' Unsubclass the window
   SetWindowLong Form.hWnd, GWL_WNDPROC, GetProp(Form.hWnd, PROP_PREVPROC)

   ' Remove the properties
   RemoveProp Form.hWnd, PROP_FORM
   RemoveProp Form.hWnd, PROP_PREVPROC
   
   ' Refresh the form
   Form.Refresh
   
End Sub

'
' AnimateWinProc
'
' Window procedure for AnimateWindow
'
Private Function AnimateWinProc( _
   ByVal hWnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Dim lPrevProc As Long
Dim lForm As Long
Dim oForm As Form

   ' Get the previous WinProc pointer
   lPrevProc = GetProp(hWnd, PROP_PREVPROC)
   
   ' Get the form object
   lForm = GetProp(hWnd, PROP_FORM)
   MoveMemory oForm, lForm, 4&
   
   Select Case Msg
   
      Case WM_PRINTCLIENT
         Dim tRect As RECT
         Dim hBr As Long

         ' Get the window client size
         GetClientRect hWnd, tRect
         
         ' Create a brush with the
         ' form background color
         hBr = CreateSolidBrush(OleTranslateColor(oForm.BackColor))

         ' Fill the DC with the
         ' background color
         FillRect wParam, tRect, hBr

         ' Delete the brush
         DeleteObject hBr

         If Not oForm.Picture Is Nothing Then
            Dim lScrDC As Long
            Dim lMemDC As Long
            Dim lPrevBMP As Long

            ' Create a compatible DC
            lScrDC = GetDC(0&)
            lMemDC = CreateCompatibleDC(lScrDC)
            ReleaseDC 0, lScrDC

            ' Select the form picture in the DC
            lPrevBMP = SelectObject(lMemDC, oForm.Picture.Handle)

            ' Draw the picture in the DC
            BitBlt wParam, _
               0, 0, _
               HM2Pix(oForm.Picture.Width), HM2Pix(oForm.Picture.Height), _
               lMemDC, 0, 0, vbSrcCopy

            ' Release the picture
            SelectObject lMemDC, lPrevBMP

            ' Delete the DC
            DeleteDC lMemDC
            
         End If
         
   End Select

   ' Release the form object
   MoveMemory oForm, 0&, 4&

   ' Call the original window procedure
   AnimateWinProc = CallWindowProc(lPrevProc, hWnd, Msg, wParam, lParam)
   
End Function

'
' HM2Pix
'
' Converts HIMETRIC to Pixel
'
Private Function HM2Pix(ByVal Value As Long) As Long

   HM2Pix = MulDiv(Value, 1440, 2540) / Screen.TwipsPerPixelX

End Function

'
' OleTranslateColor
'
' Wrapper for OleTranslateColor API
'
Private Function OleTranslateColor(ByVal Clr As Long) As Long
   apiOleTranslateColor Clr, 0, OleTranslateColor
End Function


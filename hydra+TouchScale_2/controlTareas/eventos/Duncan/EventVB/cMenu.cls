VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2307
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_HMenu As Long
Private m_Index As Long
Private m_CommandId As Long

Private mSubmenus As colMenu

Private Declare Function GetMenuItemCountApi Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenuApi Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptX As Long, ByVal ptY As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

' ##ENUMERATION_DESCRIPTION Menu information can be retrieved by location or by unique id
Public Enum enMenuIndexMethod
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_BYCOMMAND Index by unique id
    MF_BYCOMMAND = &H0&
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_BYPOSITION Index by relative position
    MF_BYPOSITION = &H400&
End Enum

Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

' ##ENUMERATION_DESCRIPTION
Public Enum enMenuStates
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_CHECKED
    MF_CHECKED = &H8&
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_ENABLED
    MF_ENABLED = &H0&
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_HILITE
    MF_HILITE = &H80&
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_DISABLED
    MF_DISABLED = &H2&
    ' ##ENUMERATION_MEMBER_DESCRIPTION MF_GRAYED
    MF_GRAYED = &H1&
End Enum

'\\ used to determine if this menu is valid...
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long

'\\ Used to notify a window when it's menu changes...
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'\\ Handle to the menu that 'owns' the menu...
Private mHwnd As Long

Private bChanged As Boolean

' ##ENUMERATION_DESCRIPTION Describes the attributes of the menu item
Public Enum MenuItemTypes
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_STRING Displays the menu using a string
     MFT_STRING = &H0
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_BITMAP Displays the menu with a bitmap
     MFT_BITMAP = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_MENUBARBREAK The menu is seperated by a solid line
     MFT_MENUBARBREAK = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_MENUBREAK The menu is seperated but has no demarkating line
     MFT_MENUBREAK = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_OWNERDRAW The menu will be drawn by the menu owner
     MFT_OWNERDRAW = &H100
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_RADIOCHECK Menus have a radio button rather than a checkbox
     MFT_RADIOCHECK = &H200
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_SEPARATOR A horizontal seperator bar
     MFT_SEPARATOR = &H800
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_RIGHTORDER Menu cascades right-to-left
     MFT_RIGHTORDER = &H2000
     ' ##ENUMERATION_MEMBER_DESCRIPTION MFT_RIGHTJUSTIFY Menu text is right justified
     MFT_RIGHTJUSTIFY = &H4000
End Enum

' ##ENUMERATION_DESCRIPTION Describes which attributes the call to Get/SetMenuItemInfo
Public Enum MenuItemIndexes
    ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_STATE Get or set the menu state
     MIIM_STATE = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_ID Get or set the menu unique id
     MIIM_ID = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_SUBMENU Get or set the submenu handle
     MIIM_SUBMENU = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_CHECKMARKS Get or set the checkmark bitmaps
     MIIM_CHECKMARKS = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_TYPE Get or set the checkmark bitmaps
     MIIM_TYPE = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_DATA Get or set the Itemdata
     MIIM_DATA = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_STRING Get or set the menu string
     MIIM_STRING = &H40
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_BITMAP Get or set the menu picture
     MIIM_BITMAP = &H80
     ' ##ENUMERATION_MEMBER_DESCRIPTION MIIM_FTYPE Get or set the %menu type:EventVB~MenuItemTypes~EN%
     MIIM_FTYPE = &H100
End Enum

' ##ENUMERATION_DESCRIPTION These can be passed to add stock pictures to the menu item
Public Enum MenuItemStockBitmaps
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_NOTAPPLICABLE The menu does not use stock bitmaps
     HBMMENU_NOTAPPLICABLE = 0
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_CALLBACK The window that owns the menu will draw the bitmap
     HBMMENU_CALLBACK = -1
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_SYSTEM The window's icon
     HBMMENU_SYSTEM = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_MBAR_RESTORE The restore stock bitmap
     HBMMENU_MBAR_RESTORE = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_MBAR_MINIMIZE The minimise menu stock bitmap
     HBMMENU_MBAR_MINIMIZE = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_MBAR_CLOSE The close menu stock bitmap
     HBMMENU_MBAR_CLOSE = 5
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_MBAR_CLOSE_D The stock bitmap for the disabled close menu
     HBMMENU_MBAR_CLOSE_D = 6
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_MBAR_MINIMIZE_D The stock bitmap for the disabled minimise menu
     HBMMENU_MBAR_MINIMIZE_D = 7
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_POPUP_CLOSE The close menu stock bitmap
     HBMMENU_POPUP_CLOSE = 8
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_POPUP_RESTORE The restore menu stock bitmap
     HBMMENU_POPUP_RESTORE = 9
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_POPUP_MAXIMIZE The maximise menu stock bitmap
     HBMMENU_POPUP_MAXIMIZE = 10
     ' ##ENUMERATION_MEMBER_DESCRIPTION HBMMENU_POPUP_MINIMIZE The minimise menu stock bitmap
     HBMMENU_POPUP_MINIMIZE = 11
End Enum

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    hbmpItem As Long
End Type
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal Index As Long, ByVal ByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal Index As Long, ByVal ByPosition As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private mMenuItemInfo As MENUITEMINFO


Private mMenuId As Long

Public Property Set Bitmap(ByVal newbmp As ApiBitmap)
Attribute Bitmap.VB_Description = "The picture used in place of a caption on the menu item  "
Attribute Bitmap.VB_HelpID = 3043

Call GetInfo(MIIM_BITMAP)
If Not newbmp Is Nothing Then
    mMenuItemInfo.hbmpItem = newbmp.hBitmap
Else
    mMenuItemInfo.hbmpItem = 0
End If

Call SetInfo

End Property

' ##BLOCK_DESCRIPTION The picture used in place of a caption on the menu item
' ##BLOCK_REMARKS This may not always be set
Public Property Get Bitmap() As ApiBitmap

Call GetInfo(MIIM_BITMAP)
Dim bmpThis As ApiBitmap

If mMenuItemInfo.hbmpItem > HBMMENU_POPUP_MINIMIZE Or mMenuItemInfo.hbmpItem < -1 Then
    Set bmpThis = New ApiBitmap
    bmpThis.hBitmap = mMenuItemInfo.hbmpItem
    Set Bitmap = bmpThis
End If

End Property

Public Property Let Caption(ByVal newCaption As String)
Attribute Caption.VB_HelpID = 3044


Dim lret As Long

Call GetInfo(MIIM_STRING)

With mMenuItemInfo
    .fMask = MIIM_STRING
    .fType = MFT_STRING
    .dwTypeData = newCaption
    .cch = Len(newCaption)
End With

Call SetInfo


End Property

Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
    '\\ Refresh the menu bar displayed
    Call DrawMenuBar(mHwnd)
End Property


Public Property Get Caption() As String

Dim sRet As String
Dim lret As Long

sRet = String$(1024, 0)

lret = GetMenuString(m_HMenu, mMenuId, sRet, 1024, MF_BYCOMMAND)
If lret = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:Caption", GetLastSystemError)
End If
If lret > 0 Then
    sRet = Left$(sRet, lret)
End If
Caption = sRet

End Property

Public Property Get CaptionFromPoint(ByVal Parent As ApiWindow, ByVal Point As APIPoint) As String
Attribute CaptionFromPoint.VB_HelpID = 3045

Dim sRet As String
Dim lret As Long

lret = IndexFromPoint(Parent, Point)
If lret > -1 Then
    sRet = String$(1024, 0)
    lret = GetMenuString(m_HMenu, lret, sRet, 1024, MF_BYPOSITION)
    If lret > 0 Then
        sRet = Left$(sRet, lret)
    End If
End If
CaptionFromPoint = sRet

End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_HelpID = 2310

    Checked = (MenuState(mMenuId, MF_BYCOMMAND) Or MF_CHECKED)

End Property


Public Property Get CheckedBitmap() As ApiBitmap
Attribute CheckedBitmap.VB_HelpID = 3046

Call GetInfo(MIIM_BITMAP)
Dim bmpThis As ApiBitmap

If mMenuItemInfo.hbmpChecked > HBMMENU_POPUP_MINIMIZE Or mMenuItemInfo.hbmpChecked < -1 Then
    Set bmpThis = New ApiBitmap
    bmpThis.hBitmap = mMenuItemInfo.hbmpChecked
    Set CheckedBitmap = bmpThis
End If

End Property

Public Property Set CheckedBitmap(ByVal newbmp As ApiBitmap)

Call GetInfo(MIIM_BITMAP)
If Not newbmp Is Nothing Then
    mMenuItemInfo.hbmpChecked = newbmp.hBitmap
Else
    mMenuItemInfo.hbmpChecked = 0
End If

Call SetInfo

End Property

Private Sub GetInfo(ByVal Index As MenuItemIndexes)

Dim lret

With mMenuItemInfo
    .cbSize = Len(mMenuItemInfo)
    .fMask = Index
End With

lret = GetMenuItemInfo(m_HMenu, mMenuId, False, mMenuItemInfo)

End Sub

Public Property Get hMenu() As Long
Attribute hMenu.VB_HelpID = 2311

    hMenu = m_HMenu
    
End Property

Public Property Let hMenu(ByVal newHandle As Long)

    If newHandle <> m_HMenu Then
        m_HMenu = newHandle
    End If
    
End Property
Public Property Get Grayed() As Boolean
Attribute Grayed.VB_HelpID = 2312

    Grayed = (MenuState(mMenuId, MF_BYCOMMAND) Or MF_GRAYED)
    
End Property

Friend Property Let id(ByVal newId As Long)

mMenuId = newId

End Property

Public Property Get id() As Long
Attribute id.VB_HelpID = 3048

id = mMenuId

End Property

Public Property Let Index(ByVal newIndex As Long)
Attribute Index.VB_HelpID = 2856

    m_Index = newIndex
    mMenuId = GetMenuItemID(m_HMenu, m_Index)

End Property

Public Property Get IndexFromPoint(ByVal Parent As ApiWindow, ByVal Point As APIPoint) As Long
Attribute IndexFromPoint.VB_HelpID = 2313

Dim lret As Long
Dim x As Long, y As Long

With Point
    x = .x
    y = .y
End With

lret = MenuItemFromPoint(Parent.hwnd, m_HMenu, x, y)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:IndexFromPoint", GetLastSystemError)
    IndexFromPoint = -1
Else
    IndexFromPoint = lret
End If

End Property

Public Property Let ItemData(ByVal newData As Long)
Attribute ItemData.VB_HelpID = 3050

Call GetInfo(MIIM_DATA)
mMenuItemInfo.dwItemData = newData
Call SetInfo

End Property

Public Property Get ItemData() As Long

Call GetInfo(MIIM_DATA)
ItemData = mMenuItemInfo.dwItemData

End Property

Public Property Let OwnerDrawn(ByVal newval As Boolean)
Attribute OwnerDrawn.VB_Description = "The menu is drawn by it's owner window  "
Attribute OwnerDrawn.VB_HelpID = 3052

Call GetInfo(MIIM_TYPE)
If newval Then
    mMenuItemInfo.fType = mMenuItemInfo.fType Or MFT_OWNERDRAW
Else
    mMenuItemInfo.fType = mMenuItemInfo.fType Xor MFT_OWNERDRAW
End If


End Property

' ##BLOCK_DESCRIPTION The menu is drawn by it's owner window
' ##BLOCK_REMARKS Set this to true to trigger a DrawOwnerDrawnMenuItem event
Public Property Get OwnerDrawn() As Boolean

Call GetInfo(MIIM_TYPE)
If mMenuItemInfo.fType And MFT_OWNERDRAW Then
    OwnerDrawn = True
Else
    OwnerDrawn = False
End If

End Property

Public Property Get Popup(ByVal nPos As Integer) As Boolean
Attribute Popup.VB_HelpID = 2314

    Popup = (ItemId(nPos) = -1)

End Property



Public Property Get Separator(ByVal nPos As Long) As Boolean
Attribute Separator.VB_HelpID = 2315

    Separator = (ItemId(nPos) = 0)

End Property

Public Property Get ItemCount() As Long
Attribute ItemCount.VB_HelpID = 2316

Dim lret As Long

lret = GetMenuItemCount(m_HMenu)
If (lret = -1) Or (Err.LastDllError > 0) Then
    Call ReportError(Err.LastDllError, "ApiMenu:ItemCount", GetLastSystemError)
End If

ItemCount = lret

End Property

Public Property Get ItemId(ByVal nPos As Long) As Long
Attribute ItemId.VB_HelpID = 2317

    Dim lret As Long
    
    lret = GetMenuItemID(m_HMenu, nPos)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiMenu:ItemId", GetLastSystemError)
    End If
    
    ItemId = lret
    
End Property


Public Property Get Key() As String
Attribute Key.VB_HelpID = 2318

Key = "HMENU:" & m_HMenu

End Property





Public Property Get MenuState(ByVal nIndex As Long, ByVal IndexMethod As enMenuIndexMethod) As Long
Attribute MenuState.VB_HelpID = 2321

Dim lret As Long

lret = GetMenuState(m_HMenu, nIndex, IndexMethod)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:MenuState", GetLastSystemError)
End If
MenuState = lret

End Property


Private Sub SetInfo()

Dim lret As Long

mMenuItemInfo.cbSize = Len(mMenuItemInfo)

lret = SetMenuItemInfo(m_HMenu, mMenuId, False, mMenuItemInfo)
Changed = True

End Sub

Public Property Let StockPicture(ByVal PictureIndex As MenuItemStockBitmaps)
Attribute StockPicture.VB_HelpID = 3054

Call GetInfo(MIIM_BITMAP)

mMenuItemInfo.hbmpItem = PictureIndex

Call SetInfo

End Property

Public Property Get StockPicture() As MenuItemStockBitmaps

Call GetInfo(MIIM_BITMAP)

If mMenuItemInfo.hbmpItem > HBMMENU_POPUP_MINIMIZE Or mMenuItemInfo.hbmpItem < -1 Then
    StockPicture = HBMMENU_NOTAPPLICABLE
Else
    StockPicture = mMenuItemInfo.hbmpItem
End If

End Property

Public Property Get Submenus() As colMenu
Attribute Submenus.VB_HelpID = 2322

Dim lSubmenus As Long
Dim lThismenu As Long
Dim hThismenu As Long

Dim mnuItem As ApiMenu

Set mSubmenus = New colMenu

lSubmenus = GetMenuItemCountApi(m_HMenu)
If lSubmenus > 0 Then
    For lThismenu = 0 To (lSubmenus - 1)
        hThismenu = GetSubMenuApi(m_HMenu, lThismenu)
        If hThismenu <> 0 And Err.LastDllError = 0 Then
            Set mnuItem = mSubmenus.AddInternal(hThismenu, mHwnd, lThismenu)
        End If
    Next lThismenu
End If

Set Submenus = mSubmenus

End Property


Public Property Set UncheckedBitmap(ByVal newbmp As ApiBitmap)
Attribute UncheckedBitmap.VB_HelpID = 3056

Call GetInfo(MIIM_BITMAP)
If Not newbmp Is Nothing Then
    mMenuItemInfo.hbmpUnchecked = newbmp.hBitmap
Else
    mMenuItemInfo.hbmpUnchecked = 0
End If

Call SetInfo

End Property

Public Property Get UncheckedBitmap() As ApiBitmap


Call GetInfo(MIIM_BITMAP)
Dim bmpThis As ApiBitmap

If mMenuItemInfo.hbmpUnchecked > HBMMENU_POPUP_MINIMIZE Or mMenuItemInfo.hbmpUnchecked < -1 Then
    Set bmpThis = New ApiBitmap
    bmpThis.hBitmap = mMenuItemInfo.hbmpUnchecked
    Set UncheckedBitmap = bmpThis
End If


End Property

Public Property Get Valid() As Boolean
Attribute Valid.VB_HelpID = 2323
Dim lret As Long

lret = IsMenu(m_HMenu)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:Valid", GetLastSystemError)
Else
    Valid = (lret <> 0)
End If

End Property

Friend Property Let WindowHandle(ByVal hwndOwner As Long)

If hwndOwner <> mHwnd Then
    mHwnd = hwndOwner
    Changed = True
End If

End Property

Private Sub Class_Terminate()

Set mSubmenus = Nothing

End Sub



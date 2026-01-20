VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2208
Attribute VB_Description = "This class provides the properties, events and methods used to manipulate any given window."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ##MODULE_DESCRIPTION This class provides the properties, events and methods used _
to manipulate any given window.

' ##MODULE_DESCRIPTION Using this class and the %subclassed windows collection:EventVB~APIFunctions~SubclassedWindows% _
provides you with access to a large number of new events in addition to those that are _
exposed by Visual basic forms and also allows you to respond to events in windows that _
were not created by Visual Basic using the familiar syntax.

Private Const CLSNAME = "ApiWindow"

'\\ New window events --------------------------------------------------------------------------------------------------------------------------------------------------------------
' ##EVENT_DESCRIPTION WindowMessageFired Triggered for any message not dealt with by the other more specific events
' ##EVENT_REMARKS WindowMessageFired
Public Event WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)
Attribute WindowMessageFired.VB_Description = "WindowMessageFired Triggered for any message not dealt with by the other more specific events"
Attribute WindowMessageFired.VB_HelpID = 2209

' ##EVENT_DESCRIPTION ActiveApplicationChanged Triggered when the subclassed window is a top level window of an application and the user swicthes to or from that application (using ALT+TAB or the mouse)
' ##EVENT_REMARKS ActiveApplicationChanged This event is only fired if the window is in the thread losing the focus or the thread gaining the focus
Public Event ActiveApplicationChanged(ByVal ActivatingThisApp As Boolean, ByVal hThread As Long, Cancel As Boolean)
Attribute ActiveApplicationChanged.VB_Description = "ActiveApplicationChanged Triggered when the subclassed window is a top level window of an application and the user swicthes to or from that application (using ALT+TAB or the mouse)"
Attribute ActiveApplicationChanged.VB_HelpID = 2210

' ##EVENT_DESCRIPTION LostCapture Triggered when the subclassed window loses the mouse capture. _
In normal operation this occurs when the mouse leaves the window.
' ##EVENT_REMARKS LostCapture This event will not be triggered if a window has explicitly captured the mouse input.
Public Event LostCapture(ByVal hwndNewCapture As Long, Cancel As Boolean)
Attribute LostCapture.VB_Description = "LostCapture Triggered when the subclassed window loses the mouse capture. In normal operation this occurs when the mouse leaves the window."
Attribute LostCapture.VB_HelpID = 2211

' ##EVENT_DESCRIPTION KeyPressed Triggered when any key is pressed whle the subclassed window has the input focus.
' ##EVENT_REMARKS KeyPressed This event provides additional information beyond the intrinsic form_Keypressed event.
Public Event KeyPressed(ByVal VKey As Long, ByVal Repetition As Long, ByVal scanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal AlreadyDown As Boolean, ByVal BeingPressed As Boolean, Cancel As Boolean)
Attribute KeyPressed.VB_Description = "KeyPressed Triggered when any key is pressed whle the subclassed window has the input focus."
Attribute KeyPressed.VB_HelpID = 2212

' ##EVENT_DESCRIPTION LowMemory Triggered when the available memory on a system is low and the system is required to try and compact the memory to free some space.
' ##EVENT_REMARKS LowMemory The TimeSpentCompacting increases as the memory situation gets worse.  You should free up any possible memory you can _
in response to this event.
Public Event LowMemory(ByVal TimeSpentCompacting As Long)
Attribute LowMemory.VB_Description = "LowMemory Triggered when the available memory on a system is low and the system is required to try and compact the memory to free some space."
Attribute LowMemory.VB_HelpID = 2213

' ##EVENT_DESCRIPTION Move Triggered when a form is moved, by dragging or through code.
' ##EVENT_REMARKS Move The x and y parameters indicate the position of the top left corner of the window.
Public Event Move(ByVal x As Long, ByVal y As Long, Cancel As Boolean)
Attribute Move.VB_Description = "Move Triggered when a form is moved, by dragging or through code."
Attribute Move.VB_HelpID = 3179

' ##EVENT_DESCRIPTION VerticalScroll Triggered when the user interacts with a vertical scrollbar on a window.
' ##EVENT_REMARKS VerticalScroll The message parameter tells you what operation the user performed.
' ##EVENT_REMARKS VerticalScroll Changing the Position parameter changes the location of the scrollbar
Public Event VerticalScroll(ByVal Message As enScrollMessages, Position As Long, Cancel As Boolean)
Attribute VerticalScroll.VB_Description = "VerticalScroll Triggered when the user interacts with a vertical scrollbar on a window."
Attribute VerticalScroll.VB_HelpID = 3013

' ##EVENT_DESCRIPTION HorizontalScroll Triggered when the user interacts with a horizontal scrollbar on a window.
' ##EVENT_REMARKS HorizontalScroll The message parameter tells you what operation the user performed.
' ##EVENT_REMARKS HorizontalScroll Changing the Position parameter changes the location of the scrollbar
Public Event HorizontalScroll(ByVal Message As enScrollMessages, Position As Long, Cancel As Boolean)
Attribute HorizontalScroll.VB_Description = "HorizontalScroll Triggered when the user interacts with a horizontal scrollbar on a window."
Attribute HorizontalScroll.VB_HelpID = 3014

' ##EVENT_DESCRIPTION WindowsSettingsChanged Triggered when the system settings changed.
' ##EVENT_REMARKS WindowsSettingsChanged newSetting describes which type of setting changed.
Public Event WindowsSettingsChanged(ByVal newSetting As enSystemParametersInfo)
Attribute WindowsSettingsChanged.VB_Description = "WindowsSettingsChanged Triggered when the system settings changed."
Attribute WindowsSettingsChanged.VB_HelpID = 2217

' ##EVENT_DESCRIPTION WindowsINIChanged Triggered when a setting stored in the WIN.INI is changed.
' ##EVENT_REMARKS WindowsINIChanged Section is the major section e.g. [INTL] that has been changed
Public Event WindowsINIChanged(ByVal Section As String)
Attribute WindowsINIChanged.VB_Description = "WindowsINIChanged Triggered when a setting stored in the WIN.INI is changed."
Attribute WindowsINIChanged.VB_HelpID = 2218

' ##EVENT_DESCRIPTION NonClientMouseMove Triggered when the mouse is moved over the non client area of the window.
' ##EVENT_REMARKS NonClientMouseMove The non client area includes the caption bar, menu bar etc.
Public Event NonClientMouseMove(ByVal Location As enHitTestResult, ByVal x As Single, ByVal y As Single)
Attribute NonClientMouseMove.VB_Description = "NonClientMouseMove Triggered when the mouse is moved over the non client area of the window."
Attribute NonClientMouseMove.VB_HelpID = 2219

' ##EVENT_DESCRIPTION NonClientMouseDown Triggered when the mouse button is pressed while the cursor is over a non client part of the window.
' ##EVENT_REMARKS NonClientMouseDown The non client area includes the caption bar, menu bar etc.
Public Event NonClientMouseDown(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Attribute NonClientMouseDown.VB_Description = "NonClientMouseDown Triggered when the mouse button is pressed while the cursor is over a non client part of the window."
Attribute NonClientMouseDown.VB_HelpID = 2220

' ##EVENT_DESCRIPTION NonClientMouseUp Triggered when the mouse button is released while the cursor is over a non client part of the window.
' ##EVENT_REMARKS NonClientMouseUp The non client area includes the caption bar, menu bar etc.
Public Event NonClientMouseUp(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Attribute NonClientMouseUp.VB_Description = "NonClientMouseUp Triggered when the mouse button is released while the cursor is over a non client part of the window."
Attribute NonClientMouseUp.VB_HelpID = 2221

' ##EVENT_DESCRIPTION NonClientDblClick Triggered when the mouse button is double-clicked while the cursor is over a non client part of the window.
' ##EVENT_REMARKS NonClientDblClick The non client area includes the caption bar, menu bar etc.
Public Event NonClientDblClick(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Attribute NonClientDblClick.VB_Description = "NonClientDblClick Triggered when the mouse button is double-clicked while the cursor is over a non client part of the window."
Attribute NonClientDblClick.VB_HelpID = 2222

' ##EVENT_DESCRIPTION MinMaxSize Triggered when the operating system wishes to know the minimum and maximum sizes of the window.
' ##EVENT_REMARKS MinMaxSize Changing any of the min or max values will prevent the window being resized outside of that size. _
This is very useful for creating docking style windows.
Public Event MinMaxSize(MaxHeight As Long, MaxWidth As Long, MaxPositionTop As Long, MaxPositionLeft As Long, MinTrackWidth As Long, MinTrackheight As Long, MaxTrackWidth As Long, MaxTrackHeight As Long)
Attribute MinMaxSize.VB_Description = "MinMaxSize Triggered when the operating system wishes to know the minimum and maximum sizes of the window."
Attribute MinMaxSize.VB_HelpID = 2223

' ##EVENT_DESCRIPTION MouseOverMenu Triggered when the mouse is moved over a top level menu bar item.
' ##EVENT_REMARKS MouseOverMenu You can use this event to provide feedback to the user about what a given menu is used for.
Public Event MouseOverMenu(ByVal Caption As String)
Attribute MouseOverMenu.VB_Description = "MouseOverMenu Triggered when the mouse is moved over a top level menu bar item."
Attribute MouseOverMenu.VB_HelpID = 2224

' ##EVENT_DESCRIPTION Hotkey Triggered when a %hotkey:EventVB~ApiHotkey% that has been registered for this window is pressed.
' ##EVENT_REMARKS Hotkey Currently a window can only have one hotkey registered to it at any ti
Public Event Hotkey(ByVal Hotkey As ApiHotkey)
Attribute Hotkey.VB_Description = "Hotkey Triggered when a hotkey:EventVB~ApiHotkey that has been registered for this window is pressed."
Attribute Hotkey.VB_HelpID = 2225

' ##EVENT_DESCRIPTION Paint Triggered when all or part of the form needs to be repainted.
' ##EVENT_REMARKS Paint Use this event for owner drawn controls, for example.
' ##EVENT_REMARKS Set Override = True to prevent default processing
' ##EVENT_REMARKS If UpdateRect is not set, do NOT call BeginPaint/EndPaint...
Public Event Paint(ByVal UpdateRect As APIRect, Override As Boolean)
Attribute Paint.VB_Description = "Paint Triggered when all or part of the form needs to be repainted."
Attribute Paint.VB_HelpID = 3015

' ##EVENT_DESCRIPTION NonClientPaint Triggered when all or part of the non client area needs to be repainted.
' ##EVENT_REMARKS NonClientPaint This event can be used to implement customised non client areas - like skins.
' ##EVENT_REMARKS NonClientPaint Setting the Override parameter to true prevents drawing the default window non client area
Public Event NonClientPaint(ByVal NonClientDeviceContext As ApiDeviceContext, Override As Boolean)
Attribute NonClientPaint.VB_Description = "NonClientPaint Triggered when all or part of the non client area needs to be repainted."
Attribute NonClientPaint.VB_HelpID = 2830

' ##EVENT_DESCRIPTION HitTest Triggered when windows wants to know what part of the form the mouse is over (client, caption., menu etc.)
' ##EVENT_REMARKS HitTest Overriding this event alters the way windows behaves.  Using this you can allow a window with no border to be _
moved or resized, for example.
Public Event HitTest(ByVal x As Long, ByVal y As Long, ReturnValue As enHitTestResult, Override As Boolean)
Attribute HitTest.VB_Description = "HitTest Triggered when windows wants to know what part of the form the mouse is over (client, caption., menu etc.)"
Attribute HitTest.VB_HelpID = 3176

' ##EVENT_DESCRIPTION Sizing event is triggered when the window is being drag-resized
' ##EVENT_REMARKS SizeEdge indicates which edge/grab handle (or combination) is being dragged
' ##EVENT_REMARKS DragRectangle is the current size of the resizing rectangle.  You _
can alter the dimensions of this rectangle in this event to control how the window is resized.
' ##EVENT_REMARKS This is useful for tool windows, snap-to-edge and other windows that must _
resize in a managed manner
Public Event Sizing(ByVal SizeEdges As WindowSizingEdges, DragRectangle As APIRect)
Attribute Sizing.VB_Description = "Sizing event is triggered when the window is being drag-resized"
Attribute Sizing.VB_HelpID = 2229

' ##EVENT_DESCRIPTION Moving event is triggered when the window is being drag-moved
' ##EVENT_REMARKS MoveEdges indicates which edge/grab handle (or combination) is being dragged
' ##EVENT_REMARKS MoveRectangle is the current size of the moving rectangle.  You _
can alter the dimensions of this rectangle in this event to control how the window is resized.
' ##EVENT_REMARKS This is useful for tool windows, snap-to-edge and other windows that must _
be moved in a managed manner
Public Event Moving(ByVal MoveEdges As WindowSizingEdges, MoveRectangle As APIRect)
Attribute Moving.VB_Description = "Moving event is triggered when the window is being drag-moved"
Attribute Moving.VB_HelpID = 2230

' ##EVENT_DESCRIPTION PowerManagement is triggered when a power management hardware event occurs
' ##EVENT_REMARKS Message is the type of power management event that occured
' ##EVENT_REMARKS Setting Cancel to true prevents the event proceeding
' ##EVENT_REMARKS Use this message with caution
Public Event PowerManagement(ByVal Message As PowerBroadcastMessageTypes, Cancel As Boolean)
Attribute PowerManagement.VB_Description = "PowerManagement is triggered when a power management hardware event occurs"
Attribute PowerManagement.VB_HelpID = 2231

' ##EVENT_DESCRIPTION MouseHover event is triggered when the mouse pointer hovers over the window for _
and amount of time specified by the %timeout:EventVB~ApiWindow~MouseHoverTimeout% property
' ##EVENT_REMARKS If the the %timeout:EventVB~ApiWindow~MouseHoverTimeout% property is zero this event will not fire
' ##EVENT_REMARKS The properties of this event are the same as those provided by the VB MouseMove event
' ##EVENT_REMARKS This event is only triggered on Win98/NT4 or higher
Public Event MouseHover(Button As Integer, Shift As Integer, x As Long, y As Long)
Attribute MouseHover.VB_Description = "MouseHover event is triggered when the mouse pointer hovers over the window for and amount of time specified by the timeout:EventVB~ApiWindow~MouseHoverTimeout property"
Attribute MouseHover.VB_HelpID = 2232

' ##EVENT_DESCRIPTION MouseLeave event is triggered when the mouse pointer leaves the window
' ##EVENT_REMARKS This event is only triggered on Win98/NT4 or higher
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "MouseLeave event is triggered when the mouse pointer leaves the window"
Attribute MouseLeave.VB_HelpID = 2233

' ##EVENT_DESCRIPTION ContextMenu is triggered when the user right-clicks or invokes the context menu through the keyboard
' ##EVENT_REMARKS x and y describe the mouse position when the event is triggered
' ##EVENT_REMARKS If x and y are -1 then the context menu was invoked by the keyboard
' ##EVENT_REMARKS Setting Override to true prevents the window from showing its default context menu (if any)
Public Event ContextMenu(ByVal x As Long, ByVal y As Long, Override As Boolean)
Attribute ContextMenu.VB_Description = "ContextMenu is triggered when the user right-clicks or invokes the context menu through the keyboard"
Attribute ContextMenu.VB_HelpID = 2831

' ##EVENT_DESCRIPTION ControlColour is triggered when a control is painting it's background
' ##EVENT_REMARKS Changing the colourbrush means that the contol will be drawn using that given brush as its background
Public Event ControlColour(ByVal ControlType As StandardControlTypes, ByVal DeviceContext As ApiDeviceContext, ByVal ControlWindow As ApiWindow, ColourBrush As ApiLogBrush)
Attribute ControlColour.VB_Description = "ControlColour is triggered when a control is painting it's background"
Attribute ControlColour.VB_HelpID = 2832

'\\ OWNER DRAWN CONTROL MESSAGES ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' ##EVENT_DESCRIPTION DrawOwnerDrawnButton
' ##EVENT_REMARKS ControlId is the unique identifier of the button control that needs painting
Public Event DrawOwnerDrawnButton(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal Default As Boolean, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoAcceleratorMnemonics As Boolean, ByVal NoFocusRect As Boolean, ByVal ButtonWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect)
Attribute DrawOwnerDrawnButton.VB_Description = "DrawOwnerDrawnButton"
Attribute DrawOwnerDrawnButton.VB_HelpID = 2833

' ##EVENT_DESCRIPTION DrawOwnerDrawnComboItem - Fired when a combo box with the owner-drawn style needs repainting
' ##EVENT_REMARKS ControlId is the unique identifier of the combo box control that needs painting
Public Event DrawOwnerDrawnComboItem(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal ItemIndex As Long, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal Combowindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnComboItem.VB_Description = "DrawOwnerDrawnComboItem"
Attribute DrawOwnerDrawnComboItem.VB_HelpID = 2834

' ##EVENT_DESCRIPTION DrawOwnerDrawnListboxItem - Fired when an owner drawn menu needs drawing
' ##EVENT_REMARKS ControlId is the unique identifier of the control that needs painting
Public Event DrawOwnerDrawnListboxItem(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal ItemIndex As Long, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal ListboxWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnListboxItem.VB_Description = "DrawOwnerDrawnListboxItem - Fired when an owner drawn menu needs drawing"
Attribute DrawOwnerDrawnListboxItem.VB_HelpID = 2835

' ##EVENT_DESCRIPTION DrawOwnerDrawnStatic - Fired when an owner drawn static control needs drawing
' ##EVENT_REMARKS ControlId is the unique identifier of the static control that needs painting
Public Event DrawOwnerDrawnStatic(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal StaticWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect)
Attribute DrawOwnerDrawnStatic.VB_Description = "DrawOwnerDrawnStatic - Fired when an owner drawn static control needs drawing"
Attribute DrawOwnerDrawnStatic.VB_HelpID = 2836

' ##EVENT_DESCRIPTION DrawOwnerDrawnTab - Fired when an owner drawn tab control needs drawing
' ##EVENT_REMARKS ControlId is the unique identifier of the tab control that needs painting
Public Event DrawOwnerDrawnTab(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal ItemIndex As Long, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal TabWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnTab.VB_Description = "DrawOwnerDrawnTab - Fired when an owner drawn tab control needs drawing"
Attribute DrawOwnerDrawnTab.VB_HelpID = 2837

' ##EVENT_DESCRIPTION DrawOwnerDrawnHeader - Fired when an owner drawn header needs drawing
' ##EVENT_REMARKS ControlId is the unique identifier of the header control that needs painting
Public Event DrawOwnerDrawnHeader(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal ItemIndex As Long, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal HeaderWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnHeader.VB_Description = "DrawOwnerDrawnHeader - Fired when an owner drawn header needs drawing"
Attribute DrawOwnerDrawnHeader.VB_HelpID = 2838

' ##EVENT_DESCRIPTION DrawOwnerDrawnListviewItem - Fired when an owner drawn listview item needs drawing
' ##EVENT_REMARKS ControlId is the unique identifier of the list view control that needs painting
Public Event DrawOwnerDrawnListviewItem(ByVal ControlId As Long, ByVal DrawAction As OwnerDrawActions, ByVal ItemIndex As Long, ByVal Disabled As Boolean, ByVal Focused As Boolean, ByVal HotTracked As Boolean, ByVal NoFocusRect As Boolean, ByVal ListviewWindow As ApiWindow, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnListviewItem.VB_Description = "DrawOwnerDrawnListviewItem - Fired when an owner drawn listview item needs drawing"
Attribute DrawOwnerDrawnListviewItem.VB_HelpID = 2839

' ##EVENT_DESCRIPTION DrawOwnerDrawnMenuItem - Fired when an owner drawn menu needs drawing
Public Event DrawOwnerDrawnMenuItem(ByVal Index As Long, ByVal DrawAction As OwnerDrawActions, ByVal Checked As Boolean, ByVal Greyed As Boolean, ByVal Selected As Boolean, ByVal Menu As ApiMenu, ByVal ItemDeviceContext As ApiDeviceContext, ByVal ItemRect As APIRect, ByVal ItemData As Long)
Attribute DrawOwnerDrawnMenuItem.VB_Description = "DrawOwnerDrawnMenuItem - Fired when an owner drawn menu needs drawing"
Attribute DrawOwnerDrawnMenuItem.VB_HelpID = 2840

'\\ Device management Events - NT4/Win95 or beyond ------------------------
' ##EVENT_DESCRIPTION DeviceManagementDiskFull - The specified disk is full
' ##EVENT_REMARKS IFS manager has detected that a drive has run out of space
' ##EVENT_REMARKS DiskIndex is 1 based
Public Event DeviceManagementDiskFull(ByVal DiskIndex As Long)
Attribute DeviceManagementDiskFull.VB_Description = "DeviceManagementDiskFull - The specified disk is full"
Attribute DeviceManagementDiskFull.VB_HelpID = 2841

' ##EVENT_DESCRIPTION DeviceManagementDiskLow - The specified disk is low on disk space
' ##EVENT_REMARKS A VFAT drive has less space than the registry setting
' ##EVENT_REMARKS DiskIndex is 1 based
Public Event DeviceManagementDiskLow(ByVal DiskIndex As Long)
Attribute DeviceManagementDiskLow.VB_Description = "DeviceManagementDiskLow - The specified disk is low on disk space"
Attribute DeviceManagementDiskLow.VB_HelpID = 2842

' ##EVENT_DESCRIPTION DeviceManagementMonitorChanged - The display monitor has changed
' ##EVENT_REMARKS xResolution is the number of pixels across the monitor
' ##EVENT_REMARKS yResolution is the number of pixels down the monitor
Public Event DeviceManagementMonitorChanged(ByVal xResolution As Integer, ByVal yResolution As Integer)
Attribute DeviceManagementMonitorChanged.VB_Description = "DeviceManagementMonitorChanged - The display monitor has changed"
Attribute DeviceManagementMonitorChanged.VB_HelpID = 2843

' ##EVENT_DESCRIPTION MouseWheelMove - The mouse wheel (on the new type of mouse) was moved
' ##EVENT_REMARKS Rotation - The amount the wheeel moved (positive is forward, negative is back)
' ##EVENT_REMARKS The other properties of this event are the same as those provided by the VB MouseMove event
Public Event MouseWheelMove(ByVal Rotation As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Attribute MouseWheelMove.VB_Description = "MouseWheelMove - The mouse wheel (on the new type of mouse) was moved"
Attribute MouseWheelMove.VB_HelpID = 2844

' ##EVENT_DESCRIPTION PaintWindow - The window has recieved a request to draw itself on a third party device context
' ##EVENT_REMARKS TargetDevice - The device context (usually another window) that this window should be drawn onto
' ##EVENT_REMARKS DrawOptions - The way to draw the window
Public Event PaintWindow(ByVal TargetDevice As ApiDeviceContext, ByVal DrawOptions As PrintClientFlags)
Attribute PaintWindow.VB_Description = "PaintWindow - The window has recieved a request to draw itself on a third party device context"
Attribute PaintWindow.VB_HelpID = 2845

' ##EVENT_DESCRIPTION EraseBackground - Called when a window needs to fill its background prior to repainting
' ##EVENT_REMARKS EraseDeviceContext - The device context that needs to be erased
' ##EVENT_REMARKS Override - Set to true if you want to erase the background and not pass the message on to the default window procedure
' ##EVENT_REMARKS This can be used to paint a cutom background for a window
Public Event EraseBackground(ByVal EraseDeviceContext As ApiDeviceContext, Override As Boolean)
Attribute EraseBackground.VB_Description = "EraseBackground - Called when a window needs to fill its background prior to repainting"
Attribute EraseBackground.VB_HelpID = 3016

'-- MENU MESSAGES -------------------------------------------------------------------------

' ##EVENT_DESCRIPTION MenuClick - A menu item on this window's menu was selected
' ##EVENT_REMARKS MenuItem - The particular menu chosen
Public Event MenuClick(ByVal MenuItem As ApiMenu)
Attribute MenuClick.VB_Description = "MenuClick - A menu item on this window's menu was selected"
Attribute MenuClick.VB_HelpID = 3017


' --CONTEXT SENSITIVE HELP ----------------------------------------------------------------

' ##EVENT_DESCRIPTION ContextSensitiveMenuHelp - The user clicked on a menu while in context sensitive help mode

' ##EVENT_REMARKS HelpId - The context sensitive help id for the menu that was clicked on
' ##EVENT_REMARKS x - The x position of the mouse click
' ##EVENT_REMARKS x - The y position of the mouse click
Public Event ContextSensitiveMenuHelp(ByVal MenuItem As ApiMenu, ByVal HelpId As Long, ByVal x As Long, ByVal y As Long)
Attribute ContextSensitiveMenuHelp.VB_Description = "ContextSensitiveMenuHelp - The user clicked on a menu while in context sensitive help mode"
Attribute ContextSensitiveMenuHelp.VB_HelpID = 3018

' ##EVENT_DESCRIPTION ContextSensitiveControlHelp - The user clicked on a control while in context sensitive help mode
' ##EVENT_REMARKS ControlWindow - The window of the control that was clicked on
' ##EVENT_REMARKS HelpId - The context sensitive help id for the control that was clicked on
' ##EVENT_REMARKS x - The x position of the mouse click
' ##EVENT_REMARKS x - The y position of the mouse click
Public Event ContextSensitiveControlHelp(ByVal ControlWindow As ApiWindow, ByVal HelpId As Long, ByVal x As Long, ByVal y As Long)
Attribute ContextSensitiveControlHelp.VB_Description = "ContextSensitiveControlHelp - The user clicked on a control while in context sensitive help mode"
Attribute ContextSensitiveControlHelp.VB_HelpID = 3019

'\\ -- CHILD WINDOW NOTIFICATIONS -------------------------------------------------------
' ##EVENT_DESCRIPTION ChildButtonNotification - A button control on this window has notified this window of an event occuring to it
' ##EVENT_REMARKS Notification - The event that happened
' ##EVENT_REMARKS ButtonId - The unique control identifier of the button that sent this notification
' ##EVENT_REMARKS ButtonWindow - The window of the button control that sent the notification
Public Event ChildButtonNotification(ByVal Notification As ButtonParentNotifications, ByVal ButtonId As Integer, ByVal ButtonWindow As ApiWindow)

' ##EVENT_DESCRIPTION ChildEditNotification - A textbox control on this window has notified this window of an event occuring to it
' ##EVENT_REMARKS Notification - The event that happened
' ##EVENT_REMARKS EditId - The unique control identifier of the textbox that sent this notification
' ##EVENT_REMARKS EditWindow - The window of the textbox control that sent the notification
Public Event ChildEditNotification(ByVal Notification As EditBoxParentNotifications, ByVal EditId As Integer, ByVal EditWindow As ApiWindow)

' ##EVENT_DESCRIPTION ChildListboxNotification - A listbox control on this window has notified this window of an event occuring to it
' ##EVENT_REMARKS Notification - The event that happened
' ##EVENT_REMARKS ListboxId - The unique control identifier of the listbox that sent this notification
' ##EVENT_REMARKS ListboxWindow - The window of the listbox control that sent the notification
Public Event ChildListboxNotification(ByVal Notification As ListBoxNotifications, ByVal ListboxId As Integer, ByVal ListboxWindow As ApiWindow)

' ##EVENT_DESCRIPTION ChildStaticControlNotification - A static (frame, picturebox etc) control on this window has notified this window of an event occuring to it
' ##EVENT_REMARKS Notification - The event that happened
' ##EVENT_REMARKS StaticId - The unique control identifier of the static that sent this notification
' ##EVENT_REMARKS StaticWindow - The window of the static control that sent the notification
Public Event ChildStaticControlNotification(ByVal Notification As StaticControlNotifications, ByVal StaticId As Integer, ByVal StaticWindow As ApiWindow)

' ##EVENT_DESCRIPTION ChildComboboxNotification - A combo box control on this window has notified this window of an event occuring to it
' ##EVENT_REMARKS Notification - The event that happened
' ##EVENT_REMARKS ComboBoxId - The unique control identifier of the combo box that sent this notification
' ##EVENT_REMARKS ComboBoxWindow - The window of the combo box control that sent the notification
Public Event ChildComboboxNotification(ByVal Notification As ComboBoxNotifications, ByVal ComboId As Integer, ByVal Combowindow As ApiWindow)


'\\ Private Member variables -------------------------------------------------------------
Private m_hwnd As Long

'\\ API Declarations -------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function IsWindowApi Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function GetParentApi Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function WindowFromPointApi Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowApi Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function IsWindowEnabledApi Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicodeApi Lib "user32" Alias "IsWindowUnicode" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisibleApi Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomedApi Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Long
'\\ Window specific information
Private Declare Function GetWindowLongApi Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongApi Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'\\ Menu specific....
Private Declare Function GetSystemMenuApi Lib "user32" Alias "GetSystemMenu" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuApi Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
'\\ Device context specific....
Private Declare Function GetDCApi Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
'\\ Window class specific...
Private Declare Function GetClassNameApi Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'\\ Send message....
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function InvalidateRectByPointerApi Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'\\ Scrolling....
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'\\ Window scrolling
' ##ENUMERATION_DESCRIPTION ScrollWindowModes - Describes how a window scrolling affects the client area
Public Enum ScrollWindowModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_ERASE - The background area scrolled is erased
    SW_ERASE = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SCROLLCHILDREN - Child windows in the scroll area are scrolled
    SW_SCROLLCHILDREN = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_INVALIDATE - The scrolled area is marked as needing painting
    SW_INVALIDATE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SW_SMOOTHSCROLL - The scroll is done in smooth steps
    SW_SMOOTHSCROLL = &H10
End Enum
Private Declare Function ScrollWindowExAPI Lib "user32" Alias "ScrollWindowEx" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As Long, lprcClip As Long, hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
Private Declare Function ScrollWindow Lib "user32" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long

'\\ Mouse message capture
Private Declare Function SetCaptureApi Lib "user32" Alias "SetCapture" (ByVal hwnd As Long) As Long
Private Declare Function GetCaptureApi Lib "user32" Alias "GetCapture" () As Long
'\\ Windows regions.....
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'\\ Refreshing the window states....
Private Declare Function DrawMenuBarApi Lib "user32" Alias "DrawMenuBar" (ByVal hwnd As Long) As Long
'\\ Window text...
Private Declare Function GetWindowTextLengthApi Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextApi Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'\\ Painting the window,
Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

'\\ Non-client area painting...
Private Declare Function GetDCEx Lib "user32" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Public Enum ExtendedDeviceContextIndexes
    DCX_WINDOW = &H1
    DCX_CACHE = &H2
    DCX_NORESETATTRS = &H4
    DCX_CLIPCHILDREN = &H8
    DCX_CLIPSIBLINGS = &H10
    DCX_PARENTCLIP = &H20
    DCX_EXCLUDERGN = &H40
    DCX_INTERSECTRGN = &H80
    DCX_EXCLUDEUPDATE = &H100
    DCX_INTERSECTUPDATE = &H200
    DCX_LOCKWINDOWUPDATE = &H400
    DCX_VALIDATE = &H200000
End Enum

'\\ Set window pos...
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Enum SetWindowPosFlags
    SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

'\\ Getting information about a window's title bar...
Public Enum WindowSystemStates
    STATE_SYSTEM_UNAVAILABLE = &H1                '\\ Disabled
    STATE_SYSTEM_SELECTED = &H2
    STATE_SYSTEM_FOCUSED = &H4
    STATE_SYSTEM_PRESSED = &H8
    STATE_SYSTEM_CHECKED = &H10
    STATE_SYSTEM_MIXED = &H20                     '\\ 3-state checkbox or toolbar button
    STATE_SYSTEM_READONLY = &H40
    STATE_SYSTEM_HOTTRACKED = &H80
    STATE_SYSTEM_DEFAULT = &H100
    STATE_SYSTEM_EXPANDED = &H200
    STATE_SYSTEM_COLLAPSED = &H400
    STATE_SYSTEM_BUSY = &H800
    STATE_SYSTEM_FLOATING = &H1000                '\\ Children "owned" not "contained" by parent
    STATE_SYSTEM_MARQUEED = &H2000
    STATE_SYSTEM_ANIMATED = &H4000
    STATE_SYSTEM_INVISIBLE = &H8000
    STATE_SYSTEM_OFFSCREEN = &H10000
    STATE_SYSTEM_SIZEABLE = &H20000
    STATE_SYSTEM_MOVEABLE = &H40000
    STATE_SYSTEM_SELFVOICING = &H80000
    STATE_SYSTEM_FOCUSABLE = &H100000
    STATE_SYSTEM_SELECTABLE = &H200000
    STATE_SYSTEM_LINKED = &H400000
    STATE_SYSTEM_TRAVERSED = &H800000
    STATE_SYSTEM_MULTISELECTABLE = &H1000000      '\\ Supports multiple selection
    STATE_SYSTEM_EXTSELECTABLE = &H2000000        '\\ Supports extended selection
    STATE_SYSTEM_ALERT_LOW = &H4000000            '\\ This information is of low priority
    STATE_SYSTEM_ALERT_MEDIUM = &H8000000         '\\ This information is of medium priority
    STATE_SYSTEM_ALERT_HIGH = &H10000000          '\\ This information is of high priority
    STATE_SYSTEM_VALID = &H1FFFFFFF
End Enum

Private Type TITLEBARINFO
    cbSize As Long '\\ Size of this structure
    rcTitleBar As RECT '\\ Rectangle of the whole title bar except menu..
    rgstate(0 To 5) As Long '\\ Each is a combination of window_system_states
End Type
Private Declare Function GetTitleBarInfo Lib "user32" (ByVal hwnd As Long, lpTileBarInfo As TITLEBARINFO) As Long


Public Enum PowerBroadcastMessageTypes
     PBT_APMQUERYSUSPEND = &H0
     PBT_APMQUERYSTANDBY = &H1
     PBT_APMQUERYSUSPENDFAILED = &H2
     PBT_APMQUERYSTANDBYFAILED = &H3
     PBT_APMSUSPEND = &H4
     PBT_APMSTANDBY = &H5
     PBT_APMRESUMECRITICAL = &H6
     PBT_APMRESUMESUSPEND = &H7
     PBT_APMRESUMESTANDBY = &H8
     PBT_APMBATTERYLOW = &H9
     PBT_APMPOWERSTATUSCHANGE = &HA
     PBT_APMOEMEVENT = &HB
     PBT_APMRESUMEAUTOMATIC = &H12
End Enum


'\\ Private member variables ----------------------------------------------------------------------------------------------------------------------------------------------------------
Private mOldProcAddress As Long
Private mPrevCapture As Long
'\\ Objects....
Private mSystemMenu As ApiMenu
Private mMenu As ApiMenu
Private mWndClass As APIWndClass
Private mRegion As ApiRegion
Private mDeviceContext As ApiDeviceContext
Private mDeviceContextNC As ApiDeviceContext

'\\ Setting of minmaxinfo....
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type


'\\ Memory manipulation routines for min/max info
Private Declare Sub CopyMemoryMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As MINMAXINFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMinMaxToMemoryInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As MINMAXINFO, ByVal Length As Long)

'\\ message passing
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'\\ Threads...
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

' ##ENUMERATION_DESCRIPTION ComboBoxNotifications - Notifications sent from a combo box to its parent when event occur to it
Public Enum ComboBoxNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_CLOSEUP - The combo drop down area closed
    CBN_CLOSEUP = 8
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_DBLCLK - The control was double-clicked
    CBN_DBLCLK = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_DROPDOWN - The combo box dropped down
    CBN_DROPDOWN = 7
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_EDITCHANGE - Combo text has changed
    CBN_EDITCHANGE = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_EDITUPDATE - Combo text is about to change
    CBN_EDITUPDATE = 6
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_ERRSPACE - The combo has run out of memory
    CBN_ERRSPACE = -1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_KILLFOCUS - The combo has lost the focus
    CBN_KILLFOCUS = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_SELCHANGE - The combo box selected item has changed
    CBN_SELCHANGE = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_SELENDCANCEL - The selection change was cancelled
    CBN_SELENDCANCEL = 10
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_SELENDOK - The combo box selection has been made correctly
    CBN_SELENDOK = 9
    ' ##ENUMERATION_MEMBER_DESCRIPTION CBN_SETFOCUS - The combo has gained the focus
    CBN_SETFOCUS = 3
End Enum



' ##ENUMERATION_DESCRIPTION CommonControlGenericNotifications - Notifications common to all the common control types
Public Enum CommonControlGenericNotifications
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_OUTOFMEMORY - The common control is out of memory
    NM_OUTOFMEMORY = -1
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_CLICK - The common control was clicked
    NM_CLICK = -2
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_DBLCLK - The common control was double clicked
    NM_DBLCLK = -3
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_RETURN - The enter key was pressed while the common control had the focus
    NM_RETURN = -4
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_RCLICK - The common control was double clicked with the right mouse button
    NM_RCLICK = -5
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_RDBLCLK - The common control was double clicked with the right mouse button
    NM_RDBLCLK = -6
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_SETFOCUS - The common control has gained the focus
    NM_SETFOCUS = -7
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_KILLFOCUS - The common control has lost the input focus
    NM_KILLFOCUS = -8
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_CUSTOMDRAW - The common control is doing a custom draw event
    NM_CUSTOMDRAW = -12
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_HOVER - The mouse has hovered over a custom control
    NM_HOVER = -13
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_NCHITTEST - The hit test is being performed on the custom control
    NM_NCHITTEST = -14
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_KEYDOWN - A key was pressed down when the custom control had the focus
    NM_KEYDOWN = -15
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_RELEASEDCAPTURE - The custom control has lost the mouse capture
    NM_RELEASEDCAPTURE = -16
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_SETCURSOR - The custom control is setting the cursor
    NM_SETCURSOR = -17
    ' ##ENUMERATION_MEMBER_DESCRIPTION NM_CHAR - A character key has been pressed while the custom control had the focus
    NM_CHAR = -18
End Enum

Public Enum ListviewNotificationMessages
    LVN_ITEMCHANGING = -100
    LVN_ITEMCHANGED = -101
    LVN_INSERTITEM = -102
    LVN_DELETEITEM = -103
    LVN_DELETEALLITEMS = -104
    LVN_BEGINLABELEDITA = -105
    LVN_BEGINLABELEDITW = -175
    LVN_ENDLABELEDITA = -106
    LVN_ENDLABELEDITW = -176
    LVN_COLUMNCLICK = -108
    LVN_BEGINDRAG = -109
    LVN_BEGINRDRAG = -111
    LVN_ODCACHEHINT = -113
    LVN_ODFINDITEMA = -152
    LVN_ODFINDITEMW = -179
    LVN_ITEMACTIVATE = -114
    LVN_ODSTATECHANGED = -115
End Enum

Private Const HDN_FIRST = -300
Public Enum enHeaderNotificationMessages
    HDN_ITEMCHANGINGA = (HDN_FIRST - 0)
    HDN_ITEMCHANGINGW = (HDN_FIRST - 20)
    HDN_ITEMCHANGEDA = (HDN_FIRST - 1)
    HDN_ITEMCHANGEDW = (HDN_FIRST - 21)
    HDN_ITEMCLICKA = (HDN_FIRST - 2)
    HDN_ITEMCLICKW = (HDN_FIRST - 22)
    HDN_ITEMDBLCLICKA = (HDN_FIRST - 3)
    HDN_ITEMDBLCLICKW = (HDN_FIRST - 23)
    HDN_DIVIDERDBLCLICKA = (HDN_FIRST - 5)
    HDN_DIVIDERDBLCLICKW = (HDN_FIRST - 25)
    HDN_BEGINTRACKA = (HDN_FIRST - 6)
    HDN_BEGINTRACKW = (HDN_FIRST - 26)
    HDN_ENDTRACKA = (HDN_FIRST - 7)
    HDN_ENDTRACKW = (HDN_FIRST - 27)
    HDN_TRACKA = (HDN_FIRST - 8)
    HDN_TRACKW = (HDN_FIRST - 28)
    HDN_GETDISPINFOA = (HDN_FIRST - 9)
    HDN_GETDISPINFOW = (HDN_FIRST - 29)
    HDN_BEGINDRAG = (HDN_FIRST - 10)
    HDN_ENDDRAG = (HDN_FIRST - 11)
End Enum

Private Const TVN_FIRST = -400
Public Enum enTreeviewNotificationMessages
    TVN_SELCHANGINGA = (TVN_FIRST - 1)
    TVN_SELCHANGINGW = (TVN_FIRST - 50)
    TVN_SELCHANGEDA = (TVN_FIRST - 2)
    TVN_SELCHANGEDW = (TVN_FIRST - 51)
    TVN_GETDISPINFOA = (TVN_FIRST - 3)
    TVN_GETDISPINFOW = (TVN_FIRST - 52)
    TVN_SETDISPINFOA = (TVN_FIRST - 4)
    TVN_SETDISPINFOW = (TVN_FIRST - 53)
    TVN_ITEMEXPANDINGA = (TVN_FIRST - 5)
    TVN_ITEMEXPANDINGW = (TVN_FIRST - 54)
    TVN_ITEMEXPANDEDA = (TVN_FIRST - 6)
    TVN_ITEMEXPANDEDW = (TVN_FIRST - 55)
    TVN_BEGINDRAGA = (TVN_FIRST - 7)
    TVN_BEGINDRAGW = (TVN_FIRST - 56)
    TVN_BEGINRDRAGA = (TVN_FIRST - 8)
    TVN_BEGINRDRAGW = (TVN_FIRST - 57)
    TVN_DELETEITEMA = (TVN_FIRST - 9)
    TVN_DELETEITEMW = (TVN_FIRST - 58)
    TVN_BEGINLABELEDITA = (TVN_FIRST - 10)
    TVN_BEGINLABELEDITW = (TVN_FIRST - 59)
    TVN_ENDLABELEDITA = (TVN_FIRST - 11)
    TVN_ENDLABELEDITW = (TVN_FIRST - 60)
    TVN_KEYDOWN = (TVN_FIRST - 12)
    TVN_GETINFOTIPA = (TVN_FIRST - 13)
    TVN_GETINFOTIPW = (TVN_FIRST - 14)
    TVN_SINGLEEXPAND = (TVN_FIRST - 15)
End Enum

Private Const TTN_FIRST = -520
Public Enum enToolTipNotificationMessages
    TTN_GETDISPINFOA = (TTN_FIRST - 0)
    TTN_GETDISPINFOW = (TTN_FIRST - 10)
    TTN_SHOW = (TTN_FIRST - 1)
    TTN_POP = (TTN_FIRST - 2)
End Enum

Private Const TCN_FIRST = -550
Public Enum enTabControlNotificationMessages
    TCN_KEYDOWN = (TCN_FIRST - 0)
    TCN_SELCHANGE = (TCN_FIRST - 1)
    TCN_SELCHANGING = (TCN_FIRST - 2)
    TCN_GETOBJECT = (TCN_FIRST - 3)
End Enum

Private Const CDN_FIRST = -601
Public Enum enCommonDialogNotificationMessages
    CDN_INITDONE = (CDN_FIRST - &H0)
    CDN_SELCHANGE = (CDN_FIRST - &H1)
    CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
    CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
    CDN_HELP = (CDN_FIRST - &H4)
    CDN_FILEOK = (CDN_FIRST - &H5)
    CDN_TYPECHANGE = (CDN_FIRST - &H6)
    CDN_INCLUDEITEM = (CDN_FIRST - &H7)
End Enum

Private Const TBN_FIRST = -700
Public Enum enToolbarNotificationMessages
    TBN_GETBUTTONINFOA = (TBN_FIRST - 0)
    TBN_GETBUTTONINFOW = (TBN_FIRST - 20)
    TBN_BEGINDRAG = (TBN_FIRST - 1)
    TBN_ENDDRAG = (TBN_FIRST - 2)
    TBN_BEGINADJUST = (TBN_FIRST - 3)
    TBN_ENDADJUST = (TBN_FIRST - 4)
    TBN_RESET = (TBN_FIRST - 5)
    TBN_QUERYINSERT = (TBN_FIRST - 6)
    TBN_QUERYDELETE = (TBN_FIRST - 7)
    TBN_TOOLBARCHANGE = (TBN_FIRST - 8)
    TBN_CUSTHELP = (TBN_FIRST - 9)
    TBN_DROPDOWN = (TBN_FIRST - 10)
    TBN_GETOBJECT = (TBN_FIRST - 12)
    TBN_HOTITEMCHANGE = (TBN_FIRST - 13)
    TBN_DRAGOUT = (TBN_FIRST - 14)            ''\\ this is sent when the user clicks down on a button then drags off the button
    TBN_DELETINGBUTTON = (TBN_FIRST - 15)     ''\\ uses TBNOTIFY
    TBN_GETDISPINFOA = (TBN_FIRST - 16)       ''\\ This is sent when the  toolbar needs  some display information
    TBN_GETDISPINFOW = (TBN_FIRST - 17)       ''\\ This is sent when the  toolbar needs  some display information
    TBN_GETINFOTIPA = (TBN_FIRST - 18)
    TBN_GETINFOTIPW = (TBN_FIRST - 19)
End Enum

Private Const UDN_FIRST = -721
Public Enum enUpDownNotificationMessages
    UDN_DELTAPOS = (UDN_FIRST - 1)
End Enum

Private Const MCN_FIRST = -750
Public Enum enMonthCalendarNotificationMessages
    MCN_SELCHANGE = (MCN_FIRST + 1)
    MCN_GETDAYSTATE = (MCN_FIRST + 3)
    MCN_SELECT = (MCN_FIRST + 4)
End Enum

Private Const DTN_FIRST = -760
Public Enum enDateTimeNotificationMessages
    DTN_DATETIMECHANGE = (DTN_FIRST + 1)
    DTN_USERSTRINGA = (DTN_FIRST + 2) ''\\ the user has entered a string
    DTN_USERSTRINGW = (DTN_FIRST + 15)
    DTN_WMKEYDOWNA = (DTN_FIRST + 3) ''\\ modify keydown on app format field (X)
    DTN_WMKEYDOWNW = (DTN_FIRST + 16)
    DTN_FORMATA = (DTN_FIRST + 4) ''\\ query display for app format field (X)
    DTN_FORMATW = (DTN_FIRST + 17)
    DTN_FORMATQUERYA = (DTN_FIRST + 5) ''\\ query formatting info for app format field (X)
    DTN_FORMATQUERYW = (DTN_FIRST + 18)
    DTN_DROPDOWN = (DTN_FIRST + 6)   ''\\ MonthCal has dropped down
    DTN_CLOSEUP = (DTN_FIRST + 7)    ''\\ MonthCal is popping up
End Enum

Private Const CBEN_FIRST = -800
Public Enum enComboBoxExtendedNotificationMessages
    CBEN_GETDISPINFO = -800
    CBEN_INSERTITEM = -801
    CBEN_DELETEITEM = -802
    CBEN_BEGINEDIT = -804
    CBEN_ENDEDITA = -805
    CBEN_ENDEDITW = -806
    CBEN_GETDISPINFOW = -807
    CBEN_DRAGBEGINA = -808
    CBEN_DRAGBEGINW = -809
End Enum

Private Const RBN_FIRST = -831
Public Enum enCoolbarNotificationMessages
    RBN_HEIGHTCHANGE = (RBN_FIRST - 0)
    RBN_GETOBJECT = (RBN_FIRST - 1)
    RBN_LAYOUTCHANGED = (RBN_FIRST - 2)
    RBN_AUTOSIZE = (RBN_FIRST - 3)
    RBN_BEGINDRAG = (RBN_FIRST - 4)
    RBN_ENDDRAG = (RBN_FIRST - 5)
    RBN_DELETINGBAND = (RBN_FIRST - 6)       ''\\ Uses NMREBAR
    RBN_DELETEDBAND = (RBN_FIRST - 7)        ''\\ Uses NMREBAR
    RBN_CHILDSIZE = (RBN_FIRST - 8)
End Enum

Private Const SBN_FIRST = -880
Public Enum enStatusBarNotificationMessages
    SBN_SIMPLEMODECHANGE = (SBN_FIRST - 0)
End Enum


Private mThreadParent As Long


'\\ Get relative window constants
Public Enum enGetWindow
    GW_HWNDFIRST = 0
    GW_HWNDLAST = 1
    GW_HWNDNEXT = 2
    GW_HWNDPREV = 3
    GW_OWNER = 4
    GW_CHILD = 5
End Enum

'\\ Window Style
' ##ENUMERATION_DESCRIPTION These style bits are combined to define what style elements the _
window that is created with them will have.
Public Enum WindowStyleBits
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_BORDER The window has a frame border around it
    WS_BORDER = &H800000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_CAPTION The window has a caption title bar
    WS_CAPTION = &HC00000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_CHILD The window is a child (non top level) window
    WS_CHILD = &H40000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_CLIPCHILDREN The window cannot draw over other windows that are its children
    WS_CLIPCHILDREN = &H2000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_CLIPSIBLINGS This window's child windows cannot draw over each other
    WS_CLIPSIBLINGS = &H4000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_DISABLED The window is disabled and cannot accept imput
    WS_DISABLED = &H8000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_DLGFRAME The window has a double border but no caption
    WS_DLGFRAME = &H400000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_GROUP The window is the first of a group of windows in a dialog box
    WS_GROUP = &H20000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_HSCROLL The window has a horizontal scroll bar
    WS_HSCROLL = &H100000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_MAXIMIZE The window is maximised when first created
    WS_MAXIMIZE = &H1000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_MAXIMIZEBOX The window has a maximise box
    WS_MAXIMIZEBOX = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_MINIMIZE The window is minimised when first created
    WS_MINIMIZE = &H20000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_MINIMIZEBOX The window has a minimise box
    WS_MINIMIZEBOX = &H20000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_OVERLAPPED The window is a standard top level (application) window
    WS_OVERLAPPED = &H0&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_POPUP The window is a popup type window
    WS_POPUP = &H80000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_SYSMENU The window has a system menu box
    WS_SYSMENU = &H80000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_TABSTOP The window accepts the focus when tabbed to
    WS_TABSTOP = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_THICKFRAME The window has a resizeable window style thick border
    WS_THICKFRAME = &H40000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_VISIBLE The window is visible
    WS_VISIBLE = &H10000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_VSCROLL The window has a vertical scroll bar
    WS_VSCROLL = &H200000
End Enum

' ##ENUMERATION_DESCRIPTION These extended style bits are combined to define what style elements the _
window that is created with them will have.
Public Enum WindowStyleExtendedBits
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_ACCEPTFILES The window will accept files being dropped on it from Explorer
    WS_EX_ACCEPTFILES = &H10&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_DLGMODALFRAME The window has a double border and may have a caption
    WS_EX_DLGMODALFRAME = &H1&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_NOPARENTNOTIFY The window will not send a message to its parent (if any) when it is destroyed
    WS_EX_NOPARENTNOTIFY = &H4&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_TOPMOST The window will float over non topmost windows
    WS_EX_TOPMOST = &H8&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_TRANSPARENT The window does not obscure windows behind it
    WS_EX_TRANSPARENT = &H20&
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_TOOLWINDOW The window has a toolbox look (smaller caption bar)
    WS_EX_TOOLWINDOW = &H80&
    
    '\\ New from 95/NT4 onwards
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_MDICHILD The window is an MDI child style window
    WS_EX_MDICHILD = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_WINDOWEDGE The window has a beveled outer edge
    WS_EX_WINDOWEDGE = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_CLIENTEDGE The window has a chiselled inner edge
    WS_EX_CLIENTEDGE = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_CONTEXTHELP The window supports context sensitive help
    WS_EX_CONTEXTHELP = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_RIGHT Text in this window is right aligned (if WS_EX_RTLREADING is set)
    WS_EX_RIGHT = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_LEFT Test in this window is left aligned
    WS_EX_LEFT = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_RTLREADING Arranges text for right to left reading (for Hebrew or Arabic for example)
    WS_EX_RTLREADING = &H2000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_LTRREADING Arranges text for left to right reading
    WS_EX_LTRREADING = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_LEFTSCROLLBAR Scrollbar on the left hand side (if WS_EX_RTLREADING is set)
    WS_EX_LEFTSCROLLBAR = &H4000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_RIGHTSCROLLBAR Scrollbar on the right hand side
    WS_EX_RIGHTSCROLLBAR = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_CONTROLPARENT causes child windows of this window to have a tab order
    WS_EX_CONTROLPARENT = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_STATICEDGE Makes the window a 3D style frame window
    WS_EX_STATICEDGE = &H20000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_APPWINDOW Minimises to the task bar
    WS_EX_APPWINDOW = &H40000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_OVERLAPPEDWINDOW Has a raised edge around it like an application window
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_PALETTEWINDOW Is a floating toolbar or palette style window
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_LAYERED Creates a layered window
    WS_EX_LAYERED = &H80000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_NOINHERITLAYOUT Does not pass on it's layout to child windows
    WS_EX_NOINHERITLAYOUT = &H100000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_LAYOUTRTL Layout ir right to left reading (where appropriate)
    WS_EX_LAYOUTRTL = &H400000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_COMPOSITED Paints the child windows in bottom to top order
    WS_EX_COMPOSITED = &H2000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION WS_EX_NOACTIVATE Window cannot be activated
    WS_EX_NOACTIVATE = &H8000000
End Enum

'\\ This is in the wParam of WM_POWER and DRV_POWER messages
Public Enum PowerStatusChangeMessages
     PWR_OK = 1
     PWR_FAIL = (-1)
     PWR_SUSPENDREQUEST = 1
     PWR_SUSPENDRESUME = 2
     PWR_CRITICALRESUME = 3
End Enum

' ##ENUMERATION_DESCRIPTION WindowSizingEdges - Indicates which part of the window is being drag-resized
Public Enum WindowSizingEdges
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_LEFT The sizing is from the left hand frame border
     WMSZ_LEFT = 1
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_RIGHT The sizing is from the right hand frame border
     WMSZ_RIGHT = 2
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_TOP The sizing is from the top frame border
     WMSZ_TOP = 3
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_TOPLEFT The sizing is from the top left corner
     WMSZ_TOPLEFT = 4
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_TOPRIGHT The sizing is from the top right corner
     WMSZ_TOPRIGHT = 5
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_BOTTOM The sizing is from the bottom frame border
     WMSZ_BOTTOM = 6
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_BOTTOMLEFT The sizing is from the bottom left corner
     WMSZ_BOTTOMLEFT = 7
     ' ##ENUMERATION_MEMBER_DESCRIPTION WMSZ_BOTTOMRIGHT The sizing is from the bottom right corner
     WMSZ_BOTTOMRIGHT = 8
End Enum

'\\ This is in wParam of WM_SIZE messages
Public Enum WindowSizeMessages
     SIZE_RESTORED = 0
     SIZE_MINIMIZED = 1
     SIZE_MAXIMIZED = 2
     SIZE_MAXSHOW = 3
     SIZE_MAXHIDE = 4
End Enum

' ##ENUMERATION_DESCRIPTION Before WM_MOUSEHOVER and WM_MOUSELEAVE messages can be generated for
' ##ENUMERATION_DESCRIPTION a window, the message tracking must be enabled on that window.
Public Enum TrackMouseEvents
    ' ##ENUMERATION_MEMBER_DESCRIPTION TME_HOVER Request WM_MOUSEHOVER message notification
     TME_HOVER = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION TME_LEAVE Request WM_MOUSELEAVE message notification
     TME_LEAVE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION TME_NONCLIENT Request WM_NCMOUSEHOVER and WM_NCMOUSEHOVER messages
     TME_NONCLIENT = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION TME_QUERY Get the mouse track status for this window
     TME_QUERY = &H40000000
    ' ##ENUMERATION_MEMBER_DESCRIPTION TME_CANCEL Cancel mouse tracking
     TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSESTRUCT
    cbSize As Long
    dwFlags As Long '\\ TrackMouseEvents
    hwndTrack As Long '\\ Target window
    dwHoverTime As Long '\\ Mouse hover timeout in milliseconds
End Type
Private Declare Function TrackMouseEvent Lib "user32" (lpTrackMouseStruct As TRACKMOUSESTRUCT) As Long
Dim mheThis As TRACKMOUSESTRUCT
Dim mheOut As TRACKMOUSESTRUCT

' ##ENUMERATION_DESCRIPTION MouseKeyMasks describes the state of various buttons when a mouse event occurs
Public Enum MouseKeyMasks
    ' ##ENUMERATION_MEMBER_DESCRIPTION MK_LBUTTON indicates that the left mouse button is pressed
     MK_LBUTTON = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_RBUTTON indicates that the right mouse button is pressed
     MK_RBUTTON = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_SHIFT indicates that the Shift key is pressed
     MK_SHIFT = &H4
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_CONTROL indicates that the Control key is pressed
     MK_CONTROL = &H8
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_MBUTTON idicates that the middle mouse button is pressed
     MK_MBUTTON = &H10
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_XBUTTON1 indicates that the first XButton is pressed
     MK_XBUTTON1 = &H20
     ' ##ENUMERATION_MEMBER_DESCRIPTION MK_XBUTTON2 indicates that the second XButton is pressed
     MK_XBUTTON2 = &H40
End Enum

'\\ Window message constants - for WndProc(wMsg).
' ##ENUMERATION_DESCRIPTION WindowMessages are the system defined messages.
Public Enum WindowMessages
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ACTIVATE the window being activated or the window being deactivated
    WM_ACTIVATE = &H6
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ACTIVATEAPP triggers the %Active Application Changed:EventVB~ApiWindow~ActiveApplicationChanged_EV% event
    WM_ACTIVATEAPP = &H1C
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ASKCBFORMATNAME
    WM_ASKCBFORMATNAME = &H30C
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CANCELJOURNAL
    WM_CANCELJOURNAL = &H4B
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CANCELMODE sent to cancel certain modes, such as mouse capture
    WM_CANCELMODE = &H1F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CAPTURECHANGED is sent when a window loses the mouse capture.  This message triggers the %LostCapture:EventVB~ApiWindow~LostCapture_EV% event
    WM_CAPTURECHANGED = &H1F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CAPTURECHANGED_R
    WM_CAPTURECHANGED_R = &H215
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHANGECBCHAIN
    WM_CHANGECBCHAIN = &H30D
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHAR
    WM_CHAR = &H102
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHARTOITEM
    WM_CHARTOITEM = &H2F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHILDACTIVATE sent to a child window when the user clicks the window's title bar or when the window is activated, moved, or sized
    WM_CHILDACTIVATE = &H22
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHOOSEFONT_GETLOGFONT
    WM_CHOOSEFONT_GETLOGFONT = &H401
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHOOSEFONT_SETFLAGS
    WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CHOOSEFONT_SETLOGFONT
    WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CLEAR
    WM_CLEAR = &H303
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CLOSE
    WM_CLOSE = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_COMMAND
    WM_COMMAND = &H111
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_COMPACTING
    WM_COMPACTING = &H41
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_COMPAREITEM
    WM_COMPAREITEM = &H39
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CONTEXTMENU Sent when the user right-clicks on a control.
    WM_CONTEXTMENU = &H7B
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CONVERTREQUESTEX
    WM_CONVERTREQUESTEX = &H108
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_COPY
    WM_COPY = &H301
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_COPYDATA
    WM_COPYDATA = &H4A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CREATE
    WM_CREATE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORBTN
    WM_CTLCOLORBTN = &H135
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORDLG
    WM_CTLCOLORDLG = &H136
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLOREDIT
    WM_CTLCOLOREDIT = &H133
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORLISTBOX
    WM_CTLCOLORLISTBOX = &H134
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORMSGBOX
    WM_CTLCOLORMSGBOX = &H132
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORSCROLLBAR
    WM_CTLCOLORSCROLLBAR = &H137
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CTLCOLORSTATIC
    WM_CTLCOLORSTATIC = &H138
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_CUT
    WM_CUT = &H300
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_ACK
    WM_DDE_ACK = (&H3E0 + 4)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_ADVISE
    WM_DDE_ADVISE = (&H3E0 + 2)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_DATA
    WM_DDE_DATA = (&H3E0 + 5)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_EXECUTE
    WM_DDE_EXECUTE = (&H3E0 + 8)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_FIRST
    WM_DDE_FIRST = &H3E0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_INITIATE
    WM_DDE_INITIATE = &H3E0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_LAST
    WM_DDE_LAST = (&H3E0 + 8)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_POKE
    WM_DDE_POKE = (&H3E0 + 7)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_REQUEST
    WM_DDE_REQUEST = (&H3E0 + 6)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_TERMINATE
    WM_DDE_TERMINATE = (&H3E0 + 1)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DDE_UNADVISE
    WM_DDE_UNADVISE = (&H3E0 + 3)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DEADCHAR
    WM_DEADCHAR = &H103
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DELETEITEM
    WM_DELETEITEM = &H2D
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DESTROY
    WM_DESTROY = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DESTROYCLIPBOARD
    WM_DESTROYCLIPBOARD = &H307
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DEVICECHANGE
    WM_DEVICECHANGE = &H219
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DEVMODECHANGE
    WM_DEVMODECHANGE = &H1B
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DRAWCLIPBOARD
    WM_DRAWCLIPBOARD = &H308
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DRAWITEM
    WM_DRAWITEM = &H2B
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_DROPFILES
    WM_DROPFILES = &H233
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ENABLE
    WM_ENABLE = &HA
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ENDSESSION
    WM_ENDSESSION = &H16
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ENTERIDLE
    WM_ENTERIDLE = &H121
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ENTERSIZEMOVE
    WM_ENTERSIZEMOVE = &H231
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ENTERMENULOOP
    WM_ENTERMENULOOP = &H211
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ERASEBKGND
    WM_ERASEBKGND = &H14
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_EXITMENULOOP
    WM_EXITMENULOOP = &H212
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_EXITSIZEMOVE
    WM_EXITSIZEMOVE = &H232
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_FONTCHANGE
    WM_FONTCHANGE = &H1D
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETDLGCODE
    WM_GETDLGCODE = &H87
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETFONT
    WM_GETFONT = &H31
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETHOTKEY
    WM_GETHOTKEY = &H33
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETICON - Gets the icon associated with a window
    WM_GETICON = &H7F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETMINMAXINFO
    WM_GETMINMAXINFO = &H24
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETTEXT
    WM_GETTEXT = &HD
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_GETTEXTLENGTH
    WM_GETTEXTLENGTH = &HE
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_HELP - Recieved by a window when F1 is pressed
    WM_HELP = &H53
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_HOTKEY
    WM_HOTKEY = &H312
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_HSCROLL
    WM_HSCROLL = &H114
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_HSCROLLCLIPBOARD
    WM_HSCROLLCLIPBOARD = &H30E
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_ICONERASEBKGND
    WM_ICONERASEBKGND = &H27
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_CHAR
    WM_IME_CHAR = &H286
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_COMPOSITION
    WM_IME_COMPOSITION = &H10F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_COMPOSITIONFULL
    WM_IME_COMPOSITIONFULL = &H284
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_CONTROL
    WM_IME_CONTROL = &H283
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_ENDCOMPOSITION
    WM_IME_ENDCOMPOSITION = &H10E
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_KEYDOWN
    WM_IME_KEYDOWN = &H290
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_KEYLAST
    WM_IME_KEYLAST = &H10F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_KEYUP
    WM_IME_KEYUP = &H291
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_NOTIFY
    WM_IME_NOTIFY = &H282
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_SELECT
    WM_IME_SELECT = &H285
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_SETCONTEXT
    WM_IME_SETCONTEXT = &H281
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_IME_STARTCOMPOSITION
    WM_IME_STARTCOMPOSITION = &H10D
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_INITDIALOG
    WM_INITDIALOG = &H110
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_INITMENU Is sent to a window when a menu is going to be shown
    WM_INITMENU = &H116
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_INITMENUPOPUP Is sent to a window when a popup menu is going to be shown
    WM_INITMENUPOPUP = &H117
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_INPUTLANGCHANGEREQUEST
    WM_INPUTLANGCHANGEREQUEST = &H50
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_INPUTLANGCHANGE
    WM_INPUTLANGCHANGE = &H51
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_KEYDOWN
    WM_KEYDOWN = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_KEYUP
    WM_KEYUP = &H101
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_KILLFOCUS
    WM_KILLFOCUS = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_LBUTTONDBLCLK
    WM_LBUTTONDBLCLK = &H203
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_LBUTTONDOWN
    WM_LBUTTONDOWN = &H201
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_LBUTTONUP
    WM_LBUTTONUP = &H202
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MBUTTONDBLCLK
    WM_MBUTTONDBLCLK = &H209
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MBUTTONDOWN
    WM_MBUTTONDOWN = &H207
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MBUTTONUP
    WM_MBUTTONUP = &H208
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIACTIVATE
    WM_MDIACTIVATE = &H222
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDICASCADE
    WM_MDICASCADE = &H227
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDICREATE
    WM_MDICREATE = &H220
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIDESTROY
    WM_MDIDESTROY = &H221
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIGETACTIVE
    WM_MDIGETACTIVE = &H229
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIICONARRANGE
    WM_MDIICONARRANGE = &H228
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIMAXIMIZE
    WM_MDIMAXIMIZE = &H225
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDINEXT
    WM_MDINEXT = &H224
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIREFRESHMENU
    WM_MDIREFRESHMENU = &H234
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDIRESTORE
    WM_MDIRESTORE = &H223
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDISETMENU
    WM_MDISETMENU = &H230
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MDITILE
    WM_MDITILE = &H226
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MEASUREITEM
    WM_MEASUREITEM = &H2C
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENUCHAR
    WM_MENUCHAR = &H120
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENUSELECT
    WM_MENUSELECT = &H11F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENURBUTTONUP
    WM_MENURBUTTONUP = &H122
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENUDRAG
    WM_MENUDRAG = &H123
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENUGETOBJECT
    WM_MENUGETOBJECT = &H124
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MENUCOMMAND
    WM_MENUCOMMAND = &H126
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOUSEACTIVATE
    WM_MOUSEACTIVATE = &H21
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOUSEHOVER
    WM_MOUSEHOVER = &H2A1
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOUSELEAVE
    WM_MOUSELEAVE = &H2A3
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOUSEMOVE
    WM_MOUSEMOVE = &H200
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOUSEWHEEL
    WM_MOUSEWHEEL = &H20A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOVE
    WM_MOVE = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_MOVING
    WM_MOVING = &H216
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCACTIVATE
    WM_NCACTIVATE = &H86
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCCALCSIZE
    WM_NCCALCSIZE = &H83
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCCREATE
    WM_NCCREATE = &H81
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCDESTROY
    WM_NCDESTROY = &H82
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCHITTEST
    WM_NCHITTEST = &H84
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCLBUTTONDBLCLK
    WM_NCLBUTTONDBLCLK = &HA3
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCLBUTTONDOWN
    WM_NCLBUTTONDOWN = &HA1
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCLBUTTONUP
    WM_NCLBUTTONUP = &HA2
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCMBUTTONDBLCLK
    WM_NCMBUTTONDBLCLK = &HA9
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCMBUTTONDOWN
    WM_NCMBUTTONDOWN = &HA7
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCMBUTTONUP
    WM_NCMBUTTONUP = &HA8
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCMOUSEMOVE
    WM_NCMOUSEMOVE = &HA0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCPAINT
    WM_NCPAINT = &H85
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCRBUTTONDBLCLK
    WM_NCRBUTTONDBLCLK = &HA6
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCRBUTTONDOWN
    WM_NCRBUTTONDOWN = &HA4
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NCRBUTTONUP
    WM_NCRBUTTONUP = &HA5
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NEXTDLGCTL
    WM_NEXTDLGCTL = &H28
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NEXTMENU
    WM_NEXTMENU = &H213
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NOTIFY - Used to notify a control parent of an event in the child control
    WM_NOTIFY = &H4E
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_NULL
    WM_NULL = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PAINT
    WM_PAINT = &HF
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PAINTCLIPBOARD
    WM_PAINTCLIPBOARD = &H309
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PAINTICON
    WM_PAINTICON = &H26
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PALETTECHANGED
    WM_PALETTECHANGED = &H311
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PALETTEISCHANGING
    WM_PALETTEISCHANGING = &H310
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PARENTNOTIFY
    WM_PARENTNOTIFY = &H210
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PASTE
    WM_PASTE = &H302
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PENWINFIRST
    WM_PENWINFIRST = &H380
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PENWINLAST
    WM_PENWINLAST = &H38F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_POWER
    WM_POWER = &H48
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_POWERBROADCAST
    WM_POWERBROADCAST = &H218
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PRINT
    WM_PRINT = &H317
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PRINTCLIENT
    WM_PRINTCLIENT = &H318
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_ENVSTAMPRECT
    WM_PSD_ENVSTAMPRECT = (&H400 + 5)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_FULLPAGERECT
    WM_PSD_FULLPAGERECT = (&H400 + 1)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_GREEKTEXTRECT
    WM_PSD_GREEKTEXTRECT = (&H400 + 4)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_MARGINRECT
    WM_PSD_MARGINRECT = (&H400 + 3)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_MINMARGINRECT
    WM_PSD_MINMARGINRECT = (&H400 + 2)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_PAGESETUPDLG
    WM_PSD_PAGESETUPDLG = (&H400)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_PSD_YAFULLPAGERECT
    WM_PSD_YAFULLPAGERECT = (&H400 + 6)
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUERYDRAGICON
    WM_QUERYDRAGICON = &H37
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUERYENDSESSION
    WM_QUERYENDSESSION = &H11
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUERYNEWPALETTE
    WM_QUERYNEWPALETTE = &H30F
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUERYOPEN
    WM_QUERYOPEN = &H13
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUEUESYNC
    WM_QUEUESYNC = &H23
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_QUIT
    WM_QUIT = &H12
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_RBUTTONDBLCLK
    WM_RBUTTONDBLCLK = &H206
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_RBUTTONDOWN
    WM_RBUTTONDOWN = &H204
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_RBUTTONUP
    WM_RBUTTONUP = &H205
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_RENDERALLFORMATS
    WM_RENDERALLFORMATS = &H306
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_RENDERFORMAT
    WM_RENDERFORMAT = &H305
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETCURSOR
    WM_SETCURSOR = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETFOCUS
    WM_SETFOCUS = &H7
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETFONT
    WM_SETFONT = &H30
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETHOTKEY
    WM_SETHOTKEY = &H32
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETICON - Sets the icon for a window
    WM_SETICON = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETREDRAW
    WM_SETREDRAW = &HB
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETTEXT
    WM_SETTEXT = &HC
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SETTINGCHANGE
    WM_SETTINGCHANGE = &H1A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SHOWWINDOW
    WM_SHOWWINDOW = &H18
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SIZE
    WM_SIZE = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SIZING
    WM_SIZING = &H214
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SIZECLIPBOARD
    WM_SIZECLIPBOARD = &H30B
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SPOOLERSTATUS
    WM_SPOOLERSTATUS = &H2A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSCHAR
    WM_SYSCHAR = &H106
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSCOLORCHANGE
    WM_SYSCOLORCHANGE = &H15
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSCOMMAND
    WM_SYSCOMMAND = &H112
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSDEADCHAR
    WM_SYSDEADCHAR = &H107
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSKEYDOWN
    WM_SYSKEYDOWN = &H104
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_SYSKEYUP
    WM_SYSKEYUP = &H105
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_TIMECHANGE
    WM_TIMECHANGE = &H1E
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_TIMER
    WM_TIMER = &H113
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_UNDO
    WM_UNDO = &H304
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_USER The base index for user defined window messages
    WM_USER = &H400
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_VKEYTOITEM
    WM_VKEYTOITEM = &H2E
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_VSCROLL A vertical scroll event occured
    WM_VSCROLL = &H115
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_VSCROLLCLIPBOARD
    WM_VSCROLLCLIPBOARD = &H30A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_WINDOWPOSCHANGED
    WM_WINDOWPOSCHANGED = &H47
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_WINDOWPOSCHANGING
    WM_WINDOWPOSCHANGING = &H46
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_WININICHANGE
    WM_WININICHANGE = &H1A
    ' ##ENUMERATION_MEMBER_DESCRIPTION WM_APPCOMMAND (Win2000 or above)
    WM_APPCOMMAND = &H319
End Enum

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

' ##ENUMERATION_DESCRIPTION These values are passed in wParam of the WM_INPUTLANGCHANGE message
Public Enum KeyboardLanguageChangeRequestModes
    ' ##ENUMERATION_MEMBER_DESCRIPTION INPUTLANGCHANGE_SYSCHARSET The new input langauge can use the system default character set
     INPUTLANGCHANGE_SYSCHARSET = &H1
     ' ##ENUMERATION_MEMBER_DESCRIPTION INPUTLANGCHANGE_FORWARD The user moved forward in the input languages collection
     INPUTLANGCHANGE_FORWARD = &H2
     ' ##ENUMERATION_MEMBER_DESCRIPTION INPUTLANGCHANGE_BACKWARD The user moved backward in the input languages collection
     INPUTLANGCHANGE_BACKWARD = &H4
End Enum

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' ##ENUMERATION_DESCRIPTION The standard types of simple windows controls
Public Enum StandardControlTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_Button - A button control
    Control_BUTTON = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_Edit - A textbox or edit control
    Control_EDIT = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_Listbox - A listbox control
    Control_LISTBOX = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_Scrollbar - A horizontal or vertical scrollbar
    Control_SCROLLBAR = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_Static - A label type control
    Control_STATIC = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION Control_ComboBox - A combo box style control
    Control_COMBOBOX = 6
End Enum


'used for owner-drawn contols....
' ##ENUMERATION_DESCRIPTION The types of controls that can be set as onwer-drawn
Public Enum OwnerDrawControlTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_MENU - An owner-drawn menu
    ODT_MENU = 1
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_LISTBOX - An owner-drawn listbox
    ODT_LISTBOX = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_COMBOBOX - An owner-drawn combo box
    ODT_COMBOBOX = 3
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_BUTTON - An owner drawn button
    ODT_BUTTON = 4
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_STATIC - An owner drawn static control
    ODT_STATIC = 5
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_HEADER - An owner drawn header control
    ODT_HEADER = 100
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_TAB - An owner drawn tab control
    ODT_TAB = 101
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODT_LISTVIEW - An owner drawn listview control
    ODT_LISTVIEW = 102
End Enum

' ##ENUMERATION_DESCRIPTION The drawing action required by an owner drawn control
Public Enum OwnerDrawActions
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODA_DRAWENTIRE - Draw the entire control
    ODA_DRAWENTIRE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODA_SELECT - Redraw to reflect changed selected item
    ODA_SELECT = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODA_FOCUS - Redraw to reflect changed focus state (gained or lost focus)
    ODA_FOCUS = &H4
End Enum

' ##ENUMERATION_DESCRIPTION The state(s) of the control that is being owner-drawn
Public Enum OwnerDrawnControlStates
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_SELECTED - The menu item is selected
    ODS_SELECTED = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_GRAYED - The menu is greyed out
    ODS_GRAYED = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_DISABLED - The control item is disabled
    ODS_DISABLED = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_CHECKED - The menu item is checked
    ODS_CHECKED = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_FOCUS - The control has the focus
    ODS_FOCUS = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_DEFAULT - The item is the default
    ODS_DEFAULT = &H20
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_COMBOBOXEDIT - The combo box edit section must be drawn
    ODS_COMBOBOXEDIT = &H1000
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_HOTLIGHT - The control is hot-tracking
    ODS_HOTLIGHT = &H40
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_INACTIVE - The contol item is inactive
    ODS_INACTIVE = &H80
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_NOACCEL - The control should be drawn without keyboard accelerator mnemonics
    ODS_NOACCEL = &H100
    ' ##ENUMERATION_MEMBER_DESCRIPTION ODS_NOFOCUSRECT - The control should be drawn without a focus rectangle
    ODS_NOFOCUSRECT = &H200
End Enum

'\\ Structures used in owner-drawn controls
Private Type DRAWITEMSTRUCT
    ctlType As OwnerDrawControlTypes
    CtlID As Long
    ItemId As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Declare Sub CopyMemoryDrawItem Lib "kernel32" Alias "RtlMoveMemory" (Destination As DRAWITEMSTRUCT, ByVal Source As Long, ByVal Length As Long)


Private Type MEASUREITEMSTRUCT
    ctlType As Long
    CtlID As Long
    ItemId As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type
Private Declare Sub CopyMemoryMeasureItem Lib "kernel32" Alias "RtlMoveMemory" (Destination As MEASUREITEMSTRUCT, ByVal Source As Long, ByVal Length As Long)

'--[ Device Management ]--------------------------------------------------------------------------------------------------------------------

' ##ENUMERATION_DESCRIPTION DeviceManagementChangeTypes
Public Enum DeviceManagementChangeTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_CONFIGCHANGECANCELED - Device management change event was cancelled
    DBT_CONFIGCHANGECANCELED = &H19
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_CONFIGCHANGED - Device management change has occured
    DBT_CONFIGCHANGED = &H18
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_CUSTOMEVENT - Custom device management event
    DBT_CUSTOMEVENT = &H8006
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICEARRIVAL - Device has been added
    DBT_DEVICEARRIVAL = &H8000
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICEQUERYREMOVE - Request to remove a device
    DBT_DEVICEQUERYREMOVE = &H8001
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICEQUERYREMOVEFAILED - Request to remove a device failed
    DBT_DEVICEQUERYREMOVEFAILED = &H8002
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICEREMOVECOMPLETE - Request to remove a device completed
    DBT_DEVICEREMOVECOMPLETE = &H8004
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICEREMOVEPENDING - Request to remove a device is pending
    DBT_DEVICEREMOVEPENDING = &H8003
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVICETYPESPECIFIC - Device specific device management event
    DBT_DEVICETYPESPECIFIC = &H8005
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_NO_DISK_SPACE - Disk is full
    DBT_NO_DISK_SPACE = &H47
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_LOW_DISK_SPACE - Disk space is running low
    DBT_LOW_DISK_SPACE = &H48
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_MONITORCHANGE - The display device has changed
    DBT_MONITORCHANGE = &H1B
End Enum

' ##ENUMERATION_DESCRIPTION DeviceManagementDeviceTypes - The classes of device handled by the device management system
Public Enum DeviceManagementDeviceTypes
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_OEM - OEM-defined device type
    DBT_DEVTYP_OEM = &H0
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_DEVNODE - Devnode number
    DBT_DEVTYP_DEVNODE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_VOLUME - Logical volume (drive or tape)
    DBT_DEVTYP_VOLUME = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_PORT - Serial or parallel port
    DBT_DEVTYP_PORT = &H3
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_NET - Network resource
    DBT_DEVTYP_NET = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_DEVICEINTERFACE - Device interface class
    DBT_DEVTYP_DEVICEINTERFACE = &H5
    ' ##ENUMERATION_MEMBER_DESCRIPTION DBT_DEVTYP_HANDLE - File system handle
    DBT_DEVTYP_HANDLE = &H6
End Enum

Private Type DEV_BROADCAST_HEADER
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
End Type
Private Declare Sub CopyMemoryDBHeader Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_HEADER, ByVal Source As Long, ByVal Length As Long)

Private Type DEV_BROADCAST_DEVNODE
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
    dwDevNode As Long
End Type
Private Declare Sub CopyMemoryDBDevNode Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_DEVNODE, ByVal Source As Long, ByVal Length As Long)

Private Type DEV_BROADCAST_VOLUME
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
    dwUnitMask As Long
    dwFlags As Long
End Type
Private Declare Sub CopyMemoryDBVolume Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_VOLUME, ByVal Source As Long, ByVal Length As Long)

Private Type DEV_BROADCAST_PORT
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
    lpName As Long
End Type
Private Declare Sub CopyMemoryDBPort Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_PORT, ByVal Source As Long, ByVal Length As Long)

Private Type DEV_BROADCAST_NET
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
    dwResource As Long
    dwFlags As Long
End Type
Private Declare Sub CopyMemoryDBNet Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_NET, ByVal Source As Long, ByVal Length As Long)

Private Type DEV_BROADCAST_DEVICEINTERFACE
    dwSize As Long
    dwDeviceType As DeviceManagementDeviceTypes
    dwReserved As Long
    dwGuid As Long
    lpName As Long
End Type
Private Declare Sub CopyMemoryDBDevInterface Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_DEVICEINTERFACE, ByVal Source As Long, ByVal Length As Long)

'----------------------------------------------------------------------------------------------------------------------------------------

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

' ##ENUMERATION_DESCRIPTION How a WM_PRINTCLIENT message is to be dealt with
Public Enum PrintClientFlags
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_CHECKVISIBLE Only draw the window if it is visible
    PRF_CHECKVISIBLE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_NONCLIENT Draw the non client part of the window
    PRF_NONCLIENT = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_CLIENT Draw the client part of the window
    PRF_CLIENT = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_ERASEBKGND Erase the target %device context:EventVB~ApiDeviceContext% before drawing the window
    PRF_ERASEBKGND = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_CHILDREN Draw the window's children as well
    PRF_CHILDREN = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION PRF_OWNED Draw any windows oned by this window as well
    PRF_OWNED = &H20
End Enum

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long

' ##ENUMERATION_DESCRIPTION Special effects for showing or hiding windows
Public Enum WindowAnimationCommands
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_HOR_POSITIVE Animate left to right horizontally
    AW_HOR_POSITIVE = &H1
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_HOR_NEGATIVE Animate right to left horizontally
    AW_HOR_NEGATIVE = &H2
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_VER_POSITIVE Animate top to bottom vertically
    AW_VER_POSITIVE = &H4
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_VER_NEGATIVE Animate bottom to top vertically
    AW_VER_NEGATIVE = &H8
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_CENTER Animate to/from the center
    AW_CENTER = &H10
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_HIDE Animate to hidden state
    AW_HIDE = &H10000
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_ACTIVATE Animate to shown state
    AW_ACTIVATE = &H20000
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_SLIDE Use a sliding animation
    AW_SLIDE = &H40000
    ' ##ENUMERATION_MEMBER_DESCRIPTION AW_BLEND Use a colour blending animation
    AW_BLEND = &H80000
End Enum

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal Filename As String, ByVal cch As Long) As Long

'\\ Overriding the drawing the window......
Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 32) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Private mPaintStruct As PAINTSTRUCT
Private bPainting As Boolean
Private bInPaintMessage As Boolean

'\\ -- Scrollable windows ---------------------------------------------------------------------------------------------------------------------------------------------------

' ##ENUMERATION_DESCRIPTION Specifies which scrollbar is changed by SetScrollPos
Public Enum ScrollbarLocations
    ' ##ENUMERATION_MEMBER_DESCRIPTION SB_CTL The scrollbar is a control
    SB_CTL = 2
    ' ##ENUMERATION_MEMBER_DESCRIPTION SB_HORZ The window's horizontal scrollbar
    SB_HORZ = 0
    ' ##ENUMERATION_MEMBER_DESCRIPTION SB_VERT The window's vertical scrollbar
    SB_VERT = 1
End Enum

'\\ Scrollbar management....
Private Type SCROLLINFO
    cbSize As Long
    fMask As ScrollbarInfoIndexes
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As ScrollbarLocations, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal nBar As ScrollbarLocations, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal nBar As ScrollbarLocations, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As ScrollbarLocations) As Long

'\\ Automatically scrolling and zooming the window...
Private bAutoScroll As Boolean
Private mlVirtualPageWidth As Long
Private mlVirtualPageHeight As Long
Private mlVirtualWidth As Long
Private mlVirtualHeight As Long

Private bChanged As Boolean

'-- CONTEXT SENSITIVE HELP ---------------------------------------------------

Public Enum ContextSensitiveHelpSources
     HELPINFO_WINDOW = &H1
     HELPINFO_MENUITEM = &H2
End Enum

Private Type HELPINFO
    cbSize As Long              '/* Size in bytes of this struct  */
    iContextType As ContextSensitiveHelpSources        '/* Either HELPINFO_WINDOW or HELPINFO_MENUITEM */
    iCtrlId As Long             '/* Control Id or a Menu item Id. */
    hItemHandle As Long         '/* hWnd of control or hMenu.     */
    dwContextId As Long         '/* Context Id associated with this item */
    MousePos As POINTAPI        '/* Mouse Position in screen co-ordinates */
End Type

Private Declare Sub CopyMemoryHelpInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As HELPINFO, ByVal Source As Long, ByVal Length As Long)

Dim mProperties As colProperties




' ##BLOCK_DESCRIPTION Animate - Apply the given animation effect to the window
' ##BLOCK_REMARKS Timespan - The duration of the effect (in milliseconds)
' ##BLOCK_REMARKS Effect - The type of effect to apply
' ##BLOCK_REMARKS Requires Win2000 or Win98 or higher
Public Sub Animate(ByVal Timespan As Long, ByVal Effect As WindowAnimationCommands)
Attribute Animate.VB_Description = "Animate - Apply the given animation effect to the window"
Attribute Animate.VB_HelpID = 2846

Dim lRet As Long
Dim lTime As Long

If Timespan = 0 Then
    lTime = 200
Else
    lTime = Timespan
End If

lRet = AnimateWindow(m_hwnd, lTime, Effect)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Animate", GetLastSystemError
End If

End Sub

Public Property Let AutoScroll(ByVal newSetting As Boolean)
Attribute AutoScroll.VB_HelpID = 3021

bAutoScroll = newSetting
'\\ For auto-scroll to work, the map mode must be mm_anisotropic..
If newSetting Then
    With DeviceContext
        .MappingMode = MM_ANISOTROPIC
    End With
End If

End Property

Public Property Get AutoScroll() As Boolean

AutoScroll = bAutoScroll

End Property

Public Property Get Button() As ApiButtonControl
Attribute Button.VB_HelpID = 2847

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_BUTTON Then
        Dim bcThis As ApiButtonControl
        Set bcThis = New ApiButtonControl
        Set bcThis.Window = Me
        Set Button = bcThis
    End If
End If

End Property



Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property

Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

'  ##BLOCK_DESCRIPTION This is the height of the window border (top and bottom) in pixels
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get BorderHeight() As Long
Attribute BorderHeight.VB_Description = "This is the height of the window border (top and bottom) in pixels"
Attribute BorderHeight.VB_HelpID = 2234

If BorderStyle = vbBSNone Then
    BorderHeight = 0
Else
    If IsWindowStyleSet(WS_THICKFRAME) Then
        BorderHeight = APIDispenser.System.Metrics(SM_CXFRAME)
    Else
        BorderHeight = APIDispenser.System.Metrics(SM_CXEDGE)
    End If
End If

End Property
' ## BLOCK_DESCRIPTION The BorderStyle allows you to get the border style of the window in VB terms
' ## BLOCK_REMARKS This returns a VB FormBorderStyleConstants value
Public Property Get BorderStyle() As FormBorderStyleConstants
Attribute BorderStyle.VB_Description = "BLOCK_DESCRIPTION The BorderStyle allows you to get the border style of the window in VB terms"
Attribute BorderStyle.VB_HelpID = 2792

Dim bsThis As FormBorderStyleConstants

    'vbSizableToolWindow
    'vbSizable
    'vbFixedToolWindow
    'vbFixedSingle
    'vbFixedDialog

If Not IsWindowStyleSet(WS_BORDER) Then
    bsThis = vbBSNone
Else
    If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
        'Window is a tool window
        If IsWindowStyleSet(WS_THICKFRAME) Then
            bsThis = vbSizableToolWindow
        Else
            bsThis = vbFixedToolWindow
        End If
    Else
        'Window is not a tool window
        If IsWindowStyleSet(WS_THICKFRAME) Then
            bsThis = vbSizable
        Else
            If IsWindowStyleSet(WS_DLGFRAME) Then
                bsThis = vbFixedDialog
            Else
                bsThis = vbFixedSingle
            End If
        End If
    End If
End If

BorderStyle = bsThis

End Property

'  ##BLOCK_DESCRIPTION This is the width of the window border (left and right) in pixels
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_Description = "This is the width of the window border (left and right) in pixels"
Attribute BorderWidth.VB_HelpID = 2236

If BorderStyle = vbBSNone Then
    BorderWidth = 0
Else
    If IsWindowStyleSet(WS_THICKFRAME) Then
        BorderWidth = APIDispenser.System.Metrics(SM_CYFRAME)
    Else
        BorderWidth = APIDispenser.System.Metrics(SM_CYEDGE)
    End If
End If

End Property

'  ##BLOCK_DESCRIPTION This is the height of the caption in pixels
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get CaptionHeight() As Long
Attribute CaptionHeight.VB_Description = "This is the height of the caption in pixels"
Attribute CaptionHeight.VB_HelpID = 2237

    If IsWindowStyleSet(WS_CAPTION) Then
        If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
            CaptionHeight = APIDispenser.System.Metrics(SM_CYSMCAPTION)
        Else
            CaptionHeight = APIDispenser.System.Metrics(SM_CYCAPTION)
        End If
    Else
        CaptionHeight = 0
    End If
    
End Property

' ##BLOCK_DESCRIPTION Child Windows are those windows owned as children of this window
' ##BLOCK_REMARKS Windows set to be owned by this window by the %set parent:EventVB~ApiWindow~Parent% function may not be added to this list
Public Property Get ChildWindows() As Collection
Attribute ChildWindows.VB_Description = "Child Windows are those windows owned as children of this window"
Attribute ChildWindows.VB_HelpID = 2238

Dim colChildren As Collection
Dim wndThis As ApiWindow
Dim PrevHwnd As Long

'\\ Initialise the collection
Set colChildren = New Collection

Set wndThis = GetWindow(GW_CHILD)
If Not (wndThis Is Nothing) Then
    colChildren.Add wndThis
    PrevHwnd = wndThis.hwnd
    Set wndThis = GetWindow(GW_HWNDNEXT, PrevHwnd)
    While Not (wndThis Is Nothing)
        If wndThis.hwnd <> PrevHwnd Then
            colChildren.Add wndThis
            PrevHwnd = wndThis.hwnd
            Set wndThis = GetWindow(GW_HWNDNEXT, PrevHwnd)
        Else
            Set wndThis = Nothing
        End If
    Wend
End If

Set ChildWindows = colChildren

End Property


' ##BLOCK_DESCRIPTION The class name uniquely identifies the class of which this window is an instance
' ##BLOCK_REMARKS Standard class names include:
' ##BLOCK_REMARKS "STATIC" for label type windows
' ##BLOCK_REMARKS "EDIT" for edit box type windows etc.
Public Property Get ClassName() As String
Attribute ClassName.VB_Description = "The class name uniquely identifies the class of which this window is an instance"
Attribute ClassName.VB_HelpID = 2239

Dim sClassName As String
Dim lRet As Long

'\\ Get the class name
sClassName = String$(1024, 0)
lRet = GetClassNameApi(m_hwnd, sClassName, 1024)
If (Err.LastDllError = 0) And (lRet > 0) Then
    ClassName = Left$(sClassName, lRet)
Else
    ReportError Err.LastDllError, CLSNAME & ":Classname", GetLastSystemError
End If

End Property
' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get AttachedClipboard() As ApiClipboard
Attribute AttachedClipboard.VB_HelpID = 2240

Dim mClipboard As ApiClipboard

Set mClipboard = New ApiClipboard
mClipboard.ParenthWnd = hwnd
Set AttachedClipboard = mClipboard

End Property

Public Property Get ClientRect() As APIRect
Dim mRect As APIRect
Dim lpRect As RECT
Dim lRet As Long

Set mRect = New APIRect
lRet = GetClientRect(m_hwnd, lpRect)
If Err.LastDllError = 0 Then
    With mRect
        .Left = lpRect.Left
        .Right = lpRect.Right
        .Top = lpRect.Top
        .Bottom = lpRect.Bottom
    End With
End If
Set ClientRect = mRect


End Property

Public Property Get ComboBox() As ApiComboControl
Attribute ComboBox.VB_HelpID = 2848

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_COMBOBOX Then
        Dim cbThis As ApiComboControl
        Set cbThis = New ApiComboControl
        Set cbThis.Window = Me
        Set ComboBox = cbThis
    End If
End If

End Property

Public Property Let ContextSensitiveHelp(ByVal newValue As Boolean)
Attribute ContextSensitiveHelp.VB_Description = "This property is true if the window has a context sensitive help button  "
Attribute ContextSensitiveHelp.VB_HelpID = 3023

If newValue Then
    SetWindowExtendedStyle WS_EX_CONTEXTHELP
Else
    UnSetWindowExtendedStyle WS_EX_CONTEXTHELP
End If

End Property

' ##BLOCK_DESCRIPTION This property is true if the window has a context sensitive help button
' ##BLOCK_REMARKS This is only supported on newer versions of the Windows operating system
' ##BLOCK_REMARKS Also, this cannot be set if the Minimise or Maximise box is set
Public Property Get ContextSensitiveHelp() As Boolean

    ContextSensitiveHelp = IsWindowExtendedStyleSet(WS_EX_CONTEXTHELP)

End Property

Public Property Let ControlBox(ByVal newSetting As Boolean)
Attribute ControlBox.VB_Description = "This property is true if the window has a control box  "
Attribute ControlBox.VB_HelpID = 2242

If newSetting Then
    SetWindowStyle WS_SYSMENU
Else
    UnSetWindowStyle WS_SYSMENU
End If

End Property

' ##BLOCK_DESCRIPTION This property is true if the window has a control box
' ##BLOCK_REMARKS On a standard window this control box contains the window's icon
Public Property Get ControlBox() As Boolean

ControlBox = IsWindowStyleSet(WS_SYSMENU)

End Property

'  ##BLOCK_DESCRIPTION This is the rectangle describing the location of the control box relative to the window as a whole
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get ControlBoxRectangle() As APIRect
Attribute ControlBoxRectangle.VB_Description = "This is the rectangle describing the location of the control box relative to the window as a whole"
Attribute ControlBoxRectangle.VB_HelpID = 2243

Dim rcThis As New APIRect

With rcThis
    .Right = (Width - BorderWidth)
    .Top = BorderHeight
    If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSMSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSMSIZE)
    Else
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSIZE)
    End If
End With

Set ControlBoxRectangle = rcThis

End Property

' ##BLOCK_DESCRIPTION Copies (and returns) any selected text in the window
' ##BLOCK_REMARKS This function may overwrite any string data in the clipboard
Public Function CopyText() As String
Attribute CopyText.VB_Description = "Copies (and returns) any selected text in the window"
Attribute CopyText.VB_HelpID = 2244

Dim lRet As Long
Dim sText As String

'\\ First copy the text into the clipboard buffer...
lRet = SendMessage(WM_COPY, 0, 0)

'\\ Now get that data out of the clipboard data area...
If Clipboard.GetFormat(vbCFText) Then
    sText = Clipboard.GetText()
End If

CopyText = sText

End Function

' ##BLOCK_DESCRIPTION CreateNewChildControl - Creates a standard button control and adds it to this window's children
' ##BLOCK_REMARKS WindowType - The base type of the control
' ##BLOCK_REMARKS Position - Where the window will be positioned relative to the client area (in pixels)
' ##BLOCK_REMARKS Style - The style bits that will be set for the window
' ##BLOCK_REMARKS ExtendedStyle - The extended style bits that will be set for the window
' ##BLOCK_REMARKS ButtonStyle - The style bits that only apply to button controls
' ##BLOCK_REMARKS EditStyle - The style bits that only apply to textbox controls
' ##BLOCK_REMARKS ListboxStyle - The style bits that only apply to listbox controls
' ##BLOCK_REMARKS ListboxStyle - The style bits that only apply to scrollbar controls
' ##BLOCK_REMARKS StaticStyle - The style bits that only apply to static controls
' ##BLOCK_REMARKS Caption  - The caption to use on the button
Public Function CreateNewChildControl(ByVal WindowType As StandardControlTypes, ByVal Position As APIRect, ByVal Style As WindowStyleBits, ByVal ExtendedStyle As WindowStyleExtendedBits, ByVal ButtonStyle As ButtonControlStyles, ByVal EditStyle As EditControlStyles, ByVal ListboxStyle As ListboxControlStyles, ByVal ScrollbarStyle As ScrollbarStyles, ByVal StaticStyle As StaticControlStyles, ByVal Caption As String) As ApiWindow
Attribute CreateNewChildControl.VB_Description = "CreateNewChildControl - Creates a standard button control and adds it to this window's children"
Attribute CreateNewChildControl.VB_HelpID = 3175

Dim sType As String
Dim wndThis As ApiWindow
Dim clsThis As APIWndClass
Dim lStyle As Long

sType = EventVB_Utils.StandardControlClassName(WindowType)

lStyle = Style
Select Case WindowType
Case Control_BUTTON
    lStyle = lStyle Or ButtonStyle
Case Control_EDIT
    lStyle = lStyle Or EditStyle
Case Control_LISTBOX
    lStyle = lStyle Or ListboxStyle
Case Control_SCROLLBAR
    lStyle = lStyle Or ScrollbarStyle
Case Control_STATIC
    lStyle = lStyle Or StaticStyle
End Select

Set clsThis = New APIWndClass
With clsThis
    .CreateFromClassname 0, sType
    .Name = sType
    .Style = lStyle
End With


With Position
    Set wndThis = APIDispenser.System.WindowDispenser.CreateWindow(ExtendedStyle, clsThis, Caption, Style Or ButtonStyle, .Left, .Top, .Right - .Left, .Bottom - .Top, m_hwnd, 0, 0, 0)
End With
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & "CreateChildButton", GetLastSystemError
End If
Set wndThis.Parent = Me
Set CreateNewChildControl = wndThis

End Function

Public Sub Destroy()
Attribute Destroy.VB_HelpID = 3024

If m_hwnd <> 0 Then
    Call DestroyWindow(m_hwnd)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":Destroy", GetLastSystemError
    End If
End If

End Sub

' ##BLOCK_DESCRIPTION The %device context:EventVB~ApiDeviceContext% is the canvas upon which drawing operations for a window client area occur
' ##BLOCK_REMARKS This also provides information about the graphical capabilities of the window
Public Property Get DeviceContext() As ApiDeviceContext
Attribute DeviceContext.VB_Description = "The device context:EventVB~ApiDeviceContext is the canvas upon which drawing operations for a window client area occur"
Attribute DeviceContext.VB_HelpID = 2793

Dim lRet As Long

If mDeviceContext Is Nothing Then
    Set mDeviceContext = New ApiDeviceContext
End If
If bPainting Then
    lRet = mPaintStruct.hdc
Else
    lRet = GetDCApi(m_hwnd)
End If
If Err.LastDllError = 0 And lRet <> 0 Then
    mDeviceContext.hdc = lRet
End If
Set DeviceContext = mDeviceContext

End Property

' ##BLOCK_DESCRIPTION The %device context:EventVB~ApiDeviceContext% is the canvas upon which drawing operations for the entire window (including non client area) occur
' ##BLOCK_REMARKS This also provides information about the graphical capabilities of the window as a whole
Public Property Get DeviceContextNC() As ApiDeviceContext
Attribute DeviceContextNC.VB_Description = "The device context:EventVB~ApiDeviceContext is the canvas upon which drawing operations for the entire window (including non client area) occur"
Attribute DeviceContextNC.VB_HelpID = 2246

Dim lRet As Long

If mDeviceContextNC Is Nothing Then
    Set mDeviceContextNC = New ApiDeviceContext
End If
lRet = GetWindowDC(m_hwnd)
If Err.LastDllError = 0 And lRet <> 0 Then
    mDeviceContextNC.hdc = lRet
End If
Set DeviceContextNC = mDeviceContextNC

End Property


' ##BLOCK_DESCRIPTION This function refreshes and repaints the area of the window occupied by the menu bar
' ##BLOCK_REMARKS If you alter the main window menu bar you should call this function to repaint it
Public Sub DrawMenuBar()
Attribute DrawMenuBar.VB_Description = "This function refreshes and repaints the area of the window occupied by the menu bar"
Attribute DrawMenuBar.VB_HelpID = 2247

Dim lRet As Long

lRet = DrawMenuBarApi(m_hwnd)
If (lRet = 0) Or (Err.LastDllError > 0) Then
    Call ReportError(Err.LastDllError, CLSNAME & ":DrawMenuBar:", GetLastSystemError)
End If

End Sub





Public Property Get Edit() As ApiEditControl
Attribute Edit.VB_HelpID = 2849

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_EDIT Then
        Dim edThis As ApiEditControl
        Set edThis = New ApiEditControl
        Set edThis.Window = Me
        Set Edit = edThis
    End If
End If

End Property

Public Property Get Filename() As String
Attribute Filename.VB_HelpID = 3025

#If PRODUCTION Then
    On Error Resume Next
#End If

Dim lRet As Long
Dim sFilename As String

sFilename = String$(1024, 0)
lRet = GetWindowModuleFileName(m_hwnd, sFilename, Len(sFilename))
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Filename", GetLastSystemError
Else
    If lRet > 0 Then
        sFilename = Left$(sFilename, lRet)
        Filename = sFilename
    End If
End If

End Property

' ##BLOCK_DESCRIPTION Returns the window relative to this one according to the %wCmd:EventVB~enGetWindow% relationship
Public Function GetWindow(wCmd As enGetWindow, Optional PrevHwnd As Long) As ApiWindow
Attribute GetWindow.VB_Description = "Returns the window relative to this one according to the wCmd:EventVB~enGetWindow relationship"
Attribute GetWindow.VB_HelpID = 2249

Dim hwnd As Long
Dim Window As ApiWindow

If PrevHwnd = 0 Then
    hwnd = GetWindowApi(m_hwnd, wCmd)
Else
    hwnd = GetWindowApi(PrevHwnd, wCmd)
End If
If IsWindowApi(hwnd) Then
    '\\ If its already subclassed, return that instance...
    On Error Resume Next
    Set Window = AllSubclassedWindows.Item(hwnd)
    On Error GoTo 0
    If Window Is Nothing Then
        Set Window = New ApiWindow
        Window.hwnd = hwnd
    End If
    Set GetWindow = Window
End If

End Function

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Function GetWindowLong(Index As enGetWindowLong) As Long
Attribute GetWindowLong.VB_HelpID = 2250

Dim lRet As Long

lRet = GetWindowLongApi(m_hwnd, Index)
If Err.LastDllError = 0 Then
    GetWindowLong = lRet
End If

End Function

' ##BLOCK_DESCRIPTION The hdc property is a handle to the %device context:EventVB~ApiDeviceContext%
' ##BLOCK_REMARKS This handle can be passed to API calls that require access to the device context
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "The hdc property is a handle to the device context:EventVB~ApiDeviceContext"
Attribute hdc.VB_HelpID = 2251

hdc = GetDCApi(m_hwnd)

End Property

' ##BLOCK_DESCRIPTION This is the height of the window
' ##BLOCK_REMARKS This is equal to the bottom minus the top of the window rectangle
' ##BLOCK_REMARKS This value is expressed in twips
Public Property Get Height() As Long
Attribute Height.VB_Description = "This is the height of the window"
Attribute Height.VB_HelpID = 2252

Dim lpRect As RECT
Call GetWindowRect(m_hwnd, lpRect)
Height = (lpRect.Bottom - lpRect.Top) * Screen.TwipsPerPixelY
    
End Property

' ##BLOCK_DESCRIPTION Returns the position of the x,y coordinate (in client coordinates)  in terms of what would be hit if the mouse were clicked there
' ##BLOCK_REMARKS This is based on the %system metrics:EventVB~ApiSystem~Metrics%
Public Function HitTest(ByVal x As Long, ByVal y As Long) As enHitTestResult
Attribute HitTest.VB_Description = "Returns the position of the x,y coordinate (in client coordinates)  in terms of what would be hit if the mouse were clicked there"
Attribute HitTest.VB_HelpID = 3185

Dim htThis As enHitTestResult

With RECT
    If x < 0 Or y < 0 Or x > Width Or y > Height Then
        htThis = HTNOWHERE
    Else
        If x < BorderWidth Then
            '\\ Somewhere in the lefthand border
            htThis = HTLEFT
        ElseIf x >= (Width - BorderWidth) Then
            '\\ Somewhere in the righthand border
            htThis = HTRIGHT
        Else
            If y < BorderHeight Then
                htThis = HTTOP
            ElseIf y >= Height - BorderHeight Then
                htThis = HTBOTTOM
            ElseIf y <= CaptionHeight Then
                htThis = HTCAPTION
                '\\ If the window has the various caption buttons, test these
                If ControlBox Then
                    If ControlBoxRectangle.ContainsPoint(x, y) Then
                        htThis = HTCLOSE
                    ElseIf SysmenuBoxRectangle.ContainsPoint(x, y) Then
                        htThis = HTSYSMENU
                    End If
                End If
                If MinButton Then
                    If MinimiseBoxRectangle.ContainsPoint(x, y) Then
                        htThis = HTMINBUTTON
                    End If
                End If
                If MaxButton Then
                    If MaximiseBoxRectangle.ContainsPoint(x, y) Then
                        htThis = HTMAXBUTTON
                    End If
                End If
            Else
                htThis = HTCLIENT
            End If
        End If
    End If
End With

HitTest = htThis

End Function

' ##BLOCK_DESCRIPTION Prepares a window for custom painting
' ##BLOCK_REMARKS This method is only valid within a %paint:EventVB~ApiWindow~Paint_EV% event
Public Sub InitialiseCustomPaint()
Attribute InitialiseCustomPaint.VB_Description = "Prepares a window for custom painting"
Attribute InitialiseCustomPaint.VB_HelpID = 3027

Dim lRet As Long

If bInPaintMessage Then
    lRet = BeginPaint(m_hwnd, mPaintStruct)
    bPainting = True
End If

End Sub

Public Property Get Listbox() As ApiListBoxControl
Attribute Listbox.VB_HelpID = 2851

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_LISTBOX Then
        Dim lbThis As ApiListBoxControl
        Set lbThis = New ApiListBoxControl
        Set lbThis.Window = Me
        Set Listbox = lbThis
    End If
End If

End Property

Public Property Set Menu(ByVal newmenu As ApiMenu)
Attribute Menu.VB_Description = "Top level and MDI child windows can have an a menu associated with them  "
Attribute Menu.VB_HelpID = 2795

Dim lRet As Long

lRet = SetMenu(m_hwnd, newmenu.hMenu)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Menu (set)", GetLastSystemError
End If

End Property

Public Property Let MouseHoverTracking(ByVal newSetting As Boolean)
Attribute MouseHoverTracking.VB_Description = "MouseHoverTracking generates mouse hover:EventVB~ApiWindow~MouseHover event  "
Attribute MouseHoverTracking.VB_HelpID = 2255

If newSetting Then
    mheThis.dwFlags = mheThis.dwFlags Or TME_HOVER
Else
    mheThis.dwFlags = mheThis.dwFlags And Not TME_HOVER
End If

End Property

' ##BLOCK_DESCRIPTION MouseHoverTracking generates %mouse hover:EventVB~ApiWindow~MouseHover% event
' ##BLOCK_REMARKS This event is only supported on Windows 95 and beyond
Public Property Get MouseHoverTracking() As Boolean

Call RefreshTrackMouseStruct
MouseHoverTracking = (mheOut.dwFlags And TME_HOVER) = mheOut.dwFlags

End Property


' ##BLOCK_DESCRIPTION This is the unique window handle identifying a window instance
' ##BLOCK_REMARKS Set this value to link the %window class:EventVB~ApiWindow% to an actual window instance
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "This is the unique window handle identifying a window instance  "
Attribute hwnd.VB_HelpID = 2256

    hwnd = m_hwnd
    
End Property

Public Property Let hwnd(ByVal newHwnd As Long)

Dim lRet As Long

    If newHwnd <> m_hwnd Then
        '\\ Must reset the proc address before closing
        If mOldProcAddress > 0 Then
            lRet = SetWindowLong(GWL_WNDPROC, mOldProcAddress)
            mOldProcAddress = 0
        End If
        '\\ Must unset the menus...
        Set mSystemMenu = Nothing
        Set mMenu = Nothing
        
        m_hwnd = newHwnd
    End If
    
End Property

' ##BLOCK_DESCRIPTION This marks the given %rectange:EventVB~APIRect% as needing repainted
' ##BLOCK_REMARKS The repainting will occur in the next %paint:EventVB~ApiWindow~Paint% event
Public Sub Invalidate(ByVal RectIn As APIRect)
Attribute Invalidate.VB_Description = "This marks the given rectange:EventVB~APIRect as needing repainted"
Attribute Invalidate.VB_HelpID = 2257

Dim rcThis As RECT
Dim lRet As Long

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Top = RectIn.Top
    .Right = RectIn.Right
End With

lRet = InvalidateRect(m_hwnd, rcThis, True)

End Sub


' ##BLOCK_DESCRIPTION Checks for the style bit specified in this window
Public Function IsWindowStyleSet(ByVal TheStyle As WindowStyleBits) As Boolean
Attribute IsWindowStyleSet.VB_Description = "Checks for the style bit specified in this window"
Attribute IsWindowStyleSet.VB_HelpID = 2258

Dim lStyle As Long

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_STYLE)
IsWindowStyleSet = (lStyle And TheStyle)

End Function





' ##BLOCK_DESCRIPTION Checks for the style bit specified in this window
Public Function IsWindowExtendedStyleSet(ByVal TheStyle As WindowStyleExtendedBits) As Boolean
Attribute IsWindowExtendedStyleSet.VB_Description = "Checks for the style bit specified in this window"
Attribute IsWindowExtendedStyleSet.VB_HelpID = 3178

Dim lStyle As Long

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_EXSTYLE)
IsWindowExtendedStyleSet = (lStyle And TheStyle)

End Function
Public Property Let MaxButton(ByVal newSetting As Boolean)
Attribute MaxButton.VB_Description = "This is true if the window has a maximise (or restore) button  "
Attribute MaxButton.VB_HelpID = 2260

If newSetting Then
    SetWindowStyle WS_MAXIMIZEBOX
Else
    UnSetWindowStyle WS_MAXIMIZEBOX
End If

End Property

' ##BLOCK_DESCRIPTION This is true if the window has a maximise (or restore) button
Public Property Get MaxButton() As Boolean

MaxButton = IsWindowStyleSet(WS_MAXIMIZEBOX)

End Property

'  ##BLOCK_DESCRIPTION This is the rectangle describing the location of the maximise (or restore) box relative to the window as a whole
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get MaximiseBoxRectangle() As APIRect
Attribute MaximiseBoxRectangle.VB_Description = "This is the rectangle describing the location of the maximise (or restore) box relative to the window as a whole"
Attribute MaximiseBoxRectangle.VB_HelpID = 2797

Dim rcThis As New APIRect

With rcThis
    .Right = (Width - BorderWidth)
    .Top = BorderHeight
    If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
        .Right = .Right - APIDispenser.System.Metrics(SM_CXSMSIZE)
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSMSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSMSIZE)
    Else
        .Right = .Right - APIDispenser.System.Metrics(SM_CXSIZE)
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSIZE)
    End If
End With

Set MaximiseBoxRectangle = rcThis

End Property

' ##BLOCK_DESCRIPTION Top level and MDI child windows can have an a menu associated with them
' ##BLOCK_REMARKS This property allows you to manipulate that menu
Public Property Get Menu() As ApiMenu

Dim lRet As Long

lRet = GetMenuApi(m_hwnd)
If Err.LastDllError = 0 And lRet > 0 Then
    If mMenu Is Nothing Then
        Set mMenu = New ApiMenu
    End If
    mMenu.hMenu = lRet
    Set Menu = mMenu
Else
    ReportError Err.LastDllError, CLSNAME & ":Menu (get)", GetLastSystemError
End If

End Property

Public Property Let MinButton(ByVal newSetting As Boolean)
Attribute MinButton.VB_Description = "This is true if the window has a minimise button  "
Attribute MinButton.VB_HelpID = 2264

If newSetting Then
    SetWindowStyle WS_MINIMIZEBOX
Else
    UnSetWindowStyle WS_MINIMIZEBOX
End If

End Property

' ##BLOCK_DESCRIPTION This is true if the window has a minimise button
Public Property Get MinButton() As Boolean

MinButton = IsWindowStyleSet(WS_MINIMIZEBOX)

End Property

'  ##BLOCK_DESCRIPTION This is the rectangle describing the location of the minimise box relative to the window as a whole
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%
Public Property Get MinimiseBoxRectangle() As APIRect
Attribute MinimiseBoxRectangle.VB_Description = "This is the rectangle describing the location of the minimise box relative to the window as a whole"
Attribute MinimiseBoxRectangle.VB_HelpID = 2265

Dim rcThis As New APIRect

With rcThis
    .Right = (Width - BorderWidth)
    .Top = BorderHeight
    If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
        .Right = .Right - (APIDispenser.System.Metrics(SM_CXSMSIZE) * 2)
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSMSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSMSIZE)
    Else
        .Right = .Right - (APIDispenser.System.Metrics(SM_CXSIZE) * 2)
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSIZE)
        .Left = .Right - APIDispenser.System.Metrics(SM_CXSIZE)
    End If
End With

Set MinimiseBoxRectangle = rcThis

End Property

Public Property Let MouseCapture(ByVal CaptureMouse As Boolean)
Attribute MouseCapture.VB_Description = "The window that has the capture will recieve all the mouse messages even when it isn't the focus window  "
Attribute MouseCapture.VB_HelpID = 2267

Dim lRet As Long

If CaptureMouse Then
    mPrevCapture = SetCaptureApi(m_hwnd)
Else
    lRet = SetCaptureApi(mPrevCapture)
End If

End Property

' ##BLOCK_DESCRIPTION The window that has the capture will recieve all the mouse messages even when it isn't the focus window
' ##BLOCK_REMARKS The capture is independent of the focus
Public Property Get MouseCapture() As Boolean

MouseCapture = (GetCaptureApi() = m_hwnd)

End Property


Public Property Let MouseHoverTimeout(ByVal newTimeout As Long)
Attribute MouseHoverTimeout.VB_Description = "The mouse hover timeout is the amount of time that the mouse must remain over a window that has mouse hover tracking:EventVB~ApiWindow~MouseHoverTracking set  "
Attribute MouseHoverTimeout.VB_HelpID = 2269

Dim lRet As Long

With mheThis
    .cbSize = Len(mheThis)
    .hwndTrack = m_hwnd
    .dwHoverTime = newTimeout
    .dwFlags = (TME_HOVER Or TME_LEAVE)
End With

lRet = TrackMouseEvent(mheThis)

End Property

' ##BLOCK_DESCRIPTION The mouse hover timeout is the amount of time that the mouse must remain over a window that has %mouse hover tracking:EventVB~ApiWindow~MouseHoverTracking% set
' ##BLOCK_REMARKS The %mouse hover:EventVB~ApiWindow~MouseHover% event is triggered when this time elapses
Public Property Get MouseHoverTimeout() As Long

Call RefreshTrackMouseStruct
MouseHoverTimeout = mheOut.dwHoverTime

End Property

Public Property Let MouseLeaveTracking(ByVal newSetting As Boolean)
Attribute MouseLeaveTracking.VB_Description = "MouseLeaveTracking generates mouse leave:EventVB~ApiWindow~MouseLeave event  "
Attribute MouseLeaveTracking.VB_HelpID = 2271

If newSetting Then
    mheThis.dwFlags = mheThis.dwFlags Or TME_LEAVE
Else
    mheThis.dwFlags = mheThis.dwFlags And Not TME_LEAVE
End If

End Property

' ##BLOCK_DESCRIPTION MouseLeaveTracking generates %mouse leave:EventVB~ApiWindow~MouseLeave% event
' ##BLOCK_REMARKS This event is only supported on Windows 95 and beyond
Public Property Get MouseLeaveTracking() As Boolean

Call RefreshTrackMouseStruct
MouseLeaveTracking = (mheOut.dwFlags And TME_LEAVE) = mheOut.dwFlags

End Property

Public Sub Move(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Repaint As Boolean)
Attribute Move.VB_HelpID = 3187

Dim lRet As Long
lRet = MoveWindow(m_hwnd, x, y, Width, Height, Repaint)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Move", GetLastSystemError
End If

End Sub

Friend Property Get OldProcAddress() As Long

    OldProcAddress = mOldProcAddress
    
End Property

Public Property Set Parent(ByVal NewParent As ApiWindow)
Attribute Parent.VB_HelpID = 2800

If Not NewParent Is Nothing Then
    Call SetParent(m_hwnd, NewParent.hwnd)
Else
    Call SetParent(m_hwnd, HWND_DESKTOP)
End If
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Parent", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Parent() As ApiWindow

Dim hWndParent As Long
Dim ParentWindow As ApiWindow

If IsWindowApi(m_hwnd) Then
    hWndParent = GetParentApi(m_hwnd)
    If IsWindowApi(hWndParent) Then
        '\\ If its already subclassed, return that instance...
        Set ParentWindow = AllSubclassedWindows.Item(hWndParent)
        If ParentWindow Is Nothing Then
            '\\ Otherwise create a new instance
            Set ParentWindow = New ApiWindow
            ParentWindow.hwnd = hWndParent
        End If
        Set Parent = ParentWindow
    End If
End If


End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Picture() As StdPicture
Attribute Picture.VB_HelpID = 2272

Dim picThis As StdPicture
Set picThis = New StdPicture

Dim hCompatibleDC As Long
Dim lRet As Long

'\\ Need to get the picture from this window...
hCompatibleDC = CreateCompatibleDC(DeviceContext.hdc)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Picture", GetLastSystemError
End If
lRet = SendMessageLong(hwnd, WM_PRINTCLIENT, hCompatibleDC, PRF_CLIENT)
'\\ Todo: How to get the device context to a StdPicture?

Set Picture = picThis

End Property

Public Property Get Properties() As colProperties
Attribute Properties.VB_HelpID = 3030

If mProperties Is Nothing Then
    Set mProperties = New colProperties
End If

Call APICallbackProcs.EnumerateProperties(Me)

Set Properties = mProperties

End Property

' ##BLOCK_DESCRIPTION RECT - The rectangle that completely encloses this window
' ##BLOCK_REMARKS The rectangle is relative to the window parent, or the screen if the window is a top level window
Public Property Get RECT() As APIRect
Attribute RECT.VB_Description = "RECT - The rectangle that completely encloses this window"
Attribute RECT.VB_HelpID = 2273

Dim mRect As APIRect
Dim lpRect As RECT
Dim lRet As Long

Set mRect = New APIRect
lRet = GetWindowRect(m_hwnd, lpRect)
If Err.LastDllError = 0 Then
    With mRect
        .Left = lpRect.Left
        .Right = lpRect.Right
        .Top = lpRect.Top
        .Bottom = lpRect.Bottom
    End With
End If
Set RECT = mRect

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Sub Refresh()
Attribute Refresh.VB_HelpID = 2274

Dim lRet As Long

lRet = InvalidateRectByPointerApi(m_hwnd, 0, True)
Call UpdateWindow(m_hwnd)

End Sub


Private Sub RefreshTrackMouseStruct()

Dim lRet As Long


    With mheOut
        .cbSize = Len(mheThis)
        .dwFlags = TME_QUERY
        .hwndTrack = m_hwnd
    End With
    
    If APIDispenser.System.IsRequiredSystem("TrackMouse", ver_Win_98) Then
        lRet = TrackMouseEvent(mheOut)
        If Err.LastDllError Then
            ReportError Err.LastDllError, CLSNAME & ":MouseHoverTimeout (get)", GetLastSystemError
        End If
    End If


End Sub



Public Property Set Region(ByVal newRegion As ApiRegion)
Attribute Region.VB_HelpID = 2276

Dim lRet As Long

If newRegion Is Nothing Then
    Set mRegion = Nothing
Else
    If newRegion.hRgn <> Region.hRgn Then
        Set mRegion = newRegion
        lRet = SetWindowRgn(m_hwnd, mRegion.hRgn, True)
    End If
End If

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Region() As ApiRegion

Dim lRet As Long

If mRegion Is Nothing Then
    Set mRegion = New ApiRegion
End If
lRet = GetWindowRgn(m_hwnd, mRegion.hRgn)
If Err.LastDllError = 0 And lRet > 0 Then
    mRegion.hRgn = lRet
End If
Set Region = mRegion

End Property

Public Property Get Scrollbar() As ApiScrollbarControl
Attribute Scrollbar.VB_HelpID = 2854

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_SCROLLBAR Then
        Dim scThis As ApiScrollbarControl
        Set scThis = New ApiScrollbarControl
        Set scThis.Window = Me
        Set Scrollbar = scThis
    End If
End If

End Property

Public Property Let Scrollbars(ByVal newSetting As ScrollBarConstants)
Attribute Scrollbars.VB_HelpID = 3032

Call UnSetWindowStyle(WS_HSCROLL)
Call UnSetWindowStyle(WS_VSCROLL)
If newSetting = vbBoth Or newSetting = vbHorizontal Then
    Call SetWindowStyle(WS_HSCROLL)
End If
If newSetting = vbBoth Or newSetting = vbVertical Then
    Call SetWindowStyle(WS_VSCROLL)
End If
'\\ Force repaint of the non client area...
Call SetWindowPos(m_hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_FRAMECHANGED)

End Property

Public Property Get Scrollbars() As ScrollBarConstants

Dim sOut As ScrollBarConstants

If IsWindowStyleSet(WS_HSCROLL) Then
    sOut = vbHorizontal + sOut
End If
If IsWindowStyleSet(WS_VSCROLL) Then
    sOut = vbVertical + sOut
End If

Scrollbars = sOut

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Sub Scroll(ByVal dx As Long, ByVal dy As Long, ScrollMethod As ScrollWindowModes)
Attribute Scroll.VB_HelpID = 2277

With DeviceContext
    Call .ScrollWindow(dx, dy)
    Call .ScrollViewport(dx, dy)
End With

'\\ Move all the child windows
Call OffsetWindows(ChildWindows, dx, dy, Me)

'With rcUpdate
'    If dX < 0 Then
'
'    Else
'
'    End If
'    If dY < 0 Then
'
'    Else
'
'    End If
'End With
'Call InvalidateRect(m_hwnd, rcUpdate, True)
Call UpdateWindow(m_hwnd)

End Sub

'\\ --[SelectedText]----------------------------------------------------------------------------
'\\ Returns the text that is in the currently selected part of this window
'\\ (Allows the imlementation of .SelText on controls that don't have that member)
' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get SelectedText() As String
Attribute SelectedText.VB_HelpID = 2801

Dim sSelectedText As String
Dim lSelStart As Integer
Dim lSelEnd As Integer
Dim lRet As Long

Dim lSelMsg As Long

'\\ Select the appropriate message to query the selection according to the window type
If WindowBaseClass = Control_EDIT Then
    lSelMsg = EM_GETSEL
ElseIf WindowBaseClass = Control_COMBOBOX Then
    lSelMsg = CB_GETEDITSEL
Else
    '\\ If a control supports selected items, it should use the same message
    lSelMsg = EM_GETSEL
End If

lRet = SendMessageByLong(hwnd, lSelMsg, 0, 0)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":SelectedText", GetLastSystemError
Else
    lSelEnd = HiWord(lRet)
    lSelStart = LoWord(lRet)
    sSelectedText = WindowText
    If lSelEnd > 0 Then
        sSelectedText = Left$(sSelectedText, lSelEnd)
    End If
    If lSelStart > 0 And lSelStart < lSelEnd Then
        sSelectedText = Mid$(sSelectedText, lSelStart)
    End If
End If

SelectedText = sSelectedText

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Function SendMessage(ByVal wMsg As WindowMessages, ByVal wParam As Long, lParam As Long) As Long
Attribute SendMessage.VB_HelpID = 2279

Dim lRet As Long
Dim lThis As Long
Dim sThis As String

lRet = SendMessageLong(m_hwnd, wMsg, wParam, lParam)

If Err.LastDllError = 0 Then
    SendMessage = lRet
End If

End Function

' ##BLOCK_DESCRIPTION Sets the style bit specified to the window specified.
' ##BLOCK_REMARKS  Note that many window style  bits cannot be used at run time :. use this with caution
' ##BLOCK REMARKS  Returns true if it succeeded
Public Function SetWindowStyle(ByVal NewStyle As WindowStyleBits) As Boolean
Attribute SetWindowStyle.VB_Description = "Sets the style bit specified to the window specified."
Attribute SetWindowStyle.VB_HelpID = 3180

Dim lStyle As Long
Dim lRet As Long

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_STYLE)

'\\ Add the new style bit to it
lStyle = lStyle Or NewStyle

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, GWL_STYLE, lStyle)

'\\ For some styles to take effect, the window must be redrawn...
lRet = UpdateWindow(m_hwnd)

SetWindowStyle = IsWindowStyleSet(lStyle)

End Function

' ##BLOCK_DESCRIPTION Sets the style bit specified to the window specified.
' ##BLOCK_REMARKS  Note that many window style  bits cannot be used at run time :. use this with caution
' ##BLOCK REMARKS  Returns true if it succeeded
Public Function SetWindowExtendedStyle(ByVal NewStyle As WindowStyleExtendedBits) As Boolean
Attribute SetWindowExtendedStyle.VB_Description = "Sets the style bit specified to the window specified."
Attribute SetWindowExtendedStyle.VB_HelpID = 3181

Dim lStyle As Long
Dim lRet As Long


'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_EXSTYLE)

'\\ Add the new style bit to it
lStyle = lStyle Or NewStyle

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, GWL_EXSTYLE, lStyle)

'\\ For some styles to take effect, the window must be redrawn...
lRet = UpdateWindow(m_hwnd)

SetWindowExtendedStyle = IsWindowExtendedStyleSet(lStyle)

End Function
Public Sub Show(ByVal ShowCmd As ShowWindowCommands)
Attribute Show.VB_HelpID = 2802

Dim lRet As Long
lRet = ShowWindow(m_hwnd, ShowCmd)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":Show", GetLastSystemError
End If

End Sub

Public Property Get StaticControl() As ApiStaticControl
Attribute StaticControl.VB_HelpID = 3188

If m_hwnd <> 0 Then
    If WindowBaseClass = Control_STATIC Then
        Dim ssThis As ApiStaticControl
        Set ssThis = New ApiStaticControl
        Set ssThis.Window = Me
        Set StaticControl = ssThis
    End If
End If

End Property

'  ##BLOCK_DESCRIPTION This is the rectangle describing the location of the system menu box relative to the window as a whole
'  ##BLOCK_REMARKS This is held in the %system metrics:EventVB~ApiSystem~Metrics%.
Public Property Get SysmenuBoxRectangle() As APIRect
Attribute SysmenuBoxRectangle.VB_Description = "This is the rectangle describing the location of the system menu box relative to the window as a whole"
Attribute SysmenuBoxRectangle.VB_HelpID = 2281

Dim rcThis As New APIRect

With rcThis
    .Left = BorderWidth
    .Top = BorderHeight
    If IsWindowExtendedStyleSet(WS_EX_TOOLWINDOW) Then
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSMSIZE)
        .Right = .Left + APIDispenser.System.Metrics(SM_CXSMSIZE)
    Else
        .Bottom = .Top + APIDispenser.System.Metrics(SM_CYSIZE)
        .Right = .Left + APIDispenser.System.Metrics(SM_CXSIZE)
    End If
End With

Set SysmenuBoxRectangle = rcThis

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get SystemMenu() As ApiMenu
Attribute SystemMenu.VB_HelpID = 2282

Dim lRet As Long

lRet = GetSystemMenuApi(m_hwnd, False)
If Err.LastDllError = 0 And lRet > 0 Then
    If mSystemMenu Is Nothing Then
        Set mSystemMenu = New ApiMenu
    End If
    mSystemMenu.hMenu = lRet
    Set SystemMenu = mSystemMenu
End If

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Thread() As ApiThread
Attribute Thread.VB_HelpID = 2283

Dim lRet As Long

Dim threadThis As ApiThread

Set threadThis = New ApiThread
lRet = GetWindowThreadProcessId(m_hwnd, ByVal 0&)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":Thread", GetLastSystemError
Else
    threadThis.ThreadId = lRet
End If

Set Thread = threadThis

End Property

Friend Property Get ThreadParent() As Long

If mThreadParent <> 0 Then
    ThreadParent = mThreadParent
Else
    mThreadParent = GetWindowThreadProcessId(m_hwnd, ByVal 0&)
    ThreadParent = mThreadParent
End If

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get TitleBarRectangle() As APIRect
Attribute TitleBarRectangle.VB_HelpID = 2284

Dim lRet As Long
Dim tiThis As TITLEBARINFO
Dim rcRet As New APIRect

tiThis.cbSize = Len(tiThis)
If APIDispenser.System.IsRequiredSystem("GetTitlebarInfo", ver_Win_Win2000) Then
    lRet = GetTitleBarInfo(m_hwnd, tiThis)
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TitleBarRect", GetLastSystemError
    End If
    rcRet.CreateFromPointer (VarPtr(tiThis.rcTitleBar))
End If
Set TitleBarRectangle = rcRet

End Property

Friend Sub TriggerEvent(wMsg As WindowMessages, wParam As Long, lParam As Long, bDiscardMsg As Boolean, lWndProcRet As Long)

Dim Cancel As Boolean

Dim lRet As Long

On Error GoTo TriggerEventError

'\\ Used in wm_vscroll and wm_hscroll processing
Dim ScrollMessage As enScrollMessages
Dim ScrollPosition As Long
Dim ScrollAmount As Long
Static LastHScroll As Long
Static LastVScroll As Long

Dim Override As Boolean

'\\ Used in WM_MOVE, WM_MOUSE*  processing
Dim x As Long
Dim y As Long

'\\ Used in WM_MOSEHOVER and WM_MOUSELEAVE
Dim Button As Integer, Shift As Integer
Static MouseInWindow As Boolean
Static winVer As WindowsOSVersions

'\\ Used in WM_CTLCOLOR... processing
Dim ControlWindow As ApiWindow
Dim ControlDC As ApiDeviceContext
Dim ControlBrush As ApiLogBrush

'\\ Used in WM_MEASUREITEM and WM_DRAWITEM
Dim ControlRect As APIRect

'\\ used in NCPAINT and ERASEBKND
Dim dcThis As ApiDeviceContext

Dim ControlMenu As ApiMenu

    If winVer = 0 Then
        winVer = APIDispenser.System.WindowsOSVersion
    End If

    If wMsg = WM_ACTIVATEAPP Then
        '\\ This app is being activetd or deactivated...
        RaiseEvent ActiveApplicationChanged(CBool(wParam), lParam, Cancel)
    
    ElseIf wMsg = WM_CAPTURECHANGED Then
        '\\This window lost mouse capture
        RaiseEvent LostCapture(lParam, Cancel)
    
    ElseIf wMsg = WM_CHAR Then
        '\\ Keypress was intercepted by this window...
        Dim lRepetition As Long
        Dim lScanCode As Long
        Dim bExtendedKey As Boolean
        Dim bAltDown As Boolean
        Dim bAlreadyPressed As Boolean
        Dim bBeingPressed As Boolean
        With APIDispenser
            lRepetition = .GetBits(lParam, 0, 15)
            lScanCode = .GetBits(lParam, 16, 23)
            bExtendedKey = .GetBits(lParam, 24, 24)
            bAltDown = .GetBits(lParam, 29, 22)
            bAlreadyPressed = .GetBits(lParam, 30, 30)
            bBeingPressed = .GetBits(lParam, 31, 31)
        End With
        RaiseEvent KeyPressed(wParam, lRepetition, lScanCode, bExtendedKey, bAltDown, bAlreadyPressed, bBeingPressed, Cancel)
    
    ElseIf wMsg = WM_COMPACTING Then
        '\\ Windows is low on memory and is compacting memory...
        RaiseEvent LowMemory(wParam)
        '\\ It is not sensible to have a cancel option on this
    
    ElseIf wMsg = WM_MOVE Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent Move(x, y, Cancel)
        
        
    ElseIf wMsg = WM_VSCROLL Or wMsg = WM_HSCROLL Then
        Dim lPageLen As Long, lTotalLen As Long
        '\\ Vertical or horizontal scrollbar event has occured
        ScrollMessage = LoWord(wParam)
        '\\ If tracking then position is given
        If ScrollMessage = SB_THUMBTRACK Then
            ScrollPosition = HiWord(wParam)
            If wMsg = WM_VSCROLL Then
                ScrollAmount = ScrollPosition - LastVScroll
            Else
                ScrollAmount = ScrollPosition - LastHScroll
            End If
        Else
            ScrollPosition = IIf(wMsg = WM_VSCROLL, LastVScroll, LastHScroll)
            If wMsg = WM_HSCROLL Then
                lPageLen = VirtualPageWidth
                lTotalLen = VirtualWidth
            Else
                lPageLen = VirtualPageHeight
                lTotalLen = VirtualHeight
            End If
            Select Case ScrollMessage
            Case SB_LINEUP
                ScrollAmount = IIf(ScrollPosition > 0, -1, 0)
            Case SB_LINEDOWN
                ScrollAmount = IIf(ScrollPosition < lTotalLen, 1, 0)
            Case SB_PAGEUP
                ScrollAmount = IIf(ScrollPosition > 0, IIf(ScrollPosition >= lPageLen, -lPageLen, -ScrollPosition), 0)
            Case SB_PAGEDOWN
                ScrollAmount = IIf(ScrollPosition < lTotalLen, IIf(ScrollPosition <= (lTotalLen - lPageLen), lPageLen, (lTotalLen - lPageLen)), 0)
            End Select
            ScrollPosition = ScrollAmount + ScrollPosition
        End If
        If wMsg = WM_VSCROLL Then
            RaiseEvent VerticalScroll(ScrollMessage, ScrollPosition, Cancel)
        Else
            RaiseEvent HorizontalScroll(ScrollMessage, ScrollPosition, Cancel)
        End If
        '\\ Scroll the window
        If bAutoScroll Then
            '\\ And scroll the contents of the window....
            '\\ the opposite direction to the scrollbar
            If wMsg = WM_VSCROLL Then
                '\\ Update the scrollbar position
                Call SetScrollPos(m_hwnd, SB_VERT, ScrollPosition, True)
                Call Scroll(0, 0 - ScrollAmount, SW_INVALIDATE + SW_ERASE)
                LastVScroll = ScrollPosition
            Else
                '\\ Update the scrollbar position
                Call SetScrollPos(m_hwnd, SB_HORZ, ScrollPosition, True)
                Call Scroll(0 - ScrollAmount, 0, SW_INVALIDATE + SW_ERASE)
                LastHScroll = ScrollPosition
            End If
        End If
        
    ElseIf wMsg = WM_HOTKEY Then
        Dim hkThis As ApiHotkey
        Set hkThis = New ApiHotkey
        hkThis.VKey = HiWord(lParam)
        hkThis.Modifier = LoWord(lParam)
        RaiseEvent Hotkey(hkThis)
        Set hkThis = Nothing
        
    ElseIf wMsg = WM_SETTINGCHANGE Then
        RaiseEvent WindowsSettingsChanged(wParam)
    
    ElseIf wMsg = WM_WININICHANGE Then
        Dim sSection As String
        sSection = StringFromPointer(lParam, 1024)
        RaiseEvent WindowsINIChanged(sSection)
        
    '\\ Mouse button down
    ElseIf wMsg = WM_NCLBUTTONDOWN Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbLeftButton, x, y)
        
    ElseIf wMsg = WM_NCMBUTTONDOWN Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbMiddleButton, x, y)
        
    ElseIf wMsg = WM_NCRBUTTONDOWN Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbRightButton, x, y)
        
    '\\ Mouse button up
    ElseIf wMsg = WM_NCLBUTTONUP Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbLeftButton, x, y)
        
    ElseIf wMsg = WM_NCMBUTTONUP Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbMiddleButton, x, y)
        
    ElseIf wMsg = WM_NCRBUTTONUP Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbRightButton, x, y)

    '\\ Mouse move
    ElseIf wMsg = WM_NCMOUSEMOVE Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent NonClientMouseMove(wParam, x, y)
        '\\ If the mouse moved over a menu...
        If wParam = HTMENU Then
            Dim sCaption As String
            Dim ptThis As APIPoint
            Set ptThis = New APIPoint
            With ptThis
                .x = x
                .y = y
            End With
            sCaption = Menu.CaptionFromPoint(Me, ptThis)
            If sCaption <> "" Then
                RaiseEvent MouseOverMenu(sCaption)
            End If
        End If

    '\\ Minimum and maximum info.....
    ElseIf wMsg = WM_GETMINMAXINFO Then
        Dim myMinMaxInfo As MINMAXINFO
        Dim lMaxHeight As Long, lMaxWidth As Long
        Dim lMaxPositionTop As Long
        Dim lMaxPositionLeft As Long
        Dim lMinTrackWidth As Long, lMinTrackheight As Long
        Dim lMaxTrackWidth As Long, lMaxTrackHeight As Long
        Call CopyMemoryMinMaxInfo(myMinMaxInfo, lParam, LenB(myMinMaxInfo))
        If Err.LastDllError <> 0 Then
            ReportError Err.LastDllError, CLSNAME & ":WM_GETMINMAXINFO", GetLastSystemError
        Else
            With myMinMaxInfo
                '\\ Maximum height and width
                lMaxWidth = .ptMaxSize.x
                lMaxHeight = .ptMaxSize.y
                '\\ Maximum left and top
                lMaxPositionLeft = .ptMaxPosition.x
                lMaxPositionTop = .ptMaxPosition.y
                '\\ Maxiumum track size
                lMaxTrackWidth = .ptMaxTrackSize.x
                lMaxTrackHeight = .ptMaxTrackSize.y
                '\\ Minimum track size
                lMinTrackWidth = .ptMinTrackSize.x
                lMinTrackheight = .ptMinTrackSize.y
            End With
            RaiseEvent MinMaxSize(lMaxHeight, lMaxWidth, lMaxPositionTop, lMaxPositionLeft, lMinTrackWidth, lMinTrackheight, lMaxTrackWidth, lMaxTrackHeight)
            With myMinMaxInfo
                '\\ Maximum height and width
                 .ptMaxSize.x = lMaxWidth
                 .ptMaxSize.y = lMaxHeight
                '\\ Maximum left and top
                 .ptMaxPosition.x = lMaxPositionLeft
                 .ptMaxPosition.y = lMaxPositionTop
                '\\ Maxiumum track size
                 .ptMaxTrackSize.x = lMaxTrackWidth
                 .ptMaxTrackSize.y = lMaxTrackHeight
                '\\ Minimum track size
                 .ptMinTrackSize.x = lMinTrackWidth
                 .ptMinTrackSize.y = lMinTrackheight
            End With
            Call CopyMinMaxToMemoryInfo(lParam, myMinMaxInfo, LenB(myMinMaxInfo))
        End If
    '\\ Paint event
    ElseIf wMsg = WM_PAINT Then
        '\\ lParam and wParam are not used
        Dim mUpdateRect As EventVB.APIRect
        Dim rcThis As RECT
        Call GetUpdateRect(m_hwnd, rcThis, False)
        With rcThis
            If .Left Or .Right Or .Top Or .Bottom Then
                Set mUpdateRect = New APIRect
                Call mUpdateRect.CreateFromPointer(VarPtr(rcThis))
            Else
                '\\ No update region....
            End If
        End With
        bInPaintMessage = True
        RaiseEvent Paint(mUpdateRect, Override)
        bInPaintMessage = False
        If Override Or bPainting Then
            Cancel = 0
            bDiscardMsg = True
            If bPainting Then
                bPainting = False
                lRet = EndPaint(m_hwnd, mPaintStruct)
            End If
        End If
        
    ElseIf wMsg = WM_NCPAINT Then
        '\\ wParam is the handle to the update region or 1 if the whole non client
        '\\ area needs to be updated
        Set dcThis = New ApiDeviceContext
        
        If wParam = 1 Then
            dcThis.hdc = DeviceContextNC.hdc
        Else
            dcThis.hdc = GetDCEx(m_hwnd, wParam, (DCX_WINDOW Or DCX_INTERSECTRGN))
        End If
        RaiseEvent NonClientPaint(dcThis, Override)
        If Override Then
            Cancel = 0
            bDiscardMsg = True
        End If
        
    '\\ Non client hit test
    ElseIf wMsg = WM_NCHITTEST Then
        Dim RetVal As enHitTestResult
        '\\ Adjust x and y to map to the non client rectangle
        x = LoWord(lParam)
        y = HiWord(lParam)
        With RECT
            x = x - .Left
            y = y - .Top
        End With
        
        RaiseEvent HitTest(x, y, RetVal, Override)
        
        If Override Then
            Cancel = RetVal
            bDiscardMsg = True
        End If

    '\\ Sizing event
    ElseIf wMsg = WM_SIZING Then
        '\\ Get the sizing rectangle
        Dim rcSizingTemp As APIRect
        Set rcSizingTemp = New APIRect
        Call rcSizingTemp.CreateFromPointer(lParam)
        RaiseEvent Sizing(wParam, rcSizingTemp)
        If rcSizingTemp.Changed Then
            Call rcSizingTemp.SaveToPointer(lParam)
        End If

    ElseIf wMsg = WM_MOVING Then
         '\\ Get the sizing rectangle
        Dim rcMovingTemp As APIRect
        Set rcMovingTemp = New APIRect
        Call rcMovingTemp.CreateFromPointer(lParam)
        RaiseEvent Moving(wParam, rcMovingTemp)
        If rcMovingTemp.Changed Then
            Call rcMovingTemp.SaveToPointer(lParam)
        End If

    ElseIf wMsg = WM_POWERBROADCAST Then
        RaiseEvent PowerManagement(wParam, Cancel)
        If Cancel Then
            lWndProcRet = BROADCAST_QUERY_DENY
            Exit Sub
        End If

    ElseIf wMsg = WM_MOUSEHOVER Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        If (wParam And MK_CONTROL) = wParam Then
            Shift = Shift Or &H2
        End If
        If (wParam And MK_SHIFT) = wParam Then
            Shift = Shift Or &H1
        End If
        If (wParam And MK_LBUTTON) = wParam Then
            Button = Button Or &H1
        End If
        If (wParam And MK_MBUTTON) = wParam Then
            Button = Button Or &H4
        End If
        If (wParam And MK_RBUTTON) = wParam Then
            Button = Button Or &H2
        End If
        RaiseEvent MouseHover(Button, Shift, x, y)
        MouseInWindow = False
        
    ElseIf wMsg = WM_MOUSELEAVE Then
        If MouseInWindow Then
            RaiseEvent MouseLeave
            MouseInWindow = False
        End If
    
    '\\ context menu handling
    ElseIf wMsg = WM_CONTEXTMENU Then
        x = LoWord(lParam)
        y = HiWord(lParam)
        RaiseEvent ContextMenu(x, y, Override)
        If Override Then
            lWndProcRet = 1
            bDiscardMsg = True
        End If
    
    
    '\\ Control colouring - ControlColour() event
    ElseIf wMsg = WM_CTLCOLORBTN Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = lParam
        Set ControlBrush = New ApiLogBrush
        RaiseEvent ControlColour(Control_BUTTON, ControlDC, ControlWindow, ControlBrush)
        If ControlBrush.Changed Then
            lWndProcRet = ControlBrush.Handle
            bDiscardMsg = True
        End If
    
    ElseIf wMsg = WM_CTLCOLOREDIT Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = lParam
        Set ControlBrush = New ApiLogBrush
        RaiseEvent ControlColour(Control_EDIT, ControlDC, ControlWindow, ControlBrush)
        If ControlBrush.Changed Then
            lWndProcRet = ControlBrush.Handle
            bDiscardMsg = True
        End If
        
    ElseIf wMsg = WM_CTLCOLORLISTBOX Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = lParam
        Set ControlBrush = New ApiLogBrush
        RaiseEvent ControlColour(Control_LISTBOX, ControlDC, ControlWindow, ControlBrush)
        If ControlBrush.Changed Then
            lWndProcRet = ControlBrush.Handle
            bDiscardMsg = True
        End If
        
    ElseIf wMsg = WM_CTLCOLORSCROLLBAR Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = lParam
        Set ControlBrush = New ApiLogBrush
        RaiseEvent ControlColour(Control_SCROLLBAR, ControlDC, ControlWindow, ControlBrush)
        If ControlBrush.Changed Then
            lWndProcRet = ControlBrush.Handle
            bDiscardMsg = True
        End If
    
    ElseIf wMsg = WM_CTLCOLORSTATIC Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = lParam
        Set ControlBrush = New ApiLogBrush
        RaiseEvent ControlColour(Control_STATIC, ControlDC, ControlWindow, ControlBrush)
        If ControlBrush.Changed Then
            lWndProcRet = ControlBrush.Handle
            bDiscardMsg = True
        End If
        
    '\\ Drawing owner-drawn items
    ElseIf wMsg = WM_DRAWITEM Then
        Dim odStruct As DRAWITEMSTRUCT
        Call CopyMemoryDrawItem(odStruct, lParam, Len(odStruct))
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = odStruct.hdc
        Set ControlWindow = New ApiWindow
        ControlWindow.hwnd = odStruct.hwndItem
        Set ControlRect = New APIRect
        With odStruct
            '\\ Get the control's item rect
            ControlRect.Bottom = .rcItem.Bottom
            ControlRect.Left = .rcItem.Left
            ControlRect.Right = .rcItem.Right
            ControlRect.Top = .rcItem.Top
            '\\ Raise the event according to control type...
            Select Case .ctlType
            Case ODT_MENU
                '\\ Draw an owner-drawn menu....
                Set ControlMenu = New ApiMenu
                ControlMenu.WindowHandle = .hwndItem
                '\\ Todo - set menu id
                RaiseEvent DrawOwnerDrawnMenuItem(.ItemId, .itemAction, (.itemState And ODS_CHECKED), (.itemState And ODS_GRAYED), (.itemState And ODS_SELECTED), ControlMenu, ControlDC, ControlRect, .ItemData)
            
            Case ODT_LISTBOX
                '\\ Draw the owner draw listbox's item
                RaiseEvent DrawOwnerDrawnListboxItem(.CtlID, .itemAction, .ItemId, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect, .ItemData)
            
            Case ODT_COMBOBOX
                '\\ Draw the owner drawn combo box's item
                RaiseEvent DrawOwnerDrawnComboItem(.CtlID, .itemAction, .ItemId, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect, .ItemData)
            
            Case ODT_BUTTON
                '\\ Draw the owner-drawn button
                RaiseEvent DrawOwnerDrawnButton(.CtlID, .itemAction, (.itemState And ODS_DEFAULT), (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOACCEL), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect)
            
            Case ODT_STATIC
                '\\ Draw the static control
                RaiseEvent DrawOwnerDrawnStatic(.CtlID, .itemAction, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect)
            
            Case ODT_HEADER
                '\\ Draw the header control
                RaiseEvent DrawOwnerDrawnHeader(.CtlID, .itemAction, .ItemId, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect, .ItemData)

            Case ODT_TAB
                '\\ Draw the tab control
                RaiseEvent DrawOwnerDrawnTab(.CtlID, .itemAction, .ItemId, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect, .ItemData)
            
            Case ODT_LISTVIEW
                '\\ Draw the listview control item
                RaiseEvent DrawOwnerDrawnListviewItem(.CtlID, .itemAction, .ItemId, (.itemState And ODS_DISABLED), (.itemState And ODS_FOCUS), (.itemState And ODS_HOTLIGHT), (.itemState And ODS_NOFOCUSRECT), ControlWindow, ControlDC, ControlRect, .ItemData)
 
            End Select
        End With
        Cancel = True
    
    ElseIf wMsg = WM_DEVICECHANGE Then
        Dim wpTranslated As DeviceManagementChangeTypes
        wpTranslated = wParam
        Select Case wpTranslated
        Case DBT_CONFIGCHANGECANCELED
        
        Case DBT_CONFIGCHANGED
        
        Case DBT_CUSTOMEVENT
        
        Case DBT_DEVICEARRIVAL
        
        Case DBT_DEVICEQUERYREMOVE
        
        Case DBT_DEVICEQUERYREMOVEFAILED
        
        Case DBT_DEVICEREMOVECOMPLETE
        
        Case DBT_DEVICEREMOVEPENDING
        
        Case DBT_DEVICETYPESPECIFIC
        
        Case DBT_NO_DISK_SPACE
            RaiseEvent DeviceManagementDiskFull(lParam)
            
        Case DBT_LOW_DISK_SPACE
            RaiseEvent DeviceManagementDiskLow(lParam)
        
        Case DBT_MONITORCHANGE
            If lParam = 0 Then
                RaiseEvent DeviceManagementMonitorChanged(Screen.Width, Screen.Height)
            Else
                RaiseEvent DeviceManagementMonitorChanged(LoWord(lParam), HiWord(lParam))
            End If
        
        End Select
    
    ElseIf wMsg = WM_MOUSEWHEEL Then
        '\\ x and y are the current cursor pos...
        x = LoWord(lParam)
        y = HiWord(lParam)
        Dim lwwParam As Integer
        Dim hiwParam As Integer
        lwwParam = LoWord(wParam)
        hiwParam = HiWord(wParam)
        If (lwwParam And MK_CONTROL) = lwwParam Then
            Shift = Shift Or &H2
        End If
        If (lwwParam And MK_SHIFT) = lwwParam Then
            Shift = Shift Or &H1
        End If
        If (lwwParam And MK_LBUTTON) = lwwParam Then
            Button = Button Or &H1
        End If
        If (lwwParam And MK_MBUTTON) = lwwParam Then
            Button = Button Or &H4
        End If
        If (lwwParam And MK_RBUTTON) = lwwParam Then
            Button = Button Or &H2
        End If
        RaiseEvent MouseWheelMove(hiwParam, Button, Shift, x, y)
    
    ElseIf wMsg = WM_PRINTCLIENT Then
        Set ControlDC = New ApiDeviceContext
        ControlDC.hdc = wParam
        RaiseEvent PaintWindow(ControlDC, lParam)
        
    ElseIf wMsg = WM_ERASEBKGND Then
        Set dcThis = New ApiDeviceContext
        dcThis.hdc = wParam
        RaiseEvent EraseBackground(dcThis, Override)
        If Override Then
            lWndProcRet = 1
            bDiscardMsg = True
        End If
    
    ElseIf wMsg = WM_COMMAND Then
        If HiWord(wParam) = 0 Then
            '\\ The command was a menu click
            Dim mnuThis As New ApiMenu
            mnuThis.hMenu = GetMenuApi(m_hwnd)
            mnuThis.id = LoWord(wParam)
            RaiseEvent MenuClick(mnuThis)
        ElseIf HiWord(wParam) = 1 Then
            '\\ The command was from an accelerator
            
        Else
            '\\ The command was from a control
            Dim nNotify As Integer, ControlClass As StandardControlTypes
            nNotify = HiWord(wParam)
            ControlClass = GetWindowBaseClass(lParam)
            Set ControlWindow = New ApiWindow
            ControlWindow.hwnd = lParam
            Select Case ControlClass
            Case Control_BUTTON
                '\\ The notification was one of the BN_ notifications
                RaiseEvent ChildButtonNotification(nNotify, LoWord(wParam), ControlWindow)
                               
            Case Control_EDIT
                '\\ The notification is one of the EN_ notifications
                RaiseEvent ChildEditNotification(nNotify, LoWord(wParam), ControlWindow)
                
            Case Control_LISTBOX
                '\\ Notification is one of the LBN_ notifications
                RaiseEvent ChildListboxNotification(nNotify, LoWord(wParam), ControlWindow)
                
            Case Control_SCROLLBAR
                '\\ No known notifications sent through WM_COMMAND...
                
            Case Control_STATIC
                '\\ Notification is one of the STN_ notifications
                RaiseEvent ChildStaticControlNotification(nNotify, LoWord(wParam), ControlWindow)
                
            Case Control_COMBOBOX
                '\\ Notifications are of the CBN_ type
                RaiseEvent ChildComboboxNotification(nNotify, LoWord(wParam), ControlWindow)
                
            End Select
            
        End If
    
    ElseIf wMsg = WM_HELP Then
        Dim hiThis As HELPINFO
        hiThis.cbSize = Len(hiThis)
        Call CopyMemoryHelpInfo(hiThis, lParam, Len(hiThis))
        With hiThis
            If .iContextType = HELPINFO_MENUITEM Then
                Set ControlMenu = New ApiMenu
                ControlMenu.id = .hItemHandle
                RaiseEvent ContextSensitiveMenuHelp(ControlMenu, .dwContextId, .MousePos.x, .MousePos.y)
            Else
                Set ControlWindow = New ApiWindow
                ControlWindow.hwnd = .hItemHandle
                RaiseEvent ContextSensitiveControlHelp(ControlWindow, .dwContextId, .MousePos.x, .MousePos.y)
            End If

        End With
        
    Else
        '\\ For all other events
        RaiseEvent WindowMessageFired(wMsg, wParam, lParam, bDiscardMsg, lWndProcRet)
    End If
    
    If Cancel Then
        lWndProcRet = Cancel
    End If
    
    '\\ Special internal process
    If wMsg = WM_MOUSEMOVE Then
        '\\ If we have sufficient version of windows then allow for mouse tracking
        If winVer <> ver_Win_95 And winVer <> ver_Win_NT351 Then
            If Not MouseInWindow Then
                MouseInWindow = True
                mheThis.cbSize = Len(mheThis)
                mheThis.hwndTrack = m_hwnd
                lRet = TrackMouseEvent(mheThis)
                If Err.LastDllError Then
                    ReportError Err.LastDllError, CLSNAME & ":MouseTracking", GetLastSystemError
                End If
            End If
        End If
    End If

Exit Sub
TriggerEventError:
    If Err.LastDllError Then
        ReportError Err.LastDllError, CLSNAME & ":TriggerEvent", GetLastSystemError
    Else
        ReportError Err.Number, CLSNAME & ":TriggerEvent", Err.Description
    End If
End Sub


' ##BLOCK_DESCRIPTION Unsets the style bit specified to the window
' ##BLOCK_REMARKS Note that many window style bits cannot be used at run time :. use this with caution
Public Function UnSetWindowStyle(ByVal NewStyle As WindowStyleBits) As Boolean
Attribute UnSetWindowStyle.VB_Description = "Unsets the style bit specified to the window"
Attribute UnSetWindowStyle.VB_HelpID = 2285

Dim lStyle As Long
Dim lRet As Long

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_STYLE)

'\\ Add the new style bit to it
If IsWindowStyleSet(NewStyle) Then
    lStyle = lStyle Xor NewStyle
End If

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, GWL_STYLE, lStyle)

'\\ Refresh the window
lRet = UpdateWindow(m_hwnd)

UnSetWindowStyle = Not (IsWindowStyleSet(lStyle))

End Function

' ##BLOCK_DESCRIPTION Unsets the extended style bit specified to the window
' ##BLOCK_REMARKS Note that many window style bits cannot be used at run time :. use this with caution
Public Function UnSetWindowExtendedStyle(ByVal NewStyle As WindowStyleExtendedBits) As Boolean
Attribute UnSetWindowExtendedStyle.VB_Description = "Unsets the extended style bit specified to the window"
Attribute UnSetWindowExtendedStyle.VB_HelpID = 3182

Dim lStyle As Long
Dim lRet As Long


'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, GWL_EXSTYLE)

'\\ Add the new style bit to it
If IsWindowExtendedStyleSet(NewStyle) Then
    lStyle = lStyle Xor NewStyle
End If

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, GWL_EXSTYLE, lStyle)

'\\ Refresh the window
lRet = UpdateWindow(m_hwnd)

UnSetWindowExtendedStyle = Not (IsWindowExtendedStyleSet(lStyle))

End Function

' ##BLOCK_DESCRIPTION This is true if this window is able to recieve the input focus
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "This is true if this window is able to recieve the input focus"
Attribute Enabled.VB_HelpID = 2286

    Enabled = IsWindowEnabledApi(m_hwnd)
    
End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Unicode() As Boolean
Attribute Unicode.VB_HelpID = 2287

    Unicode = IsWindowUnicodeApi(m_hwnd)
    
End Property

Public Property Let VirtualHeight(ByVal newval As Long)
Attribute VirtualHeight.VB_Description = "VirtualHeight - The total height of the scrollable form  "
Attribute VirtualHeight.VB_HelpID = 3034

If newval < Height Then
    newval = Height
End If

mlVirtualHeight = newval

Dim siThis As SCROLLINFO

With siThis
    .cbSize = Len(siThis)
    .nMax = mlVirtualHeight
    .fMask = SIF_RANGE
End With
Call SetScrollInfo(m_hwnd, SB_VERT, siThis, True)


End Property

' ##BLOCK_DESCRIPTION VirtualHeight - The total height of the scrollable form
' ##BLOCK_REMARKS This cannot be less than the actual form height
' ##BLOCK_REMARKS This value is measured in Pixels
Public Property Get VirtualHeight() As Long

If mlVirtualHeight < Height Then
    mlVirtualHeight = Height
End If

VirtualHeight = mlVirtualHeight

End Property

Public Property Let VirtualPageHeight(ByVal newval As Long)
Attribute VirtualPageHeight.VB_Description = "VirtualPageHeight - The height of a page of the scrollable form  "
Attribute VirtualPageHeight.VB_HelpID = 3036

If newval < 1 Then
    newval = 1
End If

mlVirtualPageHeight = newval


Dim siThis As SCROLLINFO

With siThis
    .cbSize = Len(siThis)
    .nPage = mlVirtualPageHeight
    .fMask = SIF_PAGE
End With
Call SetScrollInfo(m_hwnd, SB_VERT, siThis, True)

End Property

' ##BLOCK_DESCRIPTION VirtualPageHeight - The height of a page of the scrollable form
' ##BLOCK_REMARKS This cannot be less than 1
' ##BLOCK_REMARKS This value is measured in Pixels
Public Property Get VirtualPageHeight() As Long

If mlVirtualPageHeight < 1 Then
    mlVirtualPageHeight = 1
End If

VirtualPageHeight = mlVirtualPageHeight

End Property


Public Property Let VirtualPageWidth(ByVal newval As Long)
Attribute VirtualPageWidth.VB_Description = "VirtualPageWidth - The width of a page of the scrollable form  "
Attribute VirtualPageWidth.VB_HelpID = 3038

If newval < 1 Then
    newval = 1
End If

mlVirtualPageWidth = newval

Dim siThis As SCROLLINFO

With siThis
    .cbSize = Len(siThis)
    .nPage = mlVirtualPageWidth
    .fMask = SIF_PAGE
End With
Call SetScrollInfo(m_hwnd, SB_HORZ, siThis, True)

End Property

' ##BLOCK_DESCRIPTION VirtualPageWidth - The width of a page of the scrollable form
' ##BLOCK_REMARKS This cannot be less than the actual form width
' ##BLOCK_REMARKS This value is measured in Pixels
Public Property Get VirtualPageWidth() As Long

If mlVirtualPageWidth < 1 Then
    mlVirtualPageWidth = 1
End If

VirtualPageWidth = mlVirtualPageWidth

End Property


Public Property Let VirtualWidth(ByVal newval As Long)
Attribute VirtualWidth.VB_Description = "VirtualWidth - The total width of the scrollable form  "
Attribute VirtualWidth.VB_HelpID = 3040

If newval < 1 Then
    newval = 1
End If

mlVirtualWidth = newval

Dim siThis As SCROLLINFO

With siThis
    .cbSize = Len(siThis)
    .nMax = mlVirtualWidth
    .fMask = SIF_RANGE
End With
Call SetScrollInfo(m_hwnd, SB_HORZ, siThis, True)


End Property

' ##BLOCK_DESCRIPTION VirtualWidth - The total width of the scrollable form
' ##BLOCK_REMARKS This cannot be less than the actual form width
' ##BLOCK_REMARKS This value is measured in Pixels
Public Property Get VirtualWidth() As Long

If mlVirtualWidth < Width Then
    mlVirtualWidth = Width
End If

VirtualWidth = mlVirtualWidth

End Property


' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get Visible() As Boolean
Attribute Visible.VB_HelpID = 2288

    Visible = IsWindowVisibleApi(m_hwnd)
    
End Property



' ##BLOCK_DESCRIPTION This is true if the window is maximised
' ##BLOCK_REMARKS The Zoomed keyword is a throwback to previous versions of Windows
Public Property Get IsZoomed() As Boolean
Attribute IsZoomed.VB_Description = "This is true if the window is maximised"
Attribute IsZoomed.VB_HelpID = 2289

    IsZoomed = IsZoomedApi(m_hwnd)
    
End Property

' ##BLOCK_DESCRIPTION The key uniquely describes this window class in any collection
Public Property Get Key() As String
Attribute Key.VB_Description = "The key uniquely describes this window class in any collection"
Attribute Key.VB_HelpID = 2290

    Key = "HWND:" & m_hwnd

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Function SetWindowLong(ByVal Index As enGetWindowLong, ByVal newValue As Long) As Long
Attribute SetWindowLong.VB_HelpID = 2291

Dim lRet As Long

lRet = SetWindowLongApi(m_hwnd, Index, newValue)
If Err.LastDllError = 0 Then
    If Index = GWL_WNDPROC Then
        '\\ We must store old proc address...
        mOldProcAddress = lRet
    End If
    SetWindowLong = lRet
End If

End Function

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS This value is expressed in TWIPS
Public Property Get Width() As Long
Attribute Width.VB_HelpID = 2292

Dim lpRect As RECT
Call GetWindowRect(m_hwnd, lpRect)
Width = (lpRect.Right - lpRect.Left) * Screen.TwipsPerPixelX
    
End Property

' ##BLOCK_DESCRIPTION WindowBaseClass - Returns the base class of this window
Public Property Get WindowBaseClass() As StandardControlTypes
Attribute WindowBaseClass.VB_Description = "WindowBaseClass - Returns the base class of this window"
Attribute WindowBaseClass.VB_HelpID = 2293

WindowBaseClass = GetWindowBaseClass(m_hwnd)

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Function WindowFromPoint(ByVal xPoint As Long, ByVal yPoint As Long) As ApiWindow
Attribute WindowFromPoint.VB_HelpID = 2294

Dim hwnd As Long
Dim Window As ApiWindow

hwnd = WindowFromPointApi(xPoint, yPoint)
If IsWindowApi(hwnd) Then
    Set Window = New ApiWindow
    Window.hwnd = hwnd
    Set WindowFromPoint = Window
End If

End Function


' ## BLOCK_DESCRIPTION
' ## BLOCK_REMARKS
Public Property Let WindowText(ByVal newText As String)
Attribute WindowText.VB_Description = "BLOCK_DESCRIPTION"
Attribute WindowText.VB_HelpID = 2803

Dim lRet As Long

lRet = SetWindowText(m_hwnd, newText)
If Err.LastDllError Then
    ReportError Err.LastDllError, CLSNAME & ":WindowText", GetLastSystemError
End If

End Property

' ##BLOCK_DESCRIPTION
' ##BLOCK_REMARKS
Public Property Get WindowText() As String

Dim lRet As Long

Dim sRet As String

lRet = SendMessage(WM_GETTEXTLENGTH, 0, 0)
If lRet > 0 Then
    sRet = String$(lRet + 1, 0)
    lRet = SendMessageByString(m_hwnd, WM_GETTEXT, Len(sRet), sRet)
    '\\ Returns length up to NULL terminator
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
    End If
ElseIf Err.LastDllError > 0 Then
    ReportError Err.LastDllError, CLSNAME & ":WindowText", GetLastSystemError
Else
    sRet = String$(1024, 0)
    lRet = GetWindowTextApi(m_hwnd, sRet, Len(sRet))
    sRet = Left$(sRet, lRet)
End If

WindowText = sRet

End Property

' ##BLOCK_DESCRIPTION The %WNDCLASS:EventVB~ApiWndClass% structure describes the window class used to create this window instance
' ##BLOCK_REMARKS You can pass this structure to the %window dispenser:EventVB~ApiSystem~WindowDispenser% to create new instances of this class of window
Public Property Get WndClass() As APIWndClass
Attribute WndClass.VB_Description = "The WNDCLASS:EventVB~ApiWndClass structure describes the window class used to create this window instance"
Attribute WndClass.VB_HelpID = 2296

If mWndClass Is Nothing Then
    Set mWndClass = New APIWndClass
End If
Call mWndClass.CreateFromClassname(0, ClassName)
Set WndClass = mWndClass


End Property

Private Sub Class_Terminate()

    '\\ If the old proc address is set, unhook this...
    If mOldProcAddress > 0 Then
        Call SetWindowLong(GWL_WNDPROC, mOldProcAddress)
    End If

    
    Set mMenu = Nothing
    Set mSystemMenu = Nothing
    Set mWndClass = Nothing
    Set mRegion = Nothing
    Set mDeviceContext = Nothing
    Set mDeviceContextNC = Nothing
    
End Sub

Public Property Get Icon(ByVal Index As IconTypes) As ApiIcon
Attribute Icon.VB_HelpID = 3041

Dim icThis As New ApiIcon
'\\ Get the icon handle associated with this window
icThis.hIcon = SendMessage(WM_GETICON, Index, 0)

Set Icon = icThis

End Property

Public Property Let Icon(ByVal Index As IconTypes, ByVal icThis As ApiIcon)

'\\ Get the icon handle associated with this window
Call SendMessage(WM_SETICON, Index, icThis.hIcon)

End Property

Friend Sub AddProperty(ByVal PropertyName As String)


End Sub

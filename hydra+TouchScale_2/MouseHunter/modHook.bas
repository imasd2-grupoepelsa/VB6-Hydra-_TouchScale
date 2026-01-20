Attribute VB_Name = "modHook"
Option Explicit

'* The CopyMemory function copies a block of memory from one location to another.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'* The GetProp function retrieves a data handle from the property list of the given window. The given character string identifies the handle to be retrieved. The string and handle must have been added to the property list by a previous call to the SetProp function.
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'* The SetProp function adds a new entry or changes an existing entry in the property list of the specified window. The function adds a new entry to the list if the specified character string does not exist already in the list. The new entry contains the string and the handle. Otherwise, the function replaces the string’s current handle with the specified handle.
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'* The RemoveProp function removes an entry from the property list of the specified window. The specified character string identifies the entry to be removed.
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd&, ByVal lpString$) As Long

'* Global variable to hold the hWnd of the caller application's window
Global glnghWnd   As Long

'* Main procedure to trace the mouse events, but it is bypassed
'* to the Tracer class's MouseProc method so that the events can be raised from there.
'* As you know events can't be raised from a general module.
Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    MouseProc = TracerFromMemory(glnghWnd).MouseProc(nCode, wParam, lParam)
End Function

'* This method retrieves the object of the tracer class created in the memory
Private Function TracerFromMemory(ByVal hwnd As Long) As Tracer
    Dim MsgHookEx   As Tracer
    Dim ptrObject   As Long
    
    '* Getting the already created Object of the Tracer class
    '* from memory in ptrObjec
    ptrObject = GetProp(glnghWnd, ByVal "MouseHook")
    
    '* Setting the Tracer class object from ptrObject to MsgHookEx
    CopyMemory MsgHookEx, ptrObject, Len(ptrObject)
    
    '* return the Tracer class object
    Set TracerFromMemory = MsgHookEx
    
    '* Remove the reference of the Tracer class object from memory
    CopyMemory MsgHookEx, 0&, Len(ptrObject)
End Function

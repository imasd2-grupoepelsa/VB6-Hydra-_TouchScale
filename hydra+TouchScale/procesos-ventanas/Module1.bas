Attribute VB_Name = "Module1"
'By The-Pirat 2005    thepirat000@hotmail.com
Option Explicit

Public objWMIService As Object
Public colItems As Object
Public colProcesos As New Collection
Public colEjecutable As New Collection

'Win32 API
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Const WM_CLOSE = &H10
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOREDRAW = &H8
Public Enum eESTADOWIN
    eSW_MAXIMIZE = 3
    eSW_MINIMIZE = 6
    eSW_SHOW = 5
    eWM_CLOSE = &H10
    eSW_HIDE = 0
    eSW_RESTORE = 9
End Enum

'Modifica el padre (contenedor) de una ventana
Public Function fbModificarPadreVentana(hWnd_Hija As Long, hWnd_Padre As Long) As Boolean
Dim l As Long
l = SetParent(hWnd_Hija, hWnd_Padre)
MoveWindow hWnd_Hija, 0, 0, 100, 100, 1
fbModificarPadreVentana = (l <> 0)
End Function

'Mata un proceso
Public Function Matar(PID As Long)
Dim objitem
Set objWMIService = GetObject("winmgmts:\\")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Process", , 48)
For Each objitem In colItems
    If objitem.processid = PID Then
        objitem.Terminate
        Exit For
    End If
Next
End Function

'Función delegada que es llamada una vez por cada ventana (por la API en EnumWindows)
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim sSave As String, Ret As Long, p As Long, dev As Long, sTemp As String
    Dim cont As Long, sEstado As String
    'Obtener el título de la ventana
    Ret = GetWindowTextLength(hwnd)
    sSave = Space(Ret)
    sEstado = ""
    GetWindowText hwnd, sSave, Ret + 1
    If sSave <> "" Then
    		 'Si la ventana tiene título, obtener el proceso al que pertenece
         dev = GetWindowThreadProcessId(hwnd, p)
         sTemp = colProcesos.Item("K" & p)
         'Mostrar estatus Visible, Invisible y Enabled, Disabled.
         sEstado = sEstado & IIf(IsWindowVisible(hwnd) <> 0, "V", "I") & "-" & IIf(IsWindowEnabled(hwnd) <> 0, "E", "D")
         Form1.MSFlexGrid1.AddItem p & Chr(9) & sTemp & Chr(9) & Str$(hwnd) & Chr(9) & Trim(sSave) & Chr(9) & sEstado
         cont = cont + 1
    End If
    EnumWindowsProc = True
End Function

'Retorna información de un proceso
Public Function MasInfo(PID As String) As String
Dim s As String, objitem
s = ""
For Each objitem In colItems
    If objitem.processid = Val(PID) Then
        s = s & "Caption: " & objitem.Caption & vbCrLf
        s = s & "CreationClassName: " & objitem.CreationClassName & vbCrLf
        s = s & "CreationDate: " & objitem.CreationDate & vbCrLf
        s = s & "CSCreationClassName: " & objitem.CSCreationClassName & vbCrLf
        s = s & "CSName: " & objitem.CSName & vbCrLf
        s = s & "Description: " & objitem.Description & vbCrLf
        s = s & "ExecutablePath: " & objitem.executablepath & vbCrLf
        s = s & "ExecutionState: " & objitem.ExecutionState & vbCrLf
        s = s & "Handle: " & objitem.Handle & vbCrLf
        s = s & "HandleCount: " & objitem.HandleCount & vbCrLf
        s = s & "InstallDate: " & objitem.InstallDate & vbCrLf
        s = s & "KernelModeTime: " & objitem.KernelModeTime & vbCrLf
        s = s & "MaximumWorkingSetSize: " & objitem.MaximumWorkingSetSize & vbCrLf
        s = s & "MinimumWorkingSetSize: " & objitem.MinimumWorkingSetSize & vbCrLf
        s = s & "Name: " & objitem.Name & vbCrLf
        s = s & "OSCreationClassName: " & objitem.OSCreationClassName & vbCrLf
        s = s & "OSName: " & objitem.OSName & vbCrLf
        s = s & "OtherOperationCount: " & objitem.OtherOperationCount & vbCrLf
        s = s & "OtherTransferCount: " & objitem.OtherTransferCount & vbCrLf
        s = s & "PageFaults: " & objitem.PageFaults & vbCrLf
        s = s & "PageFileUsage: " & objitem.PageFileUsage & vbCrLf
        s = s & "ParentProcessId: " & objitem.ParentProcessId & vbCrLf
        s = s & "PeakPageFileUsage: " & objitem.PeakPageFileUsage & vbCrLf
        s = s & "PeakVirtualSize: " & objitem.PeakVirtualSize & vbCrLf
        s = s & "PeakWorkingSetSize: " & objitem.PeakWorkingSetSize & vbCrLf
        s = s & "Priority: " & objitem.Priority & vbCrLf
        s = s & "PrivatePageCount: " & objitem.PrivatePageCount & vbCrLf
        s = s & "ProcessId: " & objitem.processid & vbCrLf
        s = s & "QuotaNonPagedPoolUsage: " & objitem.QuotaNonPagedPoolUsage & vbCrLf
        s = s & "QuotaPagedPoolUsage: " & objitem.QuotaPagedPoolUsage & vbCrLf
        s = s & "QuotaPeakNonPagedPoolUsage: " & objitem.QuotaPeakNonPagedPoolUsage & vbCrLf
        s = s & "QuotaPeakPagedPoolUsage: " & objitem.QuotaPeakPagedPoolUsage & vbCrLf
        s = s & "ReadOperationCount: " & objitem.ReadOperationCount & vbCrLf
        s = s & "ReadTransferCount: " & objitem.ReadTransferCount & vbCrLf
        s = s & "SessionId: " & objitem.SessionId & vbCrLf
        s = s & "Status: " & objitem.Status & vbCrLf
        s = s & "TerminationDate: " & objitem.TerminationDate & vbCrLf
        s = s & "ThreadCount: " & objitem.ThreadCount & vbCrLf
        s = s & "UserModeTime: " & objitem.UserModeTime & vbCrLf
        s = s & "VirtualSize: " & objitem.VirtualSize & vbCrLf
        s = s & "WindowsVersion: " & objitem.WindowsVersion & vbCrLf
        s = s & "WorkingSetSize: " & objitem.WorkingSetSize & vbCrLf
        s = s & "WriteOperationCount: " & objitem.WriteOperationCount & vbCrLf
        s = s & "WriteTransferCount: " & objitem.WriteTransferCount & vbCrLf
        Exit For
    End If
Next
MasInfo = s
End Function

'Carga los procesos en ejecución en las colecciones
Public Function ActualizarObjetosWMI()
Dim proceso
Set objWMIService = GetObject("winmgmts:\\")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Process")
Set colProcesos = Nothing
Set colEjecutable = Nothing
'Meto todo en una colección con index = processid, para hacer mucho más rápida la búsqueda
For Each proceso In colItems
    colProcesos.Add proceso.Caption, "K" & proceso.processid
    colEjecutable.Add proceso.executablepath, "K" & proceso.processid
Next
End Function

'Envía un mensaje para cambiar el estado de una ventana (maximizado, etc)
Public Sub WindowHandle(ByVal hWindow As Long, ByVal mCase As eESTADOWIN)
If mCase = eWM_CLOSE Then
        SendMessage hWindow, eWM_CLOSE, 0, 0
Else
        ShowWindow hWindow, mCase
End If
End Sub

'Habilita/Deshabilita una ventana
Public Sub windowEnable(ByVal hWindow As Long, Enable As Boolean)
EnableWindow hWindow, IIf(Enable, 1, 0)
End Sub


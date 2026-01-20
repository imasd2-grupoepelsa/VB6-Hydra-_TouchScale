VERSION 5.00
Begin VB.Form FrmInactividad 
   Caption         =   "Form3"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form3"
   ScaleHeight     =   1335
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrInactivo 
      Interval        =   100
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "FrmInactividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Integer
    Y As Integer
End Type
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI)
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private posOld As POINTAPI
Private posNew As POINTAPI
Private Huella As Integer
Private StatusKeyboard As Boolean
Private i As Integer
Private RatonAnterior As Boolean
Private TecladoAnterior As Boolean
Private TiempoTotal As Long
Public Function StatusMouse() As Boolean
    Call GetCursorPos(posNew)
    If ((posNew.X <> posOld.X) Or (posNew.Y <> posOld.Y)) Then
        posOld = posNew
        StatusMouse = True
        Exit Function
    End If
    StatusMouse = False
End Function

Private Sub TmrInactivo_Timer()
    Dim Bucle As Form
    TiempoTotal = TiempoTotal + 1
    If StatusMouse Then
        If RatonAnterior = False Then
            RatonAnterior = True
            TiempoMonitor = 0
            TiempoTotal = 0
            Exit Sub
        End If
    Else
        If RatonAnterior = True Then
            RatonAnterior = False
            TiempoMonitor = 0
            TiempoTotal = 0
            Exit Sub
        End If
    End If
    For i = 0 To 255
        If (GetAsyncKeyState(i) And &H8001) <> 0 Then
            StatusKeyboard = True
            Exit For
        End If
    Next i
    If StatusKeyboard Then
        If TecladoAnterior = False Then
            TecladoAnterior = True
            TiempoMonitor = 0
            TiempoTotal = 0
            Exit Sub
        End If
    Else
        If TecladoAnterior = True Then
            TecladoAnterior = False
            TiempoMonitor = 0
            TiempoTotal = 0
            Exit Sub
        End If
    End If
    If lComunica Then
        TecladoAnterior = False
        TiempoMonitor = 0
        TiempoTotal = 0
        Exit Sub
    Else
        If TiempoTotal >= 3000 Then
            TiempoMonitor = 0
            TmrInactivo.Enabled = False
            Unload Me
            End
        End If
    End If
End Sub

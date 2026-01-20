VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trace your Mouse!"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEvents 
      BackStyle       =   0  'Transparent
      Caption         =   "lblEvents"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   4320
   End
   Begin VB.Label lblWindowTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "lblWindowTitle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'* WARNING!!!
'* PLEASE DO NOT STOP THIS APPLICATION FROM THE VB IDE
'* CLICK THE FORM'S CROSS ICON IF YOU INTEND TO CLOSE
'********************************************************

Option Explicit

'* Decraling the mouse tracer object using WithEvents so that the events
'* can be traced
Private WithEvents MyMouseHunter As MouseHunter.Tracer
Attribute MyMouseHunter.VB_VarHelpID = -1

Private Sub Form_Load()
    '* EventStealingInfo will help to steal any particular mouse event you want.
    '* This will help you to grab the event just within your application
    '* And will not pass the event to the system
    Dim EventStealingInfo As EventThief
    
    lblWindowTitle.Caption = ""
    lblEvents.Caption = ""
    
    Set MyMouseHunter = New MouseHunter.Tracer
    
    '* Selecting the particular mouse events that you want to steal
    '* from the system
    '* Here, you'll steal the Mouse's Right Button's UP & Down events
    '* From the System, so you'll see your right button of the Mouse
    '* Is not working in anywhere in the System, But your application
    '* is still tracing the events!!
    With EventStealingInfo
        .RIGHT_DOWN = True
        .RIGHT_UP = True
    End With
    
    '* Passing the EventStealingInfo to the tracer class
    MyMouseHunter.StealMouseEvents = EventStealingInfo
    '* Start tracing the system wide mouse events
    MyMouseHunter.StartMouseTracing Me.hWnd
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '* make sure the system wide mouse tracing is stopped after the application
    '* is closed
    MyMouseHunter.StopMouseTracing
End Sub

'* Tracking the system wide mouse events
Private Sub MyMouseHunter_OnSystemMouseMove()
    '* Choose anything you want to trace when your mouse moves
    '* You can trace the X and Y coordinate of your mouse position in the window
    '* Or the Window handle (hWnd) of the window under your mouse
    '* Or the Window title (Caption) under your mouse.
    
    'lblWindowTitle.Caption = MyMouseHunter.CoordinateX & "," & MyMouseHunter.CoordinateY
    'lblWindowTitle.Caption = MyMouseHunter.WindowHandleUnderMouse
    lblWindowTitle.Caption = MyMouseHunter.WindowTextUnderMouse
    
    lblEvents.Caption = "Moving..."
End Sub

Private Sub MyMouseHunter_OnSystemMouseLeftDown()
    lblEvents.Caption = "Left Down"
End Sub

Private Sub MyMouseHunter_OnSystemMouseLeftUp()
    lblEvents.Caption = "Left Up"
End Sub

Private Sub MyMouseHunter_OnSystemMouseRightDown()
    lblEvents.Caption = "Right Down"
End Sub

Private Sub MyMouseHunter_OnSystemMouseRightUp()
    lblEvents.Caption = "Right Up"
End Sub

Private Sub MyMouseHunter_OnSystemMouseMiddleDown()
    lblEvents.Caption = "Middle Down"
End Sub

Private Sub MyMouseHunter_OnSystemMouseMiddleUp()
    lblEvents.Caption = "Middle Up"
End Sub

Private Sub MyMouseHunter_OnSystemMouseWheel()
    lblEvents.Caption = "Wheel..."
End Sub

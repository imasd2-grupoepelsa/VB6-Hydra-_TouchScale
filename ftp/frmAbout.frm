VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2520
   ClientLeft      =   6225
   ClientTop       =   3645
   ClientWidth     =   2985
   DrawMode        =   12  'Nop
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1080
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "© www.recodesoftware.com    2005, All Rights Reserved"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblInfo2 
      BackStyle       =   0  'Transparent
      Caption         =   "ReCode Software"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "QuickFTP v.1.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.Shape shpAbout 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2415
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWDEFAULT = 10
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


Private Sub lblExit_Click()
    Unload Me
End Sub

'PURPOSE: Opens default browser to display URL

'RETURNS: module handle to executed application or
'Error Code ( < 32) if there is an error

'Use one of the constants in the declarations as
'the window state parameter

'can also be used to open any document associated with
'an application on the system (e.g., passing the name
'of a file with a .doc extension will open that file in Word)

Private Function OpenLocation(URL As String, _
WindowState As Long) As Long

    Dim lHWnd As Long
    Dim lAns As Long

    lAns = ShellExecute(lHWnd, "open", URL, vbNullString, _
    vbNullString, WindowState)
   
    OpenLocation = lAns

    'ALTERNATIVE: if not interested in module handle or error
    'code change return value to boolean; then the above line
    'becomes:

    'OpenLocation = (lAns > 32)

End Function


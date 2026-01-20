VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2BD7306D-8E26-11D3-BEAB-0080C8FBFF6B}#8.0#0"; "UnRarOCX.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin UnRarOCX.UnRar UnRar1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GreGleTurK@netcourrier.com"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Extract"
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "list"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject 'See references
Dim Foldr As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Private Sub Command1_Click()
List1.Clear
If fso.FileExists(Text1.Text) = False Then
  MsgBox "File doesn't exist !"
  Exit Sub
End If
If Option1(0).Value = True Then
  UnRar1.Lister Text1.Text
Else
  UnRar1.Decompress Text1.Text, Foldr
End If
End Sub

Private Sub Command2_Click()
ShellExecute Me.hWnd, "Open", "mailto:gregleturk@netcourrier.com", "", "", 1
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 1 Then Foldr = InputBox("Please enter the folder to Decompress in :", , "c:\windows\temp")
End Sub

Private Sub UnRar1_Progression(Pourcent As Integer)
ProgressBar1.Value = Pourcent
End Sub

Private Sub UnRar1_RarFileChange(FichierEnCours As UnRarOCX.RarFile)
List1.AddItem FichierEnCours.NomFichier
End Sub

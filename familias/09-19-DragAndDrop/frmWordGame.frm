VERSION 5.00
Begin VB.Form frmWordGame 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Young Challenger Word Game"
   ClientHeight    =   8115
   ClientLeft      =   1590
   ClientTop       =   1185
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWordGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7020
      Picture         =   "frmWordGame.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7020
      Picture         =   "frmWordGame.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4500
      Width           =   855
   End
   Begin VB.CommandButton cmdCheckScore 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7020
      Picture         =   "frmWordGame.frx":2406
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7020
      Picture         =   "frmWordGame.frx":2CD0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1980
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   9480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image imgNo 
      Height          =   480
      Left            =   9120
      Picture         =   "frmWordGame.frx":359A
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Directions: With the mouse, drag the picture in the left column to the slot next to the matching word in the right column."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScoreMessage 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7020
      TabIndex        =   3
      Top             =   780
      Width           =   2535
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   9
      Left            =   2760
      Top             =   6660
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   8
      Left            =   2760
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   7
      Left            =   2760
      Top             =   5355
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   6
      Left            =   2760
      Top             =   4695
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   5
      Left            =   2760
      Top             =   4050
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   4
      Left            =   2760
      Top             =   3390
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   3
      Left            =   2760
      Top             =   2745
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   2
      Left            =   2760
      Top             =   2085
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   1
      Left            =   2760
      Top             =   1440
      Width           =   555
   End
   Begin VB.Image imgMark 
      Height          =   495
      Index           =   0
      Left            =   2760
      Top             =   780
      Width           =   555
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   9
      Left            =   4140
      TabIndex        =   19
      Top             =   6720
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   8
      Left            =   4140
      TabIndex        =   18
      Top             =   6060
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   7
      Left            =   4140
      TabIndex        =   15
      Top             =   5415
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   6
      Left            =   4140
      TabIndex        =   13
      Top             =   4755
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   5
      Left            =   4140
      TabIndex        =   11
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   4
      Left            =   4140
      TabIndex        =   10
      Top             =   3450
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   3
      Left            =   4140
      TabIndex        =   7
      Top             =   2805
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   2
      Left            =   4140
      TabIndex        =   6
      Top             =   2145
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   1
      Left            =   4140
      TabIndex        =   2
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "word here"
      Height          =   285
      Index           =   0
      Left            =   4140
      TabIndex        =   1
      Top             =   840
      Width           =   1080
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   9
      Left            =   3480
      Top             =   6660
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   3480
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   3480
      Top             =   5355
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   3480
      Top             =   4695
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   3480
      Top             =   4050
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   3480
      Top             =   3390
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   3480
      Top             =   2745
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   3480
      Top             =   2085
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   3480
      Top             =   1440
      Width           =   555
   End
   Begin VB.Image imgMatchingPic 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   3480
      Top             =   780
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   9
      Left            =   60
      Top             =   780
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   60
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   60
      Top             =   5355
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   60
      Top             =   4695
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   60
      Top             =   4050
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   60
      Top             =   3390
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   60
      Top             =   2745
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   60
      Top             =   2085
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   60
      Top             =   1440
      Width           =   555
   End
   Begin VB.Image imgPicToMatch 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   60
      Top             =   6660
      Width           =   555
   End
   Begin VB.Image imgWrong 
      Height          =   480
      Left            =   9180
      Picture         =   "frmWordGame.frx":38A4
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCorrect 
      Height          =   480
      Left            =   8625
      Picture         =   "frmWordGame.frx":416E
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   285
      Left            =   7980
      TabIndex        =   17
      Top             =   6000
      Width           =   390
   End
   Begin VB.Label lblNewGame 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      Height          =   285
      Left            =   7980
      TabIndex        =   12
      Top             =   4710
      Width           =   1155
   End
   Begin VB.Label lblCheckScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Score"
      Height          =   285
      Left            =   7980
      TabIndex        =   9
      Top             =   3420
      Width           =   1290
   End
   Begin VB.Label lblReset 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      Height          =   315
      Left            =   7980
      TabIndex        =   5
      Top             =   2100
      Width           =   735
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   24
      Left            =   8010
      Picture         =   "frmWordGame.frx":4A38
      Tag             =   "star"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   23
      Left            =   6795
      Picture         =   "frmWordGame.frx":5CAA
      Tag             =   "spaceship"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   22
      Left            =   6180
      Picture         =   "frmWordGame.frx":6F1C
      Tag             =   "tree"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   21
      Left            =   7410
      Picture         =   "frmWordGame.frx":818E
      Tag             =   "traffic light"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   20
      Left            =   5565
      Picture         =   "frmWordGame.frx":9400
      Tag             =   "sun"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   19
      Left            =   4950
      Picture         =   "frmWordGame.frx":A672
      Tag             =   "submarine"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   18
      Left            =   60
      Picture         =   "frmWordGame.frx":B8E4
      Tag             =   "castle"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   17
      Left            =   4350
      Picture         =   "frmWordGame.frx":CB56
      Tag             =   "robot"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   16
      Left            =   6180
      Picture         =   "frmWordGame.frx":DDC8
      Tag             =   "heart"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   15
      Left            =   3735
      Picture         =   "frmWordGame.frx":F03A
      Tag             =   "present"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   14
      Left            =   3120
      Picture         =   "frmWordGame.frx":102AC
      Tag             =   "popsicle"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   13
      Left            =   2505
      Picture         =   "frmWordGame.frx":1151E
      Tag             =   "paddle ball"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   12
      Left            =   1890
      Picture         =   "frmWordGame.frx":12790
      Tag             =   "house"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   11
      Left            =   1290
      Picture         =   "frmWordGame.frx":13A02
      Tag             =   "computer"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   10
      Left            =   675
      Picture         =   "frmWordGame.frx":14C74
      Tag             =   "moon"
      Top             =   7560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   9
      Left            =   9120
      Picture         =   "frmWordGame.frx":15EE6
      Tag             =   "ladybug"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   8
      Left            =   8535
      Picture         =   "frmWordGame.frx":17158
      Tag             =   "ice cream cone"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   7
      Left            =   7950
      Picture         =   "frmWordGame.frx":183CA
      Tag             =   "fish"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   6
      Left            =   7350
      Picture         =   "frmWordGame.frx":1963C
      Tag             =   "fire plug"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   5
      Left            =   6765
      Picture         =   "frmWordGame.frx":1A8AE
      Tag             =   "dog"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   4
      Left            =   5580
      Picture         =   "frmWordGame.frx":1BB20
      Tag             =   "bone"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   3
      Left            =   5010
      Picture         =   "frmWordGame.frx":1CD92
      Tag             =   "pencil"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   2
      Left            =   4440
      Picture         =   "frmWordGame.frx":1E004
      Tag             =   "baseball glove"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   1
      Left            =   1890
      Picture         =   "frmWordGame.frx":1F276
      Tag             =   "balloons"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgSourcePic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   0
      Left            =   1320
      Picture         =   "frmWordGame.frx":204E8
      Tag             =   "airplane"
      Top             =   6960
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmWordGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintNumTries    As Integer

'------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------
    
    CenterForm Me
    
    NewGameInit

End Sub

'------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------
    
    Dim inUserResponse As Integer

    inUserResponse = MsgBox("Are you sure you want to exit?", _
                            vbYesNo + vbQuestion + vbDefaultButton1, _
                            "Young Challenger Word Game")
    
    If inUserResponse = vbNo Then
        Cancel = 1
    End If

End Sub

'------------------------------------------------------------------------
Private Sub imgMatchingPic_DragDrop(Index As Integer, _
                                    Source As Control, _
                                    X As Single, _
                                    Y As Single)
'------------------------------------------------------------------------

    ' This event occurs when the user has dragged an image from the
    ' left-hand column of images (the "imgPicToMatch" control array) and
    ' has released the mouse to drop the image into an empty slot in the
    ' right-hand column (the "imgMatchingPic" control array).
    ' The image being dragged from the imgPicToMatch control array is
    ' passed as the "Source" argument to this event.

    ' (This event will also occur if the user has dropped an image into
    ' a slot on the right, but changes their mind and drags it from that
    ' slot on the right to a different slot on the right.)

    ' If the user attempts to drop the image into a slot that already
    ' has a picture in it, do nothing and exit the sub ...
    If imgMatchingPic(Index).Picture <> LoadPicture() Then Exit Sub
    
    ' Populate the Picture and Tag properties of the "target" image
    ' with those of the "Source" image ...
    imgMatchingPic(Index).Picture = Source.Picture
    imgMatchingPic(Index).Tag = Source.Tag
    
    ' Set the DragIcon of the target image to be its Picture (in case
    ' the user wants to drag the picture back out because they changed
    ' their mind) ...
    imgMatchingPic(Index).DragIcon = imgMatchingPic(Index).Picture
    
    ' Clear the Picture and DragIcon from the "Source" image ...
    Source.Picture = LoadPicture()
    Source.DragIcon = LoadPicture()

End Sub

'------------------------------------------------------------------------
Private Sub imgPicToMatch_DragOver(Index As Integer, _
                                   Source As Control, _
                                   X As Single, _
                                   Y As Single, _
                                   State As Integer)
'------------------------------------------------------------------------

    ' This event occurs when a user is dragging an image over an element
    ' in the left-hand column (the imgPicToMatch control array).  If the
    ' image is "legal" for the slot being "hovered over", then the mouse
    ' icon will be the image that goes in that slot; otherwise, the mouse
    ' icon will be the "No" symbol.
    
    ' This can occur in the following situations:
    ' (1) The user is initally dragging an image from the left-hand side
    '     but hovers over another image or empty slot in the left-hand column
    '     prior to either dropping it either into an empty slot on the right
    '     or back in its original slot on the left.
    ' (2) After the user has dropped the image into a slot on the right, but
    '     then changes their mind and decides to drag it back to its original
    '     slot on the left.

    If State = 1 Then
        ' leaving (moving out of the control)
        Source.DragIcon = Source.Picture
    Else
        If imgPicToMatch(Index).Tag = Source.Tag Then
            Source.DragIcon = Source.Picture
        Else
            Source.DragIcon = imgNo.Picture
        End If
    End If

End Sub

'------------------------------------------------------------------------
Private Sub imgPicToMatch_DragDrop(Index As Integer, _
                                   Source As Control, _
                                   X As Single, _
                                   Y As Single)
'------------------------------------------------------------------------

    ' This event occurs when the user attempts to drop an image into a
    ' slot of the left-hand column (the imgPicToMatch control array).
    ' This will typically occur after the user moved an image to a slot on
    ' the right, but then decides to move it back to its original slot on
    ' the left.
    ' However, this event will also occur if the user attempts to move an
    ' image from one slot on the left to another slot on the left (but the
    ' "If" tests in the code will not allow the drop to take place).

    ' If the user attempts to drop the image into a slot that already
    ' has a picture in it, do nothing and exit the sub ...
    If imgPicToMatch(Index).Picture <> LoadPicture() Then Exit Sub
    
    ' Make sure that this is the original slot for the image being dropped ...
    If imgPicToMatch(Index).Tag = Source.Tag Then
        ' Populate the Picture and DragIcon properties of the "target" image
        ' with those of the "Source" image ...
        imgPicToMatch(Index).Picture = Source.Picture
        imgPicToMatch(Index).DragIcon = Source.Picture
        ' Clear the Picture and DragIcon from the "Source" image ...
        Source.Picture = LoadPicture()
        Source.DragIcon = LoadPicture()
    End If

End Sub

'------------------------------------------------------------------------
Private Sub cmdReset_Click()
'------------------------------------------------------------------------

    Dim intX    As Integer
    Dim intY    As Integer
    
    ' Loop thru the imgMatchingPic control array ...
    For intX = 0 To 9
        ' If the user got the answer wrong ...
        If imgMatchingPic(intX).Tag <> lblWord(intX).Caption Then
            ' If they left the image blank, then just clear the "red X" mark ...
            If imgMatchingPic(intX).Picture = LoadPicture() Then
                imgMark(intX).Picture = LoadPicture()
            Else
                ' Otherwise, move the picture back to its original slot on the
                ' left and clear it from the wrong slot on the right ...
                For intY = 0 To 9
                    If imgPicToMatch(intY).Tag = imgMatchingPic(intX).Tag Then
                        imgPicToMatch(intY).Picture = imgMatchingPic(intX).Picture
                        imgPicToMatch(intY).DragIcon = imgPicToMatch(intY).Picture
                        imgMatchingPic(intX).Picture = LoadPicture()
                        imgMark(intX).Picture = LoadPicture()
                        Exit For
                    End If
                Next
            End If
            ' Set the DragMode for this slot back to Automatic ...
            imgMatchingPic(intX).DragMode = vbAutomatic
        End If
    Next
    
    ' Remove the Reset button from the screen at this time ...
    lblReset.Visible = False
    cmdReset.Visible = False
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdCheckScore_Click()
'------------------------------------------------------------------------

    Dim intX                As Integer
    Dim intNumCorrect       As Integer
    Dim intPercentCorrect   As Integer
    Dim strCongrats         As String
    Dim strTries            As String
    
    ' Loop thru the imgMatchingPic control array and "mark" the answer
    ' as correct or wrong ...
    
    For intX = 0 To 9
        If imgMatchingPic(intX) = LoadPicture() Then
            ' If the picture was left blank, then mark it wrong (place
            ' the "red X" next to the image) ...
            imgMark(intX).Picture = imgWrong.Picture
        ElseIf imgMatchingPic(intX).Tag = lblWord(intX).Caption Then
            ' If the Tag of the image that was dropped into the slot
            ' matches the caption of the corresponding element of the
            ' lblWord Label control array, then mark it correct (place
            ' the green checkmark next to the image) and increment the
            ' counter for the number of correct answers.
            imgMark(intX).Picture = imgCorrect.Picture
            intNumCorrect = intNumCorrect + 1
        Else
            ' The Tag of the image that was dropped into the slot does
            ' NOT match the caption of the corresponding element of the
            ' lblWord Label control array, so mark it wrong (place
            ' the "red X" next to the image) ...
            imgMark(intX).Picture = imgWrong.Picture
        End If
        ' Disable automatic dragging for the moment by setting the DragMode
        ' to Manual ...
        imgMatchingPic(intX).DragMode = vbManual
    Next
    
    ' Increment the number of tries for this game ...
    mintNumTries = mintNumTries + 1
    
    ' Calculate the percent correct ...
    intPercentCorrect = (intNumCorrect / 10) * 100
    
    ' Formulate a message indicating how many answers are correct and in how
    ' many tries ...
    If mintNumTries = 1 Then
        strTries = "try"
        If intPercentCorrect = 100 Then
            strCongrats = "Congratulations! "
        End If
    Else
        strTries = "tries"
    End If
    
    Beep
    
    lblScoreMessage.Visible = True
    lblScoreMessage.Caption = strCongrats & "You scored " & _
                              intPercentCorrect & "% in " & mintNumTries & _
                              " " & strTries & "."
                              
    ' If the user did not get a perfect score, make the Reset button visible
    ' so they can go back and correct the answers they got wrong ...
    If intPercentCorrect < 100 Then
        lblReset.Visible = True
        cmdReset.Visible = True
        cmdReset.Default = True
    Else
        lblReset.Visible = False
        cmdReset.Visible = False
    End If
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdNewGame_Click()
'------------------------------------------------------------------------

    Dim intX    As Integer
    
    ' Clear the "marks" as well as the images in the right-hand column ...
    For intX = 0 To 9
        imgMark(intX).Picture = LoadPicture()
        imgMatchingPic(intX).Picture = LoadPicture()
    Next intX
    
    ' Reset number of tries counter ...
    mintNumTries = 0
    
    ' Prepare for a new game ...
    Call NewGameInit
    
    ' Make the "Check Score" button the default button ...
    cmdCheckScore.Default = True
    cmdCheckScore.SetFocus
    
End Sub

'------------------------------------------------------------------------
Private Sub cmdExit_Click()
'------------------------------------------------------------------------
    Unload Me
End Sub

'------------------------------------------------------------------------
Private Sub NewGameInit()
'------------------------------------------------------------------------

    Dim blnNumberExists         As Boolean
    Dim intRandomNumber         As Integer
    Dim aintPicIndex(0 To 9)    As Integer
    Dim aintWordIndex(0 To 9)   As Integer
    Dim intX                    As Integer
    Dim intY                    As Integer
    
    'Reseed the VB's random number generator
    Randomize
    
    'Set the 10 pictures that will be used in this game ...
    
    For intX = 0 To 9
        
        ' Get a random number between 0 and 24, make sure it is not
        ' one we already have ...
        Do
            intRandomNumber = GetRandomNumber(0, 24)
            blnNumberExists = False
            intY = 0
            
            Do Until intY >= intX Or blnNumberExists
               If intRandomNumber = aintPicIndex(intY) Then
                    blnNumberExists = True
                Else
                    intY = intY + 1
                End If
            Loop
            
        Loop While blnNumberExists
        
        ' Store the new random number in the aintPicIndex array, so we
        ' can check for the next go-round ...
        aintPicIndex(intX) = intRandomNumber
        
        ' Load the picture from the source image array to the
        ' current set for this game ...
        imgPicToMatch(intX).Picture = imgSourcePic(intRandomNumber).Picture
        imgPicToMatch(intX).Tag = imgSourcePic(intRandomNumber).Tag
        
        ' set the image's drag icon to match its picture
        imgPicToMatch(intX).DragIcon = imgPicToMatch(intX).Picture
    
        ' allow automatic dragging
        imgMatchingPic(intX).DragMode = vbAutomatic
        
    Next intX

    ' Using the Tag property of the 10 images to be used in this game, "mix 'em up" and
    ' set the "word" label that appears to the right of the second column of images.
    
    For intX = 0 To 9
        Do
            intRandomNumber = GetRandomNumber(0, 9)
            blnNumberExists = False
            intY = 0
            Do Until intY >= intX Or blnNumberExists
                If intRandomNumber = aintWordIndex(intY) Then
                    blnNumberExists = True
                Else
                    intY = intY + 1
                End If
            Loop
        Loop While blnNumberExists
        aintWordIndex(intX) = intRandomNumber
        lblWord(intX).Caption = imgPicToMatch(intRandomNumber).Tag
    Next intX

    ' No need to see the "scoring box" or Reset button at this time ...
    lblScoreMessage.Visible = False
    lblReset.Visible = False
    cmdReset.Visible = False
 
End Sub


'------------------------------------------------------------------------
Private Function GetRandomNumber(pintLowerBound As Integer, _
                                 pintUpperBound As Integer) _
As Integer
'------------------------------------------------------------------------
    
    ' This function will return a random integer that falls with the range
    ' of the two arguments passed.

    GetRandomNumber = Int((pintUpperBound - pintLowerBound + 1) * Rnd + pintLowerBound)

End Function

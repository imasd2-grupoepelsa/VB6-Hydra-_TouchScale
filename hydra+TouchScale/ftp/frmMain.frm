VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                QFTP"
   ClientHeight    =   8325
   ClientLeft      =   240
   ClientTop       =   1500
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   762
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   7950
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17542
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:36 PM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get info"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   10560
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   19
      Top             =   0
      Width           =   735
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "E&xit"
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
         Height          =   495
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picResource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   960
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   7440
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   9600
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   12
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&About"
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
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   1065
      TabIndex        =   10
      Top             =   0
      Width           =   1095
      Begin VB.Label cmdDisconnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Disconnect"
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
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   0
      Width           =   1095
      Begin VB.Label cmdConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Connect"
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
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   5895
      Left            =   5640
      TabIndex        =   6
      Top             =   1560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10398
      _Version        =   393217
      Indentation     =   973
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DriveListBox DriveList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.DirListBox DirList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00808000&
      Height          =   1665
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.FileListBox FileList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3600
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin MSComctlLib.ImageList imglMenu 
      Left            =   4920
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":005E
            Key             =   "About"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00BC
            Key             =   "Preferences"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":011A
            Key             =   "Connect"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0178
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01D6
            Key             =   "Disconnect"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglImages 
      Left            =   4320
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label cmdDelete 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblLocalFileSize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label lblTransferInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                Transfer Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Top             =   0
      Width           =   7560
   End
   Begin VB.Label cmdUpload 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Upload File"
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
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label cmdDownload 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Download File"
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
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Address:"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect...|@Connect To FTP|#Connect|^"" & vbBlue"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileDisconnect 
         Caption         =   "&Disconnect...|@Disconnect From FTP|#Disconnect|^"" & vbBlue"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit...|@Exit QFTP|#Exit|^"" & vbBlue"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsPreferences 
         Caption         =   "&Preferences...|@Preferences|#Preferences|^"" & vbBlue"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help...|@QFTP Help|#Help|^"" & vbBlue"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About...|@About|#About|^"" & vbBlue"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MouseOver
'Private WithEvents MenuEvents As CEvents

Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)

Private WithEvents m_frmTray As frmTray
Attribute m_frmTray.VB_VarHelpID = -1

'Private m_cN As cNeoCaption
'Private m_cSkin As cSkinConfiguration
Private m_iIndex As Long

Dim hOpen As Long, hConnection As Long, hFile As Long

Private UserFile As String
Private UserFileDel As String

Private lFileSize As Long

'Delete function
'============================================================
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
'=============================================================
'end of delete function

Private m_strFileName As String
Private m_lngFileSize As Long

Dim bActiveSession As Boolean
Dim bDirEmpty As Boolean
Dim dwType As Long

Dim EnumItemNameBag As New Collection
Dim EnumItemAttributeBag As New Collection

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'mouseover function
'============================================================



'=============================================================
'end of mouseover function

Private Sub FadeBack(p As Integer)
        Picture1(p).BackColor = RGB(0, 128, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(15, 134, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(25, 138, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(35, 144, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(45, 154, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(55, 166, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(65, 177, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(75, 188, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(85, 194, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(95, 199, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(105, 205, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(115, 215, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(125, 225, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(135, 228, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(145, 230, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(155, 235, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(165, 238, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(175, 240, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(185, 242, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(195, 244, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(205, 246, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(215, 248, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(225, 250, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(235, 252, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(245, 254, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(255, 255, 255)
End Sub
Private Sub Fade(p As Integer)
        Picture1(p).BackColor = RGB(255, 255, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(245, 255, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(235, 254, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(225, 252, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(215, 250, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(205, 248, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(195, 246, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(185, 244, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(175, 242, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(165, 240, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(155, 235, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(145, 230, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(135, 225, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(125, 220, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(115, 215, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(105, 205, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(95, 199, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(85, 194, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(75, 188, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(65, 177, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(55, 166, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(45, 154, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(35, 144, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(25, 138, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(15, 134, 255)
        Sleep 1
        Picture1(p).BackColor = RGB(0, 128, 255)
End Sub
Private Sub OnDataTransferProgress_(ByVal lngBytesTransferred As Long)
    '
    Dim strCaption As String
    '
    
    '
    lblTransferInfo.Caption = "Downloading" & m_strFileName & " (" & _
                 lngBytesTransferred & " bytes transferred)..."
    '
    'If Not hConnection = ftpRetrievingDirectoryInfo Then
        ProgressBar1.Value = lngBytesTransferred / (m_lngFileSize / 100)
    'End If
    '
End Sub

Private Sub cmdConnect_Click()

    cmdConnect.Caption = "Connecting..."
    With sbStatus.Panels(1)
        .Text = "Connecting to " & txtAddress.Text
    End With
    Me.MousePointer = vbHourglass
    'ftp://ftp.freebsd.org/pub/
    'Open FTP Site
            
    If Not bActiveSession And hOpen <> 0 Then
        If txtAddress.Text = "" Then
            MsgBox "Please enter a server name!"
            Exit Sub
        End If
        
        Dim nFlag As Long
        
        Select Case True
        Case frmOptions.optConmode(1).Value
            nFlag = INTERNET_FLAG_PASSIVE
        Case frmOptions.optConmode(0).Value
            nFlag = INTERNET_FLAG_ACTIVE
        End Select
        
        Select Case True
        Case frmOptions.optTranMode(0).Value
            dwType = FTP_TRANSFER_TYPE_BINARY
        Case frmOptions.optTranMode(1).Value
            dwType = FTP_TRANSFER_TYPE_ASCII
        End Select
        
        hConnection = InternetConnect(hOpen, txtAddress.Text, INTERNET_INVALID_PORT_NUMBER, _
        frmOptions.txtUser, frmOptions.txtPassword, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            bActiveSession = False
            
            Me.MousePointer = vbDefault
            With sbStatus.Panels(1)
                '.Text = "| Not Connected |"
            End With
        Else
            bActiveSession = True
            'EnableUI (CBool(hOpen))
            FillTreeViewControl (txtAddress.Text)
            FtpEnumDirectory ("")
            If EnumItemNameBag.Count = 0 Then Exit Sub
            FillTreeViewControl (txtAddress.Text)
            Me.MousePointer = vbDefault
            cmdConnect.Caption = "Connected"
            With sbStatus.Panels(1)
                '.Text = "Connected to " & txtAddress.Text & " as " & frmOptions.txtUser.Text
            End With
       End If
    End If
      
    SaveSetting App.EXEName, "Settings", "Server Address", txtAddress.Text

End Sub
Private Sub Clear()
    Dim Num As Integer
    For Num = 1 To EnumItemNameBag.Count
        EnumItemNameBag.Remove 1
    Next Num
    For Num = 1 To EnumItemAttributeBag.Count
        EnumItemAttributeBag.Remove 1
    Next Num
End Sub
Private Sub rcd(pszDir As String)
    If pszDir = "" Then
        MsgBox "Please enter the directory to CD"
        Exit Sub
    Else
        Dim sPathFromRoot As String
        Dim bRet As Boolean
        If InStr(1, pszDir, txtAddress.Text) Then
        sPathFromRoot = Mid(pszDir, Len(txtAddress.Text) + 1, Len(pszDir) - Len(txtAddress.Text))
        Else
        sPathFromRoot = pszDir
        End If
        If sPathFromRoot = "" Then sPathFromRoot = "/"
        bRet = FtpSetCurrentDirectory(hConnection, sPathFromRoot)
        If bRet = False Then Exit Sub
    End If
End Sub

Private Sub FtpEnumDirectory(strDirectory As String)
    
    Clear
    Dim hFind As Long
    Dim nLastError As Long
    Dim dError As Long
    Dim ptr As Long
    Dim pData As WIN32_FIND_DATA
    
    If Len(strDirectory) > 0 Then rcd (strDirectory)
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    nLastError = Err.LastDllError
    
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            MsgBox "This directory is empty!"
        Else
            Exit Sub
        End If
        Exit Sub
    End If
    
    dError = NO_ERROR
    Dim bRet As Boolean
    Dim strItemName As String
    
    EnumItemAttributeBag.Add pData.dwFileAttributes
    strItemName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    EnumItemNameBag.Add strItemName
    Do
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
               
                InternetCloseHandle (hFind)
               Exit Sub
            End If
        Else
            EnumItemAttributeBag.Add pData.dwFileAttributes
            strItemName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            EnumItemNameBag.Add strItemName
       End If
    Loop
    
    InternetCloseHandle (hFind)
End Sub

Private Sub cmdDelete_Click()
    Dim op As SHFILEOPSTRUCT

    With FileList
      If Right(.Path, 1) = "\" Then
         UserFileDel = .Path & .FileName
      Else
         UserFileDel = .Path & "\" & .FileName
      End If
    End With

    With op
        .wFunc = FO_DELETE
        .pFrom = UserFileDel
        .fFlags = FOF_ALLOWUNDO
    End With
    
    SHFileOperation op
    FileList.Refresh
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdDelete.BackColor = &H80000013
End Sub

Private Sub cmdDisconnect_Click()
cmdConnect.Caption = "Connect"

bDirEmpty = True
  If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
    Clear
    TreeView.Nodes.Remove txtAddress.Text
    bActiveSession = False
    
With sbStatus.Panels(1)
    .Text = "| Not Connected |"
End With

End Sub

Private Sub cmdDownload_Click()
    Dim bRet As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim nPos As Long, nTemp As Long
    Dim nodX As Node
    Dim Size As Long
    Dim Sum As Long
    Sum = 0
    
    Set nodX = TreeView.SelectedItem
    If bActiveSession Then
        If nodX Is Nothing Then
            MsgBox "Please select the item to Download!"
            Exit Sub
        End If
        
       pb.Visible = True
       
       
    End If
        szTempString = TreeView.SelectedItem.Text
        
        szFileRemote = szTempString
        nPos = 0
        nTemp = 0
        Do
            nTemp = InStr(1, szTempString, "/", vbBinaryCompare)
            If nTemp = 0 Then Exit Do
            szTempString = Right(szTempString, Len(szTempString) - nTemp)
            nPos = nTemp + nPos
        Loop
        szDirRemote = Left(szFileRemote, nPos)
        szFileRemote = Right(szFileRemote, Len(szFileRemote) - nPos)
        szFileLocal = FileList.Path
        rcd szDirRemote
        bRet = FtpGetFile(hConnection, szFileRemote, szFileLocal & "/" & szFileRemote, False, _
        INTERNET_FLAG_RELOAD, dwType, 0)
        FileList.Refresh
        If bRet = False Then
            Exit Sub
        Else
            MsgBox "Not in session"
        End If
End Sub


Private Sub cmdUpload_Click()
Dim bRet As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim nPos As Long, nTemp As Long
    Dim nodX As Node
    Set nodX = TreeView.SelectedItem
  
    If bActiveSession Then
        If nodX Is Nothing Then
            MsgBox "Please select a remote directory to Upload to!"
            Exit Sub
        End If
        If nodX.Image = "leaf" Then
            MsgBox "Please select a remote directory to Upload to!"
            Exit Sub
        End If
        If FileList.FileName = "" Then
            MsgBox "Please select a local file to Upload"
            Exit Sub
        End If
        szTempString = nodX.Text
        szDirRemote = Right(szTempString, Len(szTempString) - Len(txtAddress.Text))
        szFileRemote = FileList.FileName
        szFileLocal = FileList.Path & "\" & FileList.FileName
        If (szDirRemote = "") Then szDirRemote = "\"
        rcd szDirRemote
        
        bRet = FtpPutFile(hConnection, szFileLocal, szFileRemote, _
         dwType, 0)
        If bRet = False Then
            MsgBox "could not upload file!"
            Exit Sub
        End If
        
        Dim nodChild As Node, nodNextChild As Node
        Set nodChild = nodX.Child
        Do
          If nodChild Is Nothing Then Exit Do
          Set nodNextChild = nodChild.Next
            TreeView.Nodes.Remove nodChild.Index
            If nodNextChild Is Nothing Then Exit Do
            Set nodChild = nodNextChild
        Loop
        If nodX.Image = "closed" Then
            nodX.Image = "open"
        End If
        FtpEnumDirectory (nodX.Text)
        FillTreeViewControl (nodX.Text)
   End If
End Sub

Private Sub DirList_Change()
FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
On Error GoTo ErrProc
    DirList.Path = DriveList.Drive
    Exit Sub
ErrProc:
    DriveList.Drive = "c:"
    DirList.Path = DriveList.Drive
End Sub

Private Sub FileList_Click()

 With FileList
      If Right(.Path, 1) = "\" Then
         UserFile = .Path & .FileName
      Else
         UserFile = .Path & "\" & .FileName
      End If
 End With

 lFileSize = FileLen(UserFile)
 lblLocalFileSize.Caption = lFileSize & " KB"
 
End Sub

Private Sub Form_Load()
Dim Index As Integer
Dim Click As Boolean
    
    
    'Set m_cN = New cNeoCaption

    'SelectNewSkin Index
       ' m_cN.Attach2 Me, m_cSkin
       ' setColors
       ' If Index = 5 Then
        '    frmMain.BackColor = RGB(140, 139, 139)
        'End If
        
    bActiveSession = False
    hOpen = 0
    hConnection = 0

If Len(frmOptions.txtPort.Text) <> 0 Then
        hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PROXY, frmOptions.txtPort.Text, vbNullString, 0)
    Else
        hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    End If
    If hOpen = 0 Then
    Exit Sub
    End If

Dim imgI As ListImage
    Set imgI = imglImages.ListImages.Add(, "open", LoadPicture(App.Path & "\" & "\open.bmp"))
    Set imgI = imglImages.ListImages.Add(, "closed", LoadPicture(App.Path & "\" & "\closed.bmp"))
    Set imgI = imglImages.ListImages.Add(, "leaf", LoadPicture(App.Path & "\" & "\leaf.bmp"))
    Set imgI = imglImages.ListImages.Add(, "root", LoadPicture(App.Path & "\" & "\root.bmp"))
    TreeView.ImageList = imglImages
    TreeView.Style = tvwTreelinesPictureText

Set m_frmTray = New frmTray
        With m_frmTray
            .AddMenuItem "&Restore QFTP", "Restore", True
            .AddMenuItem "-"
            .AddMenuItem "QFTP on the Web", "QFTPWeb"
            .AddMenuItem "&About...", "About"
            .AddMenuItem "-"
            .AddMenuItem "E&xit", "Exit"
            .ToolTip = "SysTray QFTP!"
        End With
        SetIcon
        
Const MENU_XP_2003 = 3
Dim sFileIni As String

  
'Set MenuEvents = New CEvents
'Set objMenuEx = New cMenuEx
 

sFileIni = App.Path & "\" & Me.Name
'Call objMenuEx.Install(Me.hWnd, _
''    sFileIni, _
'    imglMenu, _
 '   MENU_XP_2003, _
 '   MenuEvents)
Refresh
frmTray.Hide
' Get text1's value
txtAddress.Text = GetSetting(App.EXEName, "Settings", "Server Address")

End Sub
Private Sub SetIcon()
    
    m_frmTray.IconHandle = frmTray.imgTwo.Picture.Handle
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload m_frmTray
    Set m_frmTray = Nothing
End Sub

Private Sub Form_Resize()
If frmMain.WindowState = vbMinimized Then
frmMain.Hide
m_frmTray.ShowBalloonTip "You have minimized QFTP to the system tray, to restore the application just right click on this icon then click on 'Restore QFTP'.", "QFTP is still running!", NIIF_INFO
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lblAbout_Click()
frmAbout.Show
End Sub


Private Sub lblExit_Click()
bDirEmpty = True
  If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
    Clear
    bActiveSession = False
    
Unload frmTray
Unload frmMain
End Sub


Private Sub mnuFileConnect_Click()
With sbStatus.Panels(1)
        .Text = "Connecting to " & txtAddress.Text
    End With
    Me.MousePointer = vbHourglass
    'ftp://ftp.freebsd.org/pub/
    'Open FTP Site
            
    If Not bActiveSession And hOpen <> 0 Then
        If txtAddress.Text = "" Then
            MsgBox "Please enter a server name!"
            Exit Sub
        End If
        
        Dim nFlag As Long
        
        Select Case True
        Case frmOptions.optConmode(1).Value
            nFlag = INTERNET_FLAG_PASSIVE
        Case frmOptions.optConmode(0).Value
            nFlag = INTERNET_FLAG_ACTIVE
        End Select
        
        Select Case True
        Case frmOptions.optTranMode(0).Value
            dwType = FTP_TRANSFER_TYPE_BINARY
        Case frmOptions.optTranMode(1).Value
            dwType = FTP_TRANSFER_TYPE_ASCII
        End Select
        
        hConnection = InternetConnect(hOpen, txtAddress.Text, INTERNET_INVALID_PORT_NUMBER, _
        frmOptions.txtUser, frmOptions.txtPassword, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            bActiveSession = False
            
            Me.MousePointer = vbDefault
            With sbStatus.Panels(1)
                .Text = "| Not Connected |"
            End With
        Else
            bActiveSession = True
            'EnableUI (CBool(hOpen))
            FillTreeViewControl (txtAddress.Text)
            FtpEnumDirectory ("")
            If EnumItemNameBag.Count = 0 Then Exit Sub
            FillTreeViewControl (txtAddress.Text)
            Me.MousePointer = vbDefault
            cmdConnect.Caption = "Connected"
            With sbStatus.Panels(1)
                .Text = "Connected to " & txtAddress.Text & " as " & frmOptions.txtUser.Text
            End With
       End If
    End If
      
    SaveSetting App.EXEName, "Settings", "Server Address", txtAddress.Text
End Sub

Private Sub mnuFileDisconnect_Click()

cmdConnect.BackColor = vbWhite
cmdConnect.Caption = "Connect"

TreeView.Nodes.Clear
With sbStatus.Panels(1)
    .Text = "| Not Connected |"
End With
End Sub

Private Sub mnuFileExit_Click()
bDirEmpty = True
  If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
    Clear
    bActiveSession = False
Unload frmTray
Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuToolsPreferences_Click()
frmOptions.Show
frmMain.Enabled = False
End Sub

Private Sub FillTreeViewControl(strParentKey As String)
    Dim nodX As Node
    Dim strImg As String
    Dim nCount As Integer, i As Integer
    Dim nAttr As Integer
    Dim strItem As String
    
    If EnumItemNameBag.Count = 0 And strParentKey = txtAddress.Text Then
        Set nodX = TreeView.Nodes.Add(, tvwFirst, txtAddress.Text, txtAddress.Text, "root")
        Exit Sub
    End If
    nCount = EnumItemAttributeBag.Count
    If nCount = 0 Then Exit Sub
    For i = 1 To nCount
        nAttr = EnumItemAttributeBag.Item(i)
        strItem = EnumItemNameBag(i)
        If nAttr = FILE_ATTRIBUTE_DIRECTORY Then
            strImg = "closed"
        Else
            strImg = "leaf"
        End If
        Set nodX = TreeView.Nodes.Add(strParentKey, tvwChild, strParentKey & "/" & strItem, _
            strParentKey & "/" & strItem, strImg)
    Next
    nodX.EnsureVisible
End Sub

Public Sub SelectNewSkin(ByVal iIndex As Long)
Dim sFont As New StdFont
   
   Set m_cSkin = New cSkinConfiguration
   With m_cSkin
         .Name = "Neo vbAccelerator"
         .ActiveCaptionColor = &HFFFFFF
         .InActiveCaptionColor = RGB(47, 122, 151)
         .ActiveMenuColor = &H0&
         .ActiveMenuColorOver = &HFFFFFF
         .InActiveMenuColor = RGB(47, 122, 151)
         .MenuBackgroundColor = &H80000013    'RGB(123, 189, 214)
         Set .CaptionFont = Me.Font
         Set .MenuFont = Me.Font
         Set .Caption = picResource(0).Picture
         Set .Borders = picResource(1).Picture
         .ButtonWidth = 19
         .ButtonHeight = 20
         .ActiveLeftEnd = 90
         .ActiveRightStart = 140
         .ActiveRightEnd = 240
         .InactiveOffset = 400
         .CaptionFont.Bold = True
   End With
   
End Sub

Private Sub setColors()
  
      Me.BackColor = m_cN.MenuBackgroundColor
      frmMain.BackColor = m_cN.MenuBackgroundColor
      frmMain.ForeColor = m_cN.ActiveMenuColorOver
  
End Sub


Private Sub TreeView_Click()
       ' m_strFileName = TreeView.SelectedItem.Text
        'm_lngFileSize = CLng(TreeView.SelectedItem.subitem(1))

        'lblTransferInfo.Caption = m_strFileName & m_lngFileSize
End Sub

Private Sub TreeView_DblClick()
Dim nodX As Node
    Set nodX = TreeView.SelectedItem
    If Not bActiveSession Then
        MsgBox "No in session!"
        Exit Sub
    End If
    If nodX Is Nothing Then
        MsgBox "no Selection to enumerate"
    End If
    If nodX.Image = "closed" Then
        nodX.Image = "open"
        FtpEnumDirectory (nodX.Text)
        FillTreeViewControl (nodX.Text)
    Else
        If nodX.Image = "open" Then
            nodX.Image = "closed"
            Dim nodChild As Node, nodNextChild As Node
            Set nodChild = nodX.Child
            Do
            Set nodNextChild = nodChild.Next
                TreeView.Nodes.Remove nodChild.Index
                If nodNextChild Is Nothing Then Exit Do
                Set nodChild = nodNextChild
            Loop
        End If
    End If
End Sub

Private Sub m_frmTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Dim Result As Long

   Select Case sKey
    Case "Restore"
        frmMain.WindowState = vbNormal
        Result = SetForegroundWindow(frmMain.hWnd)
        If frmMain.WindowState = vbMinimized Then
            frmMain.WindowState = vbNormal
        End If
        frmMain.Show
    Case "Exit"
        Unload frmMain
        Unload Me
    Case "About"
        frmAbout.Show
    
   End Select
    
End Sub

Private Sub m_frmTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    Me.Show
    Me.ZOrder
End Sub

Private Sub m_frmTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmTray.ShowMenu
    End If
End Sub


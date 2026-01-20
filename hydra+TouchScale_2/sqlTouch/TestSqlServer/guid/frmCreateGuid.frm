VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " Demo Generating a Globally Unique Identifier (GUID)"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   6300
   Icon            =   "frmCreateGuid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6300
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy to Clipboard"
      Height          =   495
      Index           =   1
      Left            =   3180
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate New GUID"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click..."
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6075
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Test Selected GUID"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Length of Selected GUID"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Create GUID with Prefix"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowItem 
         Caption         =   "mnuWindowItem"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmCreateGUID - Demo creating a Globally Unique Identifier (GUID)
' 1997/07/11 Copyright 1997, Larry Rebich, The Bridge, Inc.

    Option Explicit
    DefLng A-Z
    
    Dim mlMinWidth As Long  '2005/12/18 Added
    Dim mlMinHeight As Long '2005/12/18 Added
    '
    
Private Sub Command1_Click(Index As Integer)
    Dim sTemp As String
    Dim o As New cCreateGUID    '2005/12/18 Use class
    Select Case Index
        Case 0          'generate
            With List1
                .AddItem " " & o.CreateGUID()   'generate it
                .Selected(.ListCount - 1) = True
            End With
            Command1_Click 1        'copy to the clipboard
        Case 1          'copy
            With List1
                sTemp = Trim$(.Text)
                Clipboard.Clear         '2001/10/25
                Clipboard.SetText sTemp
            End With
    End Select
End Sub

Private Sub Form_Load()
    SetupForm
    Command1_Click 0    'generate
End Sub

Private Sub SetupForm()
' 2005/12/18 Sub added by Larry Rebich using the DELL8500 while in Indio, CA
    Dim l, t, w, h
    Dim iState As Integer
    Dim sCap As String
    
    gsCompany = "Buygold"
    gsProductName = App.ProductName
    
    sCap = "Click 'Generate...' to create a new GUID and copy it to the clipboard." & vbCrLf
    sCap = sCap & "Select a GUID then click 'Copy to...' to copy the GUID to the clipboard."
    
    Me.Label1.Caption = sCap
    Me.Command1(0).ToolTipText = "Generate a GUID and copy it to the clipboard."
    Me.Command1(1).ToolTipText = "Copy the selected GUID to the clipboard."
    
    mlMinWidth = Me.Width
    mlMinHeight = Me.Height
    
    TB_BuildWindowMenu Me, True, True, , True
    
    If TB_GetFormInformation(Me, iState, l, t, w, h) Then
        Me.Move l, t, w, h
    Else
        TB_CenterForm32 Me
    End If
End Sub

Private Sub Form_Resize()
    Dim l, t, w, h
    ' Commands
        
    If Me.WindowState <> vbMinimized Then
        If Me.WindowState = vbNormal Then
            If Me.Width < mlMinWidth Then
                Me.Width = mlMinWidth
            End If
            If Me.Height < mlMinHeight Then
                Me.Height = mlMinHeight
            End If
        End If
        With Command1(0)
            w = ScaleWidth * 0.48
            t = ScaleHeight - .Height * 1.2
            l = ((ScaleWidth \ 2) - w) \ 2
            .Move l, t, w
        End With
        With Command1(1)
            l = ScaleWidth - w - l
            .Move l, t, w
        End With
        ' Label1
        With Label1
            l = (ScaleWidth - .Width) \ 2
            .Left = l
        End With
        ' List1
        With List1
            w = ScaleWidth * 0.96
            l = (ScaleWidth - w) \ 2
            t = .Top
            h = ScaleHeight - t - Command1(0).Height - 300
            If h < 100 Then
                h = 100
            End If
            .Move l, t, w, h
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TB_SaveFormInformation Me
End Sub

Private Sub DoTestGUID()
    Dim sTemp As String
    Dim o As New cCreateGUID
    sTemp = Trim$(Me.List1.Text)
    If o.IsGUIDValid(sTemp) Then
        MsgBox "GUID '" & sTemp & "' is valid.", vbInformation, "OK - " & sTemp
    Else
        MsgBox "GUID '" & sTemp & "' is not valid.", vbExclamation, "Invalid - " & sTemp
    End If
End Sub

Private Sub DoLengthGUID()
    Dim sTemp As String
    Dim o As New cCreateGUID
    sTemp = Trim$(Me.List1.Text)
    MsgBox sTemp & " length is: " & o.LenGUID(sTemp), vbInformation, "Length"
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0  'test Selected GUID
            DoTestGUID
        Case 1  'len Selected GUID
            DoLengthGUID
        Case 2  'with prefix
            DoCreateWithPrefix
        Case 3
            Unload Me
    End Select
End Sub

Private Sub DoCreateWithPrefix()
' 2005/12/18 Sub added by Larry Rebich using the DELL8500 while in Indio, CA
    Dim sPrefix As String
    Dim o As New cCreateGUID
    Dim sGUID As String
    
    sPrefix = InputBox("Prefix?", , "MyGuid:")
    If sPrefix <> "" Then
        sGUID = o.CreateGUIDWithPrefix(sPrefix)
        MsgBox "GUID with prefix: '" & sGUID & "'", vbInformation, "GUID With Prefix"
    End If
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            DoAbout
    End Select
End Sub

Private Sub DoAbout()
    Dim sMsg As String
    sMsg = App.EXEName & vbCr & App.LegalCopyright
    MsgBox sMsg, vbInformation, " " & App.Title
End Sub

Private Sub mnuWindow_Click()
    TB_Window Me
End Sub

Private Sub mnuWindowItem_Click(Index As Integer)
    TB_WindowItem Me, Index
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "G.Epelsa. Alimerka CFG"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14460
   Begin VB.CommandButton CmdEnvMail 
      Caption         =   "Envío Correo"
      Height          =   255
      Left            =   12600
      TabIndex        =   21
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CmdMailLog 
      Caption         =   "Datos Mail y Log."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton CmdKlog 
      Caption         =   "Eliminar Fichero Incidencias."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   18
      Top             =   5320
      Width           =   2775
   End
   Begin VB.ListBox ListIncidencias 
      Height          =   2400
      Left            =   7080
      TabIndex        =   17
      Top             =   5760
      Width           =   7335
   End
   Begin VB.CheckBox ChkSend 
      Caption         =   "ENVIAR CORREO"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CmdSMail 
      Caption         =   "Salvar datos Correo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   5320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   5760
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   6480
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   6840
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   7200
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   7560
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   7920
      Width           =   4455
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4905
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8652
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12648447
      BackColor       =   8388608
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   6960
      X2              =   6960
      Y1              =   5400
      Y2              =   8160
   End
   Begin VB.Label LblInci 
      Caption         =   "Registro incidencias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   19
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14400
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label1 
      Caption         =   "Servidor SMTP:"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre Remitente:"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Correo Electrónico Remitente:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre Receptor:"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Correo Electrónico Receptor:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Asunto:"
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Mensaje:"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Datos Correo Electrónico Envío Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   4215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
Private bSendFailed     As Boolean

Private Sub ChkSend_Click()
Dim Arch As Integer
    If ChkSend.Value = vbChecked Then
        Arch = FreeFile()
        Open App.Path & "\lsendmail" For Output As #Arch
        Close #Arch
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        CmdEnvMail.Enabled = True
    Else
        If Dir(App.Path & "\lsendmail") <> "" Then
            Kill App.Path & "\lsendmail"
        End If
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text7.Enabled = False
        CmdEnvMail.Enabled = False
    End If
End Sub

Private Sub CmdMailLog_Click()
    If Me.Height <> 8640 Then
        Me.Height = 8640
        CmdMailLog.Caption = "^Datos Mail y Log."
    Else
        Me.Height = CmdSMail.Top + CmdSMail.Height 'Me.Height - CmdMailLog.Top + CmdMailLog.Height
        CmdMailLog.Caption = "_Datos Mail y Log."
    End If
End Sub

Private Sub CmdSMail_Click()
Dim Arch As Integer
    Arch = FreeFile()
    Open App.Path & "\alarmail.txt" For Output As #Arch
    Print #Arch, "SMTPSERVER" & ",", Trim(Text1.Text)
    Print #Arch, "SENDERNAME" & ",", Trim(Text2.Text)
    Print #Arch, "SENDEREMAIL" & ",", Trim(Text3.Text)
    Print #Arch, "RECIPIENTNAME" & ",", Trim(Text4.Text)
    Print #Arch, "RECIPIENTEMAIL" & ",", Trim(Text5.Text)
    Print #Arch, "SUBJECT" & ",", Trim(Text6.Text)
    Print #Arch, "MESSAGE" & ",", Trim(Text7.Text)
    Print #Arch, "ATTACHMENT" & ",", sAttachment
    Close #Arch
End Sub

Private Sub CmdEnvMail_Click()
    'If lSendMail And Dir(App.Path & "\" & sFichMail) <> "" And ((Alarmas_CAJAS1(1) = True And Not (CAJAS1.Exc)) Or (Alarmas_CAJAS2(1) = True And Not (CAJAS2.Exc))) Then
    '    lSendMail = False
        Set poSendMail = New clsSendMail
        With poSendMail
    
            ' **************************************************************************
            ' Set the basic properties common to all messages to be sent
            ' **************************************************************************
            .SMTPHost = sSmtpServer                  ' Required the fist time, optional thereafter
            .From = sSenderMail                      ' Required the fist time, optional thereafter
            .FromDisplayName = sSenderName           ' Optional, saved after first use
            .Message = sMessage                      ' Optional
            .Attachment = Trim(App.Path & "\" & "informe_alarmas.txt") ' Optional, separate multiple entries with delimiter character
    
            ' **************************************************************************
            ' Send the mail in a loop. In a real app you would need to load a new
            ' recipient from a file or database each pass through the loop.
            ' **************************************************************************
            'send method only (normal button)
            .Recipient = sRecipientEmail
            .RecipientDisplayName = sRecipientName
            .Subject = sSubject
            .Send
            
            ' connect, send, & disconnect methods (bulk send button)
            'If .Connect Then
            '    .Recipient = sRecipientName
            '    .RecipientDisplayName = sRecipientName
            '    .Subject = sSubject
            '    .Send
            '    .Disconnect
            'End If

    
        End With
        Set poSendMail = Nothing
    'Else
    '    lSendMail = False
    'End If

End Sub

Private Sub Form_Load()
Dim Arch As Integer
Dim Buf1, Buf2 As String
    If Dir(App.Path & "\alarmail.txt") <> "" Then
        Arch = FreeFile()
        Open App.Path & "\alarmail.txt" For Input As #Arch
        Do While Not EOF(Arch)
            Input #Arch, Buf1, Buf2
            Select Case Trim(UCase(Buf1))
            Case "SMTPSERVER"
                sSmtpServer = Trim(Buf2)
            Case "SENDERNAME"
                sSenderName = Trim(Buf2)
            Case "SENDEREMAIL"
                sSenderMail = Trim(Buf2)
            Case "RECIPIENTNAME"
                sRecipientName = Trim(Buf2)
            Case "RECIPIENTEMAIL"
                sRecipientEmail = Trim(Buf2)
            Case "SUBJECT"
                sSubject = Trim(Buf2)
            Case "MESSAGE"
                sMessage = Trim(Buf2)
            Case "ATTACHMENT"
                sAttachment = Trim(Buf2)
            End Select
        Loop
        Close #Arch
    Else
        sSmtpServer = "smtp.alimerka.es"
        sSenderName = "PC BALANZAS EPELSA"
        sSenderMail = "pcbalanzas@alimerka.es"
        sRecipientName = "Servicio Atención Balanzas"
        sRecipientEmail = "ccampos@grupoepelsa.com"
        sSubject = "Estado General Alarmas Balanzas."
        sMessage = "Se adjunta fichero Informe_Alarmas.txt"
        sAttachment = "informe_alarmas.txt"
    End If

    Text1.Text = sSmtpServer
    Text2.Text = sSenderName
    Text3.Text = sSenderMail
    Text4.Text = sRecipientName
    Text5.Text = sRecipientEmail
    Text6.Text = sSubject
    Text7.Text = sMessage

    If Dir(App.Path & "\lsendmail") = "" Then
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text7.Enabled = False
        CmdEnvMail.Enabled = False
    Else
        ChkSend.Value = vbChecked
    End If
    CmdMailLog.Caption = "_Datos Mail y Log."
    Me.Height = CmdSMail.Top + CmdSMail.Height 'Me.Height - CmdMailLog.Top + CmdMailLog.Height
End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

'Private Sub poSendMail_Progress(lPercentCompete As Long)

'    ' vbSendMail 'Progress Event'
'
'    lblProgress = lPercentCompete & "% complete"'
'
'End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event'
    Dim nF As Integer
    nF = FreeFile()
    Open App.Path & "\sendmail.nok" For Output As #nF
    Print #nF, Now(); "-->" & Explanation
    Close #nF
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    Dim nF As Integer
    nF = FreeFile()
    Open App.Path & "\sendmail.ok" For Output As #nF
    Print #nF, Now(); "-->" & "Enviado mail a: " & sRecipientEmail & " .Correctamente."
    Close #nF

End Sub

Private Sub DatosLista()
    Lista.ListItems.Clear
    Lista.ListItems.Add , "L00", "Sistemas"
    Lista.ListItems.Add , "L01", "Mostradores"
    Lista.ListItems.Add , "L02", "SC10"
    If ChkSC10.Value = vbChecked Or ChkGamaAlta.Value = vbChecked Then
        Lista.ListItems.Add , "L03", CargaCadena(602)
    End If
    Lista.ListItems.Add , "L04", CargaCadena(601)
    If ChkEcoPrint.Value = vbChecked Then
        Lista.ListItems.Add , "L05", "IV4/EcoPrint"
    End If
    If Not HayMulti Then Lista.ListItems.Add , "L06", CargaCadena(603)
    Lista.ListItems.Add , "L07", CargaCadena(604)
    Lista.ListItems.Add , "L08", CargaCadena(605)
    If Not HayMulti Then Lista.ListItems.Add , "L09", "Master/Slave"
    Lista.ListItems.Add , "L10", CargaCadena(1052)
    If ChkModem.Value = vbChecked Then
        Lista.ListItems.Add , "L11", "Modem"
    End If
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim MyVal As Integer
    Dim Bucle As Integer
    MyVal = Val(Mid(Item.Key, 2))

    For Bucle = 0 To 11
        
        If Bucle <> MyVal Then
            Tabla.TabVisible(Bucle) = False
        Else
            Tabla.TabVisible(Bucle) = True
            Tabla.Tab = Bucle
        End If
    Next Bucle
End Sub


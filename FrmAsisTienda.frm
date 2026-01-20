VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAsisTienda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Modem Tienda"
   ClientHeight    =   5130
   ClientLeft      =   2820
   ClientTop       =   2070
   ClientWidth     =   5220
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   2160
      Top             =   4440
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.ComboBox Combo3 
      Height          =   288
      ItemData        =   "FrmAsisTienda.frx":0000
      Left            =   2640
      List            =   "FrmAsisTienda.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   4932
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   4932
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4932
   End
   Begin VB.Timer TmrMyErr 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   4440
   End
   Begin MSCommLib.MSComm Serial 
      Left            =   2640
      Top             =   4440
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1452
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "FrmAsisTienda.frx":0004
      Left            =   3600
      List            =   "FrmAsisTienda.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   1452
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "FrmAsisTienda.frx":0008
      Left            =   3600
      List            =   "FrmAsisTienda.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "Si tiene necesidades especiales, indique aquí la cadena de configuración del modem  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4932
   End
   Begin VB.Label Label3 
      Caption         =   "Velocidad de las comunicaciones :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3492
   End
   Begin VB.Label Label2 
      Caption         =   "Puerto en el que ha conectado el modem de la Tienda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3372
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   $"FrmAsisTienda.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5172
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir Configuración existente"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar Esta Configuración"
      End
      Begin VB.Menu mnuRestaurar 
         Caption         =   "Restaurar Configuración por Defecto"
      End
   End
End
Attribute VB_Name = "FrmAsisTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sSeHaMo As Boolean
Private MyErr As Boolean
Private Const ModCad1 = "ATQ1V0ES0=2B0Y0&C0&S0&D0&Y0&A0&H0&I0&K0&M0"
Private Const ModCad2 = "AT&R1&G0&W0"
Private Const ModCad3 = ""
Private Const ModCad4 = "ATQ1V0ES0=2B0Y0&C0&S0&D0&Y0&A0&H0&I0&K0&M0"
Private Const ModCad5 = "AT&R1&G0&W0"
Private Const ModCad6 = ""

Private Sub Combo2_Click()
    If Not sSeHaMo Then
        Select Case Combo2.ListIndex
            Case 0
                Text1(0).TexT = ModCad1 & "&N2"
            Case 1
                Text1(0).TexT = ModCad1 & "&N3"
            Case 2
                Text1(0).TexT = ModCad1 & "&N6"
            Case 3
                Text1(0).TexT = ModCad1 & "&N10"
        End Select
    End If
    sSeHaMo = False
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        Me.Enabled = False
        If Not sSeHaMo Then
            Text1(0).TexT = Trim(Text1(0).TexT)
            Text1(1).TexT = Trim(Text1(1).TexT)
            Text1(2).TexT = Trim(Text1(2).TexT)
        End If
        Config_Modem
        Me.Enabled = True
        Command1(0).Enabled = True
        Command1(1).Enabled = True
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = Form2.Icon
    CambiarIdioma
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Text1(0).TexT = ModCad1 & "&N2"
    Text1(1).TexT = ModCad2
    Text1(2).TexT = ModCad3
    sSeHaMo = False
End Sub
Private Sub CambiarIdioma()
    With Me
        .Caption = CargaCadena(926)
        .mnuArchivo.Caption = CargaCadena(1)
        .mnuAbrir.Caption = CargaCadena(1062)
        .mnuGuardar.Caption = CargaCadena(1063)
        .mnuRestaurar.Caption = CargaCadena(1064)
        .Label1.Caption = CargaCadena(1058)
        .Label2.Caption = CargaCadena(1059)
        .Label3.Caption = CargaCadena(1060)
        .Label5.Caption = CargaCadena(55)
        .Label4.Caption = CargaCadena(1061)
        .Command1(0).Caption = CargaCadena(287)
        .Command1(1).Caption = CargaCadena(7)
        .Combo1.AddItem "COM1"
        .Combo1.AddItem "COM2"
        .Combo1.AddItem "COM3"
        .Combo1.AddItem "COM4"
        .Combo2.AddItem "1200  " & db.CargaCadena(261)
        .Combo2.AddItem "2400  " & db.CargaCadena(261)
        .Combo2.AddItem "9600  " & db.CargaCadena(261)
        .Combo2.AddItem "19200  " & db.CargaCadena(261)
        .Combo3.AddItem db.CargaCadena(1255)
        .Combo3.AddItem db.CargaCadena(1256)
    End With
End Sub
Private Sub Config_Modem()
    Dim resp As String
    Dim MiCadena As String
    Dim Bucle As Integer
    If Serial.PortOpen = True Then Serial.PortOpen = False
    On Error Resume Next
    Serial.CommPort = Combo1.ListIndex + 1
    Serial.InputMode = comInputModeBinary
    If Err.Number <> 0 Then
        MsgBox CargaCadena(965), vbExclamation
    Else
        On Error Resume Next
        Serial.PortOpen = True
        If Err.Number <> 0 Then
            MsgBox CargaCadena(965), vbExclamation
        Else
            resp = MyComunica("ATQ0V0")
            If resp = "0" Then
                For Bucle = 0 To 2
                    If Trim(Text1(Bucle).TexT) <> "" Then
                        resp = MyComunica(Text1(Bucle).TexT)
                    End If
                Next Bucle
                If resp = "0" Then
                    MsgBox CargaCadena(966), vbInformation
                Else
                    If resp = "" Then
                        MsgBox CargaCadena(966), vbInformation
                    Else
                        MsgBox CargaCadena(967) & " " & resp, vbInformation
                    End If
                End If
            Else
                MsgBox CargaCadena(968), vbExclamation
            End If
        End If
    End If
    On Error Resume Next
    If Serial.PortOpen = True Then Serial.PortOpen = False
    On Error GoTo 0
End Sub
Private Function MyComunica(SendMe As String) As String
    Dim MyBuf As String
    Dim RetryCR As Integer
    For RetryCR = 0 To 1
        MyBuf = ""
        If RetryCR = 0 Then
            Serial.Output = SendMe & Chr(13)
        Else
            Serial.Output = SendMe & Chr(13) & Chr(10)
        End If
        MyErr = False
        TmrMyErr.Interval = 3000
        TmrMyErr.Enabled = True
        Do Until InStr(1, MyBuf, Chr(13)) > 0 Or MyErr = True
            If Serial.InBufferCount > 0 Then
                MyBuf = MyBuf & StrConv(Serial.Input, vbUnicode)
            End If
            Do_Events
        Loop
        If MyBuf <> "" Then Exit For
    Next RetryCR
    If Len(MyBuf) > 1 Then MyBuf = Mid(MyBuf, 1, Len(MyBuf) - 1)
    TmrMyErr.Enabled = False
    If MyErr Then MyBuf = ""
    MyComunica = MyBuf
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Serial.PortOpen = True Then Serial.PortOpen = False
    On Error GoTo 0
End Sub

Private Sub mnuAbrir_Click()
    Dim Arch As Integer
    Dim Buf(3) As String
    Dialogo.DefaultExt = ".txt"
    Dialogo.Filter = "*.txt"
    Dialogo.InitDir = App.Path & "\etc"
    Dialogo.ShowOpen
    If Dialogo.FileName <> "" Then
        Arch = FreeFile()
        On Error Resume Next
        Open Dialogo.FileName For Input As #Arch
        If Err.Number = 0 Then
            If Not EOF(Arch) Then Line Input #Arch, Buf(0)
            If Not EOF(Arch) Then Line Input #Arch, Buf(1)
            If Not EOF(Arch) Then Line Input #Arch, Buf(2)
            Text1(0).TexT = Buf(0)
            Text1(1).TexT = Buf(1)
            Text1(2).TexT = Buf(2)
            Close #Arch
            sSeHaMo = True
        Else
            MsgBox CargaCadena(969), vbCritical
        End If
    End If
End Sub

Private Sub mnuGuardar_Click()
    Dim Arch As Integer
    Dialogo.DefaultExt = ".txt"
    Dialogo.Filter = "*.txt"
    Dialogo.InitDir = App.Path & "\etc"
    Dialogo.ShowSave
    If Dialogo.FileName <> "" Then
        Arch = FreeFile()
        On Error Resume Next
        Open Dialogo.FileName For Output As #Arch
        If Err.Number = 0 Then
            Print #Arch, Text1(0).TexT
            Print #Arch, Text1(1).TexT
            Print #Arch, Text1(2).TexT
            Close #Arch
            sSeHaMo = True
            MsgBox CargaCadena(800), vbInformation
        Else
            MsgBox CargaCadena(969), vbCritical
        End If
    End If
End Sub

Private Sub mnuRestaurar_Click()
    Text1(1).TexT = ModCad2
    Text1(2).TexT = ModCad3
    sSeHaMo = False
    Combo2_Click
    sSeHaMo = False
End Sub

Private Sub Text1_Change(Index As Integer)
    sSeHaMo = True
End Sub

Private Sub TmrMyErr_Timer()
    MyErr = True
End Sub

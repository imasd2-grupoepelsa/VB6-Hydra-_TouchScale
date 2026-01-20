VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   6210
   ClientLeft      =   5715
   ClientTop       =   2190
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7260
   Begin VB.ListBox List3 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   4812
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2880
      Width           =   6972
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   6972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restaurar Tiendas Seleccionadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   4800
      Width           =   1812
   End
   Begin VB.Image Command1 
      Height          =   612
      Index           =   2
      Left            =   5760
      Picture         =   "FrmBackup.frx":0000
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   612
   End
   Begin VB.Image Command1 
      Height          =   612
      Index           =   1
      Left            =   3120
      Picture         =   "FrmBackup.frx":08CA
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   612
   End
   Begin VB.Image Command1 
      Height          =   612
      Index           =   0
      Left            =   720
      Picture         =   "FrmBackup.frx":1194
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   5760
      Width           =   852
   End
   Begin VB.Image CmdSalir 
      Height          =   384
      Left            =   5880
      Picture         =   "FrmBackup.frx":1A5E
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   384
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de la Copia Seleccionada :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copias Disponibles :"
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Realizar Copia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1812
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Realizar_Backup
            refresca_copias
        Case 1
            Seleccionar_Todas
        Case 2
            Restaurar_Copias
    End Select
    
End Sub
Private Sub Seleccionar_Todas()
    Dim Bucle As Integer
    If List2.ListCount > 0 Then
        For Bucle = 0 To List2.ListCount - 1
            List2.Selected(Bucle) = True
        Next Bucle
    End If
End Sub
Private Sub CambiarIdioma()
    Label1(0).Caption = CargaCadena(1155)
    Label1(1).Caption = CargaCadena(1156)
    Label2(0).Caption = CargaCadena(7)
    Label2(1).Caption = CargaCadena(1157)
    Label2(2).Caption = CargaCadena(1158)
    Label2(3).Caption = CargaCadena(1159)
End Sub

Private Sub Form_Activate()
    If UCase(Trim(Command)) = "/BACKUP" Then
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        Me.Refresh
        Realizar_Backup
        'Descarga = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'de momento
    CambiarIdioma
    List3.Clear
    Me.Icon = Form2.Icon
    List2.Clear
    DrawGradient Me, 12, 81, 179, 0, 50, 0, Me.ScaleHeight, 1, True, 1, 1, 1
    '
    On Error GoTo error1
    If Dir(Backup_path, vbDirectory) = "" Then
        MkDir Backup_path
    End If
    'If UCase(Trim(Command)) = "/BACKUP" Then
    '    Realizar_Backup
    '    'Descarga = True
    '    Unload Me
    'End If
    On Error GoTo 0
    '
    refresca_copias
    Exit Sub
error1:
    MsgBox CargaCadena(847), vbCritical
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
End Sub
Public Sub Realizar_Backup()
    Dim MyResul As Long
    Dim MyProc As Long
    Dim Bucle As Integer
    Dim B2 As Integer
    Dim MyArch As String
    Dim RutaZIP As String
    Dim Salida As String
    Dim NombreCorto As String
    Dim i As Long
    Dim sBuf As String * 260
    ''
    Me.Enabled = False
    ChDir App.Path
    'If Dir(App.Path & "\tiendas", vbDirectory) <> "" Then
        FrmTools.Dir1 = App.Path
        FrmTools.Dir1.Refresh
        For Bucle = 0 To FrmTools.Dir1.ListCount - 1
            If (IsNumeric(Right(FrmTools.Dir1.List(Bucle), 2)) Or IsNumeric(Right(FrmTools.Dir1.List(Bucle), 3))) Then
                If (Dir(FrmTools.Dir1.List(Bucle) & "\hydratouch.ini") <> "") And (Dir(FrmTools.Dir1.List(Bucle) & "\dbasetouch.mdb") <> "") Then
                    '
                    If IsNumeric(Right(FrmTools.Dir1.List(Bucle), 2)) Then List3.AddItem CargaCadena(1160) & " " & (Right(FrmTools.Dir1.List(Bucle), 2))
                    If IsNumeric(Right(FrmTools.Dir1.List(Bucle), 3)) Then List3.AddItem CargaCadena(1160) & " " & (Right(FrmTools.Dir1.List(Bucle), 3))
                    Do_Events
                    For B2 = 1 To 3
                        Select Case B2
                            Case 1
                                MyArch = "\hydratouch.ini"
                            Case 2
                                MyArch = "\dbasetouch.mdb"
                            Case 3
                                MyArch = "\tienda.cfg"
                        End Select
                        RutaZIP = Mid(FrmTools.Dir1.List(Bucle), Len(App.Path) + 2)
                        '
                        i = GetShortPathName(RutaZIP, sBuf, Len(sBuf))
                        If i = 0 Then Exit Sub
                        NombreCorto = Left$(sBuf, i)
                        RutaZIP = NombreCorto
                        '
                        i = GetShortPathName(Backup_path, sBuf, Len(sBuf))
                        If i = 0 Then Exit Sub
                        NombreCorto = Left$(sBuf, i)
                        
                        '
                        Salida = Ejecuta("mizip zip.exe " & NombreCorto & "\backuptouch_" & Format(Now, "yymmdd") & ".zip -r " & RutaZIP & MyArch)
                        
                        If B2 = 3 Then
                            If InStr(1, UCase(Salida), "DEFLATED") <> 0 Or InStr(1, UCase(Salida), "STORED") <> 0 Then
                                List3.AddItem "OK"
                            Else
                                List3.AddItem "ERROR"
                            End If
                            Do_Events
                        End If
                    Next B2
                    '
                End If
            End If
        Next Bucle
    'End If
    '''
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Descarga = True
End Sub
Private Function Ejecuta(ByVal CmdLine As String) As String
    'Ejecuta el comando indicado, espera a que termine
    'y redirige la salida hacia VB
    
    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    ChDir App.Path
    '=== Longitud de la cadena, en teoría 64 KB,
    '   pero no en la práctica
    'mybuff = String(64 * 1024, Chr$(65))
    '
    mybuff = String(10 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        Ejecuta = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SW_SHOWMINNOACTIVE
    
    ' Start the shelled application:
    
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If
    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    'En el original, sólo leian 100 caracteres
    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = Left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
    End If
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    
    Ejecuta = UCase(sReturnStr)
End Function
Private Sub refresca_copias()
    Dim Bucle As Integer
    Dim MiCad As String
    List1.Clear
    FrmTools.File1.Path = Backup_path
    FrmTools.File1.FileName = "backup*.zip"
    FrmTools.File1.Refresh
    For Bucle = 0 To FrmTools.File1.ListCount - 1
        If Len(FrmTools.File1.List(Bucle)) = 17 Then
            If IsNumeric(Mid(FrmTools.File1.List(Bucle), 8, 6)) Then
                MiCad = CargaCadena(434) & " : " & Mid(FrmTools.File1.List(Bucle), 12, 2) & "/" & Mid(FrmTools.File1.List(Bucle), 10, 2) & "/" & Mid(FrmTools.File1.List(Bucle), 8, 2)
                MiCad = MiCad & CargaCadena(552) & " : " & FrmTools.File1.List(Bucle)
                List1.AddItem MiCad
            End If
        End If
    Next Bucle
    
End Sub
Private Sub Lista_Copias(LaCopia As String)
    Dim MiCad As String
    Dim MiCad2 As String
    Dim Exito As Boolean
    Dim Bucle As Integer
    Dim V_Tienda(256) As Integer
    Dim NombreCorto As String
    Dim i As Long
    Dim sBuf As String * 260
    List2.Clear
    MiCad = Backup_path & "\" & LaCopia
    '
    '
    i = GetShortPathName(MiCad, sBuf, Len(sBuf))
    If i = 0 Then Exit Sub
    NombreCorto = Left$(sBuf, i)
                        
    
    '
    MiCad = Ejecuta("mizip unzip -l " & NombreCorto)
    If InStr(1, MiCad, "--------    ----   ----    ----" & vbCrLf) > 0 Then
        MiCad = Mid(MiCad, 33 + InStr(1, MiCad, "--------    ----   ----    ----" & vbCrLf))
        If InStr(1, MiCad, "--------                   -------" & vbCrLf) > 0 Then
            MiCad = Left(MiCad, InStr(1, MiCad, " --------                   -------" & vbCrLf))
            Do While Trim(MiCad) <> ""
                MiCad2 = Left(MiCad, InStr(1, MiCad, vbCrLf))
                MiCad = Mid(MiCad, Len(MiCad2) + 2)
                For Bucle = 1 To 255
                    If InStr(1, UCase(MiCad2), "T" & Format(Bucle, "00") & "/hydratouch.ini") > 0 Or InStr(1, UCase(MiCad2), "T" & Format(Bucle, "00") & "/dbasetouch.mdb") > 0 Then
                        V_Tienda(Bucle) = V_Tienda(Bucle) + 1
                        Exit For
                    End If
                Next Bucle
            Loop
            For Bucle = 1 To 255
                If V_Tienda(Bucle) > 1 Then
                    List2.AddItem CargaCadena(46) & " " & Bucle
                End If
            Next Bucle
        End If
    End If
End Sub
Private Sub Restaurar_Copias()
    Dim Bucle As Integer
    Dim MiCopia As String
    Dim Encontrado As Boolean
    Dim MyOK As Boolean
    Dim MiTienda As String
    Dim Resul As String
    Dim NombreCorto As String
    Dim i As Long
    Dim sBuf As String * 260
    If List1.ListCount > 0 Then
        For Bucle = 0 To List1.ListCount - 1
            If List1.Selected(Bucle) Then
                MiCopia = Trim(Right(List1.List(Bucle), 17))
                Exit For
            End If
        Next Bucle
    End If
    If MiCopia = "" Then
        MsgBox CargaCadena(1161)
    Else
        Encontrado = False
        If List2.ListCount > 0 Then
            For Bucle = 0 To List2.ListCount - 1
                If List2.Selected(Bucle) Then
                    Encontrado = True
                    Exit For
                End If
            Next Bucle
        End If
        If Encontrado Then
            If MsgBox(CargaCadena(1162), vbExclamation + vbYesNo + vbDefaultButton2) Then
                ChDir App.Path
            
                For Bucle = 0 To List2.ListCount - 1
                    If List2.Selected(Bucle) Then
                        MyOK = True
                        List3.AddItem CargaCadena(1164) & " " & List2.List(Bucle)
                        '
                        i = GetShortPathName(Backup_path & "\" & MiCopia, sBuf, Len(sBuf))
                        If i = 0 Then Exit Sub
                        NombreCorto = Left$(sBuf, i)
                        '
                        MiTienda = "t" & Format(Val(Mid(List2.List(Bucle), InStr(1, List2.List(Bucle), " "))), "00") & "\"
                        Resul = ""
                        Resul = Ejecuta("mizip unzip -oC " & NombreCorto & " " & MiTienda & "dbasetouch.mdb")
                        
                        If InStr(1, UCase(Resul), "INFLATING") = 0 And InStr(1, UCase(Resul), "EXTRACTING") = 0 Then MyOK = False
                        
                        If MyOK Then
                            Resul = ""
                            Resul = Ejecuta("mizip unzip -oC " & NombreCorto & " " & MiTienda & "hydratouch.ini")
                        End If
                        If InStr(1, UCase(Resul), "INFLATING") = 0 And InStr(1, UCase(Resul), "EXTRACTING") = 0 Then MyOK = False
                        
                        If MyOK Then
                            Resul = ""
                            Resul = Ejecuta("mizip unzip -oC " & NombreCorto & " " & MiTienda & "tienda.cfg")
                        End If
                        If InStr(1, UCase(Resul), "INFLATING") = 0 And InStr(1, UCase(Resul), "EXTRACTING") = 0 Then MyOK = False
                    
                        If MyOK Then
                            List3.AddItem "OK"
                        Else
                            List3.AddItem "ERROR"
                        End If
                    End If
                Next Bucle
            End If
        Else
            MsgBox CargaCadena(1163)
        End If
    End If
End Sub

Private Sub List1_ItemCheck(Item As Integer)
    Dim Bucle As Integer
    If List1.ListCount > 0 Then
        For Bucle = 0 To List1.ListCount - 1
            If Bucle <> Item Then List1.Selected(Bucle) = False
        Next Bucle
        If List1.Selected(Item) = True Then
            Lista_Copias FrmTools.File1.List(Item)
        Else
            List2.Clear
        End If
    End If
End Sub

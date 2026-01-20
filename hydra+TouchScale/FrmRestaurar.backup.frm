VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmResturarBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S"
   ClientHeight    =   4980
   ClientLeft      =   4845
   ClientTop       =   2190
   ClientWidth     =   5835
   Icon            =   "FrmRestaurar.backup.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "E"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "B"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Width           =   1572
   End
   Begin VB.CommandButton CmdRestaurar 
      Caption         =   "R"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Width           =   1572
   End
   Begin VB.CommandButton CmhHacer 
      Caption         =   "A"
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
      TabIndex        =   2
      Top             =   3960
      Width           =   1572
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3132
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5652
      _ExtentX        =   9975
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   3492
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "R"
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
      TabIndex        =   6
      Top             =   240
      Width           =   5652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "A"
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
      TabIndex        =   5
      Top             =   3600
      Width           =   5652
   End
End
Attribute VB_Name = "FrmResturarBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MIERROR As Boolean
Private Sub CambiarIdioma()
    Label2.Caption = CargaCadena(730)
    Label1.Caption = CargaCadena(731)
    CmhHacer.Caption = CargaCadena(729)
    CmdRestaurar.Caption = CargaCadena(111)
    CmdEliminar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)
    FrmResturarBackup.Caption = CargaCadena(32)
End Sub
Private Sub cmdEliminar_Click()
    Dim Resp
    CmdRestaurar.Enabled = False
    CmdEliminar.Enabled = False
    Resp = MsgBox(CargaCadena(732), vbExclamation + vbYesNo)
    If Resp = vbYes Then
        On Error GoTo errorred
        Kill pathred & "\" & Lista.SelectedItem.SubItems(1)
        On Error GoTo 0
    End If
    refresca_lista
    Exit Sub
errorred:
    On Error GoTo 0
    MsgBox CargaCadena(733)
    
End Sub

Private Sub CmdRestaurar_Click()
    Dim Resp
    
    CmdRestaurar.Enabled = False
    CmdEliminar.Enabled = False
    Resp = MsgBox(CargaCadena(734), vbExclamation + vbYesNo)
    If Resp = vbYes Then
        On Error GoTo errorred
        FileCopy pathred & "\" & Lista.SelectedItem.SubItems(1), Miruta & "\dbasetouch.mdb"
        On Error GoTo 0
        MsgBox CargaCadena(735)
    End If
    
    Exit Sub
errorred:
    On Error GoTo 0
    MsgBox CargaCadena(733)
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmhHacer_Click()
    
    CmdRestaurar.Enabled = False
    CmdEliminar.Enabled = False
    Realizar_Backup True
    refresca_lista
    
End Sub

Private Sub Form_Load()
  CambiarIdioma
  Me.Icon = Form2.Icon
  Image1.Picture = Form2.Imagen.Picture
  MIERROR = False
  refresca_lista
End Sub
'''''''
'''''''
Public Sub Borrar_Copias(Antigua As Integer)
    Dim MiFecha As String
    Dim bucle As Integer
    MiFecha = Format((Now - Antigua), "yymmdd")
    On Error GoTo errorred
    File1.Path = pathred
    File1.Pattern = "dbase_??????_????.mdb"
    File1.Refresh
    For bucle = 0 To File1.ListCount - 1
        If Val(Mid(File1.List(bucle), 7, 6)) < Val(MiFecha) Then
            Kill pathred & "\" & File1.List(bucle)
        End If
    Next bucle
    On Error GoTo 0
    Exit Sub
errorred:
    If DebugActivo Then
        If id = 0 Then
            CadenadeLog "Error de Acceso a Ruta de Backup. No se pueden borrar copias antiguas"
        Else
            CadenadeLog "Error Accessing Backup Path. Can't erase old Backups"
        End If
    End If
End Sub
Public Function Realizar_Backup(MostrarMsg As Boolean) As Integer
    Dim nombrecopia As String
    Dim extension As String
    Dim Contador As Long
    Dim Resp As Integer
    
    'C2F 29/07/02-->intento no dejar colgada copia seguridad.
    Dim MiProceso As Long
    Dim MiResultado As Long
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    lCogeTiquet = False
    Resp = 0
    If FrmResturarBackup.Visible = False Then
        frmMonitor.MostrarDato CargaCadena(737)
        frmMonitor.Command1.Enabled = False
   End If
   If MIERROR = True Then
        If frmMonitor.Visible = True Then
            Resp = 1
            frmMonitor.MostrarDato CargaCadena(267)  '"FALLO EN COPIA DE SEGURIDAD"
            frmMonitor.Label1.Caption = CargaCadena(267)  '"FALLO EN COPIA DE SEGURIDAD"
            HabilitarSalir
        End If
        lCogeTiquet = True
        Realizar_Backup = Resp
        Unload Me
        Exit Function
    End If
    CmdRestaurar.Enabled = False
    CmdEliminar.Enabled = False
    
    If Trim(pathred) = "" Then
        If MostrarMsg Then MsgBox CargaCadena(736)
        If FrmResturarBackup.Visible = False Then
            frmMonitor.MostrarDato CargaCadena(107) & " --> NOK"
            HabilitarSalir
        End If
        Resp = 1
        lCogeTiquet = True
        Realizar_Backup = Resp
        Exit Function
    End If
    Contador = 1
    nombrecopia = pathred & "\dbase_" & Format(Now(), "yymmdd")
    On Error GoTo fin
    extension = "_" & Format(Contador, "0000") & ".mdb"
    Do Until (Dir(nombrecopia & extension)) = "" And ConfirmarNombre(Contador)
        If Contador > 9999 Then
            If MostrarMsg Then MsgBox CargaCadena(267)  '"No se puede realizar la copia"
            If FrmResturarBackup.Visible = False Then
                frmMonitor.MostrarDato CargaCadena(107) & " --> NOK"
                HabilitarSalir
            End If
            Resp = 2
            lCogeTiquet = True
            Realizar_Backup = Resp
            Exit Function
        End If
        Contador = Contador + 1
        extension = "_" & Format(Contador, "0000") & ".mdb"
    Loop
    extension = "_" & Format(Contador, "0000") & ".mdb"
    nombrecopia = nombrecopia & extension
    'If Dir(MiruTa & "\mastercopy.exe") <> "" Then
    '    ChDir pathred
    '    Me.Enabled = False
    '    'C2F 29/08/02 --> intento solucionar cuelgue copia a CD...
    '    'se comenta lo que sigue hasta linea continua de comentario...
    '    'Shell (MiruTa & "\mastercopy.exe /C " & nombrecopia)
    '    'Sleep (10)
    '    'Do While Dir(MiruTa & "\etc\copiando") <> ""
    '    '    Do_Events
    '    'Loop
    '    ''''''''''''''''''''
    '    'se sustituye por...
    '    MiResultado = STILL_ACTIVE
    '    MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiruTa & "\mastercopy.exe /C " & nombrecopia, vbNormalFocus))
    '    If MiProceso <> 0 Then
    '       Do While MiResultado = STILL_ACTIVE
    '          GetExitCodeProcess MiProceso, MiResultado
    '          Do_Events
    '       Loop
    '    End If
    '    Me.Enabled = True
    '    ChDir MiruTa
    'Else
    'CadenadeLog "INTENTO COPIA Anterior:" & Miruta & "\dbasetouch.mdb" & " a " & nombrecopia
    If TiendaActual = 0 Then
        CadenadeLog "INTENTO COPIA Actual:" & Miruta & "\dbasetouch.mdb" & " a " & nombrecopia
        FileCopy Miruta & "\dbasetouch.mdb", nombrecopia
    Else
        If TiendaActual < 100 Then
            CadenadeLog "INTENTO COPIA Actual:" & App.Path & "t" & Format(TiendaActual, "00") & "\dbasetouch.mdb" & " a " & nombrecopia
            FileCopy App.Path & "\t" & Format(TiendaActual, "00") & "\dbasetouch.mdb", nombrecopia
        Else
            CadenadeLog "INTENTO COPIA Actual:" & App.Path & "t" & Format(TiendaActual, "000") & "\dbasetouch.mdb" & " a " & nombrecopia
            FileCopy App.Path & "\t" & Format(TiendaActual, "000") & "\dbasetouch.mdb", nombrecopia
        End If
    End If
    'FileCopy Miruta & "\dbasetouch.mdb", nombrecopia
    'End If
    On Error GoTo 0
    If FrmResturarBackup.Visible = False Then
        frmMonitor.MostrarDato CargaCadena(107) & " --> OK"
        HabilitarSalir
        frmMonitor.MostrarDato CargaCadena(816) & " " & nombrecopia
    Else
        If MostrarMsg Then MsgBox CargaCadena(816) & " " & nombrecopia
    End If
    lCogeTiquet = True
    Realizar_Backup = 0
    Exit Function
fin:
    On Error GoTo 0
    ChDir Miruta
    If MostrarMsg Then MsgBox CargaCadena(733)
    If FrmResturarBackup.Visible = False Then
        frmMonitor.MostrarDato CargaCadena(107) & " --> NOK"
        HabilitarSalir
    End If
    lCogeTiquet = True
    Realizar_Backup = 3
    Exit Function

End Function
Private Function ConfirmarNombre(ElNumero As Long) As Boolean
Dim Devuelve As Boolean
Dim Encontrado As Boolean
Dim bucle As Integer
Dim Buclelite As Long
Dim extension As String
Dim nombrecopia As String
Devuelve = True
    
        For bucle = ElNumero + 1 To 999
            extension = "_" & Format(bucle, "0000") & ".mdb"
            nombrecopia = pathred & "\dbase_" & Format(Now(), "yymmdd")
            Encontrado = False
            For Buclelite = 1 To File1.ListCount
                If pathred & "\" & File1.List(Buclelite) = nombrecopia & extension Then Encontrado = True
            Next Buclelite
            
                If Encontrado = True Then
                    Devuelve = False
                    Exit For
                End If
        Next bucle
            
           
            
            
        ConfirmarNombre = Devuelve
End Function
Private Sub refresca_lista()
    Dim bucle As Long
    Dim lacopia As String
    Dim CuentaBucle As Long
    MIERROR = False
    On Error GoTo errorred
    If Right(pathred, 1) <> "\" Then
        File1.Path = pathred & "\"
    Else
        File1.Path = pathred
    End If
    File1.Pattern = "*.mdb"
    File1.Refresh
    Lista.ListItems.Clear
    Lista.ColumnHeaders.Clear
     
    Lista.ColumnHeaders.Add 1, , CargaCadena(434)
    Lista.ColumnHeaders.Add 2, , CargaCadena(552), 4130
    CuentaBucle = 0
    For bucle = 0 To (File1.ListCount - 1)
        lacopia = File1.List(bucle)
        If Len(lacopia) = 21 Then
            If IsNumeric(Mid(lacopia, 7, 6)) Then
                Lista.ListItems.Add CuentaBucle + 1, , Mid(lacopia, 11, 2) & "/" & Mid(lacopia, 9, 2) & "/" & Mid(lacopia, 7, 2)
                Lista.ListItems(CuentaBucle + 1).SubItems(1) = lacopia
                CuentaBucle = CuentaBucle + 1
            End If
        End If
    Next bucle
    On Error GoTo 0
    ordenar_copias
    Exit Sub
errorred:
    On Error GoTo 0
    MsgBox CargaCadena(847)
    MIERROR = True
End Sub
Private Sub ordenar_copias()
    Dim Bucle1 As Long
    Dim Bucle2 As Long
    Dim Fecha1 As Double
    Dim Fecha2 As Double
    Dim Buffer As String
    
    For Bucle1 = 1 To (Lista.ListItems.Count - 1)
        Fecha1 = Val(Right(Lista.ListItems(Bucle1).TexT, 2) & Mid(Lista.ListItems(Bucle1).TexT, 4, 2) & left(Lista.ListItems(Bucle1).TexT, 2) & Mid(Lista.ListItems(Bucle1).SubItems(1), 14, 4))
        For Bucle2 = (Bucle1 + 1) To Lista.ListItems.Count
            Fecha2 = Fecha1 = Val(Right(Lista.ListItems(Bucle1).TexT, 2) & Mid(Lista.ListItems(Bucle1).TexT, 4, 2) & left(Lista.ListItems(Bucle1).TexT, 2) & Mid(Lista.ListItems(Bucle1).SubItems(1), 14, 4))
            If Fecha2 < Fecha1 Then
                Buffer = Lista.ListItems(Bucle1).TexT
                Lista.ListItems(Bucle1).TexT = Lista.ListItems(Bucle2).TexT
                Lista.ListItems(Bucle2).TexT = Buffer
                Buffer = Lista.ListItems(Bucle1).SubItems(1)
                Lista.ListItems(Bucle1).SubItems(1) = Lista.ListItems(Bucle2).SubItems(1)
                Lista.ListItems(Bucle2).SubItems(1) = Buffer
            End If
        Next Bucle2
    Next Bucle1
    Exit Sub

        
End Sub



Private Sub Lista_Click()
    
    CmdRestaurar.Enabled = True
    CmdEliminar.Enabled = True
End Sub

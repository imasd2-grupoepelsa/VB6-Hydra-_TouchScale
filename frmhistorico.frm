VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmhistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico"
   ClientHeight    =   2220
   ClientLeft      =   6990
   ClientTop       =   2190
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   372
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Aceptar"
      Default         =   -1  'True
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
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Label2"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Introduzca aquí la fecha en formato dd/mm/aa"
      Height          =   252
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3492
   End
End
Attribute VB_Name = "frmhistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdgo_Click()
    Dim antiguo As String
    Dim MiPath As String
    Dim MiResultado As Long
    Dim MiProceso As Long
    '
    ' Recoge la fecha del cuadro de texto
    ' y la convierte al formato del nombre del fichero
    '
    Me.Enabled = False
    Me.cmdgo.Enabled = False
    Me.cmdcancel.Enabled = False
    Me.Label2.Caption = "Trabajando..."
    Do_Events
    If Not IsNumeric(Right(TxtFecha.TexT, 2)) Or Not IsNumeric(left(TxtFecha.TexT, 2)) _
    Or Not IsNumeric(Mid(TxtFecha.TexT, 4, 2)) Then
        MsgBox CargaCadena(506), vbInformation
    End If
    If frmhistorico.Tag = "" Then
        frmhistorico.Visible = False
        Exit Sub
    End If
    If Dir(App.Path & "\gigante.txt") <> "" Then
        'histórico gigante
        Call GeneraGianT(left(TxtFecha.TexT, 2) & "/" & Mid(TxtFecha.TexT, 4, 2) & "/" & Right(TxtFecha.TexT, 2))
        'frmhistorico.Visible = False
        GoTo SigueTrasGian
        'Exit Sub
    End If
    
    Select Case frmhistorico.Tag
        Case 0
            antiguo = Miruta & "\" & Right(TxtFecha.TexT, 2) & Mid(TxtFecha.TexT, 4, 2) & left(TxtFecha.TexT, 2)
            '
            ' Comprueba si existe el fichero
            '
            If Dir(antiguo) = "" Then
                MsgBox CargaCadena(864)
                Me.Enabled = True
                cmdcancel.Enabled = True
                cmdcancel.Caption = "Salir"
                Me.Label2.Caption = "Finalizado"
                Exit Sub
            End If
            '
            ' Copia el fichero con el nombre totales.dat
            '
            On Error GoTo manejaerror
            FileCopy antiguo, PathOrden & "\" & "totales.dat"
            On Error GoTo 0
        Case 1
            If lAlba Then
                Alba_EnviaTiquetsPendientes True, Mid(TxtFecha.TexT, 4, 2) & "/" & left(TxtFecha.TexT, 2) & "/" & Right(TxtFecha.TexT, 2)
            Else
                If lBoka Then
                    FrmExportKuups.ExportarFichero_BOKA 0, "boka.txt", Mid(TxtFecha.TexT, 4, 2) & "/" & left(TxtFecha.TexT, 2) & "/" & Right(TxtFecha.TexT, 2), True, True
                    'FrmExportKuups.ExportarFichero_BOKA 0, "boka.txt", 0, True, True
                Else
                    '''
                    If lFornes Then
                        Call export_BL(True)
                    Else
                        ExportaTotalesDat False, Mid(TxtFecha.TexT, 4, 2) & "/" & left(TxtFecha.TexT, 2) & "/" & Right(TxtFecha.TexT, 2)
                    End If
                End If
            End If
            
   End Select
            '*******************************
            ' tras generar el archivo
            ' llama a msb100
            '*******************************
 
SigueTrasGian:

 Me.Enabled = True
 cmdcancel.Enabled = True
 cmdcancel.Caption = "Salir"
 Me.Label2.Caption = "Finalizado"
 Do_Events
If Me.Tag = 0 Then
If Dir(MiPath & "\xmsb100.exe") <> "" Or Dir(Miruta & "\fedeiv8.exe") <> "" Or Dir(Miruta & "\fede386.exe") <> "" Then
    ChDir MiPath
    If Dir(MiPath & "\xmsb100.exe") <> "" Then
        MiResultado = STILL_ACTIVE
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiPath & "\" & "xmsb100.exe", vbHide))
        Do While MiResultado = STILL_ACTIVE
            GetExitCodeProcess MiProceso, MiResultado
            Do_Events
        Loop
    End If
    If Dir(MiPath & "\fedeiv8.exe") <> "" Then
        MiResultado = STILL_ACTIVE
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiPath & "\" & "fedeiv8.exe", vbHide))
        Do While MiResultado = STILL_ACTIVE
            GetExitCodeProcess MiProceso, MiResultado
            Do_Events
        Loop
    End If
    If Dir(MiPath & "\fede386.exe") <> "" Then
        MiResultado = STILL_ACTIVE
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiPath & "\" & "fede386.exe", vbHide))
        Do While MiResultado = STILL_ACTIVE
            GetExitCodeProcess MiProceso, MiResultado
            Do_Events
        Loop
    End If
    
    If Dir(MiPath & "\msb100") = "" Then
        MsgBox (CargaCadena(861))
        lCogeTiquet = True
        Exit Sub
    End If
    
    lCogeTiquet = True
End If
End If
    'Unload Me
    If lAlba Then
        MsgBox "Generado TQGEN para la fecha indicada.", vbExclamation
        frmEpelsa.Enabled = True
        Unload Me
    End If
    
    Exit Sub
    '
    ' Maneja errores de archivo
    '
manejaerror:
        MsgBox (Error)
    
End Sub

Private Sub Form_Load()
        frmhistorico.TxtFecha = AdaptaFecha(Format(Date, "dd/mm/yy"))
        cmdgo.Caption = CargaCadena(287)
        cmdcancel.Caption = CargaCadena(288)
        Label1.Caption = CargaCadena(831)
        If lFornes Then
            TxtFecha.Enabled = False
            Label1.Caption = ""
        Else
            Label1.Caption = CargaCadena(831)
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEpelsa.Enabled = True
End Sub

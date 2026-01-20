VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmimportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación"
   ClientHeight    =   5670
   ClientLeft      =   5775
   ClientTop       =   510
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmimportar.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4950
   Begin VB.CommandButton CmdClientes 
      Caption         =   "Clientes.dat"
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1290
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   720
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   1320
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   7
      Left            =   3720
      TabIndex        =   15
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   6
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   5
      Left            =   1320
      TabIndex        =   13
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   2
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton CmdFichas 
      Caption         =   "Fichas.dat"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1092
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdcancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdimportar 
      Caption         =   "&Importar"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtfichero 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label LblProgreso 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "100%"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label LblAyuda 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elija un Fichero Predefinido, o escríbalo explícitamente, y pulse importar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4812
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmimportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub cmdcancelar_Click()
    Cancelar = True
    Success = 2
    If LogActivo Then
        If id = 0 Then
            CadenadeLog "Importación interrumpida por Usuario"
        Else
            CadenadeLog "Import Cancelled by User"
        End If
    End If
    Unload Me
End Sub


Private Sub CmdClientes_Click()
    ' *************************************
    ' Función específica para Importar Clientes.
    '5  dig --> Cod. Cliente
    '25 Car --> Nombre Cliente
    '25 Car --> Extra 1
    '25 Car --> Extra 2
    '25 Car --> Extra 3
    '25 Car --> Extra 4
    ' Este formato no es configurable
    ' a diferencia del de Artículos
    ' *************************************
    Dim nFichClientes As Integer
    Dim codigo As Long
    Dim scodigo As String
    Dim sNombre As String
    Dim sAdicional1 As String
    Dim sAdicional2 As String
    Dim sAdicional3 As String
    Dim sAdicional4 As String
    
    Dim Modificacion As Boolean
    Dim Base As Database
    Dim buffer() As String
    Dim contador As Long
    Dim bucle As Long
    Dim Registro As Recordset
    Success = 4
    Modificacion = False
    cmdimportar.Enabled = False
    For bucle = 0 To 7
        Command1(bucle).Enabled = False
    Next bucle
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    nFichClientes = FreeFile
    If Dir(PathImport & "\clientes.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo ErrorLectura
    Open (PathImport & "\clientes.dat") For Input As #nFichClientes
    contador = 1
    Do Until EOF(nFichClientes)
        ReDim Preserve buffer(contador)
        Line Input #nFichClientes, buffer(contador - 1)
        contador = contador + 1
    Loop
    Close nFichClientes
    On Error GoTo 0
    '
    ' trata los datos
    '
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cliente")
    '
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    For bucle = 0 To contador - 2
        '
        Modificacion = False
        BarraProgreso.Value = bucle * (100 / contador)
        MiProgreso = "Linea : " & (bucle + 1)
      '  ControlImport.LblProgreso.Caption = "Linea : " & (bucle + 1)
        DoEvents
      '
        codigo = 0
        sNombre = Space(25)
        sAdicional1 = Space(25)
        sAdicional2 = Space(25)
        sAdicional3 = Space(25)
        sAdicional4 = Space(25)
        
        If Trim(buffer(bucle)) <> "" Then
            scodigo = (Left(buffer(bucle), 5))
            codigo = Val(scodigo)
            sNombre = Mid(buffer(bucle), 6, 25)
            sAdicional1 = Mid(buffer(bucle), 31, 25)
            sAdicional2 = Mid(buffer(bucle), 56, 25)
            sAdicional3 = Mid(buffer(bucle), 81, 25)
            sAdicional4 = Mid(buffer(bucle), 106, 25)
            
            With Registro
                If Not .EOF Then .MoveLast
                If Not .EOF Then .FindFirst "cod_cli = " & scodigo
                If .EOF Or .NoMatch Then
                    .AddNew
                    Modificacion = True
                Else
                    Edit_Record Registro
                    If !cod_cli <> codigo Then Modificacion = True
                    If !nom_cli <> sNombre Then Modificacion = True
                    If !Data1 <> sAdicional1 Then Modificacion = True
                    If !Data2 <> sAdicional2 Then Modificacion = True
                    If !Data3 <> sAdicional3 Then Modificacion = True
                    If !Data4 <> sAdicional4 Then Modificacion = True
                
                End If
                !cod_cli = codigo
                !nom_cli = sNombre
                !Data1 = sAdicional1
                !Data2 = sAdicional2
                !Data3 = sAdicional3
                !Data4 = sAdicional4
                
                If Modificacion Then !tran_cli = ""
                .Update
            End With
        End If
    Next bucle
    Success = 0
    Unload Me
    Exit Sub
ErrorLectura:
    Success = 12
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : CLIENTES.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo CLIENTES.dat"
        End If
    End If
    Unload Me
End Sub



Public Sub CmdFichas_Click()
    ' *************************************
    ' Función específica para Importar Fichas de Vacuno
    ' Este formato no es configurable
    ' a diferencia del de Artículos
    ' *************************************
    Dim Modificacion As Boolean
    Dim codigo As String
    Dim Identificador As String
    Dim codNacimiento As Long
    Dim codCrianza As Long
    Dim codSacrificio As Long
    Dim codDespiece As Long
    Dim codProduccion As Long
    Dim RSIsacrificio As String
    Dim RSIdespiece As String
    Dim Categoria As Long
    Dim Raza As Long
    ' version d
    Dim Sexo As String
    Dim Edad As Integer
    Dim Peso As Double
    Dim Texto As String
    Dim Lote As String
    Dim Fecha_Nac As String
    Dim Fecha_Sac As String
    Dim Fecha_Des As String
    ' /version d
    Dim Fichas As Integer
    Dim Base As Database
    Dim buffer() As String
    Dim contador As Long
    Dim bucle As Long
    Dim Registro As Recordset
    Success = 4
    Modificacion = False
    cmdimportar.Enabled = False
    For bucle = 0 To 7
        Command1(bucle).Enabled = False
    Next bucle
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    Fichas = FreeFile
    If Dir(PathImport & "\fichas.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo ErrorLectura
    Open (PathImport & "\fichas.dat") For Input As Fichas
    contador = 1
    Do Until EOF(Fichas)
        ReDim Preserve buffer(contador)
        Line Input #Fichas, buffer(contador - 1)
        contador = contador + 1
    Loop
    Close Fichas
    On Error GoTo 0
    '
    ' trata los datos
    '
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from fichavacuno")
    '
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    For bucle = 0 To contador - 2
        '
        Modificacion = False
        BarraProgreso.Value = bucle * (100 / contador)
        MiProgreso = "Linea : " & (bucle + 1)
      '  ControlImport.LblProgreso.Caption = "Linea : " & (bucle + 1)
        DoEvents
            '
        '/ version d
        Sexo = ""
        Edad = 0
        Peso = 0
        Texto = ""
        Lote = ""
        Fecha_Nac = "010180"
        Fecha_Sac = "010180"
        Fecha_Des = "010180"
        '/ version d
        If Trim(buffer(bucle)) <> "" Then
            codigo = Left(buffer(bucle), 3)
            Identificador = Mid(buffer(bucle), 4, 10)
            codNacimiento = Mid(buffer(bucle), 14, 3)
            codCrianza = Mid(buffer(bucle), 17, 3)
            codSacrificio = Mid(buffer(bucle), 20, 3)
            codDespiece = Mid(buffer(bucle), 23, 3)
            codProduccion = Mid(buffer(bucle), 26, 3)
            RSIsacrificio = Mid(buffer(bucle), 29, 10)
            RSIdespiece = Mid(buffer(bucle), 39, 10)
            Categoria = Mid(buffer(bucle), 49, 5)
            Raza = Mid(buffer(bucle), 54, 5)
            ' vacuno d
            If Len(buffer(bucle)) >= 84 Then
                Identificador = Identificador & Mid(buffer(bucle), 55, 10)
                RSIsacrificio = RSIsacrificio & Mid(buffer(bucle), 65, 10)
                RSIdespiece = RSIdespiece & Mid(buffer(bucle), 75, 10)
            End If
            '
            If Len(buffer(bucle)) >= 129 Then
                Lote = Mid(buffer(bucle), 85, 20)
                Texto = Mid(buffer(bucle), 105, 25)
            End If
            '
            If Len(buffer(bucle)) >= 139 Then
                Edad = Mid(buffer(bucle), 130, 3)
                Sexo = Mid(buffer(bucle), 133, 1)
                Peso = Mid(buffer(bucle), 134, 6)
            End If
            '
            If Len(buffer(bucle)) >= 152 Then
                Fecha_Sac = Mid(buffer(bucle), 135, 6)
                Fecha_Nac = Mid(buffer(bucle), 141, 6)
                Fecha_Des = Mid(buffer(bucle), 147, 6)
            End If
            '
            With Registro
                If Not .EOF Then .MoveLast
                If Not .EOF Then .FindFirst "codigo = " & Chr(34) & Val(codigo - 1) & Chr(34)
                If .EOF Or .NoMatch Then
                    .AddNew
                    Modificacion = True
                Else
                    Edit_Record Registro
                    If !codigo <> (codigo - 1) Then Modificacion = True
                    If !Identificador <> Identificador Then Modificacion = True
                    If !codNacimiento <> codNacimiento Then Modificacion = True
                    If !codCrianza <> codCrianza Then Modificacion = True
                    If !codSacrificio <> codSacrificio Then Modificacion = True
                    If !codDespiece <> codDespiece Then Modificacion = True
                    If !codProduccion <> codProduccion Then Modificacion = True
                    If !RSIsacrificio <> RSIsacrificio Then Modificacion = True
                    If !RSIdespiece <> RSIdespiece Then Modificacion = True
                    If !Categoria <> Categoria Then Modificacion = True
                    If !Raza <> Raza Then Modificacion = True
                End If
                !codigo = codigo - 1
                !Identificador = Identificador
                !codNacimiento = codNacimiento
                !codCrianza = codCrianza
                !codSacrificio = codSacrificio
                !codDespiece = codDespiece
                !codProduccion = codProduccion
                !RSIsacrificio = RSIsacrificio
                !RSIdespiece = RSIdespiece
                !Categoria = Categoria
                !Raza = Raza
                '*****************
                ' version D
                '*****************
                !Edad = Edad
                !Sexo = Sexo
                !Peso = Peso
                !Texto = Texto
                !Lote = Lote
                !Fecha_Nac = DateSerial(Right(Fecha_Nac, 2), Mid(Fecha_Nac, 3, 2), Left(Fecha_Nac, 2))
                !Fecha_Sac = DateSerial(Right(Fecha_Sac, 2), Mid(Fecha_Sac, 3, 2), Left(Fecha_Sac, 2))
                !Fecha_Des = DateSerial(Right(Fecha_Des, 2), Mid(Fecha_Des, 3, 2), Left(Fecha_Des, 2))
                '*****************
                ' / version D
                '*****************
                If Modificacion Then !tranficha = ""
                .Update
            End With
        End If
    Next bucle
    Success = 0
    Unload Me
    Exit Sub
ErrorLectura:
    Success = 12
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : FICHAS.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo fichas.dat"
        End If
    End If
    Unload Me
End Sub

Public Sub cmdimportar_Click()
' ***************************************************************
' Daniel Campos Fernández
' 25-7-2001
' Esta función lee el fichero de importación que se ha indicado
' en el cuadro de texto, basandose en la información,
' del fichero de configuración correspondiente (.cfg),y envía los datos a la base,
' contando con la ayuda de la función "tratarRegistro" de este mismo formulario
' *****************************************************************************
    Dim tipo(DimensionArticulo) As String
    Dim Inicio(DimensionArticulo) As Long
    Dim Final(DimensionArticulo) As Long
    Dim posicion(DimensionArticulo) As Integer
    Dim contador As Long
    Dim dato As String
    Dim buffer As String
    Dim fileConfig As Integer
    Dim fileImport As Integer
    Dim bucle As Long
    Success = 4
    Cancelar = False
    fileConfig = FreeFile
    fileImport = FreeFile
    FicheroImport = Trim(txtfichero.Text)
    cmdimportar.Enabled = False
    For bucle = 0 To 7
        Command1(bucle).Enabled = False
    Next bucle
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    DoEvents
    '*********************
    ' Lanza la importación
    If StrConv(FicheroImport, vbUpperCase) <> "FICHAS.DAT" Then
        If StrConv(FicheroImport, vbUpperCase) <> "CLIENTES.DAT" Then
            If LogActivo Then
                If id = 0 Then
                    CadenadeLog "Importando Fichero : " & FicheroImport
                Else
                    CadenadeLog "Importing file : " & FicheroImport
                End If
            End If
            ImportaElFichero FicheroImport
        Else
            If LogActivo Then
                If id = 0 Then
                    CadenadeLog "Importando Fichero : Clientes.DAT"
                Else
                    CadenadeLog "Importing file : Clientes.DAT"
                End If
            End If
            CmdClientes_Click
        
        End If
        
    Else
        If LogActivo Then
            If id = 0 Then
                CadenadeLog "Importando Fichero : FICHAS.DAT"
            Else
                CadenadeLog "Importing file : FICHAS.DAT"
            End If
        End If
        CmdFichas_Click
    End If
    '*************************
    cmdimportar.Enabled = True
    Command1.Item(0).Enabled = True
    Command1.Item(1).Enabled = True
    Command1.Item(2).Enabled = True
    CmdFichas.Enabled = True
    txtfichero.Enabled = True
    Unload Me
End Sub


Private Sub Command1_Click(Index As Integer)
    txtfichero.Text = Command1(Index).Caption
    cmdimportar_Click
End Sub





Private Sub Form_Load()
    Dim bucle As Integer
    Lista.Clear
    If id <> 0 Then frmimportar.Caption = "Import"
    Label1.Caption = LoadResString(id + 745)
    If id <> 0 Then LblAyuda.Caption = LoadResString(id + 633)
    cmdimportar.Caption = "Import"
    cmdcancelar.Caption = LoadResString(id + 288)
    File1.Path = PathLocal & "\etc\"
    File1.FileName = "*.fia"
    For bucle = 1 To 8
        If File1.ListCount < bucle Then
            Command1(bucle - 1).Caption = ""
            Command1(bucle - 1).Enabled = False
        Else
            Command1(bucle - 1).Caption = Left(File1.List(bucle - 1), Len(File1.List(bucle - 1)) - 4)
            Command1(bucle - 1).Enabled = True
        End If
    Next bucle
    txtfichero.Text = FicheroImport
    If FicheroImport <> "" Then
        Timer1.Enabled = True
        cmdimportar.Enabled = False
    End If
End Sub



Private Sub Timer1_Timer()
    Timer1.Enabled = False
    frmimportar.cmdimportar_Click
End Sub

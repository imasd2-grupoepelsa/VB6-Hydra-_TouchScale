VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmimportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación"
   ClientHeight    =   5670
   ClientLeft      =   4815
   ClientTop       =   2535
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmimportar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVend 
      Caption         =   "VenPos.dat"
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdFamilias 
      Caption         =   "FamPos.dat"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdTaras 
      Caption         =   "Taras.dat"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
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
      Left            =   3360
      Top             =   480
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   1320
      TabIndex        =   16
      Top             =   2400
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
      Top             =   480
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
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4815
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

Private Type miGigante
    nSec As Integer
    nPLu As Integer
    nFam As Integer
    sDes As String
    sPes As String
    nPre As Double
    
End Type
Dim regGigante As miGigante


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
    '100 Car --> Nombre Cliente
    '25 Car --> Extra 1 --> NIF
    '100 Car --> Extra 2 --> Dirección
    '25 Car --> Extra 3 --> Código Postal
    '25 Car --> Extra 4 --> Teléfono
    ' Este formato no es configurable
    ' a diferencia del de Artículos
    ' *************************************
    '5  dig --> Cod. Cliente
    '100 Car --> Nombre Cliente
    '25 Car --> Extra 1 --> Dirección
    '100 Car --> Extra 2 --> Población
    '25 Car --> Extra 3 --> CIF
    '25 Car --> Extra 4 --> Teléfono
    '1 dig --> Factura simplificada (0=no, 1=Si)
    '4 dig --> descuento (ejemplo:2025=20,25%)
    Dim nFichClientes As Integer
    Dim codigo As Long
    Dim scodigo As String
    Dim sNombre As String
    Dim sAdicional1 As String
    Dim sAdicional2 As String
    Dim sAdicional3 As String
    Dim sAdicional4 As String
    Dim nFactura As Integer
    Dim nDescuento As Double
    Dim Modificacion As Boolean
    Dim Base As Database
    Dim buffer() As String
    Dim contador As Long
    Dim Bucle As Long
    Dim Registro As Recordset
    Dim lX As Boolean
    
    Success = 4
    Modificacion = False
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    'CmdTramos.Enabled = False
    CmdFamilias.Enabled = False
    CmdVend.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    nFichClientes = FreeFile
    If Dir(PathImport & "\clientes.dat") = "" And Dir(PathImport & "\clientesx.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo errorlectura
    If Dir(PathImport & "\clientes.dat") <> "" Then
        Open (PathImport & "\clientes.dat") For Input As #nFichClientes
    Else
        Open (PathImport & "\clientesx.dat") For Input As #nFichClientes
        lX = True
    End If
    contador = 1
    Do Until EOF(nFichClientes)
        ReDim Preserve buffer(contador)
        Line Input #nFichClientes, buffer(contador - 1)
        contador = contador + 1
    Loop
    Close nFichClientes

    '
    ' trata los datos
    '
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cliente")
    '
    CadenadeLog "Registros:" & CStr(contador)
    
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    For Bucle = 0 To contador - 2
        '
        
        Modificacion = False
        BarraProgreso.Value = Bucle * (100 / (contador))
        MiProgreso = "Linea : " & (Bucle + 1)
        frmimportar.LblProgreso.Caption = "Linea : " & (Bucle + 1)
        Do_Events
      '
        codigo = 0
        sNombre = Space(25)
        sAdicional1 = Space(25)
        sAdicional2 = Space(25)
        sAdicional3 = Space(25)
        sAdicional4 = Space(25)
        nFactura = 0
        nDescuento = 0
        
        If Trim(buffer(Bucle)) <> "" Then
            If lX = False Then
                scodigo = (Left(buffer(Bucle), 5))
                codigo = Val(scodigo)
                sNombre = Mid(buffer(Bucle), 6, 100)
                sAdicional1 = Mid(buffer(Bucle), 106, 25)
                sAdicional2 = Mid(buffer(Bucle), 131, 100)
                sAdicional3 = Mid(buffer(Bucle), 231, 25)
                sAdicional4 = Mid(buffer(Bucle), 256, 25)
                
                CadenadeLog scodigo & " " & sNombre & " " & sAdicional1 & " " & sAdicional2 & " " & sAdicional3 & " " & sAdicional4
                If Len(buffer(Bucle)) > 281 Then
                    nFactura = Val(Mid(buffer(Bucle), 281, 1))
                    nDescuento = Val(Mid(buffer(Bucle), 282)) / 100
                Else
                    nFactura = 0
                    nDescuento = 0
                End If
            Else
                scodigo = (Left(buffer(Bucle), 4))
                codigo = Val(scodigo)
                sNombre = Trim(Mid(buffer(Bucle), 5, 90))
                sAdicional1 = Mid(buffer(Bucle), 95, 30)
                sAdicional2 = Mid(buffer(Bucle), 125)
                sAdicional3 = scodigo
                sAdicional4 = ""
                
                CadenadeLog scodigo & " " & sNombre & " " & sAdicional1 & " " & sAdicional2 & " " & sAdicional3 & " " & sAdicional4
            
            End If
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
                    If !discount <> nDescuento Then Modificacion = True
                    If !factura <> nFactura Then Modificacion = True
                End If
                !cod_cli = codigo
                !nom_cli = sNombre
                !Data1 = sAdicional1
                !Data2 = sAdicional2
                !Data3 = sAdicional3
                !Data4 = sAdicional4
                !discount = nDescuento
                !factura = nFactura
                If Modificacion Then !tran_cli = ""
                .Update
            End With
        End If
    Next Bucle
    'On Error Resume Next
    'Registro.Close  'c2f/caspiunza
    'Base.Close 'c2f/caspiunza
    'On Error GoTo 0
    Success = 0
    'Unload Me
    
    cmdimportar.Enabled = True
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = True
    Next Bucle
    
    CmdFichas.Enabled = True
    CmdClientes.Enabled = True
    CmdTaras.Enabled = True
    CmdFamilias.Enabled = True
    CmdVend.Enabled = True
    
    Exit Sub
errorlectura:
    'Base.Close 'c2f/caspiunza
    Success = 12
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : CLIENTES.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo CLIENTES.dat"
        End If
    End If
    'Unload Me
End Sub


Private Sub CmdFamilias_Click()
    ' *****************************************
    ' Función específica para Importar Familias
    'Nombre del Fichero: FAMPOS.DAT
    'Campo   Longitud    Contenido
    '1       2 dig.      Número de Sección.
    '2       3 dig.      Número de Familia.
    '3       35 Car.     Descripción de la Familia.
    '4       Max.250 Car.    Imagen a signar a la Familia.
    '5       2 Car.      CR+LF
    
    Dim nFichFamilias As Integer
    Dim codi_idn As Integer
    Dim codi_Fam As Integer
    Dim sNombre As String
    Dim sImagen As String
    'Dim Ident(99) As Integer
    Dim Maqui(99) As Integer
    
    Dim Modificacion As Boolean
    Dim Base As Database
    Dim buffer() As String
    Dim contador As Long
    Dim Bucle As Long
    Dim Registro As Recordset
    Dim MyRecSec As Recordset
    Dim n As Integer
    
    Success = 4
    Modificacion = False
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    CmdFamilias.Enabled = False
    CmdVend.Enabled = False
    
    'CmdTramos.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    nFichFamilias = FreeFile()
    If Dir(PathImport & "\fampos.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo errorlectura
        
    Set Base = OpenDatabase(Base_General)
    Set MyRecSec = Base.OpenRecordset("select distinct(codi_ident),secc_maqui from seccion where borrado=false")
    If MyRecSec.EOF Then
        MyRecSec.Close
        Set MyRecSec = Nothing
        Base.Close
        Set Base = Nothing
        Success = 11
        Exit Sub
    End If
    For n = 0 To 99
        Maqui(n) = 0
    Next n
    
    MyRecSec.MoveFirst
    Do While Not MyRecSec.EOF
        Maqui(MyRecSec.Fields("codi_ident")) = MyRecSec.Fields("secc_maqui")
        MyRecSec.MoveNext
    Loop
    MyRecSec.Close
    Set MyRecSec = Nothing
    
    Open (PathImport & "\fampos.dat") For Input As #nFichFamilias
    contador = 1
    Do Until EOF(nFichFamilias)
        ReDim Preserve buffer(contador)
        Line Input #nFichFamilias, buffer(contador - 1)
        contador = contador + 1
    Loop
    Close nFichFamilias

    '
    ' trata los datos
    '
    Set Registro = Base.OpenRecordset("select * from fam_code")
    '
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    For Bucle = 0 To contador - 2
        '
        Modificacion = False
        BarraProgreso.Value = Bucle * (100 / (contador - 2))
        MiProgreso = "Linea : " & (Bucle + 1)
        frmimportar.LblProgreso.Caption = "Linea : " & (Bucle + 1)
        Do_Events
      '
        codi_idn = 0
        codi_Fam = 0
        sNombre = Space(25)
        sImagen = ""
        'Campo   Longitud    Contenido
        '1       2 dig.      Número de Sección.
        '2       3 dig.      Número de Familia.
        '3       35 Car.     Descripción de la Familia.
        '4       Max.250 Car.    Imagen a signar a la Familia.
        
        If Trim(buffer(Bucle)) <> "" Then
            codi_idn = Val(Mid(buffer(Bucle), 1, 2))
            codi_Fam = Val(Mid(buffer(Bucle), 3, 3))
            sNombre = Mid(buffer(Bucle), 6, 35)
            sImagen = Mid(buffer(Bucle), 41)
            
            With Registro
                If Not .EOF Then .MoveLast
                If Not .EOF Then .FindFirst "codi_ident = " & CStr(codi_idn) & " and codi_fam=" & CStr(codi_Fam)
                If .EOF Or .NoMatch Then
                    .AddNew
                    If lFamPluPos = False Then
                        !posicion = 0
                    End If
                Else
                    Edit_Record Registro
                End If
                !codi_ident = codi_idn
                !secc_maqui = Maqui(codi_idn)
                !codi_Fam = codi_Fam
                !txt_fam = Trim(sNombre)
                If lFamPluPos Then
                    !posicion = codi_Fam
                End If
                If Trim(sImagen) <> "NO IMAGE" Then
                    If IsNumeric(sImagen) Then
                        !imagen = APaleta(Val(sImagen))
                    Else
                        If Trim(sImagen) <> "" Then
                            If Dir(Trim(sImagen)) <> "" Then
                                FileCopy Trim(sImagen), App.Path & "\images\families\" & CStr(codi_Fam)
                                !imagen = App.Path & "\images\families\" & CStr(codi_Fam)
                            Else
                                If Dir(App.Path & "\fmlimg.jpg") <> "" Then
                                    FileCopy App.Path & "\fmlimg.jpg", App.Path & "\images\families\" & CStr(codi_Fam)
                                    !imagen = App.Path & "\images\families\" & CStr(codi_Fam)
                                Else
                                    !imagen = "#FFE20000"
                                End If
                            End If
                        Else
                            If Dir(App.Path & "\fmlimg.jpg") <> "" Then
                                FileCopy App.Path & "\fmlimg.jpg", App.Path & "\images\families\" & CStr(codi_Fam)
                                !imagen = App.Path & "\images\families\" & CStr(codi_Fam)
                            Else
                                !imagen = "#FFE20000"
                            End If
                        
                        End If
                    End If
                'Else
                '    FileCopy App.Path & "\vacio.jpg", App.Path & "\images\families\" & CStr(codi_Fam)
                '    !imagen = App.Path & "\images\families\" & CStr(codi_Fam)
                End If
                
                .Update
            End With
        End If
    Next Bucle
    On Error Resume Next
    'Registro.Close  'c2f/caspiunza
    Base.Close 'c2f/caspiunza
    Set Base = Nothing
    Registro.Close
    Set Registro = Nothing
    On Error GoTo 0
    Success = 0
    'Unload Me
    cmdimportar.Enabled = True
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = True
    Next Bucle
    
    CmdFichas.Enabled = True
    CmdClientes.Enabled = True
    CmdTaras.Enabled = True
    CmdFamilias.Enabled = True
    CmdVend.Enabled = True
    
    Exit Sub
errorlectura:
    'Base.Close 'c2f/caspiunza
    CadenadeLog "Error:" & CStr(Err.Number) & " --> " & Err.Description
    On Error GoTo 0
    
    Success = 12
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : FAMPOS.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo FAMPOS.DAT"
        End If
    End If
    'Unload Me

End Sub

Private Sub CmdTaras_Click()
'función específica para importar taras numéricas...
'fichero TARAS.DAT
'Tipo   Valor   CR+LF
'2Dig.  4Dig.   ---> Debe estar de acuerdo con step Balanza...
    
    Dim nFichTara As Integer
    Dim Modificacion As Boolean
    Dim Registro As Recordset
    Dim sTara As String
    Dim Base As Database
    Dim nT As Long
    Dim nV As Long
    Dim Bucle As Integer
    
    Success = 4
    Modificacion = False
    
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    'CmdTramos.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    If Dir(PathImport & "\taras.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo errorlectura

    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from tipotara")

    nFichTara = FreeFile()
    CadenadeLog "Importando Fichero : " & PathImport & "\taras.dat"

    Open (PathImport & "\taras.dat") For Input As nFichTara
    Do While Not EOF(nFichTara)
        Line Input #nFichTara, sTara
        If IsNumeric(Mid(sTara, 1, 2)) And IsNumeric(Mid(sTara, 3)) Then
            nT = Val(Mid(sTara, 1, 2))
            nV = Val(Mid(sTara, 3))
            If nT > 0 And nV < 10000 Then
                Set Registro = Base.OpenRecordset("select * from tipotara where tipo=" & CStr(nT))
                If Registro.EOF Then
                    Registro.AddNew
                    Registro.Fields("tipo") = nT
                    Registro.Fields("valor") = nV
                Else
                    Registro.Edit
                    Registro.Fields("tipo") = nT
                    Registro.Fields("valor") = nV
                End If
            
                Registro.Fields("tran_el") = " "
                Registro.Update
            
            End If
            
        
        End If
    Loop

errorlectura:
    Close #nFichTara
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
        Success = 12
        If MostrarMensajes Then
            If id <> 0 Then
                MsgBox "ERROR : TARAS.DAT"
            Else
                MsgBox "Fallo en la lectura del archivo taras.dat"
            End If
        End If
    Else
        On Error GoTo 0
        Success = 0
    End If
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
    Dim Bucle As Long
    Dim Registro As Recordset
    Success = 4
    Modificacion = False
    
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    'CmdTramos.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    
    If Trim(PathImport) = "" Then PathImport = App.Path
    
    Fichas = FreeFile
    If Dir(PathImport & "\fichas.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    
    On Error GoTo errorlectura
    Open (PathImport & "\fichas.dat") For Input As Fichas
    
    CadenadeLog "Fichero:" & PathImport & "\fichas.dat"
    
    contador = 1
    Do Until EOF(Fichas)
        ReDim Preserve buffer(contador)
        Line Input #Fichas, buffer(contador - 1)
        CadenadeLog buffer(contador - 1)
        contador = contador + 1
    Loop
    Close Fichas

    '
    ' trata los datos
    '
    Set Base = OpenDatabase(Base_General)
    'Set Registro = Base.OpenRecordset("select * from fichavacuno")
    '
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    CadenadeLog "Registros:" & CStr(contador)
    
    For Bucle = 0 To contador - 2
        '
        Modificacion = False
        BarraProgreso.Value = Bucle * (100 / (contador - 2 + 1))
        MiProgreso = "Linea : " & (Bucle + 1)
        frmimportar.LblProgreso.Caption = "Linea : " & (Bucle + 1)
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

'Campo Longitud
'Código de la ficha                  3 caracteres (valores 0 a 999, según capacidad del modelo de balanza)
'Identificador del Animal                     20 caracteres
'Código País de Nacimiento   3 caracteres
'Código País de Crianza                     3 caracteres
'Código País de Sacrificio      3 caracteres
'Código País de Despiece      3 caracteres
'Código País de Producción   3 caracteres
'R.S.I. Sacrificio                                 20 caracteres
'R.S.I. Despiece                      20 caracteres
'Categoría del animal              5 caracteres
'Raza del Animal                     5 caracteres
'Lote                                        20 caracteres
'Texto                                      25 caracteres
'Edad                                       3 caracteres
'Sexo                                       1 carácter
'Peso                                       6 caracteres
'Fecha Nacimiento                           6
'Fecha Sacrificio                           6
'Fecha Despiece                             6

        '/ version d
        If Trim(buffer(Bucle)) <> "" Then
            codigo = Left(buffer(Bucle), 3)
            CadenadeLog codigo
            Identificador = Mid(buffer(Bucle), 4, 20)
            CadenadeLog Identificador
            codNacimiento = Mid(buffer(Bucle), 24, 3)
            CadenadeLog CStr(codNacimiento)
            codCrianza = Mid(buffer(Bucle), 27, 3)
            CadenadeLog CStr(codCrianza)
            codSacrificio = Mid(buffer(Bucle), 30, 3)
            CadenadeLog CStr(codSacrificio)
            codDespiece = Mid(buffer(Bucle), 33, 3)
            CadenadeLog CStr(codDespiece)
            codProduccion = Mid(buffer(Bucle), 36, 3)
            CadenadeLog CStr(codProduccion)
            RSIsacrificio = Mid(buffer(Bucle), 39, 20)
            CadenadeLog RSIsacrificio
            RSIdespiece = Mid(buffer(Bucle), 59, 20)
            CadenadeLog RSIdespiece
            Categoria = Mid(buffer(Bucle), 79, 5)
            CadenadeLog CStr(Categoria)
            Raza = Mid(buffer(Bucle), 84, 5)
            CadenadeLog CStr(Raza)
            Lote = Mid(buffer(Bucle), 89, 20)
            CadenadeLog Lote
            Texto = Mid(buffer(Bucle), 109, 25)
            CadenadeLog Texto
            Edad = Mid(buffer(Bucle), 134, 3)
            CadenadeLog CStr(Edad)
            Sexo = Mid(buffer(Bucle), 137, 1)
            CadenadeLog Sexo
            Peso = Mid(buffer(Bucle), 138, 6)
            CadenadeLog CStr(Peso)
            Fecha_Sac = Mid(buffer(Bucle), 144, 6)
            CadenadeLog Fecha_Sac
            Fecha_Nac = Mid(buffer(Bucle), 150, 6)
            CadenadeLog Fecha_Nac
            Fecha_Des = Mid(buffer(Bucle), 156, 6)
            CadenadeLog Fecha_Des
            
            CadenadeLog "Código:" & CStr(codigo)
            CadenadeLog "Texto:" & Texto
            CadenadeLog "Fecha_Des:" & Fecha_Des
            
            Set Registro = Base.OpenRecordset("select * from fichavacuno where codigo='" & CStr(codigo) & "'")
            If Registro.EOF Then
                Registro.AddNew
            Else
                Registro.Edit
            End If
            CadenadeLog "1.1"
            'With Registro
                'If Not .EOF Then .MoveLast
                'If Not .EOF Then .FindFirst "codigo = " & Chr(34) & CStr(Val(codigo)) & Chr(34)
                'If .EOF Or .NoMatch Then
                '    .AddNew
                '    Modificacion = True
                'Else
                '    Edit_Record Registro
                    'If !codigo <> CStr(Val(codigo)) Then Modificacion = True
                    'If !Identificador <> Identificador Then Modificacion = True
                    
                    'If !codNacimiento <> Val(codNacimiento) Then Modificacion = True
                    'If !codCrianza <> Val(codCrianza) Then Modificacion = True
                    'If !codSacrificio <> Val(codSacrificio) Then Modificacion = True
                    'If !codDespiece <> Val(codDespiece) Then Modificacion = True
                    'If !codProduccion <> Val(codProduccion) Then Modificacion = True
                    
                    'If !RSIsacrificio <> RSIsacrificio Then Modificacion = True
                    'If !RSIdespiece <> RSIdespiece Then Modificacion = True
                    
                    'If !Categoria <> Val(Categoria) Then Modificacion = True
                    'If !Raza <> Val(Raza) Then Modificacion = True
                '    Modificacion = True
                'End If
                Registro.Fields("codigo") = CStr(Val(codigo))
                Registro.Fields("Identificador") = Identificador
                CadenadeLog "1.2"
                Registro.Fields("codNacimiento") = Val(codNacimiento)
                Registro.Fields("codCrianza") = Val(codCrianza)
                CadenadeLog "1.3"
                Registro.Fields("codSacrificio") = Val(codSacrificio)
                Registro.Fields("codDespiece") = Val(codDespiece)
                CadenadeLog "1.4"
                Registro.Fields("codProduccion") = Val(codProduccion)
                Registro.Fields("RSIsacrificio") = RSIsacrificio
                CadenadeLog "1.5"
                Registro.Fields("RSIdespiece") = RSIdespiece
                Registro.Fields("Categoria") = Val(Categoria)
                CadenadeLog "1.6"
                Registro.Fields("Raza") = Val(Raza)
                '*****************
                ' version D
                '*****************
                Registro.Fields("Edad") = Val(Edad)
CadenadeLog "1.7"
                Registro.Fields("Sexo") = Sexo
                Registro.Fields("Peso") = Val(Peso)
CadenadeLog "1.8"
                Registro.Fields("Texto") = Texto
                Registro.Fields("Lote") = Lote
CadenadeLog "1.9"
                Registro.Fields("Fecha_Nac") = DateSerial(Right(Fecha_Nac, 2), Mid(Fecha_Nac, 3, 2), Left(Fecha_Nac, 2))
                Registro.Fields("Fecha_Sac") = DateSerial(Right(Fecha_Sac, 2), Mid(Fecha_Sac, 3, 2), Left(Fecha_Sac, 2))
CadenadeLog "2.0"
                Registro.Fields("Fecha_Des") = DateSerial(Right(Fecha_Des, 2), Mid(Fecha_Des, 3, 2), Left(Fecha_Des, 2))
                '*****************
                ' / version D
                '*****************
                'If Modificacion Then !tranficha = ""
                Registro.Fields("tranficha") = ""
CadenadeLog "2.1"
                Registro.Update
            'End With
        End If
    Next Bucle
    
    'On Error Resume Next
    'Registro.Close 'c2f/caspiunza
    'Base.Close 'c2f/caspiunza
    'On Error GoTo 0
    
    Success = 0
    'Unload Me
    Exit Sub
errorlectura:
    
    Success = 12
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : FICHAS.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo fichas.dat"
        End If
    End If
    'Unload Me
End Sub

Public Sub cmdimportar_Click()
' ***************************************************************
' Daniel Campos Fernández
' 25-7-2001
' Esta función lee el fichero de importación que se ha indicado
' en el cuadro de texto, basandose en la información,
' del fichero de configuración correspondiente (.cfg),y envía los datos a la base,
' contando con la ayuda de la función "tratarRegistro" de este mismo formulario
' ***************************************************************
    Dim tipo(DimensionArticulo) As String
    Dim Inicio(DimensionArticulo) As Long
    Dim Final(DimensionArticulo) As Long
    Dim posicion(DimensionArticulo) As Integer
    Dim contador As Long
    Dim dato As String
    Dim buffer As String
    Dim fileConfig As Integer
    Dim fileImport As Integer
    Dim filegi As Integer
    Dim sTmp1 As String
    Dim sTmp2 As String
    Dim nTmpC As Integer
    Dim Bucle As Long
    
    Bucle = 0
    If Dir(App.Path & "\plrm.txt") <> "" And StrConv(Trim(txtfichero.Text), vbUpperCase) = "PLRM.ASC" Then
        fileConfig = FreeFile()
        Open App.Path & "\plrm.txt" For Input As #fileConfig
        Do While Not EOF(fileConfig)
            Line Input #fileConfig, sTmp1
            sTmp1 = Mid(Trim(sTmp1), 3)
            If Dir(sTmp1 & "plrm.asc") <> "" Then
                If Bucle = 0 Then
                    If Dir(App.Path & "\plrm.asc") <> "" Then
                        Kill App.Path & "\plrm.asc"
                    End If
                    Bucle = 69
                End If
                filegi = FreeFile()
                Open sTmp1 & "plrm.asc" For Input As #filegi
                fileImport = FreeFile()
                Open App.Path & "\plrm.asc" For Append As #fileImport
                Do While Not EOF(filegi)
                    Line Input #filegi, sTmp2
                    Print #fileImport, sTmp2
                
                Loop
                Close #filegi
                Close #fileImport
            
            End If
        
        Loop
        Close #fileConfig
    End If
    
    If Dir(App.Path & "\gianshop.txt") <> "" Then
        fileConfig = FreeFile()
        Open App.Path & "\gianshop.txt" For Input As #fileConfig
        Line Input #fileConfig, buffer
        TiendaActual = Val(buffer)
        Line Input #fileConfig, buffer
        GrupoActual = Val(buffer)
        Close #fileConfig
    Else
        TiendaActual = 0
        GrupoActual = 0
    End If
    
    
    Success = 4
    Cancelar = False
    fileConfig = FreeFile
    fileImport = FreeFile + 1
    filegi = FreeFile + 2
    FicheroImport = Trim(txtfichero.Text)
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    txtfichero.Enabled = False
    CmdTaras.Enabled = False
    'CmdTramos.Enabled = False
    CmdFamilias.Enabled = False
    CmdVend.Enabled = False
    
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    Do_Events
    CadenadeLog LblAyuda.Caption & "...:" & FicheroImport
    ' *******************
    ' Lanza la importación
    If Dir(App.Path & "\gigante.txt") <> "" Then 'c2f gigante (1.7.12-1)
        sPathGigante = App.Path & "\"
        sTipoImpGigante = "STD"
        nTmpC = 0
        Open App.Path & "\gigante.txt" For Input As #filegi
        Do While Not EOF(filegi)
            Select Case nTmpC
            Case 0
                Line Input #filegi, sTmp1
            Case 1
                Line Input #filegi, sTmp2
            Case Else
                Exit Do
            End Select

            nTmpC = nTmpC + 1

        Loop
        Close #filegi
        If nTmpC = 2 Then
            sPathGigante = Trim(sTmp1)
            If Right(sPathGigante, 1) <> "\" Then sPathGigante = sPathGigante & "\"
            sTipoImpGigante = Trim(sTmp2)
            If sTipoImpGigante <> "STD" And sTipoImpGigante <> "ESP" Then sTipoImpGigante = "STD"
        'else deja los valores por defecto...
        End If
        If GrupoActual < 2 Then
            buffer = ""
        Else
            buffer = Format(GrupoActual, "00")
        End If
        adaptaGigante sPathGigante & buffer & "fart.sec"
        ImportaElFichero "fart.dat"
  
        
    Else
        sPathGigante = ""
        sTipoImpGigante = ""
        
        If StrConv(FicheroImport, vbUpperCase) <> "TARAS.DAT" Then
        
            If StrConv(FicheroImport, vbUpperCase) <> "FICHAS.DAT" Then
                
                If StrConv(FicheroImport, vbUpperCase) <> "CLIENTES.DAT" Then
                    
                    If StrConv(FicheroImport, vbUpperCase) <> "TRAMOS.DAT" Then
                    
                        If StrConv(FicheroImport, vbUpperCase) <> "FAMPOS.DAT" Then
                        
                            If StrConv(FicheroImport, vbUpperCase) <> "VENPOS.DAT" Then
                        
                        
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
                                        CadenadeLog "Importando Fichero : VENPOS.DAT"
                                    Else
                                        CadenadeLog "Importing file : VENPOS.DAT"
                                    End If
                                End If
                                CmdVend_Click
                            
                            End If
                        Else
                            If LogActivo Then
                                If id = 0 Then
                                    CadenadeLog "Importando Fichero : FAMPOS.DAT"
                                Else
                                    CadenadeLog "Importing file : FAMPOS.DAT"
                                End If
                            End If
                            CmdFamilias_Click
                        
                        End If
                    Else
                
                        If LogActivo Then
                            If id = 0 Then
                                CadenadeLog "Importando Fichero : Tramos.DAT"
                            Else
                                CadenadeLog "Importing file : Tramos.DAT"
                            End If
                        End If
                        CmdTramos_Click
                    
                    
                    End If
                    
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
        
        Else
            If LogActivo Then
                If id = 0 Then
                    CadenadeLog "Importando Fichero : TARAS.DAT"
                Else
                    CadenadeLog "Importing file : TARAS.DAT"
                End If
            End If
            Call CmdTaras_Click
        End If
        
    End If
    '*************************
    cmdimportar.Enabled = True
    Command1.Item(0).Enabled = True
    Command1.Item(1).Enabled = True
    Command1.Item(2).Enabled = True
    CmdFichas.Enabled = True
    CmdTaras.Enabled = True
    'CmdTramos.Enabled = True
    
    txtfichero.Enabled = True

End Sub

Private Sub BorrarArticulo()
' ***************************************************************
' Daniel Campos Fernández
' 25-7-2001
' Esta función se utiliza para poner a -1 todos los valores
' de la estructura que almacena temporalmente los datos de
' cada línea del fichero de configuración leída.
' Es necesario poner todos los valores a -1 cada vez que
' Se lee una nueva línea para no mezclar datos de las anteriores.
' Si posteriormente se encuentra un dato con valor -1, se entiende
' que no estaba especificado en el fichero de configuración
' y no se modifica
' ***************************************************************
    Dim Bucle As Integer
    With articulo
    .cod = 0
    .plu = 0
    .sec = 0
    .sub = -1
    .fam = -1
    .prc = -1
    
    .wgh = -1
    .cad = -1
    .pre = -1
    .tar = -1
    .etq = -1
    .cba = ""
    For Bucle = 0 To 10
        .tlt(Bucle) = 0
    Next Bucle
    For Bucle = 0 To 20
        .txt(Bucle) = ""
    Next Bucle
    ''''''''''''''''''''''''''cas->v142
    For Bucle = 0 To 59
        .txt2040(Bucle) = ""
    Next Bucle
    ''''''''''''''''''''''''''cas->v142
    For Bucle = 1 To 10
        .tarifa(Bucle) = 0
    Next Bucle
    .typ = -1
    .gpc = -1
    .etq = -1
    .prc1 = 0
    .prc2 = 0
    .trm1 = 0
    .trm2 = 0
    .iva = 0
    .tecla = -1
    .JuegoTecla = -1
    .Merma = 0
    .Presec = 0
    .prc100g = False
    '//
    .m_merma = False
    .m_presec = False
    .m_cod = False
    .m_plu = False
    .m_sec = False
    .m_sub = False
    .m_fam = False
    .m_prc = False
    .m_wgh = False
    .m_cad = False
    .m_pre = False
    .m_tar = False
    .m_etq = False
    .m_cba = False
    For Bucle = 0 To 10
        .m_tlt(Bucle) = False
    Next Bucle
    For Bucle = 0 To 20
        .m_txt(Bucle) = False
    Next Bucle
    For Bucle = 1 To 10
        .m_tarifa(Bucle) = False
    Next Bucle
    .m_typ = False
    .m_gpc = False
    .m_etq = False
    .m_prc1 = False
    .m_prc2 = False
    .m_trm1 = False
    .m_trm2 = False
    .m_iva = False
    .m_tecla = False
    .m_JuegoTecla = False
    .m_prc100g = False
    .lmix = False
    .npeso = 0
    
    End With
End Sub

Public Sub CmdTramos_Click()

Dim nFichTramos As Integer
Dim fileConfig As Integer
Dim RegArt As Recordset
Dim RegTramo As Recordset

Dim tipo() As String
Dim Inicio() As Long
Dim Final() As Long
Dim posicion() As Integer
Dim contador As Long
Dim MiContador As Long
Dim MiBucle As Integer
Dim Bucle As Long
Dim sDrena As String
Dim buffer() As String
Dim Buffer2 As String
Dim sRuta As String
Dim dato As String
Dim IncrementoBarra As Double
Dim Total_lineas As Long
Dim CtEvents As Integer
Dim Base As Database
Dim Balenv As Boolean
Dim MiWgh As Boolean
    Success = 4

    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    'CmdTramos.Enabled = False
    
    txtfichero.Enabled = False
    
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If

    If Dir(PathImport & "\tramos.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    If Dir(App.Path & "\etc\tramos.dat.fia") = "" Then
        Success = 11
        Exit Sub
    End If
    
    CadenadeLog "Fichero a importar:" & PathImport & "\tramos.dat"
    CadenadeLog "Fichero de Configuración:" & App.Path & "\etc\tramos.dat.fia"
    
    On Error GoTo errorlectura
    
    Set Base = OpenDatabase(Base_General)
    Set RegArt = Base.OpenRecordset("select * from articulo")
    Set RegTramo = Base.OpenRecordset("select * from tramos")
    
    'CadenadeLog "Abierta Base de Datos y Recordsets"
    
    inicia
    
    'CadenadeLog "Iniciada estructura articulo"
    
    
    fileConfig = FreeFile()
    ' Abre fichero de configuración y lo lee
    'Do_Events
    'Sleep (1)
    Open App.Path & "\etc\tramos.dat.fia" For Input As fileConfig
    
    Input #fileConfig, sDrena
    
    contador = 0
    If sDrena = "1" Then
        Balenv = True
    Else
        Balenv = False
    End If
    
    Do While Not EOF(fileConfig)
        ReDim Preserve tipo(contador + 1)
        ReDim Preserve Inicio(contador + 1)
        ReDim Preserve Final(contador + 1)
        ReDim Preserve posicion(contador + 1)
        Input #fileConfig, tipo(contador), Buffer2, Inicio(contador), Final(contador), posicion(contador)
        'cadenadelog "...Tipo:" & tipo(contador)
        'cadenadelog "...Expl:" & Buffer2
        'cadenadelog "...Inicio:" & Inicio(contador)
        'cadenadelog "...Final:" & Final(contador)
        'cadenadelog "...Indice:" & posicion(contador)
        contador = contador + 1
    Loop
    Close #fileConfig
    
    nFichTramos = FreeFile()
    
    Open (PathImport & "\tramos.dat") For Input As #nFichTramos
    
    '***************************
    ' Cuenta el nro de líneas del fichero
    ' Para la barra de progreso
    '***************************
    Do Until EOF(nFichTramos)
        IncrementoBarra = IncrementoBarra + 1
        Line Input #nFichTramos, Buffer2
    
        'CadenadeLog Buffer2
        
    Loop
    IncrementoBarra = 100 / IncrementoBarra
    frmimportar.BarraProgreso.Value = 0
    '*********************
    Close #nFichTramos
    
    nFichTramos = FreeFile()
    sRuta = PathImport & "\" & "tramos.dat"
    Open sRuta For Input As nFichTramos
    
    '*****************************
    ' número de líneas del fichero
    '*****************************
    Total_lineas = 1
    frmimportar.LblProgreso.Visible = True
    ' Lee las líneas
    Do While (Not EOF(nFichTramos)) And Cancelar = False
        MiContador = 0
        MiBucle = 0
        Do Until MiBucle = 100 Or EOF(nFichTramos)
            ReDim Preserve buffer(MiContador + 1)
            Line Input #nFichTramos, buffer(MiContador)
            MiContador = MiContador + 1
            MiBucle = MiBucle + 1
        Loop
        For MiBucle = 0 To MiContador - 1
            If Cancelar = True Then Exit For
            frmimportar.LblProgreso.Caption = LoadResString(id + 537) & " : " & Total_lineas
            MiProgreso = LoadResString(id + 537) & " : " & Total_lineas
            If (frmimportar.BarraProgreso.Value + IncrementoBarra) < 100 Then
                frmimportar.BarraProgreso.Value = frmimportar.BarraProgreso.Value + IncrementoBarra
            End If
            If CtEvents >= 50 Then
                Sleep (1)
                Do_Events
                If frmimportar.Lista.ListCount > 10 Then
                    frmimportar.Lista.TopIndex = frmimportar.Lista.ListCount - 10
                End If
                CtEvents = 0
            Else
                CtEvents = CtEvents + 1
            End If
            Total_lineas = Total_lineas + 1
            ' Almacena la información de la línea y comprueba que es correcta
            If Val(Len(buffer(MiBucle))) <> 0 Then
                
                BorrarArticulo
                
                'cadenadelog "...Registro:" & buffer(MiBucle)
                
                For Bucle = 0 To contador - 1
                
                    'cadenadelog "...Tipo:" & tipo(Bucle)
                    'cadenadelog "...Inicio:" & Inicio(Bucle)
                    'cadenadelog "...Final:" & Final(Bucle)
                    'cadenadelog "...Indice:" & posicion(Bucle)
                    
                    dato = Mid(buffer(MiBucle), Inicio(Bucle), Final(Bucle) - Inicio(Bucle) + 1)
                    
                    'cadenadelog "...DATO:" & dato
                    
                    Select Case tratadato(tipo(Bucle), dato, posicion(Bucle))
                    Case 1
                    Case 0
                        If MostrarMensajes Then
                            If id <> 0 Then
                                MsgBox LoadResString(id + 392), vbCritical
                            Else
                                MsgBox "Datos incorrectos en fichero de configuración"
                            End If
                            Success = 18
                        End If
                        Base.Close
                        Exit Sub
                    Case -1
                        If MostrarMensajes Then
                            If id <> 0 Then
                                MsgBox LoadResString(id + 392), vbCritical
                            Else
                                MsgBox "Datos incorrectos en fichero de Precio por Tramos."
                            End If
                            Success = 18
                        End If
                        Base.Close
                        Exit Sub
                    End Select
                Next Bucle
                ' Aquí se introduce el registro en la base
                CadenadeLog "Test"
                If articulo.m_cod And (articulo.m_prc1 Or articulo.m_prc2) Then
                CadenadeLog "pasa."
                'CadenadeLog "Código:" & CStr(articulo.cod)
                
                Set RegArt = Base.OpenRecordset("select * from articulo where codigo=" & CStr(articulo.cod))
                
                If Not RegArt.EOF Then
                    
                    'CadenadeLog "Artículo Localizado..."
                    
                    'If ((articulo.trm1 > 0) And (articulo.trm2 > articulo.trm1) And _
                    '    (articulo.prc2 < articulo.prc1)) Then
                        
                        'CadenadeLog "Modificar Tabla Articulo"
                        
                        RegArt.Edit
                        RegArt.Fields("usatramos") = True
                        RegArt.Fields("tran_plu") = ""
                        If articulo.m_wgh = True Then
                            If articulo.wgh = "0" Then
                                MiWgh = True
                            Else
                                If articulo.wgh = "P" Then
                                    MiWgh = True
                                Else
                                    If articulo.wgh = "W" Then
                                        MiWgh = True
                                    Else
                                        MiWgh = False
                                    End If
                                End If
                            End If
                            If MiWgh Then
                                RegArt.Fields("codi_pes") = "W"
                            Else
                                RegArt.Fields("codi_pes") = "U"
                            End If
                        End If
                        RegArt.Fields("tran_plu") = ""
                        RegArt.Update
                        
'    If UsaEuro Then
'        MiEuro = articulo.prc / 100
'        MiPrecio = EUROtoPTA(articulo.prc / 100)
'    Else
'        MiPrecio = articulo.prc / (10 ^ Decimales)
'        MiEuro = PTAtoEURO(articulo.prc)
'    End If
                        
                        Set RegTramo = Base.OpenRecordset("select * from tramos where codigo=" & CStr(articulo.cod))
                        If RegTramo.EOF Then
                                
                            CadenadeLog "Se añade nuevo registro tabla Tramos"
                            
                            RegTramo.AddNew
                            
                            RegTramo.Fields("codigo") = articulo.cod
                        Else
                            
                            CadenadeLog "Se modifica registro tabla Tramos"
                            
                            RegTramo.Edit
                        
                        End If
                            
                       If articulo.trm1 = -1 Then articulo.trm1 = 0
                        
                        RegTramo.Fields("tramo1") = articulo.trm1 / 1000
                        If articulo.m_trm2 Then
                            RegTramo.Fields("tramo2") = articulo.trm2 / 1000
                        Else
                            RegTramo.Fields("tramo2") = 0
                        End If
                        If UsaEuro Then
                            
                            RegTramo.Fields("euros1") = articulo.prc1 / 100
                            RegTramo.Fields("precio1") = EUROtoPTA(articulo.prc1 / 100)
                            If articulo.m_trm2 Then
                                RegTramo.Fields("euros2") = articulo.prc2 / 100
                                RegTramo.Fields("precio2") = EUROtoPTA(articulo.prc2 / 100)
                            Else
                                RegTramo.Fields("euros2") = 0
                                RegTramo.Fields("precio2") = 0
                            
                            End If
                        Else
                            RegTramo.Fields("euros1") = PTAtoEURO(articulo.prc1)
                            RegTramo.Fields("precio1") = articulo.prc1 / (10 ^ Decimales)
                            If articulo.m_trm2 Then
                                RegTramo.Fields("euros2") = PTAtoEURO(articulo.prc2)
                                RegTramo.Fields("precio2") = articulo.prc2 / (10 ^ Decimales)
                            Else
                                RegTramo.Fields("euros2") = 0
                                RegTramo.Fields("precio2") = 0
                            
                            End If
                        End If

                        RegTramo.Fields("tran_tramo") = ""
                        RegTramo.Fields("borrado") = False
                        RegTramo.Update
                
                
                    'End If
                    
                End If
        
        
                End If
                
            End If
        
        Next MiBucle
        
    
    Loop


errorlectura:
    Close #nFichTramos
    RegArt.Close
    Set RegArt = Nothing
    RegTramo.Close
    Set RegTramo = Nothing
    Base.Close
    Set Base = Nothing
    If Err.Number <> 0 Then
        On Error GoTo 0
        Success = 12
        If MostrarMensajes Then
            If id <> 0 Then
                MsgBox "ERROR : TRAMOS.DAT"
            Else
                MsgBox "Fallo en la lectura del archivo Tramos.dat"
            End If
        End If
    Else
        On Error GoTo 0
        Success = 0
    End If

End Sub

Private Sub CmdVend_Click()
    ' *****************************************
    ' Función específica para Importar Vendedores
    'Nombre del Fichero: VENPOS.DAT
    'Campo   Longitud    Contenido
    '1       2 dig.      Número de Sección.
    '2       4 dig.      Número de Vendedor.
    '3       30 Car.     Nombre de Vendedor.
    '4       2 dig.      Tecla (Posición) Vendedor.
    '5       Max.250 Car.    Imagen a signar a la Familia.
    '6       2 Car.      CR+LF
    
    Dim nFichVend As Integer
    Dim codi_idn As Integer
    Dim codi_Vnd As Integer
    Dim sNombre As String
    Dim nTecla As Integer
    Dim sImagen As String
    'Dim Ident(99) As Integer
    Dim Maqui(99) As Integer
    
    Dim Modificacion As Boolean
    Dim Base As Database
    Dim buffer() As String
    Dim contador As Long
    Dim Bucle As Long
    Dim Registro As Recordset
    Dim MyRecSec As Recordset
    Dim n As Integer
    
    Success = 4
    Modificacion = False
    cmdimportar.Enabled = False
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = False
    Next Bucle
    
    CmdFichas.Enabled = False
    CmdClientes.Enabled = False
    CmdTaras.Enabled = False
    CmdFamilias.Enabled = False
    CmdVend.Enabled = False
    
    'CmdTramos.Enabled = False
    
    txtfichero.Enabled = False
    If id = 0 Then
        LblAyuda.Caption = "Importando el Fichero. Puede tardar varios minutos. Por favor, espere"
    Else
        LblAyuda.Caption = LoadResString(id + 152)
    End If
    nFichVend = FreeFile()
    If Dir(PathImport & "\venpos.dat") = "" Then
        Success = 11
        Exit Sub
    End If
    On Error GoTo errorlectura
        
    Set Base = OpenDatabase(Base_General)
    Set MyRecSec = Base.OpenRecordset("select distinct(codi_ident),secc_maqui from seccion where borrado=false")
    If MyRecSec.EOF Then
        MyRecSec.Close
        Set MyRecSec = Nothing
        Base.Close
        Set Base = Nothing
        Success = 11
        Exit Sub
    End If
    For n = 0 To 99
        Maqui(n) = 0
    Next n
    
    MyRecSec.MoveFirst
    Do While Not MyRecSec.EOF
        Maqui(MyRecSec.Fields("codi_ident")) = MyRecSec.Fields("secc_maqui")
        MyRecSec.MoveNext
    Loop
    MyRecSec.Close
    Set MyRecSec = Nothing
    
    Open (PathImport & "\venpos.dat") For Input As #nFichVend
    contador = 1
    Do Until EOF(nFichVend)
        ReDim Preserve buffer(contador)
        Line Input #nFichVend, buffer(contador - 1)
        contador = contador + 1
    Loop
    Close nFichVend

    '
    ' trata los datos
    '
    'Set Registro = Base.OpenRecordset("select * from vendedor")
    '
    BarraProgreso.Value = 0
    LblProgreso.Visible = True
    For Bucle = 0 To contador - 2
        '
        Modificacion = False
        
        On Error Resume Next
        BarraProgreso.Value = (Bucle + 1) * (100 / (contador - 2))
        If Err.Number <> 0 Then
            'CadenadeLog "Value:" & CStr((bucle + 1) * (100 / (contador - 2)))
            On Error GoTo 0
            On Error GoTo errorlectura
        End If
        
        MiProgreso = "Linea : " & (Bucle + 1)
        frmimportar.LblProgreso.Caption = "Linea : " & (Bucle + 1)
        Do_Events
      '
        codi_idn = 0
        codi_Vnd = 0
        sNombre = Space(30)
        nTecla = 0
        sImagen = ""
        '1       2 dig.      Número de Sección.
        '2       4 dig.      Número de Vendedor.
        '3       30 Car.     Nombre de Vendedor.
        '4       2 dig.      Tecla (Posición) Vendedor.
        '5       Max.250 Car.    Imagen a signar a la Familia.
        
        If Trim(buffer(Bucle)) <> "" Then
            
            codi_idn = Val(Mid(buffer(Bucle), 1, 2))
            codi_Vnd = Val(Mid(buffer(Bucle), 3, 4))
            sNombre = Mid(buffer(Bucle), 7, 30)
            nTecla = Val(Mid(buffer(Bucle), 37, 2))
            sImagen = Mid(buffer(Bucle), 39)
            
            'CadenadeLog "Sec:" & CStr(codi_idn) & " Vnd:" & CStr(codi_Vnd) & " Img:" & sImagen
            Set Registro = Base.OpenRecordset("select * from vendedor where ident_vend=" & CStr(codi_Vnd))
            'With Registro
                'If Not .EOF Then .MoveLast
                'If Not .EOF Then .FindFirst "ident_vend = " & CStr(codi_Vnd)
                If Registro.EOF Then
                    Registro.AddNew
                    'CadenadeLog "Añadido registro..."
                Else
                    Registro.Edit
                    'CadenadeLog "Editado registro"
                End If
                
                Registro.Fields("tienda") = 0
                'CadenadeLog "Tecla"
                Registro.Fields("tec_vend") = nTecla
                'CadenadeLog "Sección"
                Registro.Fields("codi_ident") = codi_idn
                Registro.Fields("secc_maqui") = Maqui(codi_idn)
                'CadenadeLog "Ident. Vend."
                Registro.Fields("ident_vend") = codi_Vnd
                Registro.Fields("codi_vend") = codi_Vnd
                Registro.Fields("nombre") = Trim(sNombre)
                Registro.Fields("prog_vend") = "*"
                Registro.Fields("tran_vend") = "-"
                Registro.Fields("borrado") = False
                Registro.Fields("fecha") = Now()
                
                If Trim(sImagen) <> "NO IMAGE" Then
                    If IsNumeric(sImagen) Then
                        If Val(sImagen) > 32 Then sImagen = "32"
                        If Dir(App.Path & "\images\nn" & Format(Val(sImagen), "00") & ".jpg") <> "" Then
                            FileCopy App.Path & "\images\nn" & Format(Val(sImagen), "00") & ".jpg", App.Path & "\images\vendors\" & CStr(codi_Vnd)
                            Registro.Fields("imagen") = App.Path & "\images\vendors\" & CStr(codi_Vnd)
                        End If
                    Else
                        If Dir(Trim(sImagen)) <> "" Then
                            FileCopy Trim(sImagen), App.Path & "\images\vendors\" & CStr(codi_Vnd)
                            Registro.Fields("imagen") = App.Path & "\images\vendors\" & CStr(codi_Vnd)
                        Else
                            If Dir(App.Path & "\vndimg.jpg") <> "" Then
                                FileCopy App.Path & "\vndimg.jpg", App.Path & "\images\vendors\" & CStr(codi_Vnd)
                                Registro.Fields("imagen") = App.Path & "\images\vendors\" & CStr(codi_Vnd)
                            End If
                        End If
                    End If
                End If
                
                'CadenadeLog "Grabar..."
                Registro.Update
                
            'End With
        End If
    Next Bucle
    On Error Resume Next
    'Registro.Close  'c2f/caspiunza
    Base.Close 'c2f/caspiunza
    Set Base = Nothing
    Registro.Close
    Set Registro = Nothing
    On Error GoTo 0
    Success = 0
    'Unload Me
    
    cmdimportar.Enabled = True
    For Bucle = 0 To 7
        Command1(Bucle).Enabled = True
    Next Bucle
    
    CmdFichas.Enabled = True
    CmdClientes.Enabled = True
    CmdTaras.Enabled = True
    CmdFamilias.Enabled = True
    CmdVend.Enabled = True
    
    Exit Sub
errorlectura:
    'Base.Close 'c2f/caspiunza
    Success = 12
    CadenadeLog "Error:" & CStr(Err.Number) & " --> " & Err.Description
    On Error GoTo 0
    If MostrarMensajes Then
        If id <> 0 Then
            MsgBox "ERROR : VENPOS.DAT"
        Else
            MsgBox "Fallo en la lectura del archivo VENPOS.DAT"
        End If
    End If
    'Unload Me


End Sub

Private Sub Command1_Click(Index As Integer)
    txtfichero.Text = Command1(Index).Caption
    cmdimportar_Click
End Sub
Private Sub Form_Load()
    Dim Bucle As Integer
    Lista.Clear
    If id <> 0 Then frmimportar.Caption = "Import"
    Label1.Caption = LoadResString(id + 745)
    If id <> 0 Then LblAyuda.Caption = LoadResString(id + 633)
    cmdimportar.Caption = "Import"
    cmdcancelar.Caption = LoadResString(id + 288)
    File1.Path = PathLocal & "\etc\"
    File1.FileName = "*.fia"
    For Bucle = 1 To 8
        If File1.ListCount < Bucle Then
            Command1(Bucle - 1).Caption = ""
            Command1(Bucle - 1).Enabled = False
        Else
            Command1(Bucle - 1).Caption = Left(File1.List(Bucle - 1), Len(File1.List(Bucle - 1)) - 4)
            Command1(Bucle - 1).Enabled = True
        End If
    Next Bucle
    txtfichero.Text = FicheroImport
    CadenadeLog "Se espera Timer importación.:" & FicheroImport
    If FicheroImport <> "" Then
        Timer1.Enabled = True
        cmdimportar.Enabled = False
        Do While Timer1.Enabled
            Do_Events
        Loop
    End If
End Sub



Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Do_Events
    CadenadeLog "Entrada Timer Importación..."
    frmimportar.cmdimportar_Click
End Sub

Public Sub adaptaGigante(ByVal sFart As String)
    Dim nFichO As Integer
    Dim nFichD As Integer
    Dim nPosS As Integer
    Dim nPosO As Integer
    Dim sOrg As String
    Dim sParte As String
    Dim nCod As Long
    Dim Base As Database
    Dim nPluTmp As Long
    
    'Tipo de importación borrando artículos previamente.
    'Todos los artículos se considerarán altas y por lo tanto se enviarán.
    'Solo se borran en Hydra.
    If sTipoImpGigante = "ESP" Then
        Set Base = OpenDatabase(Base_General)
        Borrar_todos_Articulo Base
        Base.Close
        Set Base = Nothing
    End If
    
    nFichO = FreeFile()
    nFichD = FreeFile() + 1
    
    Open sFart For Input As #nFichO
    Open App.Path & "\fart.dat" For Output As #nFichD
    Do While Not EOF(nFichO)
    
        nPosS = 1
        nPosO = 1
        
        Line Input #nFichO, sOrg
        'llegan registros que no empiezan con 2...¿?
        If Mid(sOrg, 1, 2) = Chr(34) & "2" Then
            nPosS = InStr(nPosO + 1, sOrg, "@")
            sParte = Mid(sOrg, nPosO, nPosS - 1)
            nPosO = nPosS + 1
            regGigante.nSec = Val(Mid(Trim(sParte), 2, 2))
            regGigante.nPLu = Val(Mid(Trim(sParte), 4, 4))
            'Se rechazan con Sección 0 o PLU 0
            If (regGigante.nSec = 0 Or regGigante.nPLu = 0) Then GoTo siguiente
            'el código se crea con SSPPPP ... NO pero sí porque si no se repite...
            nCod = Val(Mid(sParte, 2, 6))
            
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            nPosO = nPosS + 1
            regGigante.nFam = Val(Mid(Trim(sParte), 3, 2))
            
            'salto
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            nPosO = nPosS + 1
            'código de artículo...viene igual PLU con lo que se repite se coge SSPPPP
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            nPosO = nPosS + 1
            nPluTmp = Val(sParte)
            'If nCod = 0 Then GoTo siguiente
            If ((nPluTmp <> regGigante.nPLu) And (nPluTmp > 0) And (nPluTmp < 10000)) Then '1.7.14
                nCod = nPluTmp
            End If
            
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            sParte = Trim(sParte)
            nPosO = nPosS + 1
            regGigante.sDes = sParte & Space(25 - Len(sParte))
            
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            nPosO = nPosS + 1
            regGigante.sPes = Mid(sParte, 1, 1)
            If regGigante.sPes = "N" Then
                regGigante.sPes = "U"
            Else
                regGigante.sPes = "W"
            End If
            
            nPosS = InStr(nPosO + 1, sOrg, "@")
            If nPosS = 0 Then GoTo siguiente
            sParte = Mid(sOrg, nPosO, (nPosS - nPosO))
            nPosO = nPosS + 1
            regGigante.nPre = Val(Trim(sParte))
            
            'compone cadena
            sParte = Format(regGigante.nSec, "00") & Format(regGigante.nPLu, "0000") & Format(nCod, "000000")
            sParte = sParte & Format(regGigante.nFam, "00") & V12Char(regGigante.sDes) & Format((regGigante.nPre) * 100, "0000000")
            sParte = sParte & regGigante.sPes
            
            Print #nFichD, sParte
            

    
        End If
siguiente:
    Loop
    
    Close #nFichO
    Close #nFichD
    
    

End Sub

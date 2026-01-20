VERSION 5.00
Begin VB.UserControl ControlImportTouch 
   BackColor       =   &H00C00000&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ScaleHeight     =   375
   ScaleWidth      =   1680
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblProgreso 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFF00&
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1692
   End
End
Attribute VB_Name = "ControlImportTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'//////////////
'//////////////
'//////////////
Public Property Let RechazarPrecioCero(yes_or_not As Boolean)
    RechazarPlu0 = yes_or_not
End Property
Public Property Let FactorEuro(Cantidad As Double)
    ValorEuro = Cantidad
End Property
Public Property Let MuestraMensajes(Msg As Boolean)
    MostrarMensajes = Msg
End Property
Public Property Let NumeroDecimales(ByVal Numero As Integer)
    Decimales = Numero
End Property
Public Property Let MostrarFormulario(si_o_no As Boolean)
    ImportVisible = si_o_no
End Property
Public Property Let Euros(ByVal si_o_no As Boolean)
    UsaEuro = si_o_no
End Property
Public Property Let ActivarLog(yes_or_not As Boolean)
    LogActivo = yes_or_not
End Property
Public Property Let No_modWGH(ByVal verdadero As Boolean)
    Conf_Al = verdadero
End Property
Public Property Let Supermercado(ByVal Cliente As String)
    'If Cliente = "Ahorramas" Then
    '    Ahorramas = True
    'Else
    '    Ahorramas = False
    'End If
End Property
Public Property Let FamVacuno(ByVal VacunoFamily As String)
    FamiliaVacuno = VacunoFamily
End Property
Public Property Let EtiquetaEspecial(ByVal EtiquetaVacuno As String)
    EtiquetaEsp = EtiquetaVacuno
End Property
Public Property Let PathNTQ(ByVal fimport As String)
    PathLocal = fimport
End Property
Public Property Let PathFile(ByVal fimport As String)
    PathImport = fimport
End Property
Public Property Let PathBase(ByVal sPath As String)
    Base_General = sPath
End Property
Public Property Let Fichero(ByVal fimport As String)
    FicheroImport = fimport
End Property
Public Property Let activarC9COD(ByVal estado As Integer)
    VarPublics.C9Cod = estado
End Property



Public Property Get Porcentaje() As String
    If Success <> 4 Then
        Porcentaje = MiProgreso
    Else
        Porcentaje = 0
    End If
End Property
Public Property Get Exito() As Integer
'Public Function exito() As Integer
'*****************
' 4 --> Trabajando
' 0 --> OK, importado
' 1 --> OK, interrumpido por usuario
' 10 --> No hay definición del fichero
' 11 --> No Existe el fichero
'****************************
    'On Error Resume Next
    'If Err.Number <> 0 Then
    '    CadenadeLog "EXITO ERROR:" & CStr(Err.Number) & "desc.:" & Err.Description
    '    On Error GoTo 0
    'End If
    'CadenadeLog "EXITO:" & CStr(Success)
    'CadenadeLog "lEsBack:" & CStr(lEsBacK)
    Exito = Success
    
    Sleep (200)
    If lEsBacK Then '1.7.19 And Exito = 0 Then
        If Exito <> 4 Then
            On Error Resume Next
            Exito = 0
            Unload frmimportar
            If Err.Number <> 0 Then
                CadenadeLog "error en descarga frmimportar..."
            End If
            On Error GoTo 0
        Else
            If Dir(App.Path & "\tiendamt.txt") <> "" Then
                On Error Resume Next
                Exito = 0
                Unload frmimportar
                If Err.Number <> 0 Then
                    CadenadeLog "error en descarga frmimportar..."
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Property
Public Property Get version() As String
    version = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Property



Public Sub configurar()
    lTransPluAsc = False
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        lTransPluAsc = True
    End If

    frmconfigurator.Show vbModal
End Sub
Public Sub importar()
    Success = 4
    FicheroImport = ""
    lTransPluAsc = False
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        lTransPluAsc = True
    End If
    
    If ImportVisible Then
        frmimportar.Show
    Else
        Load frmimportar
    End If
    
    'Unload frmimportar 'c2f/caspiunza
    'Sleep (1)
    'do_events
    'Sleep (1)
    'do_events
End Sub
Public Sub ImportarFichaVacuno()
    Success = 4
    If ImportVisible Then
        frmimportar.Show
    Else
        Load frmimportar
    End If
    frmimportar.CmdFichas_Click
    'Unload frmimportar 'c2f/caspiunza
    'Sleep (1)
    'do_events
    'Sleep (1)
    'do_events
End Sub
Public Sub ImportarTramos()
    Success = 4
    If ImportVisible Then
        frmimportar.Show
    Else
        Load frmimportar
    End If
    frmimportar.CmdTramos_Click
    'Unload frmimportar 'c2f/caspiunza
    'Sleep (1)
    'do_events
    'Sleep (1)
    'do_events
End Sub

Public Sub CancelarImportacion()
    Cancelar = True
    Call UserControl_Terminate
End Sub
Public Sub Importar_Background(ByVal MiFichero As String)
    lEsBacK = True
    ImportVisible = False
    Success = 4
    lTransPluAsc = False
    If Dir(App.Path & "\transpluasc.exe") <> "" Then
        lTransPluAsc = True
    End If
    FicheroImport = MiFichero
    CadenadeLog "importar_background:" & FicheroImport
    If ImportVisible Then
        Load frmimportar
        frmimportar.Show
    Else
        Load frmimportar
    End If
   Unload frmimportar
   Call UserControl_Terminate
End Sub
Public Sub LblProgreso_Click()
    'C2f anular tras pruebas
    'frmimportar.Show
    'frmconfigurator.Show
    'UsaEuro = False
    'Decimales = 2
    'PathLocal = "c:\trabajo"
    'PathImport = "c:\trabajo"
    'Base_General = "C:\trabajo\dbasetouch.mdb"
    'frmimportar.Show
End Sub
Public Property Let Idioma(Lengua As Integer)
id = Lengua
If id > 2000 Then id = 1000
End Property
Public Property Let FamFijas(bFijas As Boolean)
    FamiliasFijas = bFijas
End Property
Public Property Let SubFijas(bFijas As Boolean)
    SubseccionesFijas = bFijas
End Property
Private Sub UserControl_Click()
    'PathLocal = "c:\v5"
    'PathImport = "c:\v5"
    'frmimportar.Show
    'ImportarFichaVacuno
    'frmconfigurator.Show
End Sub
Private Sub UserControl_Initialize()
    SubseccionesFijas = False
    FamiliasFijas = False
    RechazarPlu0 = False
    MostrarMensajes = True
    Decimales = 0
    UsaEuro = False
    id = 1000
    ValorEuro = 166.386
    Success = 4
    ImportVisible = True
    If Dir(App.Path & "\SUPECO.SUP") <> "" Then
        SUPECO = True
    Else
        SUPECO = False
    End If
    If Dir(App.Path & "\NOTEXT.TXT") <> "" Then
        AHOLD = True
    Else
        AHOLD = False
    End If
    If Dir(App.Path & "\bdp.txt") <> "" Then
        lBdp = True
    End If
    If Dir(App.Path & "\bdpnot.txt") <> "" Then
        lBdpNot = True
    End If
    If Dir(App.Path & "\test.tst") <> "" Then
        lEsTest = True
    Else
        lEsTest = False
    End If
    If (Dir(App.Path & "\cocobio.cfg") <> "") Or (Dir(App.Path & "\asigfam.cfg") <> "") Then
        lCOCOBIO = True
    Else
        lCOCOBIO = False
    End If
    If Dir(App.Path & "\famplu.pos") <> "" Then
        lFamPluPos = True
    Else
        lFamPluPos = False
    End If
    lEsGestMag = False
    If Dir(App.Path & "\gestmag.imp") <> "" Then
        lEsGestMag = True
    End If
    lBelros = False
    If Dir(App.Path & "\etc\modifica.bal.fia") <> "" And Dir(App.Path & "\etc\articulo.bal.fia") <> "" Then
        lBelros = True
    End If
    C9Cod = 0
End Sub

Public Sub UserControl_Terminate()
    On Error Resume Next
    Unload frmconfigurator
    Set frmconfigurator = Nothing
    If Err.Number <> 0 Then
        CadenadeLog "Err. Cerrando Configurator: " & CStr(Err.Number) & " Desc.:" & Err.Description
    End If
    On Error GoTo 0
    On Error Resume Next
    Unload frmimportar
    Set frmimportar = Nothing
    If Err.Number <> 0 Then
        CadenadeLog "Err. Cerrando Importar: " & CStr(Err.Number) & " Desc.:" & Err.Description
    End If
    On Error GoTo 0
    On Error Resume Next
    If DAO.Workspaces(0).Databases.Count <> 0 Then
        DAO.Workspaces(0).Close
    End If
    If Err.Number <> 0 Then
        CadenadeLog "Err. Cerrando WorkSpace: " & CStr(Err.Number) & " Desc.:" & Err.Description
    End If
    On Error GoTo 0
End Sub

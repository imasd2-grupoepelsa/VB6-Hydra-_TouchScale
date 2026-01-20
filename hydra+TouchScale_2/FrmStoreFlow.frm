VERSION 5.00
Begin VB.Form FrmStoreFlow 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HYDRA BASIC - Modo Servidor"
   ClientHeight    =   2424
   ClientLeft      =   120
   ClientTop       =   600
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdMinimizar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Minimizar >>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton CmdPausa 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pausa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Finalizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1812
   End
   Begin VB.TextBox TxtStoreFlow 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   4572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HYDRA BASIC -ENLACE SOCKET"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   492
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6132
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Storeflow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "SC10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Euroscale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Shape LuzStoreFlow 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   252
      Left            =   0
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape LuzSC10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   252
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   252
   End
   Begin VB.Shape LuzEuroscale 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   252
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   252
   End
   Begin VB.Menu menuTienda 
      Caption         =   "Tienda"
      Begin VB.Menu menuNavegador 
         Caption         =   "Mostradores y Equipos"
      End
      Begin VB.Menu menuConservacion 
         Caption         =   "Grupos de Conservación..."
         Begin VB.Menu MenuGrupoE 
            Caption         =   "Grupos de Euroscale"
         End
         Begin VB.Menu MenuGruposSC10 
            Caption         =   "Grupos de SC10"
         End
      End
      Begin VB.Menu MenuVacuno 
         Caption         =   "Vacuno..."
         Begin VB.Menu MenuPaises 
            Caption         =   "Países"
         End
         Begin VB.Menu MenuFichas 
            Caption         =   "Fichas"
         End
      End
      Begin VB.Menu MenuIVA 
         Caption         =   "Tipos de IVA"
      End
      Begin VB.Menu MenuArticulos 
         Caption         =   "Artículos"
      End
      Begin VB.Menu mnuCambios 
         Caption         =   "Cambios de precio"
      End
      Begin VB.Menu mnunada1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCodBar 
         Caption         =   "Códigos de Barras"
      End
      Begin VB.Menu MenuCabley 
         Caption         =   "Cabeceras y Leyendas"
      End
      Begin VB.Menu MenuPubli 
         Caption         =   "Publicidad y Ofertas"
      End
      Begin VB.Menu mnutextlibre 
         Caption         =   "Textos Libres"
      End
      Begin VB.Menu mnunada18 
         Caption         =   "-"
      End
      Begin VB.Menu MenuVendeores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu MenuTeclas 
         Caption         =   "Teclas Rápidas"
      End
   End
   Begin VB.Menu mnuComunicaciones 
      Caption         =   "Comunicaciones"
      Begin VB.Menu MnuGranTotal 
         Caption         =   "Grandes Totales"
      End
      Begin VB.Menu mnuVoid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBaseMaestra 
         Caption         =   "Base Maestra"
         Index           =   0
      End
      Begin VB.Menu mnuBaseMaestra 
         Caption         =   "Modificaciones"
         Index           =   1
      End
      Begin VB.Menu mnuSeperado 
         Caption         =   "Envíos por separado"
         Begin VB.Menu mnuEnvSecciones 
            Caption         =   "Secciones"
         End
         Begin VB.Menu mnuEnvSubsecciones 
            Caption         =   "Subsecciones"
         End
         Begin VB.Menu mnuEnvFamilias 
            Caption         =   "Familias"
         End
         Begin VB.Menu mnuEnvEquipos 
            Caption         =   "Equipos"
         End
         Begin VB.Menu mnuEnvPaises 
            Caption         =   "Países"
         End
         Begin VB.Menu mnuEnvFichas 
            Caption         =   "Fichas de Vacuno"
         End
         Begin VB.Menu mnuEnvIVA 
            Caption         =   "Tipos de IVA"
         End
         Begin VB.Menu mnuEnvArticulos 
            Caption         =   "Artículos"
         End
         Begin VB.Menu mnuEnvCodigos 
            Caption         =   "Códigos de Barras"
         End
         Begin VB.Menu mnuEnvCabley 
            Caption         =   "Cabeceras y Leyendas"
         End
         Begin VB.Menu mnuEnvPublicidad 
            Caption         =   "Publicidad y ofertas"
         End
         Begin VB.Menu mnuEnvVendedores 
            Caption         =   "Vendedores"
         End
         Begin VB.Menu mnuEnvTeclas 
            Caption         =   "Teclas Rápidas"
         End
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu cmdconfiguracion 
         Caption         =   "Configuración"
      End
      Begin VB.Menu mnunada23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProtocolo 
         Caption         =   "Detalles de comunicación"
      End
      Begin VB.Menu mnunada55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaTiquets 
         Caption         =   "Consulta de Tiquets"
      End
      Begin VB.Menu mnunada234 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoEnlace 
         Caption         =   "Paso a Modo sin Enlace"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmStoreFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdconfiguracion_Click()
    FrmConfiguracion2.Show vbModal, Me
    StoreFlow = True
    Activar_Menues
End Sub

Private Sub CmdMinimizar_Click()
    Form_DblClick
End Sub

Private Sub CmdPausa_Click()
    If PausaTiquets = False Then
        PausaTiquets = True
        FrmStoreFlow.CmdPausa.Enabled = False
        FrmStoreFlow.CmdPausa.Caption = "Deteniendo..."
        If CI_EstadoSistema = 1 Or CI_EstadoSistema = 0 Then
            FrmStoreFlow.mnuComunicaciones.Enabled = True
            FrmStoreFlow.CmdPausa.Enabled = True
            FrmStoreFlow.cmdconfiguracion.Enabled = True
            FrmStoreFlow.CmdPausa.Caption = "En Pausa"
            mnuNoEnlace.Enabled = True
        End If
    Else
        FrmStoreFlow.mnuComunicaciones.Enabled = False
        FrmStoreFlow.CmdPausa.Enabled = True
        FrmStoreFlow.cmdconfiguracion.Enabled = False
        PausaTiquets = False
        FrmStoreFlow.CmdPausa.Caption = "Pausa"
        mnuNoEnlace.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    CI_Salir = True
    If CI_EstadoSistema = 0 Or CI_EstadoSistema = 1 Then SALIRPROGRAMA
End Sub



Private Sub Form_DblClick()
    FrmStoreFlow.Visible = False
    
End Sub



Private Sub Form_Load()
    If PausaTiquets Then
        cmdconfiguracion.Enabled = True
        mnuComunicaciones.Enabled = True
    Else
        cmdconfiguracion.Enabled = False
        mnuComunicaciones.Enabled = False
    End If
    Activar_Menues
End Sub
Private Sub Activar_Menues()
    If UsaGamaAlta Then
        mnuEnvFichas.Visible = True
        mnuEnvIVA.Visible = True
        mnuEnvPaises.Visible = True
        MenuVacuno.Visible = True
        MenuIVA.Visible = True
        MenuGrupoE.Visible = True
        Label1.Visible = True
        LuzEuroscale.Visible = True
    Else
        mnuEnvFichas.Visible = False
        mnuEnvIVA.Visible = False
        mnuEnvPaises.Visible = False
        MenuVacuno.Visible = False
        MenuIVA.Visible = False
        MenuGrupoE.Visible = False
        Label1.Visible = False
        LuzEuroscale.Visible = False
    End If
    If usaSC10 Then
        mnuEnvSubsecciones.Visible = True
        mnuEnvFamilias.Visible = True
        mnuEnvEquipos.Visible = True
        MenuGruposSC10.Visible = True
        Label2.Visible = True
        LuzSC10.Visible = True
    Else
        mnuEnvSubsecciones.Visible = False
        mnuEnvFamilias.Visible = False
        mnuEnvEquipos.Visible = False
        MenuGruposSC10.Visible = False
        Label2.Visible = False
        LuzSC10.Visible = False
    End If
End Sub
Private Sub MenuArticulos_Click()
    FrmArticulo2.Show
End Sub

Private Sub MenuCabley_Click()
    FrmCabley2.Show
End Sub

Private Sub MenuCodBar_Click()
    FrmCodBar2.Show
End Sub

Private Sub MenuFichas_Click()
    frmFichasVacuno.Show
End Sub

Private Sub MenuGrupoE_Click()
    FrmConserv386.Show
End Sub

Private Sub MenuGruposSC10_Click()
    FrmConservSC10.Show
End Sub

Private Sub MenuIVA_Click()
    FrmIVA.Show
End Sub

Private Sub menuNavegador_Click()
    FrmSeccionyMaquina.Show
End Sub

Private Sub MenuPaises_Click()
    FrmPaises.Show
End Sub

Private Sub MenuPubli_Click()
    FrmPublicidad2.Show
End Sub

Private Sub MenuTeclas_Click()
    FrmTeclas2.Show
End Sub

Private Sub MenuVendeores_Click()
    FrmVendedor2.Show
End Sub

Private Sub mnuBaseMaestra_Click(Index As Integer)
    MostrarMonitor
    If Index = 0 Then
        AN_BaseMaestra False
    Else
        AN_BaseMaestra True
    End If
End Sub

Private Sub mnuCambios_Click()
    FrmCambiosPrecio.Show
End Sub

Private Sub mnuClientes_Click()
    FrmClientes.Show
End Sub

Private Sub mnuConsultaTiquets_Click()
    FrmConsultaTiquets.Show
End Sub

Private Sub mnuEnvArticulos_Click()
    MostrarMonitor
    AN_Articulos False
End Sub

Private Sub mnuEnvCabley_Click()
    MostrarMonitor
    AN_Cabeceras_Leyendas False
End Sub

Private Sub mnuEnvCodigos_Click()
    MostrarMonitor
    AN_CodigosBarras False
    
End Sub

Private Sub mnuEnvEquipos_Click()
    MostrarMonitor
    AN_Equipos False
    
End Sub

Private Sub mnuEnvFamilias_Click()
    MostrarMonitor
    AN_Familias False
    
End Sub

Private Sub mnuEnvFichas_Click()
    MostrarMonitor
    AN_FichasVacuno False
    
End Sub

Private Sub mnuEnvIVA_Click()
    MostrarMonitor
    AN_IVAS False
End Sub

Private Sub mnuEnvPaises_Click()
    MostrarMonitor
    AN_Paises False
End Sub

Private Sub mnuEnvPublicidad_Click()
    MostrarMonitor
    AN_publicidad False
    
End Sub

Private Sub mnuEnvSecciones_Click()
    MostrarMonitor
    AN_Secciones False
    
End Sub

Private Sub mnuEnvSubsecciones_Click()
    MostrarMonitor
    AN_Subsecciones False
    
End Sub

Private Sub mnuEnvTeclas_Click()
    MostrarMonitor
    AN_TeclasRapidas False
    
End Sub

Private Sub mnuEnvVendedores_Click()
    MostrarMonitor
    AN_Vendedores False
    
End Sub

Private Sub MnuGranTotal_Click()
    MostrarMonitor
    AN_RecogerTotales True
End Sub


Private Sub mnuNoEnlace_Click()
    Dim arch As Integer
    arch = FreeFile()
    Open App.Path & "\stflow.back" For Output As #arch
    Close #arch
    If Dir(App.Path & "\laststart") <> "" Then Kill App.Path & "\laststart"
    
    '////////////////////////////////
    If Form1.tcpclient.State <> sckClosed Then
        Form1.tcpclient.Close
        Do Until Form1.tcpclient.State = sckClosed
            Do_Events
        Loop
    End If
    If Form1.datacliente.State <> sckClosed Then
        Form1.datacliente.Close
        Do Until Form1.datacliente.State = sckClosed
            Do_Events
        Loop
    End If
    Shell App.Path & "\hydra.exe"
    End
End Sub

Private Sub mnuProtocolo_Click()
    Form1.Balanza.Show_Monitor
End Sub

Private Sub mnutextlibre_Click()
    FrmTextoLibre.Show
End Sub

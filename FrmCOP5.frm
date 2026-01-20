VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCOP5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperación de Datos de COP5"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "FoxPro 2.5;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdAccion 
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
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CmdAccion 
      Caption         =   "Recuperar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2778
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tienda"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Moneda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Factor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Decimales"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Indique la Carpeta donde está instalado COP5 :"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "FrmCOP5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RutaCop As String
Private RutaTienda As String
Private TipoTienda As Integer
Private TEuro As Boolean
Private Sub CmdAccion_Click(Index As Integer)
    Select Case Index
    Case 1
        Unload Me
        Exit Sub
    Case Else
        RutaTienda = Lista.SelectedItem.SubItems(3)
        If Lista.SelectedItem.SubItems(2) = "Euroscale" Then
            TipoTienda = 0
        Else
            TipoTienda = 1
        End If
        If Trim(Lista.SelectedItem.SubItems(4)) = "Euro" Then
            TEuro = True
        Else
            TEuro = False
        End If
        RecuperaDatos
    End Select
    
End Sub

Private Sub Cmdaceptar_Click()
    
    Dim Registro As Recordset
    Text1.Text = Trim(Text1.Text)
    If Right(Text1.Text, 1) = "\" Then
        Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    End If
    If Dir(Text1.Text, vbDirectory) = "" Then
        MsgBox "La carpeta indicada no existe", vbCritical
        Exit Sub
    End If
    RutaCop = Text1.Text
    If Dir(RutaCop & "\tienda.dbf") = "" Then
        MsgBox "No existen datos de tiendas en la ubicación seleccionada"
        Exit Sub
    End If
    Data1.DatabaseName = RutaCop
    Data1.RecordSource = "select * from tienda where tipo_con=1 or tipo_con=4 or tipo_con=0"
    Data1.Refresh
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Lista.ListItems.Add , , !tienda
                If Not IsNull(!Nombre) Then
                    Lista.ListItems(Lista.ListItems.Count).SubItems(1) = !Nombre
                Else
                    Lista.ListItems(Lista.ListItems.Count).SubItems(1) = ""
                End If
                Select Case !tipo_con
                    Case 4
                        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = "Euroscale"
                    Case 1
                        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = "SC10"
                    Case 0
                        Lista.ListItems(Lista.ListItems.Count).SubItems(2) = "C9/C10"
                End If
                Lista.ListItems(Lista.ListItems.Count).SubItems(3) = !path_tien
                On Error Resume Next
                Lista.ListItems(Lista.ListItems.Count).SubItems(4) = !moneda
                Lista.ListItems(Lista.ListItems.Count).SubItems(5) = !euroeq
                Lista.ListItems(Lista.ListItems.Count).SubItems(6) = !dece_expor
                On Error GoTo 0
                .MoveNext
            Loop
        End If
    End With
    
    If Lista.ListItems.Count > 0 Then
        Lista.ListItems(1).Selected = True
        CmdAccion(0).Enabled = True
        Text1.Enabled = False
        CmdAceptar.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = Form2.Icon
    CmdAccion(0).Enabled = False
End Sub
Private Sub RecuperaDatos()
    Dim MiSec As DB_mostrador
    Dim MiEqp As DB_Balanzas
    Dim MiArt As DB_Articulo
    Dim MiCB As DB_CodigoBarras
    Dim MiCL As DB_Cabeceras
    Dim MiTec As DB_Teclas
    Dim Base As Database
    Dim Tabla As TableDef
    Dim Registro As Recordset
    Dim RegArt As Recordset
    Dim B As Integer
    Dim C As Integer
    '*****************
    
    '*****************
    Data1.DatabaseName = RutaTienda
    Data1.RecordSource = "select * from seccion"
    Data1.Refresh
    CmdAccion(0).Enabled = False
    FrmCOP5.Enabled = False
    Set Base = AbrirBase()
    Set RegArt = Base.OpenRecordset("select * from articulo where borrado=false")
    '***************
    ' sección
    '***************
    C = 0
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s1
                C = C + 1
                L2.Caption = "Sección : " & C
                Do_Events
                MiSec.NMostrador = !codi_ident
                MiSec.NSeccion = !secc_maqui
                If Not IsNull(!descripcio) Then
                    MiSec.descripcion = !descripcio
                Else
                    MiSec.descripcion = ""
                End If
                MiSec.Tipo = TipoTienda
                Alta_Mostrador MiSec
s1:
                .MoveNext
            Loop
        End If
    End With
    '***************
    ' equipos
    '***************
    C = 0
    Data1.RecordSource = "select * from equipos"
    Data1.Refresh
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s2
                C = C + 1
                L2.Caption = "Equipos : " & C
                Do_Events
                MiEqp.NMostrador = !codi_ident
                MiEqp.NBalanza = !numero_eqp
                If Not IsNull(!descripcio) Then
                    MiEqp.descripcion = !descripcio
                Else
                    MiEqp.descripcion = ""
                End If
                If TipoTienda = 0 Then
                    MiEqp.Tipo = 0
                    If tcpip Then
                        MiEqp.DireccionIP = "127.000.000.001"
                    End If
                Else
                    MiEqp.Tipo = 100
                End If
                Alta_Balanza MiEqp
s2:
                .MoveNext
            Loop
        End If
    End With
    '***************
    ' articulos
    '***************
    C = 0
    Data1.RecordSource = "select * from articulo"
    Data1.Refresh
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s3
                C = C + 1
                L2.Caption = "Articulos : " & C
                Do_Events
                MiArt.PRC1 = -1
                MiArt.PRC2 = -1
                MiArt.TRM1 = -1
                MiArt.TRM2 = -1
                MiArt.Mostrador = !codi_ident
                MiArt.subsec = !codi_sub
                MiArt.familia = !codi_fam
                MiArt.plu = !plu
                MiArt.codigo = !codigo
                If UCase(Trim(!codi_pes)) = "S" Then
                    MiArt.WGH = False
                Else
                    MiArt.WGH = True
                End If
                MiArt.IVA = !tipo_iva
                If UsaEuro Then
                    MiArt.precio = !Euros
                Else
                    MiArt.precio = !precio
                End If
                
                MiArt.caducidad = !caducidad
                MiArt.Preferente = !conprefe
                MiArt.tara = !tara
                If !Etiqueta <> 10 Then
                    MiArt.Etiqueta = !Etiqueta
                Else
                    MiArt.Etiqueta = 255
                End If
                MiArt.TipoLetra(0) = !tip_let1
                MiArt.TipoLetra(1) = !tip_let2
                MiArt.TipoLetra(2) = !tip_let3
                MiArt.TipoLetra(3) = !tip_let4
                MiArt.TipoLetra(4) = !tip_let5
                MiArt.TipoLetra(5) = !tip_let6
                MiArt.TipoLetra(6) = !tip_let7
                MiArt.TipoLetra(7) = !tip_let8
                MiArt.TipoLetra(8) = !tip_let9
                MiArt.TipoLetra(9) = !tip_let0
                For B = 0 To 20
                    MiArt.Descriptivos(B) = ""
                Next B
                If Not IsNull(!des_plu1) Then MiArt.Descriptivos(0) = !des_plu1
                If Not IsNull(!des_plu2) Then MiArt.Descriptivos(1) = !des_plu2
                If Not IsNull(!des_plu3) Then MiArt.Descriptivos(2) = !des_plu3
                If Not IsNull(!des_plu4) Then MiArt.Descriptivos(3) = !des_plu4
                If Not IsNull(!des_plu5) Then MiArt.Descriptivos(4) = !des_plu5
                If Not IsNull(!des_plu6) Then MiArt.Descriptivos(5) = !des_plu6
                If Not IsNull(!des_plu7) Then MiArt.Descriptivos(6) = !des_plu7
                If Not IsNull(!des_plu8) Then MiArt.Descriptivos(7) = !des_plu8
                If Not IsNull(!des_plu9) Then MiArt.Descriptivos(8) = !des_plu9
                If Not IsNull(!des_plu0) Then MiArt.Descriptivos(9) = !des_plu0
                If Not IsNull(!des_plua) Then MiArt.Descriptivos(10) = !des_plua
                If Not IsNull(!des_plub) Then MiArt.Descriptivos(11) = !des_plub
                If Not IsNull(!des_pluc) Then MiArt.Descriptivos(12) = !des_pluc
                If Not IsNull(!des_plud) Then MiArt.Descriptivos(13) = !des_plud
                If Not IsNull(!des_plue) Then MiArt.Descriptivos(14) = !des_plue
                If Not IsNull(!des_pluf) Then MiArt.Descriptivos(15) = !des_pluf
                If Not IsNull(!des_plug) Then MiArt.Descriptivos(16) = !des_plug
                If Not IsNull(!des_pluh) Then MiArt.Descriptivos(17) = !des_pluh
                If Not IsNull(!des_plui) Then MiArt.Descriptivos(18) = !des_plui
                If Not IsNull(!des_pluj) Then MiArt.Descriptivos(19) = !des_pluj
                If Not IsNull(!des_pluk) Then MiArt.Descriptivos(20) = !des_pluk
                If MiArt.codigo <> 0 Then
                    If MiArt.plu <> 0 Then
                        If MiArt.Mostrador <> 0 Then
                            Alta_Articulo MiArt, Base, RegArt
                        End If
                    End If
                End If
s3:
                .MoveNext
            Loop
        End If
    End With
    Base.Close
    '***************
    ' codbar
    '***************
    Data1.RecordSource = "select * from codbar"
    Data1.Refresh
    C = 0
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s4
                C = C + 1
                L2.Caption = "Cod. Bar. : " & C
                Do_Events
                MiCB.equipo = !numero_eqp
                MiCB.Mostrador = !codi_ident
                If Not IsNull(!cb_vnt) Then
                    MiCB.Datos(0) = !cb_vnt
                Else
                    MiCB.Datos(0) = ""
                End If
                If Not IsNull(!cb_sup) Then
                    MiCB.Datos(1) = !cb_sup
                Else
                    MiCB.Datos(1) = ""
                End If
                If Not IsNull(!cb_mix) Then
                    MiCB.Datos(2) = !cb_mix
                Else
                    MiCB.Datos(2) = ""
                End If
                If Not IsNull(!cb_ev1) Then
                    MiCB.Datos(3) = !cb_ev1
                Else
                    MiCB.Datos(3) = ""
                End If
                If Not IsNull(!cb_ev2) Then
                    MiCB.Datos(4) = !cb_ev2
                Else
                    MiCB.Datos(4) = ""
                End If
                Alta_CodigoBarras MiCB
s4:
                .MoveNext
            Loop
        End If
    End With
    '***************
    ' cabley
    '***************
    Data1.RecordSource = "select * from cabley"
    Data1.Refresh
    C = 0
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s5
                C = C + 1
                L2.Caption = "Cab./Ley. : " & C
                Do_Events
                MiCL.Mostrador = !codi_ident
                MiCL.equipo = !numero_eqp
                If Not IsNull(!cabecera01) Then
                    MiCL.Cabecera(0) = !cabecera01
                End If
                If Not IsNull(!cabecera02) Then
                    MiCL.Cabecera(1) = !cabecera02
                End If
                If Not IsNull(!cabecera03) Then
                    MiCL.Cabecera(2) = !cabecera03
                End If
                If Not IsNull(!cabecera04) Then
                    MiCL.Cabecera(3) = !cabecera04
                End If
                If Not IsNull(!cabecera05) Then
                    MiCL.Cabecera(4) = !cabecera05
                End If
                If Not IsNull(!leyenda01) Then
                    MiCL.leyenda(0) = !leyenda01
                End If
                If Not IsNull(!leyenda02) Then
                    MiCL.leyenda(1) = !leyenda02
                End If
                If Not IsNull(!leyenda03) Then
                    MiCL.leyenda(2) = !leyenda03
                End If
                If Not IsNull(!leyenda04) Then
                    MiCL.leyenda(3) = !leyenda04
                End If
                If Not IsNull(!leyenda05) Then
                    MiCL.leyenda(4) = !leyenda05
                End If
                MiCL.TLCabecera(0) = !tletracab1
                MiCL.TLCabecera(1) = !tletracab2
                MiCL.TLCabecera(2) = !tletracab3
                MiCL.TLCabecera(3) = !tletracab4
                MiCL.TLCabecera(4) = !tletracab5
                MiCL.TLLeyenda(0) = !tletraley1
                MiCL.TLLeyenda(1) = !tletraley2
                MiCL.TLLeyenda(2) = !tletraley3
                MiCL.TLLeyenda(3) = !tletraley4
                MiCL.TLLeyenda(4) = !tletraley5
                Alta_Cabeceras MiCL
s5:
                .MoveNext
            Loop
        End If
    End With
    '***************
    ' teclas
    '***************
    C = 0
    Data1.RecordSource = "select * from teclas"
    Data1.Refresh
    Set Registro = Data1.Recordset
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                On Error GoTo s6
                C = C + 1
                L2.Caption = "Teclas : " & C
                Do_Events
                MiTec.CodArticulo = !codigo
                MiTec.PluArticulo = 0
                MiTec.SecArticulo = 0
                MiTec.Destino_CodiTec = !codi_tec
                MiTec.Destino_NBalanza = !numero_eqp
                MiTec.Destino_NMostrador = !codi_ident
                MiTec.Destino_Tabla = !juego
                B = Alta_Teclas(MiTec)
s6:
                .MoveNext
            Loop
        End If
    End With
    On Error GoTo 0
    MsgBox "Proceso Finalizado", vbInformation
    Unload Me
    Exit Sub
End Sub

   
    
    




Attribute VB_Name = "Arranque"
Option Explicit
Public Descarga As Boolean
Public Backup_path As String
Global Fin_Totales As Boolean
Global TiendaActual As Integer
'*****************************
'*****************************
'*****************************
Sub Main()
Dim Arch As Integer
Dim orden As String
Dim P1 As String
Dim Buf1 As String
Dim Buf2 As String
Dim Buf3 As String
Dim sParte As String

'*****************
' /DEBUG
'*****************
'SubseccionesFijas = False
'FamiliasFijas = False
'CargaIdiomas
'Base_General = "e:\_lab\dbasetouch.mdb"
'FrmTotales.Show
'Do While FrmTotales.Visible = True
'    Do_Events
'Loop
'End
'*******
' /DEBUG
'*******
    Sleep (3000)
    varpublics.version = tools.versionar
''''''''''
                'Backup_path = App.Path & "\backup"
                'MostrarModal FrmBackup
                'End

''''''''''
    If Dir(App.Path & "\auxiliartouch.tmp") <> "" Then
        On Error GoTo errorf
        Arch = FreeFile()
        Open App.Path & "\auxiliartouch.tmp" For Input As #Arch
        Line Input #Arch, Base_General
        Line Input #Arch, Hydra_INI
        Line Input #Arch, orden
        Line Input #Arch, P1
        Line Input #Arch, Buf1
        Line Input #Arch, Buf2
        Line Input #Arch, Buf3
        Close (Arch)
        
        Call obtenPATHngt(Base_General)
        sOrdenNGT = orden
        
        'Kill App.Path & "\auxiliar.tmp"
        On Error GoTo 0
    Else
        If Trim(Command) <> "" Then
            If UCase(Trim(Command)) = "/BACKUP" Then
                Backup_path = App.Path & "\backup"
                MostrarModal FrmBackup
                End
            Else
                MsgBox "This program must be executed by HydraTouch", vbCritical
                End
            End If
        Else
            MsgBox "This program must be executed by HydraTouch", vbCritical
            End
        End If
    End If
    If UCase(Buf1) = "Y" Then
        HaySeguridad = True
        UsuarioActual.Nombre = Buf2
        UsuarioActual.Asociado = Val(Buf3)
    Else
        HaySeguridad = False
        UsuarioActual.Nombre = "admin"
        UsuarioActual.Asociado = 0
    End If
    'c2f 1.7.8
    'If Dir(App.Path & "\ngt.txt") <> "" Then
        Arch = FreeFile()
        If Dir(sPathNGT & "ngtord") = "" Then
            Open sPathNGT & "ngtord" For Output As #Arch
            Print #Arch, "0001"
        Else
            Open sPathNGT & "ngtord" For Input As #Arch
            Line Input #Arch, sParte
            If Val(sParte) > 9999 Then sParte = "0001"
            
            nlastGT = Val(sParte)
        End If
        Close #Arch
    'End If
    ''''''''''
    ChDir (App.Path)
    MiruTa = CurDir
    Check_Especiales
    '//DataEnvironment1.Connection1.ConnectionString = "Provider=MSDataShape.1;Persist Security Info=False;Mode=Read;Data Source=" & Base_General & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
    'DataEnvironment1.Connection2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=" & Chr(34) & Chr(34) & ";Data Source=" & MiruTa & "\dirtmptouch\tmp.mdb" & ";Mode=Read;Persist Security Info=True"
    If LCase(orden) <> "listado_tiendas" And LCase(orden) <> "backupmulti" And LCase(orden) <> "hazbackupmulti" Then LeerParametrosInicio
debugi:
    CargaIdiomas
    Select Case LCase(orden)
        Case "rpart"
           Informe_Articulos_Seccion 0
        Case "rpartsec"
            Informe_Articulos_Seccion Val(P1)
        Case "rpvendor"
            Informe_Vendedores
        Case "rpmostrador"
            Informe_Mostradores
        Case "rpequipos"
            Informe_Equipos
        Case "rpcprecios"
            
            Informe_CambiosPrecio
            'provisional HiperUsera:
            'ver Call Informe_Etq8
        Case "rpcptemporal"
            Informe_CambiosTemporales
        Case "rartbal"
            Informe_Articulos_Balanzas
        Case "rteclas"
            Informe_Teclas P1
        Case "rrechazos"
            Informe_Rechazos
        Case "gtarti"
            TiendaActual = Val(P1)
            FrmTotales.Show
            DoEvents
            FrmTotales.CmbTipo.ListIndex = 1
            Do Until Fin_Totales
                Do_Events
            Loop
        Case "gtvend"
            TiendaActual = Val(P1)
            FrmTotales.Show
            DoEvents
            FrmTotales.CmbTipo.ListIndex = 2
            Do Until Fin_Totales
                Do_Events
            Loop
        Case "gtsecs"
            TiendaActual = Val(P1)
            FrmTotales.Show
            DoEvents
            FrmTotales.CmbTipo.ListIndex = 0
            Do Until Fin_Totales
                Do_Events
            Loop
        Case "pedido"
            Informe_Pedido P1
        Case "inventario"
            Informe_Inventario P1
        Case "pedinv"
            Informe_PedidoInventario P1
        Case "backupmulti"
            Backup_path = P1
            MostrarModal FrmBackup
        Case "hazbackupmulti"
            Backup_path = P1
            FrmBackup.Realizar_Backup
        Case "listado_tiendas"
            Lista_Las_Tiendas
        Case "rclientes"
            informe_clientes
        Case "gtinfven"
            informe_control_vendedores
        Case "gtinfdes"
            informe_descuentos
        Case "gtinfhora"
            informe_horario
        Case Else
            MsgBox "Este programa sólo puede ser ejecutado desde HydraTouch", vbCritical
    End Select
    End
errorf:
    MsgBox "Este programa sólo puede ser ejecutado desde HydraTouch", vbCritical
    End
End Sub
Private Sub Check_Especiales()
    If Dir(MiruTa & "\sabeco.sbc") <> "" Then Sabeco = True
    If Dir(MiruTa & "\ahorramas.cfg") <> "" Then AhorraMas = True
    If Dir(MiruTa & "\pathorde.ini") <> "" Then Alcampo = True
End Sub
Public Sub MostrarModal(Formulario As Form)
    Descarga = False
    Formulario.Show
    Do Until Descarga
        Do_Events
    Loop
End Sub
Public Sub MostrarModal2()
    Do Until Descarga
        Do_Events
    Loop
End Sub

Attribute VB_Name = "LineaComandos"
Option Explicit
Public Type TipoDefinicion
    texto As String
    accion As String
    Adicional As String
End Type
Public FicheroImportacion As String
Public Sub CrearParametrosDefecto()
    Dim Fichero As Integer
    Dim MyNum As String
    If (HayMulti Or Dir(App.Path & "\hydramulti.exe") <> "") Then
        MyNum = "###"
    Else
        MyNum = "01"
    End If
    Fichero = FreeFile
    Open Miruta & "\cmdlinetouch.cfg" For Output As Fichero
    Print #Fichero, "/" & MyNum & "U"
    Print #Fichero, "1   Envío de Base Maestra"
    Print #Fichero, "N"
    Print #Fichero, "/" & MyNum & "UZ"
    Print #Fichero, "1   Envío de Base Maestra"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "T"
    Print #Fichero, "2   Envío de Modificaciones"
    Print #Fichero, "N"
    Print #Fichero, "/" & MyNum & "TZ"
    Print #Fichero, "2   Envío de Modificaciones"
    Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "SECZ"
    'Print #Fichero, "3   Envío de Secciones "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "SUBZ"
    'Print #Fichero, "4   Envío de SubSecciones "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "FAMZ"
    'Print #Fichero, "5   Envío de Familias "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "EQPZ"
    'Print #Fichero, "6   Envío de Equipos "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "PAIZ"
    'Print #Fichero, "7   Envío de Países "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "VAC"
    'Print #Fichero, "8   Envío de Fichas de Vacuno "
    'Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "ARTZ"
    Print #Fichero, "3   Envío de Artículos "
    Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "PUBZ"
    'Print #Fichero, "10  Envío de Publicidad "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "CABLEYZ"
    'Print #Fichero, "11  Envío de Cabeceras y Leyendas "
    'Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "CODBARZ"
    Print #Fichero, "4  Envío de Códigos de Barras "
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "VENDZ"
    Print #Fichero, "5  Envío de Vendedores "
    Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "TECRAPZ"
    'Print #Fichero, "14  Envío de Teclas Rápidas "
    'Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "S"
    Print #Fichero, "6  Recepción de Grandes Totales "
    Print #Fichero, "N"
    Print #Fichero, "/" & MyNum & "SZ"
    Print #Fichero, "6  Recepción de Grandes Totales "
    Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "H"
    'Print #Fichero, "15  Recepción de Grandes Totales "
    'Print #Fichero, "N"
    'Print #Fichero, "/" & MyNum & "HZ"
    'Print #Fichero, "15  Recepción de Grandes Totales "
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "V"
    'Print #Fichero, "16  Recepción de Fichas de Vacuno "
    'Print #Fichero, "Y"
    'If HayMulti Then
    '    Print #Fichero, "/" & MyNum & "VAC"
    'Else
    '    Print #Fichero, "/VAC"
    'End If
    'Print #Fichero, "18  Importar Fichas de Vacuno"
    'Print #Fichero, "Y"
    'If Dir(App.Path & "\gigante.txt") <> "" Then
    '    Print #Fichero, "/" & MyNum & "FZ"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Yfart.dat"
    '    Print #Fichero, "/" & MyNum & "F"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Nfart.dat"
    '    Print #Fichero, "/" & MyNum & "AZ"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Yfart.dat"
    '    Print #Fichero, "/" & MyNum & "A"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Nfart.dat"
    '    Print #Fichero, "/" & MyNum & "GZ"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Yfart.dat"
    '    Print #Fichero, "/" & MyNum & "G"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Nfart.dat"
    '    Print #Fichero, "/" & MyNum & "JZ"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Yfart.dat"
    '    Print #Fichero, "/" & MyNum & "J"
    '    Print #Fichero, "17  Importar Artículos"
    '    Print #Fichero, "Nfart.dat"
    
    'Else
        Print #Fichero, "/" & MyNum & "FZ"
        Print #Fichero, "7  Importar Artículos"
        Print #Fichero, "Yfampos.dat\artcon.dat\venpos.dat\clientes.dat\bal.dat\clientesx.dat"
        'Print #Fichero, "/" & MyNum & "F"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Nbal.dat\balenv.dat\ballss.dat"
        'Print #Fichero, "/" & MyNum & "AZ"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Ybal.dat"
        'Print #Fichero, "/" & MyNum & "A"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Nbal.dat"
        'Print #Fichero, "/" & MyNum & "GZ"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Ybalenv.dat"
        'Print #Fichero, "/" & MyNum & "G"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Nbalenv.dat"
        'Print #Fichero, "/" & MyNum & "JZ"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Yballss.dat"
        'Print #Fichero, "/" & MyNum & "J"
        'Print #Fichero, "17  Importar Artículos"
        'Print #Fichero, "Nballss.dat"
    'End If
    Print #Fichero, "/MOS" & MyNum
    Print #Fichero, "8  Mostrar Pantalla de Mostradores y Equipos"
    Print #Fichero, "Y"
    Print #Fichero, "/AR" & MyNum
    Print #Fichero, "9  Mostrar Pantalla de Configuración de Artículos"
    Print #Fichero, "Y"
    Print #Fichero, "/VN" & MyNum
    Print #Fichero, "10  Mostrar Pantalla de Configuración de Vendedores"
    Print #Fichero, "Y"
    'Print #Fichero, "/PU" & MyNum
    'Print #Fichero, "23  Mostrar Pantalla de Configuración de Publicidad"
    'Print #Fichero, "Y"
    'Print #Fichero, "/CL" & MyNum
    'Print #Fichero, "24  Mostrar Pantalla de Configuración de Cabeceras y Leyendas"
    'Print #Fichero, "Y"
    Print #Fichero, "/CB" & MyNum
    Print #Fichero, "11  Mostrar Pantalla de Configuración de Códigos de Barras"
    Print #Fichero, "Y"
    'Print #Fichero, "/TE" & MyNum
    'Print #Fichero, "26  Mostrar Pantalla de Configuración de Teclas"
    'Print #Fichero, "Y"
    'Print #Fichero, "/GCE" & MyNum
    'Print #Fichero, "27  Mostrar Pantalla de Configuración de G. Conservación Euroscale"
    'Print #Fichero, "Y"
    'Print #Fichero, "/GCV" & MyNum
    'Print #Fichero, "28  Mostrar Pantalla de Configuración de G. Conservación SC10"
    'Print #Fichero, "Y"
    'Print #Fichero, "/PA" & MyNum
    'Print #Fichero, "29  Mostrar Pantalla de Configuración de Países"
    'Print #Fichero, "Y"
    'Print #Fichero, "/FV" & MyNum
    'Print #Fichero, "30  Mostrar Pantalla de Configuración de Fichas de Vacuno"
    'Print #Fichero, "Y"
    Print #Fichero, "/CP" & MyNum
    Print #Fichero, "12  Mostrar Pantalla de Cambios de Precio"
    Print #Fichero, "Y"
    Print #Fichero, "/TI" & MyNum
    Print #Fichero, "13  Mostrar Pantalla de Tipos de IVA"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "IVAZ"
    Print #Fichero, "14  Enviar Tipos de IVA"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "EZ"
    Print #Fichero, "15  Importar Fichero + Envío Modificaciones"
    Print #Fichero, "Yfampos.dat\artcon.dat\venpos.dat\clientes.dat\bal.dat\clientesx.dat"
    'Print #Fichero, "/" & MyNum & "HOR"
    'Print #Fichero, "35  Programación de Fecha y Hora"
    'Print #Fichero, "N"
    Print #Fichero, "/CREA" & MyNum
    Print #Fichero, "16  Importa datos de balanza y genera fichero"
    Print #Fichero, "Yartcon.dat"
    'If HayMulti Then
        Print #Fichero, "/BACKUP"
        Print #Fichero, "17  Backup Multitienda"
        Print #Fichero, "Y"
    'End If
    'Print #Fichero, "/" & MyNum & "PEDINV"
    'Print #Fichero, "41  Solicitar Pedido + Inventario"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "PED"
    'Print #Fichero, "42  Solicitar Pedido"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "INV"
    'Print #Fichero, "43  Solicitar Inventario"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "GLOBAL"
    'Print #Fichero, "44  Importar GLOBAL.DAT"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "GLOBAL2"
    'Print #Fichero, "45  Importar GLOBAL.DAT + envío modificaciones"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/PTTE" & MyNum
    'Print #Fichero, "46  Mostrar Pantalla de Tipos de Tara Marte IC-ILC"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/PTLE" & MyNum
    'Print #Fichero, "47  Mostrar Pantalla de Textos Marte IC-ILC"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "TTEZ"
    'Print #Fichero, "48  Envío de Tipos de Tara Marte IC-ILC"
    'Print #Fichero, "Y"
    
    'Print #Fichero, "/" & MyNum & "TLEZ"
    'Print #Fichero, "49  Envío de Textos Marte IC-ILC"
    'Print #Fichero, "Y"
    
    Print #Fichero, "/" & MyNum & "UPDATE"
    Print #Fichero, "18  Recoger totales + envío de modificaciones"
    Print #Fichero, "Y"
    
    Print #Fichero, "/" & MyNum & "PEEKTOT"
    Print #Fichero, "19  Recoger totales sin borrado"
    Print #Fichero, "Y"
    'c2f invicta b.totales
    Print #Fichero, "/" & MyNum & "ERASETOT"
    Print #Fichero, "20  ERASE TOTALS TABLES"
    Print #Fichero, "Y"
    'c2f importar Clientes.dat
    Print #Fichero, "/" & "CLIENTES"
    Print #Fichero, "21  Importar CLIENTES.DAT"
    Print #Fichero, "Y"
    Print #Fichero, "/" & "FAMPOS"
    Print #Fichero, "22  Importar FAMPOS.DAT"
    Print #Fichero, "Y"
    Print #Fichero, "/" & "VENPOS"
    Print #Fichero, "23  Importar VENPOS.DAT"
    Print #Fichero, "Y"
    
    'c2f exportar totales.dat y convertirlo para Infoservice a través de fedeiv8.exe/fede386.exe
    Print #Fichero, "/" & MyNum & "TOTDAT1"
    Print #Fichero, "24  Exportar Totales.dat (no exportados)"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "TOTDAT2"
    Print #Fichero, "25  Exportar Totales.dat (todos)"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "END_DAY"
    Print #Fichero, "26  Exportar End_Day.ini (todos)"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "KILLSRV"
    Print #Fichero, "27  Borrar Datos Balanza"
    Print #Fichero, "Y"
    Print #Fichero, "/" & MyNum & "END_DAY1"
    Print #Fichero, "28  Exportar End_Day.ini (No Exportados)"
    Print #Fichero, "Y"
   
    Print #Fichero, "/" & MyNum & "GLOBAL"
    Print #Fichero, "29  Importar GLOBAL.DAT"
    Print #Fichero, "Y"
    
    Print #Fichero, "/" & MyNum & "GLOBAL2"
    Print #Fichero, "30  Importar GLOBAL.DAT + envío modificaciones"
    Print #Fichero, "Y"
    
    Print #Fichero, "/" & MyNum & "PED"
    Print #Fichero, "31  Solicitar Pedido"
    Print #Fichero, "Y"
    
    Print #Fichero, "/" & MyNum & "INV"
    Print #Fichero, "32  Solicitar Inventario"
    Print #Fichero, "Y"
    
    Print #Fichero, "/TE" & MyNum
    Print #Fichero, "33  Mostrar Pantalla de Configuración de Teclas"
    Print #Fichero, "Y"
    
    'c2f importar taras.dat
    'Print #Fichero, "/" & "TARAS"
    'Print #Fichero, "57  Importar TARAS.DAT"
    'Print #Fichero, "Y"
    'c2f importar Tramos.dat
    Print #Fichero, "/" & MyNum & "TRAMOS"
    Print #Fichero, "58  Importar TRAMOS.DAT"
    Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "INVBOR"
    'Print #Fichero, "59  Recoge y borra INVENTARIO"
    'Print #Fichero, "Y"
    'Print #Fichero, "/" & MyNum & "ADICIONALZ"
    'Print #Fichero, "60  Importar Fichero Adicional + Envío Modificaciones"
    'Print #Fichero, "Yadicional.dat"
    
    Close #Fichero
End Sub


Public Sub Marcar_cap(ByVal sInicio As String, ByVal sFin As String)
Dim Resp As Integer
Dim cm As String

    If IsDate(sInicio) And IsDate(sFin) Then
        CadenadeLog "Atención se van a marcar en Balanza los Tiques para ser capturados de nuevo..."
        'If Resp = vbYes Then
            'Resp = MsgBox("Confirme de nuevo esta acción...", vbOKCancel + vbDefaultButton2 + vbExclamation)
            'If Resp = vbOK Then
                'Call Absorver_Estructura
                frmControl.Switch_mode 0
                If EstadoBal = 0 Then
                    If CDate(sInicio) <= CDate(sFin) Then
                        cm = "update pcscale.htickets set captured='',CapturedDateTime='1970-01-01 00:00:00' where substr(dattim,1,10) >= '" & Format(sInicio, "yyyy-mm-dd") & "' and substr(dattim,1,10) <= '" & Format(sFin, "yyyy-mm-dd") & "'"
                    Else
                        cm = "update pcscale.htickets set captured='',CapturedDateTime='1970-01-01 00:00:00' where substr(dattim,1,10) >= '" & Format(sFin, "yyyy-mm-dd") & "' and substr(dattim,1,10) <= '" & Format(sInicio, "yyyy-mm-dd") & "'"
                    End If
                    On Error Resume Next
                    MyConnObj.Execute cm
                    If Err.Number = 0 Then
                        CadenadeLog "Tiques marcados en balanza. Reciba tiques..."
                    Else
                        On Error GoTo 0
                        CadenadeLog "No ha sido posible el marcado de tiques para su recepción..."
                    End If
                    On Error GoTo 0
                Else
                    CadenadeLog "No hay conexión con Balanza..."
                End If
            'End If
        'End If
    Else
        CadenadeLog "Fechas erroneas..."
    End If

End Sub

Public Sub LeerLineaComandos(Optional ByVal sSs As String)
    Dim archerror As Integer
    Dim ErrorEnvio As Integer
    Dim ErrorRecibe As Integer
    Dim Buffer As String
    Dim Buffer2 As String
    Dim Fichero As String
    Dim accion As Integer
    Dim eNCONTRADO As Boolean
    Dim Definiciones() As TipoDefinicion
    Dim CuentaDefiniciones As Integer
    Dim bucle As Integer
    Dim b As Integer
    Dim BufDef As String
    Dim BufDef2 As String
    Dim BufLit As String
    Dim MiAux As Integer
    Dim Arch As Integer
    Dim BufQuestion As String
    Dim BufMulti1 As String, BufMulti2 As String
    Dim nMyFich As Integer
    Dim sMiFecha As String
    Dim myRec As dao.Recordset
    Dim Mybase As dao.Database
    Dim sPort, sVelo, sBal As String
    Dim MiProceso, MiResultado, nnCont As Long
    Dim nf As Integer
    Dim MyP As String
    
    If Trim(Command) <> "" Then
        Form1.Text1.TexT = "HydraTouch --> Proceso:" & Trim(Command)
        Form1.Show
    End If
    If Not IsNull(sSs) Then
        'If Trim(sss) <> "" Then
        '    Form1.Text1.TexT = "HydraTouch --> Proceso:" & Trim(sss)
        '    Form1.Show
        'End If
    End If
    If Not HayMulti Then
        If Dir(Miruta & "\enverr.dat") <> "" Then Kill Miruta & "\enverr.dat"
        If Dir(Miruta & "\envok.dat") <> "" Then Kill Miruta & "\envok.dat"
        If Dir(Miruta & "\recerr.dat") <> "" Then Kill Miruta & "\recerr.dat"
        If Dir(Miruta & "\recok.dat") <> "" Then Kill Miruta & "\recok.dat"
    Else
        On Error Resume Next
        If TiendaActual < 100 Then
            If Dir(App.Path & "\T" & Format(TiendaActual, "00") & "\enverr.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "00") & "\enverr.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "00") & "\envok.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "00") & "\envok.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "00") & "\recerr.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "00") & "\recerr.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "00") & "\recok.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "00") & "\rekok.dat"
        Else
            If Dir(App.Path & "\T" & Format(TiendaActual, "000") & "\enverr.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "000") & "\enverr.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "000") & "\envok.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "000") & "\envok.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "000") & "\recerr.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "000") & "\recerr.dat"
            If Dir(App.Path & "\T" & Format(TiendaActual, "000") & "\recok.dat") <> "" Then Kill App.Path & "\T" & Format(TiendaActual, "000") & "\rekok.dat"
        End If
        On Error GoTo 0
    End If
    'temporal hasta añadir seguridad
    admin = True
    '**************
    If (Dir(Miruta & "\cmdlinetouch.cfg") = "") And (Not HaySeguridad) Then
       CrearParametrosDefecto
    End If
    Buffer = Trim(Command)
    If Not IsNull(sSs) Then
        If Trim(sSs) <> "" Then
            Buffer = Trim(sSs)
        End If
    End If
    'c2f iv4 ic+
    
    '2.0.32 --> Mipel...
    If (UCase(Mid(Buffer, 1, 7)) = "ARTKILL") And Len(Buffer) = 8 Then
        If Right(Buffer, 1) <> "1" Then
            AN_BorrarBases 2, False
        Else
            AN_BorrarBases 2, True
        End If
        Unload Form1
        salir_programa
    
    End If
    'transferencia de datos entre tiendas
    If (UCase(Mid(Buffer, 1, 5)) = "TRANS") And Len(Buffer) = 8 And IsNumeric(Val(Mid(Buffer, 6, 3))) Then
        'Call LeerParametrosInicio
        CadenadeLog "TransferShop:" & Mid(Buffer, 6, 3)
        transferSHOP (Val(Mid(Buffer, 6, 3)))
        Unload Form1
        salir_programa
    End If
    ''''''''''''''''--> Documentar parámetros directos nuevos...
    '2.0.33 --> Exa --> Ver total de artículos... Solo MonoTienda.
    If (UCase(Buffer) = "/TOTART") Then
        frmEpelsa.Enabled = False
        Aux_Ejecuta "gtarti", "0"
        frmEpelsa.Enabled = True
        Unload Form1
        salir_programa
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'If UCase(Mid(Buffer, 1, 8)) = "/ENVLOGO" And Len(Buffer) >= 15 Then
    '    sPort = Mid(Buffer, 9, 1)
    '    sBal = Mid(Buffer, 10, 2)
    '    sVelo = Trim(Mid(Buffer, 12))
    '    If Dir(App.Path & "\winme") <> "" Then
    '        FileCopy App.Path & "\winme", App.Path & "\winmed.exe"
    '        Sleep (1000)
    '        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("winmed 0" & " " & App.Path & "\logov4.bmp" & " " & sPort & " " & sVelo & " " & sBal & " 2 1 0 1 1 1 0 0 0 0 0", vbNormalFocus))
    '        Do While MiResultado = STILL_ACTIVE
    '            GetExitCodeProcess MiProceso, MiResultado
    '            Do_Events
    '        Loop
    '        nnCont = 0
    '        Do While nnCont < 5
    '            Sleep (1000)
    '            If Dir(App.Path & "\winmed.exe") <> "" Then
    '                On Error Resume Next
    '                Kill App.Path & "\winmed.exe"
    '                On Error GoTo 0
    '            Else
    '                Exit Do
    '            End If
    '            nnCont = nnCont + 1
    '        Loop
    '    End If
    '
    '    salir_programa
   '
    'End If
    'gescoWin
    'If UCase(Mid(Buffer, 1, 6)) = "/GESCO" And Len(Buffer) = 14 Then
    '    sFechaGesco = Mid(Buffer, 7)
    '    FrmGesco.Show
    '    Do While Dir(App.Path & "\gescofin") = ""
    '        Do_Events
    '    Loop
    '    If Dir(App.Path & "\gescofin") <> "" Then
    '        Kill App.Path & "\gescofin"
    '    End If
    '    End
    'End If
   
    'para 2.0.a Automatización consulta al sistema
    ' /CONSULTAXY  --> X=Sistema(0=EuroScale, 1=Sc10, 3=MarteIC-ILC, 4=IV4/LCD) (debe estar de acuerdo con configuración)
    '                  Y=Completa o no (0=Completa, 1=No)
    'If UCase(Mid(Buffer, 1, 9)) = "/CONSULTA" And Len(Buffer) = 11 Then
    '    nF = FreeFile()
    '    Open App.Path & "\consulta.xy" For Output As #nF
    '    Print #nF, Mid(Buffer, 10, 2)
    '    Close #nF
    '    'FrmConsultaArt.Show
    '    Do_Events
    '    Do While Dir(App.Path & "\consulta.fin") = ""
    '        Do_Events
    '    Loop
    '    If Dir(App.Path & "\consulta.xy") <> "" Then Kill App.Path & "\consulta.xy"
    '    If Dir(App.Path & "\consulta.fin") <> "" Then Kill App.Path & "\consulta.fin"
    '    salir_programa
    'End If
    
    'Buffer = "/RECIBE" 'ojo quitar
    If UCase(Mid(Command, 4, 6)) = "TOTDAT" And ((Len(Trim(Command))) > 10) Then
        sMiFecha = Mid(Command, 11, 8)
        sMiFecha = Mid(sMiFecha, 4, 2) & "/" & Mid(sMiFecha, 1, 2) & "/" & Mid(sMiFecha, 7, 2)
        Buffer = Mid(Command, 1, 10)
    End If
    'c2f para importación clientes.dat ... Documentar.
    If UCase(Buffer) = "/CLIENTES" Or UCase(Buffer) = "/CLIENTESX" Or UCase(Buffer) = "/TARAS" Or UCase(Buffer) = "/FAMPOS" Or UCase(Buffer) = "/VENPOS" Then
        Buffer = Mid(Buffer, 2) & ".dat"
        frmControl.ini_import
        frmControl.ControlImport1.FamFijas = FamiliasFijas
        frmControl.ControlImport1.SubFijas = SubSeccionesFijas
        frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
        frmControl.ControlImport1.PathBase = Base_General
        frmControl.ControlImport1.PathFile = Miruta
        frmControl.ControlImport1.PathNTQ = Miruta
        frmControl.ControlImport1.ActivarLog = DebugActivo
        frmControl.ControlImport1.MostrarFormulario = True
        frmControl.ControlImport1.Fichero = Buffer
        frmControl.ControlImport1.FamVacuno = FamiliaEspecial
        frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
        frmControl.ControlImport1.MuestraMensajes = False
        frmControl.ControlImport1.FactorEuro = ValorEuro
        frmControl.ControlImport1.Importar_Background Buffer
        Do Until frmControl.ControlImport1.Exito <> 4
            Do_Events
        Loop
        Unload frmControl
        ErrorEnvio = 1
        
        Corregir_Nulos
        
        AN_BaseMaestra True
        
        If canCelar = 1 Then ErrorEnvio = 2
        On Error Resume Next
        If ErrorEnvio <> 0 Then
            archerror = FreeFile()
            If HayMulti Then
                If ErrorEnvio = 1 Then
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\envok.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\envok.dat" For Output As #archerror
                    End If
                Else
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\enverr.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\enverr.dat" For Output As #archerror
                    End If
                End If
            Else
                If ErrorEnvio = 1 Then
                    Open App.Path & "\envok.dat" For Output As #archerror
                Else
                    Open App.Path & "\enverr.dat" For Output As #archerror
                End If
            End If
            Close #archerror
        End If
        Unload Form1
        If Dir(App.Path & "\asign.tst") <> "" Then
            Call AsignTec
        End If
        
        salir_programa
    End If
    
    
    'c2f para exportacion CABTQT.SDF,TIQUET.SDF DASA/ASYSGON ... Documentar.
    ' COMPATIBILIDAD UNISCALE...
    If UCase(Buffer) = "/RECIBE" Then
        CadenadeLog "Recepción de totales..."
        FrmExportar.AN_RecogerTotales False
        CadenadeLog "tiques recibidos..."
        CadenadeLog "Exportar Ficheros..."
        export_CABLIN False, UsaGamaAlta
        CadenadeLog "Exportados..."
        Unload Form1
        salir_programa
    End If
    If UCase(Buffer) = "/RECIBEFIN" Then
        FrmExportar.AN_RecogerTotales True
        export_CABLIN False, UsaGamaAlta
        Unload Form1
        salir_programa
    End If
    If UCase(Buffer) = "/BOKACON" Then 'boka.txt
        FrmExportar.AN_RecogerTotales False
        Trata_Boka
        Unload Form1
        salir_programa
    End If
    If UCase(Buffer) = "/BOKAFIN" Then 'boka.txt
        FrmExportar.AN_RecogerTotales True
        Trata_Boka
        Unload Form1
        salir_programa
    End If
    If UCase(Buffer) = "/SINCRONIZA" Then
        CadenadeLog "parametro SINCRONIZA"
        Call FrmExportar.KILLSRV
        frmEpelsa.mnuBaseMaestra_Click
        salir_programa
    End If
    If UCase(Mid(Buffer, 5)) = "KILLTRA" Or UCase(Mid(Buffer, 2)) = "KILLTRA" Then
        Set Mybase = AbrirBase
            Mybase.Execute "delete * from cabecera where typtic='12'"
            Mybase.Execute "delete * from tickets where typtic='12'"
            Mybase.Execute "delete * from gtsecs where modo=12"
            Mybase.Execute "delete * from gtarti where modo=12"
            Mybase.Execute "delete * from gtinfven where modo=12"
            Mybase.Execute "delete * from gtvend where modo=12"
            Mybase.Execute "delete * from gthora"
        'End If
        Mybase.Close
        Set Mybase = Nothing
        salir_programa
    End If
    '"yyyy-mm-dd"            /001NULCAPyyyy-mm-ddyyyy-mm-dd
    If UCase(Mid(Buffer, 5)) = "NULCAP" And Len(Buffer) = 30 Then
        Marcar_cap Mid(Buffer, 11, 10), Mid(Buffer, 21, 10)
        salir_programa
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    CuentaDefiniciones = 1
    If Buffer = "" Then Exit Sub
    
       '2.0.37 adaptación parámetro a 3 dígitos tienda...
       If left(Buffer, 1) = "/" And IsNumeric(Mid(Buffer, 2, 1)) Then
           If IsNumeric(Mid(Buffer, 3, 1)) Then
               CadenadeLog "Adaptación parámetro:" & Buffer
               If Not IsNumeric(Mid(Buffer, 4, 1)) Then
                   Buffer = "/0" & Mid(Buffer, 2)
               End If
               CadenadeLog "               Queda:" & Buffer
           End If
           
       End If
       ''''''''''''''''''''''''''''''''''''''''''''''''''
       Fichero = FreeFile()
       Open App.Path & "\cmdlinetouch.cfg" For Input As Fichero
       If Not EOF(Fichero) Then
            Do Until EOF(Fichero)
                ReDim Preserve Definiciones(CuentaDefiniciones + 1)
                Line Input #Fichero, Definiciones(CuentaDefiniciones).texto
                Line Input #Fichero, Definiciones(CuentaDefiniciones).accion
                Line Input #Fichero, Definiciones(CuentaDefiniciones).Adicional
                If Mid(Definiciones(CuentaDefiniciones).texto, 1, 1) = "/" And IsNumeric(Mid(Definiciones(CuentaDefiniciones).texto, 2, 1)) Then
                    If IsNumeric(Mid(Definiciones(CuentaDefiniciones).texto, 3, 1)) Then
                        If Not IsNumeric(Mid(Definiciones(CuentaDefiniciones).texto, 4, 1)) Then
                            Definiciones(CuentaDefiniciones).texto = "/0" & Mid(Definiciones(CuentaDefiniciones).texto, 2)
                        End If
                    End If
                End If
                CuentaDefiniciones = CuentaDefiniciones + 1
            Loop
        End If
        Close #Fichero
        If Mid(Buffer, 1, 1) = "/" Then
            Buffer2 = "/"
            For bucle = 2 To Len(Buffer)
                If Mid(Buffer, bucle, 1) = " " Or Mid(Buffer, bucle, 1) = "/" Then
                    Exit For
                Else
                    Buffer2 = Buffer2 & Mid(Buffer, bucle, 1)
                End If
            Next bucle
        End If
        Buffer2 = StrConv(Buffer2, vbUpperCase)
        eNCONTRADO = False
        For bucle = 1 To CuentaDefiniciones
            If InStr(1, Definiciones(bucle).texto, "?") > 0 Then
                BufDef = ""
                BufQuestion = Definiciones(bucle).texto
                For b = 1 To Len(Definiciones(bucle).texto)
                    If b <= Len(Buffer2) Then
                        If Mid(Definiciones(bucle).texto, b, 1) = "?" Then
                            BufDef = BufDef & Mid(Buffer2, b, 1)
                        Else
                            BufDef = BufDef & Mid(Definiciones(bucle).texto, b, 1)
                        End If
                    End If
                Next b
            Else
                BufDef = Definiciones(bucle).texto
            End If
            If HayMulti Then
                If InStr(1, BufDef, "#") > 0 Then
                    BufDef2 = ""
                    For b = 1 To Len(BufDef)
                        If b <= Len(Buffer2) Then
                            If Mid(BufDef, b, 1) = "#" Then
                                BufDef2 = BufDef2 & Mid(Buffer2, b, 1)
                            Else
                                BufDef2 = BufDef2 & Mid(BufDef, b, 1)
                            End If
                        End If
                    Next b
                Else
                    BufDef2 = BufDef
                End If
            Else
                BufDef2 = BufDef
            End If
            'CadenadeLog "Buffer2:" & Buffer2 & "  Bufdef2:" & BufDef2
            If Buffer2 = BufDef2 Then
                accion = Val(left(Definiciones(bucle).accion, 3))
                eNCONTRADO = True
                Exit For
            End If
        Next bucle
        If eNCONTRADO = False Then
            MsgBox CargaCadena(909)
            salir_programa
        End If
        If HaySeguridad Then
            UsuarioActual.Nombre = "cmdline"
            UsuarioActual.password = "cmdline"
            Secure_Login UsuarioActual
        Else
            UsuarioActual.Nombre = "admin"
        End If
        HayComandos = True
        
        'If Dir(App.Path & "\ahorramas.cfg") <> "" And accion <> 51 Then
        '   nMyFich = FreeFile()
        '   Open App.Path & "\hydrarun.txt" For Output As #nMyFich
        '   Close #nMyFich
        'End If
        CadenadeLog "Acción:" & CStr(accion)
        Select Case accion
            Case 1
                ErrorEnvio = 1
                HayComandos = True
                If AhorraMas Then HayComandos = False
                frmEpelsa.mnuBaseMaestra_Click
                If canCelar = 1 Then ErrorEnvio = 2
            Case 2
                ErrorEnvio = 1
                HayComandos = True
                frmEpelsa.mnuEnvio_modificaciones_Click
                If canCelar = 1 Then ErrorEnvio = 2
            'Case 3
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnuEnviarSecciones_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 4
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnuenviarSubsecciones_Click
            '    'If canCelar = 1 Then ErrorEnvio = 2
            'Case 5
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnuenviarFamilias_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 6
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnu_enviarEquipos_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 7
            '    ErrorEnvio = 1
            '    frmEpelsa.mnuEnviarPaises_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 8
            '    ErrorEnvio = 1
            '    frmEpelsa.mnuEnviarFichasVacuno_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            Case 3
                ErrorEnvio = 1
                HayComandos = False
                frmEpelsa.mnu_envio_articulos_Click
                If canCelar = 1 Then ErrorEnvio = 2
                HayComandos = True
            'Case 10
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnu_envio_publicidad_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 11
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnu_envio_cabley_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            Case 4
                ErrorEnvio = 1
                frmEpelsa.mnu_envio_codigobar_Click
                If canCelar = 1 Then ErrorEnvio = 2
            Case 5
                ErrorEnvio = 1
                frmEpelsa.mnu_envio_vendedor_Click
                If canCelar = 1 Then ErrorEnvio = 2
            'Case 14
            '    ErrorEnvio = 1
            '    'frmEpelsa.mnu_envio_teclas_rapidas_Click
            '    If canCelar = 1 Then ErrorEnvio = 2
            Case 6
                ErrorRecibe = 1
                frmEpelsa.mnuGrandes_totales_Click
If Dir(App.Path & "\softandservice.txt") <> "" Then
                export_BL True
                salir_programa
End If
                If canCelar = 1 Then ErrorRecibe = 2
            'Case 16
            '    ErrorRecibe = 1
            '    frmEpelsa.mnuTotalFichaVAcuno_Click
            '    If canCelar = 1 Then ErrorRecibe = 2
            Case 7
                CadenadeLog "Importación:" & Definiciones(bucle).Adicional
                'MsgBox "quitar"
                RealizarImportacion (Definiciones(bucle).Adicional)
            'Case 8
            '    frmControl.ini_import
            '    frmControl.ControlImport1.FamFijas = FamiliasFijas
            '    frmControl.ControlImport1.SubFijas = SubSeccionesFijas
            '    frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
            '    frmControl.ControlImport1.PathBase = Base_General
            '    frmControl.ControlImport1.PathFile = MiruTa
            '    frmControl.ControlImport1.PathNTQ = MiruTa
            '    frmControl.ControlImport1.ActivarLog = DebugActivo
            '    frmControl.ControlImport1.MostrarFormulario = True
            '    frmControl.ControlImport1.Fichero = "fichas.dat"
            '    frmControl.ControlImport1.FamVacuno = FamiliaEspecial
            '    frmControl.ControlImport1.MuestraMensajes = False
            '    frmControl.ControlImport1.FactorEuro = ValorEuro
            '    frmControl.ControlImport1.Importar_Background "fichas.dat"
            '    Do Until frmControl.ControlImport1.Exito <> 4
            '        Do_Events
            '    Loop
            '    Unload frmControl
            'Case 19
            '    If Sabeco Then
            '        MostrarMonitor
            '        FrmExportar.AN_RecogerTotales False
            '    End If
            '    realizarExportacion (Definiciones(Bucle).Adicional)
            Case 8
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    Kill App.Path & "\hydrarun.txt"
                End If
                FrmSeccionyMaquina.Show vbModal
            Case 9
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    Kill App.Path & "\hydrarun.txt"
                End If
                FrmArticulo2.Show vbModal
            Case 10
                FrmVendedor2.Show vbModal
            'Case 23
            '    'FrmPublicidad2.Show vbModal
            'Case 24
            '    'FrmCabley2.Show vbModal
            Case 11
                FrmCodBar2.Show vbModal
            'Case 26
            '    FrmTeclas2T.Show vbModal
            'Case 27
            '    'FrmConserv386.Show vbModal
            'Case 28
            '    'FrmConservSC10.Show vbModal
            'Case 29
            '    FrmPaises.Show vbModal
            'Case 30
            '    frmFichasVacuno.Show vbModal
            Case 12
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    Kill App.Path & "\hydrarun.txt"
                End If
                FrmCambiosPrecio.Show vbModal
            Case 13
                If Dir(App.Path & "\hydrarun.txt") <> "" Then
                    Kill App.Path & "\hydrarun.txt"
                End If
                FrmIVA.Show vbModal
            Case 14
                ErrorEnvio = 1
                frmEpelsa.mnu_Envio_IVA_Click
                If canCelar = 1 Then ErrorEnvio = 2
            Case 15
If Dir(App.Path & "\softandservice.txt") <> "" Then
                If TiendaActual < 100 Then
                    MyP = App.Path & "\t" & Format(TiendaActual, "00") & "\"
                Else
                    MyP = App.Path & "\t" & Format(TiendaActual, "000") & "\"
                End If
                CadenadeLog "Parámetro Importar y Enviar..." & " Se realiza importación Fichero:" & MyP & "articulo.dat"
                If Dir(MyP & "articulo.dat") <> "" Then
                    
                    FileCopy MyP & "Articulo.dat", App.Path & "\articulo.dat"
                    Sleep (200)
                    CadenadeLog "Copiado:" & MyP & "articulo.dat" & " a " & App.Path & "\articulo.dat"
                    CadenadeLog "ruta:" & Miruta
                    'frmControl.ControlImport1.Trazabilidad = False
                    frmControl.ControlImport1.FamFijas = FamiliasFijas
                    frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                    frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                    frmControl.ControlImport1.PathBase = Base_General
                    frmControl.ControlImport1.PathFile = App.Path
                    frmControl.ControlImport1.PathNTQ = App.Path
                    frmControl.ControlImport1.ActivarLog = DebugActivo
                    frmControl.ControlImport1.MostrarFormulario = True
                    frmControl.ControlImport1.Fichero = "articulo.dat"
                    frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                    frmControl.ControlImport1.MuestraMensajes = False
                    frmControl.ControlImport1.FactorEuro = 1
                    frmControl.ControlImport1.Importar_Background "articulo.dat"
                    frmControl.ControlImport1.Euros = True
                    'frmControl.ControlImport1.FactorEuro = 1
                    frmControl.ControlImport1.NumeroDecimales = decimales
                    frmControl.ControlImport1.Euros = True
                    
                    Do Until frmControl.ControlImport1.Exito <> 4
                        Do_Events
                    Loop
                    Unload frmControl
                
                    CadenadeLog "Se procede a enviar modificaciones..."
                    frmEpelsa.mnuEnvio_modificaciones_Click
                    CadenadeLog "Terminado envio de modificaciones."
                    salir_programa
                
                Else
                    CadenadeLog "No existe:" & MyP & "articulo.dat"
                    salir_programa
                End If
                
Else
                ErrorEnvio = 1
                CadenadeLog "Parámetro Importar y Enviar..." & " Se realiza importación Fichero:" & Definiciones(bucle).Adicional
                RealizarImportacion (Definiciones(bucle).Adicional)
                CadenadeLog "Se procede a enviar modificaciones..."
                frmEpelsa.mnuEnvio_modificaciones_Click
                CadenadeLog "Terminado envio de modificaciones."
                If Dir(App.Path & "\sal.txt") <> "" Then
                    salir_programa
                End If

End If
            
                If canCelar = 1 Then ErrorEnvio = 2
            'Case 35
            '    Unload frmMonitor
            'Case 16
            '    ErrorRecibe = 1
            '    If AhorraMas Then
            '        Dim NnF As Integer
            '        'If UsaGamaAlta And usaSC10 Then
            '        '    FrmConsultaArt.Option1(2).Value = True
            '        'Else
            '        '    If UsaGamaAlta Then
            '        '    FrmConsultaArt.Option1(0).Value = True
            '        '    Else
            '        '    FrmConsultaArt.Option1(1).Value = True
            '        '    End If
            '        'End If
            '        NnF = FreeFile()
            '        Open App.Path & "\ARTCON.DAT" For Output As #NnF
            '        'FrmConsultaArt.ConsultaLosArticulos False, True, NnF
            '        Close #NnF
            '        If canCelar <> 0 Then
            '            ErrorRecibe = 2
            '        End If
            '        '**************
            '
            '    Else
            '        'FrmConsultaArt.Show
            '        'FrmConsultaArt.Command1(0).Caption = CargaCadena(288)
            '        'FrmConsultaArt.CmbFiles.TexT = Mid(Definiciones(Bucle).Adicional, 2)
            '        'FrmConsultaArt.ConsultaLosArticulos False
            '        'If canCelar = 0 Then
            '        '    FrmConsultaArt.CmdExportar_Click
            '        'Else
            '        '    ErrorRecibe = 2
            '        'End If
            '        '**************
            '        ' sabeco : copia a carpeta CP
            '        If Sabeco Then
            '            ChDir MiruTa
            '            Shell "findia.exe /CP"
            '        End If
            '        '**************
            '    End If
            'Case 37
            '    For b = InStr(1, BufQuestion, "?") To Len(BufQuestion)
            '        If Mid(BufQuestion, b, 1) = "?" Then
            '            BufLit = BufLit & Mid(Buffer2, b, 1)
            '        Else
            '            Exit For
            '        End If
            '    Next b
            '    If IsNumeric(BufLit) Then
            '        MostrarMonitor
            '        AN_BorrarSeccion (Val(BufLit))
            '    End If
            'Case 38
            '    For b = InStr(1, BufQuestion, "?") To Len(BufQuestion)
            '        If Mid(BufQuestion, b, 1) = "?" Then
            '            BufLit = BufLit & Mid(Buffer2, b, 1)
            '        Else
            '            Exit For
            '        End If
            '    Next b
            '    If IsNumeric(BufLit) And Len(BufLit) = 6 Then
            '        realizarExportacion (Definiciones(Bucle).Adicional), BufLit
            '    Else
            '        MsgBox CargaCadena(506), vbCritical
            '    End If
            'Case 39
            '
            '    For b = InStr(1, BufQuestion, "?") To Len(BufQuestion)
            '        If Mid(BufQuestion, b, 1) = "?" Then
            '            BufLit = BufLit & Mid(Buffer2, b, 1)
            '        Else
            '            Exit For
            '        End If
            '    Next b
            '    If IsNumeric(BufLit) And Len(BufLit) = 2 Then
            '        If left(BufLit, 1) = "0" Or left(BufLit, 1) = "1" Then
            '            If Val(Right(BufLit, 1)) <= 8 Then
            '                Select Case Val(Right(BufLit, 1))
            '                    Case 0 To 5
            '                        MiAux = Val(Right(BufLit, 1))
            '                    Case 6
            '                        MiAux = 7
            '                    Case 7
            '                        MiAux = 9
            '                    Case 8
            '                        MiAux = 24
            '                End Select
            '                If left(BufLit, 1) = "0" Then
            '                    AN_BorrarBases MiAux, False
            '                Else
            '                    AN_BorrarBases MiAux, True
            '                End If
            '            Else
            '                MsgBox CargaCadena(1016), vbCritical
            '            End If
            '        Else
            '            MsgBox CargaCadena(1016), vbCritical
            '        End If
            '    Else
            '        MsgBox CargaCadena(1016), vbCritical
            '    End If
            'Case 40
            '    If Dir(App.Path & "\hydramulti.exe") <> "" Then
            '        If Dir(App.Path & "\hydramulti.ini") <> "" Then
            '            Arch = FreeFile()
            '            Open App.Path & "\hydramulti.ini" For Input As #Arch
            '            For Bucle = 1 To 3
            '                Input #Arch, BufMulti1, BufMulti2
            '            Next Bucle
            '            Close #Arch
            '            Aux_Ejecuta "hazbackupmulti", BufMulti1
            '        End If
            '    End If
            'Case 41
            '    ErrorRecibe = 1
            '    lCogeTiquet = False
            '    'FrmExportar.AN_PedidoInventario
            '    If canCelar = 1 Then ErrorRecibe = 2
            '    lCogeTiquet = True
            Case 31
                ErrorRecibe = 1
                lCogeTiquet = False
                MUE_SolicitaPedido 0, False
                If canCelar = 1 Then ErrorRecibe = 2
                lCogeTiquet = True
            Case 32
                ErrorRecibe = 1
                lCogeTiquet = False
                MUE_SolicitaInventario 0, False
                If canCelar = 1 Then ErrorRecibe = 2
                lCogeTiquet = True
            ''Upper / Asysgon
            'Case 44
            '    Importa_Global_Dat
            'Case 45
            '    ErrorEnvio = 1
            '    HayComandos = True
            '
            '    Importa_Global_Dat
            '    'AN_BaseMaestra True
            '    frmEpelsa.mnuEnvio_modificaciones_Click
            '
            '    If canCelar = 1 Then ErrorEnvio = 2
            ''''''''''''''''''''
            'Case 46
            '    'FrmTipoTARA.Show vbModal
            'Case 47
            ''FrmTextoECO.Show vbModal
            'Case 48
            '    ErrorEnvio = 1
            '    FrmExportar.AN_TiposTara
            '    If canCelar = 1 Then ErrorEnvio = 2
            'Case 49
            '    ErrorEnvio = 1
            '    'FrmExportar.AN_TextosECO
            '    If canCelar = 1 Then ErrorEnvio = 2
            Case 18
                ErrorEnvio = 1
                MostrarMonitor
                lCogeTiquet = False
                lesUpdate = True
                FrmExportar.AN_RecogerTotales True, True
                If EstadoBal < 1000 Then AN_BaseMaestra True, True
                If canCelar = 1 Then ErrorEnvio = 2
                lCogeTiquet = True
                '1.9.0
                lesUpdate = False
                ''''''''''''''''''''''''
            Case 19 'c2f jordi...
                If Dir(App.Path & "\softandservice.txt") = "" And Taquion = False Then
                    'If Sabeco Then
                    '    MostrarMonitor
                    '    FrmExportar.AN_RecogerTotales False
                    'End If
                    realizarExportacion (Definiciones(bucle).Adicional)
                
                Else
                    ErrorRecibe = 1
                    MostrarMonitor
                    
                    If Dir(App.Path & "\softandservice.txt") <> "" Then
                        FrmExportar.AN_RecogerTotales False
                        export_BL True
                        salir_programa
                        Else
                        FrmExportar.AN_RecogerTotales False
                        
                        'c2f jordi...
                        
                        '**********************************
                        ' TAQUION : ' mira si existe(n) fichero(s) de importación, y lo(s) procesa
                        '**********************************
                        If Taquion Then
                            Call FrmExportar.TQ_ImportaFicheros
                        End If
                        '******
                        ' /Recogida
                        '************
                        ' Taquion / Storeflow, envía
                        ' los tiquets pendientes, por Socket o ASCII
                        '***********************************************
                        If (Taquion And TaquionTiquets) Or Dir(App.Path & "\ahorramas.cfg") <> "" Then
                            frmEpelsa.Enabled = False
                            frmEpelsa.LblGA.Caption = Format(Now, "hh:mm:ss") & " " & CargaCadena(905)
                            If Dir(App.Path & "\ahorramas.cfg") <> "" Then
                                FrmExportar.Taquion_EnviaTiquetsPendientes False
                            Else
                                If usaSC10 Then FrmExportar.Taquion_EnviaTiquetsPendientes False
                                If UsaGamaAlta Then FrmExportar.Taquion_EnviaTiquetsPendientes True
                            End If
                            frmEpelsa.LblGA.Caption = CargaCadena(904)
                            frmEpelsa.Enabled = True
                        End If
                    End If
                End If
                '............
                If canCelar = 1 Then ErrorRecibe = 2
            Case 20
                'c2f invicta b.tables
                ErrorEnvio = 1
                frmMantenimiento.Opttotal(0).Value = False
                lNoSi = True
                frmMantenimiento.cmdespere_Click
                lNoSi = False
                If canCelar = 1 Then ErrorEnvio = 2
            Case 21
                'importar Clientes.dat
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "fichas.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                If Dir(App.Path & "\clientes.dat") <> "" Then
                    frmControl.ControlImport1.Importar_Background "clientes.dat"
                Else
                    If Dir(App.Path & "\clientesx.dat") <> "" Then
                        frmControl.ControlImport1.Importar_Background "clientesx.dat"
                    Else
                        frmControl.ControlImport1.Importar_Background "clientes.dat"
                    End If
                End If
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
                Corregir_Nulos
            Case 22
                'importar FAMPOS.dat
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "fichas.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                frmControl.ControlImport1.Importar_Background "FAMPOS.dat"
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
                Corregir_Nulos
            Case 23
                'importar VENPOS.dat
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "fichas.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                frmControl.ControlImport1.Importar_Background "VENPOS.dat"
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
                Corregir_Nulos
            Case 24
                ' bucle para cada fecha existente
                Set Mybase = OpenDatabase(Base_General)
                Set myRec = Mybase.OpenRecordset("select distinct fecha from cabecera where reserv2<>" & Chr(34) & "*" & Chr(34))
                If Not myRec.EOF Then
                With myRec
                .MoveFirst
                Do While Not .EOF
                    sMiFecha = Trim(!Fecha)
                    CadenadeLog sMiFecha
                    sMiFecha = Mid(sMiFecha, 4, 2) & "/" & Mid(sMiFecha, 1, 2) & "/" & Mid(sMiFecha, 7, 2)

                    ExportaTotalesDat False, sMiFecha, "-"
                    .Movenext
                Loop
                End With
                End If
                myRec.Close
                Set myRec = Nothing
                Mybase.Close
                Set Mybase = Nothing
                'Workspaces(0).close
                
            Case 25
                ExportaTotalesDat False, sMiFecha, "*"
            Case 26 '2.0.2
                Call Trata_Fin_Dia
            
            Case 27
                CadenadeLog "parametro KILLSRV"
                Call FrmExportar.KILLSRV
                
            Case 28 '2.0.2
                Call Trata_Fin_Dia_1
            
            ''Upper / Asysgon
            Case 29
                Importa_Global_Dat
            Case 30
                ErrorEnvio = 1
                HayComandos = True
            
                Importa_Global_Dat
                'AN_BaseMaestra True
                frmEpelsa.mnuEnvio_modificaciones_Click
            
                If canCelar = 1 Then ErrorEnvio = 2
            ''''''''''''''''''''
            Case 33
                'FrmTeclas2T.Show vbModal
                If Dir(App.Path & "\test.tst") <> "" Then
                    FrmOrdTec.Show vbModal
                Else
                    FrmTeclas2T.Show vbModal
                End If
            Case 34
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "fichas.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                frmControl.ControlImport1.Importar_Background "fichas.dat"
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
                Corregir_Nulos
            
            'Case 57
            ''importar taras.dat
            '    frmControl.ini_import
            '    frmControl.ControlImport1.FamFijas = FamiliasFijas
            '    frmControl.ControlImport1.SubFijas = SubSeccionesFijas
            '    frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
            '    frmControl.ControlImport1.PathBase = Base_General
            '    frmControl.ControlImport1.PathFile = MiruTa
            '    frmControl.ControlImport1.PathNTQ = MiruTa
            '    frmControl.ControlImport1.ActivarLog = DebugActivo
            '    frmControl.ControlImport1.MostrarFormulario = True
            '    frmControl.ControlImport1.Fichero = "taras.dat"
            '    frmControl.ControlImport1.FamVacuno = FamiliaEspecial
            '    frmControl.ControlImport1.MuestraMensajes = False
            '    frmControl.ControlImport1.FactorEuro = ValorEuro
            '    frmControl.ControlImport1.Importar_Background "taras.dat"
            '    Do Until frmControl.ControlImport1.Exito <> 4
            '        Do_Events
            '    Loop
            '    Unload frmControl
            Case 58
            ''importar tramos.dat
                CadenadeLog "importar tramos..."
                frmControl.ini_import
                frmControl.ControlImport1.FamFijas = FamiliasFijas
                frmControl.ControlImport1.SubFijas = SubSeccionesFijas
                frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
                frmControl.ControlImport1.PathBase = Base_General
                frmControl.ControlImport1.PathFile = Miruta
                frmControl.ControlImport1.PathNTQ = Miruta
                frmControl.ControlImport1.ActivarLog = DebugActivo
                frmControl.ControlImport1.MostrarFormulario = True
                frmControl.ControlImport1.Fichero = "tramos.dat"
                frmControl.ControlImport1.FamVacuno = FamiliaEspecial
                frmControl.ControlImport1.MuestraMensajes = False
                frmControl.ControlImport1.FactorEuro = ValorEuro
                frmControl.ControlImport1.Importar_Background "tramos.dat"
                Do Until frmControl.ControlImport1.Exito <> 4
                    Do_Events
                Loop
                Unload frmControl
            'Case 59
            '    ErrorRecibe = 1
            '    lCogeTiquet = False
            '    'AN_Inventario
            '    If canCelar = 1 Then ErrorRecibe = 2
            '    lCogeTiquet = True
            'Case 60
            '    ErrorEnvio = 1
            '    CadenadeLog "Parámetro Importar Adicional y Enviar..." & " Se realiza importación..."
            '    RealizarImportacion (Definiciones(Bucle).Adicional)
            '    CadenadeLog "Se procede a enviar modificaciones..."
            '    frmEpelsa.mnuEnvio_modificaciones_Click
            '    CadenadeLog "Terminado envio de modificaciones."
            '    If canCelar = 1 Then ErrorEnvio = 2
            
            Case Else
                MsgBox CargaCadena(909), , "HydraTouch"
                Unload Form1
                salir_programa
        End Select
        '
        On Error Resume Next
        If ErrorEnvio <> 0 Then
            'c2f/caspiunza ** se deja log de comunicaciones con las tiendas en resumen.txt
            If HayMulti Then
                On Error Resume Next
                If Dir(App.Path & "\resumen.txt") = "" Then
                    archerror = FreeFile()
                    Open App.Path & "\resumen.txt" For Output As #archerror
                Else
                    If FileLen(App.Path & "\resumen.txt") > 1000000 Then
                        Kill App.Path & "\resumen.txt"
                        Open App.Path & "\resumen.txt" For Output As #archerror
                    Else
                        Open App.Path & "\resumen.txt" For Append As #archerror
                    End If
                End If
                Print #archerror, Now()
                Print #archerror, "---------------------------"
                If ErrorEnvio = 1 Then
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Envio Correcto"
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Envio Correcto"
                    End If
                Else
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Envio Erroneo"
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Envio Erroneo"
                    End If
                End If
                
                
                On Error GoTo 0
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            archerror = FreeFile()
            If HayMulti Then
                If ErrorEnvio = 1 Then
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\envok.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\envok.dat" For Output As #archerror
                    End If
                Else
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\enverr.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\enverr.dat" For Output As #archerror
                    End If
                End If
            Else
                If ErrorEnvio = 1 Then
                    Open App.Path & "\envok.dat" For Output As #archerror
                Else
                    Open App.Path & "\enverr.dat" For Output As #archerror
                End If
            End If
            Close #archerror
        End If
        If ErrorRecibe <> 0 Then
            
            'c2f/caspiunza ** se deja log de comunicaciones con las tiendas en resumen.txt
            If HayMulti Then
                On Error Resume Next
                If Dir(App.Path & "\resumen.txt") = "" Then
                    archerror = FreeFile()
                    Open App.Path & "\resumen.txt" For Output As #archerror
                Else
                    If FileLen(App.Path & "\resumen.txt") > 1000000 Then
                        Kill App.Path & "\resumen.txt"
                        archerror = FreeFile()
                        Open App.Path & "\resumen.txt" For Output As #archerror
                    Else
                        archerror = FreeFile()
                        Open App.Path & "\resumen.txt" For Append As #archerror
                    End If
                End If
                Print #archerror, Now()
                Print #archerror, "-------------------------------------------------------"
                If ErrorRecibe = 1 Then
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Recepción Correcta"
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Recepción Correcta"
                    End If
                Else
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Recepción Erronea."
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Recepción Erronea."
                    End If
                End If
                
                Close #archerror
                On Error GoTo 0
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            archerror = FreeFile()
            If HayMulti Then
                If ErrorRecibe = 1 Then
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\recok.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\recok.dat" For Output As #archerror
                    End If
                Else
                    If TiendaActual < 100 Then
                        Open App.Path & "\T" & Format(TiendaActual, "00") & "\recerr.dat" For Output As #archerror
                    Else
                        Open App.Path & "\T" & Format(TiendaActual, "000") & "\recerr.dat" For Output As #archerror
                    End If
                End If
            Else
                If ErrorRecibe = 1 Then
                    Open App.Path & "\recok.dat" For Output As #archerror
                Else
                    Open App.Path & "\recerr.dat" For Output As #archerror
                End If
            End If
            Close #archerror
        End If
        On Error GoTo 0
      
        HayComandos = False
        If frmEpelsa.Visible = True Then frmEpelsa.Visible = False
        If Mid(Definiciones(bucle).Adicional, 1, 1) = "Y" Or AhorraMas Then
            
            If Dir(App.Path & "\shoptras.txt") <> "" And Dir(App.Path & "\gigante.txt") <> "" Then
                Call transferSHOP(TiendaActual)
            End If
            Unload Form1
            salir_programa
        Else
        
            Unload Form1
            
        End If
        
End Sub
'c2f/caspiunza **
Public Sub ResumenTXT(ByVal n As Integer)
        Dim ErrorEnvio As Integer
        Dim ErrorRecibe As Integer
        Dim archerror As Integer
        
        
        ErrorEnvio = 0
        ErrorRecibe = 0
        If n > 10000 Then
            ErrorEnvio = n - 10000
        Else
            ErrorRecibe = n
        End If
        On Error Resume Next
        If ErrorEnvio <> 0 Then
            'c2f/caspiunza ** se deja log de comunicaciones con las tiendas en resumen.txt
            If HayMulti Then
                On Error Resume Next
                If Dir(App.Path & "\resumen.txt") = "" Then
                    archerror = FreeFile()
                    Open App.Path & "\resumen.txt" For Output As #archerror
                Else
                    If FileLen(App.Path & "\resumen.txt") > 1000000 Then
                        Kill App.Path & "\resumen.txt"
                        archerror = FreeFile()
                        Open App.Path & "\resumen.txt" For Output As #archerror
                    Else
                        archerror = FreeFile()
                        Open App.Path & "\resumen.txt" For Append As #archerror
                    End If
                End If
                Print #archerror, Now()
                Print #archerror, "---------------------------"
                If ErrorEnvio = 1 Then
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Envio Correcto"
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Envio Correcto"
                    End If
                Else
                    If TiendaActual < 100 Then
                        Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Envio Erroneo"
                    Else
                        Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Envio Erroneo"
                    End If
                End If
                
                Close #archerror
                On Error GoTo 0
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            archerror = FreeFile()
            If HayMulti Then
                If ErrorEnvio = 1 Then
                    If TiendaActual < 100 Then
                    Open App.Path & "\T" & Format(TiendaActual, "00") & "\envok.dat" For Output As #archerror
                    Else
                    Open App.Path & "\T" & Format(TiendaActual, "000") & "\envok.dat" For Output As #archerror
                    End If
                Else
                    If TiendaActual < 100 Then
                    Open App.Path & "\T" & Format(TiendaActual, "00") & "\enverr.dat" For Output As #archerror
                    Else
                    Open App.Path & "\T" & Format(TiendaActual, "000") & "\enverr.dat" For Output As #archerror
                    End If
                End If
            Else
                If ErrorEnvio = 1 Then
                    Open App.Path & "\envok.dat" For Output As #archerror
                Else
                    Open App.Path & "\enverr.dat" For Output As #archerror
                End If
            End If
            Close #archerror
        End If
        If ErrorRecibe <> 0 Then
            
            'c2f/caspiunza ** se deja log de comunicaciones con las tiendas en resumen.txt
            If HayMulti Then
                On Error Resume Next
                If Dir(App.Path & "\resumen.txt") = "" Then
                    archerror = FreeFile()
                    Open App.Path & "\resumen.txt" For Output As #archerror
                Else
                    If FileLen(App.Path & "\resumen.txt") > 1000000 Then
                        Kill App.Path & "\resumen.txt"
                        Open App.Path & "\resumen.txt" For Output As #archerror
                    Else
                        Open App.Path & "\resumen.txt" For Append As #archerror
                    End If
                End If
                Print #archerror, Now()
                Print #archerror, "-------------------------------------------------------"
                If ErrorRecibe = 1 Then
                    If TiendaActual < 100 Then
                    Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Recepción Correcta"
                    Else
                    Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Recepción Correcta"
                    End If
                Else
                    If TiendaActual < 100 Then
                    Print #archerror, "Tienda " & Format(TiendaActual, "00") & " : Recepción Erronea."
                    Else
                    Print #archerror, "Tienda " & Format(TiendaActual, "000") & " : Recepción Erronea."
                    End If
                End If
                
                
                On Error GoTo 0
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            archerror = FreeFile()
            If HayMulti Then
                If ErrorRecibe = 1 Then
                    If TiendaActual < 100 Then
                    Open App.Path & "\T" & Format(TiendaActual, "00") & "\recok.dat" For Output As #archerror
                    Else
                    Open App.Path & "\T" & Format(TiendaActual, "000") & "\recok.dat" For Output As #archerror
                    End If
                Else
                    If TiendaActual < 100 Then
                    Open App.Path & "\T" & Format(TiendaActual, "00") & "\recerr.dat" For Output As #archerror
                    Else
                    Open App.Path & "\T" & Format(TiendaActual, "000") & "\recerr.dat" For Output As #archerror
                    End If
                End If
            Else
                If ErrorRecibe = 1 Then
                    Open App.Path & "\recok.dat" For Output As #archerror
                Else
                    Open App.Path & "\recerr.dat" For Output As #archerror
                End If
            End If
            Close #archerror
        End If
        On Error GoTo 0

End Sub

Public Function RealizarImportacion(MiCadena As String) As Integer
    Dim Buffer As String
    Dim Cadenon As String
    Dim Fichero As String
    Dim bucle As Integer
    Dim NoCont As Boolean
    Dim nMiFich As Integer
    Dim nLcLr As Integer
    
    'If versionG Then
    '    MiCadena = MiCadena & "\baletl.dat"
    'End If
    
    'leclerc....
    CadenadeLog "RealizarImportacion.:" & CStr(lClR) & " :" & sPathMaj
    If lClR And sPathMaj <> "" Then
        nLcLr = fVerLeclerc()
        CadenadeLog "Verificar si se ha Generado BAL.DAT"
        If Dir(App.Path & "\bal.dat") <> "" Then
            CadenadeLog "Generado bal.dat con longitud:" & CStr(FileLen(App.Path & "\bal.dat"))
        End If
        
        Sleep (2000)
    End If
    '...........
    
    If MiCadena = "" And Dir(App.Path & "\multiimp.ord") <> "" Then
       nMiFich = FreeFile()
       Open App.Path & "\multiimp.ord" For Input As #nMiFich
       Input #nMiFich, MiCadena
       Close #nMiFich
       'Kill App.Path & "\multiimp.ord"
    End If
    Cadenon = Trim(Right(MiCadena, Len(MiCadena) - 1))
    If Cadenon = "" Then Exit Function
    frmControl.ini_import
    canCelar = 0
    For bucle = 1 To Len(Cadenon)
        Fichero = ""
        Do Until Mid(Cadenon, bucle, 1) = "\" Or bucle > Len(Cadenon)
            Fichero = Fichero & Mid(Cadenon, bucle, 1)
            bucle = bucle + 1
        Loop
        frmControl.ControlImport1.FamFijas = FamiliasFijas
        frmControl.ControlImport1.SubFijas = SubSeccionesFijas
        frmControl.ControlImport1.RechazarPrecioCero = RechazarPlu0
        frmControl.ControlImport1.PathBase = Base_General
        frmControl.ControlImport1.FamVacuno = FamiliaEspecial
        frmControl.ControlImport1.EtiquetaEspecial = EtiquetaEspecial
        frmControl.ControlImport1.PathFile = Miruta
        frmControl.ControlImport1.PathNTQ = Miruta
        frmControl.ControlImport1.Fichero = Fichero
        frmControl.ControlImport1.MuestraMensajes = False
        frmControl.ControlImport1.ActivarLog = LogArticulos
        frmControl.ControlImport1.MostrarFormulario = True
        frmControl.ControlImport1.Euros = UsaEuro
        frmControl.ControlImport1.FactorEuro = 1
        frmControl.ControlImport1.NumeroDecimales = decimales
        If UsaEuro Then
            frmControl.ControlImport1.Euros = True
        Else
            frmControl.ControlImport1.Euros = False
        End If
        If HayMulti Then
            If (Dir(App.Path & "\" & Format(TiendaActual, "00") & Fichero) <> "" And TiendaActual < 100) Or (Dir(App.Path & "\" & Format(TiendaActual, "000") & Fichero) <> "" And TiendaActual > 99) Then
                NoCont = True
                On Error Resume Next
                If TiendaActual < 100 Then
                FileCopy App.Path & "\etc\" & Fichero & ".fia", App.Path & "\etc\" & Format(TiendaActual, "00") & Fichero & ".fia"
                Else
                FileCopy App.Path & "\etc\" & Fichero & ".fia", App.Path & "\etc\" & Format(TiendaActual, "000") & Fichero & ".fia"
                End If
                On Error GoTo 0
                If TiendaActual < 100 Then
                frmControl.ControlImport1.Importar_Background Format(TiendaActual, "00") & Fichero
                Else
                frmControl.ControlImport1.Importar_Background Format(TiendaActual, "000") & Fichero
                End If
                Do Until frmControl.ControlImport1.Exito <> 4
                    If canCelar = 1 Then frmControl.ControlImport1.CancelarImportacion
                    Do_Events
                Loop
                On Error Resume Next
                If TiendaActual < 100 Then
                Kill App.Path & "\etc\" & Format(TiendaActual, "00") & Fichero & ".fia"
                Else
                Kill App.Path & "\etc\" & Format(TiendaActual, "000") & Fichero & ".fia"
                End If
                Corregir_Nulos
                On Error GoTo 0
            End If
        End If
        If ((Not NoCont) And (Dir(App.Path & "\" & Fichero) <> "")) Or (HayMulti = False) Then
            frmControl.ControlImport1.Importar_Background Fichero
            Do Until frmControl.ControlImport1.Exito <> 4
                If canCelar = 1 Then frmControl.ControlImport1.CancelarImportacion
                Do_Events
            Loop
            Corregir_Nulos
        End If
        On Error Resume Next
        On Error GoTo 0
        '*******************
        ' fin de proceso estandar
        '*******************
        '*******************
        ' Ahold
        '*******************
        If (AHOLD Or lClR) And (Fichero = "bal.dat") Then
            If Dir(App.Path & "\bal.dat") <> "" Then
                If lClR Then
                    FileCopy App.Path & "\bal.dat", App.Path & "\bal.bck"
                    Sleep (3000)
                End If
                On Error Resume Next
                Kill App.Path & "\bal.dat"
                On Error GoTo 0
            End If
        End If
        '*******************
        ' Sabeco
        '*******************
        'If Sabeco And UCase(Fichero) = "SABECO.DAT" Then
        'If frmControl.ControlImport1.Exito = 0 Or frmControl.ControlImport1.Exito = 1 _
        'Or frmControl.ControlImport1.Exito = 10 Then
        '    ChDir MiruTa
        '    Shell ("XMSB200 /RETURN")
        '    Sleep (3000)
        '    If Dir(MiruTa & "\balanza.res") <> "" Then
        '        If Dir(MiruTa & "\control", vbDirectory) = "" Then MkDir MiruTa & "\control"
        '        If Dir(MiruTa & "\balanza.res") <> "" Then FileCopy MiruTa & "\balanza.res", MiruTa & "\control\balanza.res"
        '    End If
        'End If
        '*******************
        'End If
    Next bucle
    
    If Dir(App.Path & "\asign.tst") <> "" Then
        Call AsignTec
    End If
    
    RealizarImportacion = frmControl.ControlImport1.Exito
    
End Function
Public Sub realizarExportacion(ByVal cadena As String, Optional Fecha As String)
    Dim Buffer As String
    Dim Cadenon As String
    Dim Fichero As String
    Dim bucle As Integer
    Cadenon = Trim(Right(cadena, Len(cadena) - 1))
    If Cadenon = "" Then Exit Sub
    For bucle = 1 To Len(Cadenon)
        Fichero = ""
        Do Until Mid(Cadenon, bucle, 1) = "\" Or bucle > Len(Cadenon)
            Fichero = Fichero & Mid(Cadenon, bucle, 1)
            bucle = bucle + 1
        Loop
        Load FrmExportar
        FrmExportar.TxtNombre.TexT = Fichero
        FrmExportar.MostrarMensajes = False
        If Fecha <> "" Then
            FrmExportar.Check1.Value = vbChecked
            FrmExportar.TxtFecha.TexT = left(Fecha, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Right(Fecha, 2)
        End If
        FrmExportar.Show
        Do_Events
        
        FrmExportar.CmdExportar_Click
    Next bucle
    'If Sabeco And Fecha = "" Then
    '    ChDir (MiruTa)
    '    Shell ("findia.exe temp temp")
    'End If
End Sub


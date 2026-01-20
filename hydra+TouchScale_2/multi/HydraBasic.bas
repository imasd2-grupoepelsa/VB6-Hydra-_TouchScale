Attribute VB_Name = "HydraBasic"
Option Explicit
Public Type Tipo_User
    Nombre As String
    Password As String
    Permisos(8) As Byte
    Asociado As Integer
End Type
Public Type tDatosConfig
    Numero As Integer
    GA_TCPIP As Boolean
    SC10_NETPATH As Boolean
    UsaModem As Boolean
    UsaGamaAlta As Boolean
    UsaSC10 As Boolean
    UsaEco1 As Boolean
    UsaEco2 As Boolean
    UsaEcoPrint As Boolean
    TipoC9 As Boolean
    PuertoGA As Long
    PuertoECO As Long
    SC10_RS232 As String
    GA_RS232 As String
    SC10_PATH As String
    modoc9 As Integer
    protocolo As Integer
    Modem_Puerto As Integer
    Modem_Velocidad As Integer
    Modem_Telefono As String
    Modem_Cadena As String
    lScalePos As Boolean
End Type
Public UsuarioActual As Tipo_User
Public Sub ActivaSingle(NTienda As Integer, Nombre As String, Accion As Integer, Optional sPath2 As String, Optional lScalePos As Boolean)
    Dim Arch As Integer
    Dim MyProc As Long
    Dim MyResul As Long
    Dim nBas As Database
    Dim nRes As Recordset
    Dim sLocal As String
    Dim sLocal1 As String
    
    If NTienda < 100 Then
        sLocal1 = App.Path & "\t" & Format(NTienda, "00")
    Else
        sLocal1 = App.Path & "\t" & Format(NTienda, "000")
    End If
    
    Arch = FreeFile
    If NTienda <> 0 Then
        If sPath2 = "" Then
            Set nBas = OpenDatabase(App.Path & "\tiendas.mdb")
            Set nRes = nBas.OpenRecordset("select * from grupos where tienda=" & CStr(NTienda) & " order by tienda")
            If nRes.EOF Then
                nRes.Close
                nBas.Close
                Set nRes = Nothing
                Set nBas = Nothing
                Exit Sub
            Else
                nRes.Edit
                sLocal = Mid(Trim(nRes.Fields("path")), 1, Len(Trim(nRes.Fields("path"))) - 1)
                nRes.Close
                nBas.Close
                Set nRes = Nothing
                Set nBas = Nothing
                
            End If
            
            'If NTienda < 100 Then
            '    If Dir(App.Path & "\t" & Format(NTienda, "00"), vbDirectory) = "" Then
            '        MkDir App.Path & "\t" & Format(NTienda, "00")
            '    End If
           '
           '     If Dir(App.Path & "\t" & Format(NTienda, "00") & "\S", vbDirectory) = "" Then
           '         MkDir App.Path & "\t" & Format(NTienda, "00") & "\S"
           '     End If
           ' Else
           '     If Dir(App.Path & "\t" & Format(NTienda, "000"), vbDirectory) = "" Then
           '         MkDir App.Path & "\t" & Format(NTienda, "000")
           '     End If
           '
           '     If Dir(App.Path & "\t" & Format(NTienda, "000") & "\S", vbDirectory) = "" Then
           '         MkDir App.Path & "\t" & Format(NTienda, "000") & "\S"
           '     End If
           '
           ' End If
    
        Else
        
                If Dir(Mid(sPath2, 1, Len(sPath2) - 1), vbDirectory) = "" Then
                    If NTienda < 100 Then
                        MkDir App.Path & "\t" & Format(NTienda, "00")
                    Else
                        MkDir App.Path & "\t" & Format(NTienda, "000")
                    End If
                End If
            
                If Dir(Mid(sPath2, 1, Len(sPath2) - 1) & "\S", vbDirectory) = "" Then
                    If NTienda < 100 Then
                        MkDir App.Path & "\t" & Format(NTienda, "00") & "\S"
                    Else
                        MkDir App.Path & "\t" & Format(NTienda, "000") & "\S"
                    End If
                End If
        
        End If
        
    End If
    Open MiRuta & "\multistart.ord" For Output As #Arch
    Print #Arch, (ID * 1000)
    If NTienda <> 0 Then
        
        'MsgBox "Prepara multistart.ord para Tienda:" & CStr(NTienda)
        
        '1.8.9
        If sPath2 <> "" Then
            'MsgBox (sPath2 & "hydratouch.ini")
            Print #Arch, sPath2 & "hydratouch.ini"
            Print #Arch, sPath2 & "dbasetouch.mdb"
        Else
            
            Print #Arch, sLocal1 & "\hydratouch.ini"
            If Dir(sLocal & "\dbasetouch.mdb") <> "" Then
               Print #Arch, sLocal & "\dbasetouch.mdb"
            Else
               If NTienda < 100 Then
                   Print #Arch, MiRuta & "\t" & Format(NTienda, "00") & "\dbasetouch.mdb"
                Else
                   Print #Arch, MiRuta & "\t" & Format(NTienda, "000") & "\dbasetouch.mdb"
               End If
            End If
        End If
    
    Else
        Print #Arch, ""
        Print #Arch, ""
    End If
    Print #Arch, Nombre
    'if accion=0 then accion
    Print #Arch, Format(Accion, "00")
    Print #Arch, Format(NTienda, "000")
    If HaySeguridad Then
        Print #Arch, UsuarioActual.Nombre
    Else
        Print #Arch, ""
    End If
    Close #Arch
    MyResul = STILL_ACTIVE
    On Error Resume Next
    'If lScalePos Then
        MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiRuta & "\" & "hydratouch.exe", vbNormalFocus))
    'Else
    '    MyProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiRuta & "\" & "hydra.exe", vbNormalFocus))
    'End If
    If Err.Number <> 0 Then
        'If lScalePos Then
            MsgBox "Can't execute " & App.Path & "\hydratouch.exe. Can't continue", vbCritical
        'Else
        '    MsgBox "Can't execute " & App.Path & "\hydra.exe. Can't continue", vbCritical
        'End If
        End
    End If
    On Error GoTo 0
    Do While MyResul = STILL_ACTIVE
        GetExitCodeProcess MyProc, MyResul
        Do_Events
    Loop
    If Accion = 2 And Dir(App.Path & "\gigante.txt") <> "" Then
        FrmHecho.Show
        Sleep (100)
        FrmHecho.lblHecho.Caption = "Transmisión datos a Tienda: " & CStr(NTienda) & " Realizado."
        FrmHecho.Refresh
        Sleep (2000)
        FrmHecho.Hide
    End If
    
End Sub
Public Sub EscribirParametrosInicio(MyINI As String, MiTienda As tDatosConfig)
    ' Esta función depende de FrmConfiguracion2
    Dim Fichero As Integer
    Dim ParametrosSerie As String
    Dim SC10ParSerie As String
    Dim TCPEuroscale As Long
    Dim MiPathRed As String
    Dim MiCadTel As String
    Dim MiTelefono As String
    Dim MiValor As Integer
    Dim MiValor2 As Integer
    On Error GoTo 0
    'Select Case MiTienda.Recurso
    '    Case 0
    '        SC10ParSerie = "9600,N,8,1,1"
    '        ParametrosSerie = "19200,E,8,1,1"
    '        MiTelefono = ""
    '        MiCadTel = ""
    '        TCPEuroscale = MiTienda.Config
    '        MiPathRed = "z:\sc10"
    '    Case 1
    '        If MiTienda.Tipo = 0 Then
    '            ParametrosSerie = MiTienda.Config
    '        Else
    '            SC10ParSerie = MiTienda.Config
    '        End If
    '        MiTelefono = ""
    '        MiCadTel = ""
    '        TCPEuroscale = 6000
    '        MiPathRed = "z:\sc10"
    '    Case 2
    '        SC10ParSerie = "9600,N,8,1,1"
    '        ParametrosSerie = "19200,E,8,1,1"
    '        MiTelefono = ""
    '        MiCadTel = ""
    '        TCPEuroscale = 6000
    '        MiPathRed = MiTienda.Config
    '    Case 3
    '        SC10ParSerie = "9600,N,8,1,1"
    '        ParametrosSerie = "19200,E,8,1,1"
    '        MiTelefono = MiTienda.Telefono
    '        MiCadTel = MiTienda.Config
    '        TCPEuroscale = 6000
    '        MiPathRed = "z:\sc10"
    'End Select
    Fichero = FreeFile
    Open MyINI For Output As Fichero
    '**********
    ' sec
    '**********
    Randomize Time
    MiValor = Int((4500 * Rnd) + 1)
    If HaySeguridad Then
        MiValor2 = MiValor + 1
    Else
        MiValor2 = MiValor + 3
    End If
    Write #Fichero, Encripta(Format(MiValor, "0000") & Format(MiValor2, "0000")), "NEVER EDIT NOR REMOVE THIS LINE"
    ' revisar idioma
    Write #Fichero, 0, "Idioma "
    Write #Fichero, "", "Nombre de la Empresa"
    If MiTienda.UsaGamaAlta Or MiTienda.lScalePos Then
        Write #Fichero, 1, "Utilizar Euroscale"
    Else
        Write #Fichero, 0, "Utilizar Euroscale"
    End If
    If MiTienda.UsaSC10 Then
        Write #Fichero, 1, "Utilizar Concentrador SC-10"
    Else
        Write #Fichero, 0, "Utilizar Concentrador SC-10"
    End If
    Write #Fichero, 0, "Utilizar Fichas de Vacuno"
    Write #Fichero, 0, "Utilizar Grupos de Conservacion"
    Write #Fichero, 0, "Utilizar Gedit"
    If MiTienda.GA_TCPIP Then
        Write #Fichero, True, "Ethernet(#TRUE#) o RS-232(#FALSE#)"
    Else
        Write #Fichero, False, "Ethernet(#TRUE#) o RS-232(#FALSE#)"
    End If
    Write #Fichero, MiTienda.GA_RS232, "Parametros del Puerto Serie RS-232 Euroscale"
    Write #Fichero, MiTienda.PuertoGA, "Puerto TCP/IP"
    Write #Fichero, MiTienda.SC10_PATH, "Ruta de Red del Concentrador SC-10"
    Write #Fichero, 0, "Recepción contínua de Tiquets"
    Write #Fichero, "", "Intervalo de Recepción de Tiquets"
    Write #Fichero, 1, "Euroscale:Recibir Totales Venta"
    Write #Fichero, 0, "Euroscale:Recibir Totales Super"
    Write #Fichero, 0, "Euroscale:Recibir Totales Envasado"
    Write #Fichero, 0, "Euroscale:Recibir Totales Autoservicio"
    Write #Fichero, 1, "Euroscale:Recibir Totales Venta Local"
    Write #Fichero, 0, "Euroscale:Recibir Totales Super Local"
    Write #Fichero, 0, "Euroscale:Recibir Totales Envasado Local"
    Write #Fichero, 0, "Euroscale:Recibir Totales Autoservicio Local"
    Write #Fichero, 1, "Euroscale:Borrar Totales"
    Write #Fichero, 1, "SC10:Borrar Totales"
    Write #Fichero, 1, "SC10:Recibir Totales Venta"
    Write #Fichero, 0, "SC10:Recibir Totales Super"
    Write #Fichero, 0, "SC10:Recibir Totales Envasado"
    Write #Fichero, 0, "SC10:Recibir Totales Autoservicio"
    Write #Fichero, 1, "SC10:Borrar Ticket.dat"
    Write #Fichero, 1, "SC10:Borrar Eti.dat"
    Write #Fichero, 0, "Recibir Totales de Fichas de Vacuno"
    Write #Fichero, 0, "Borrar Totales de Fichas de Vacuno"
    Write #Fichero, 2, "Decimales de la Moneda"
    Write #Fichero, "pta.", "Símbolo de Moneda Nacional"
    Write #Fichero, "166,386", "Factor Euro (paridad nacional/euro)"
    Write #Fichero, "Euro", "Moneda a Utilizar (Nacional/Euro)"
    Write #Fichero, 0, "Realizar Backup de la Base"
    Write #Fichero, "", "Ruta de Backup"
    Write #Fichero, "", "Familia Especial"
    Write #Fichero, "", "Etiqueta Especial"
    Write #Fichero, "", "AH - (Reservado)"
    Write #Fichero, 2, "Nro. de Decimales moneda nacional"
    Write #Fichero, "pta", "Símbolo de la moneda Nacional"
    Write #Fichero, 1, "Usar 20 líneas de Descriptivo"
    Write #Fichero, "0", "ENVDIR - (Reservado)"
    If MiTienda.SC10_NETPATH Then
        Write #Fichero, False, "Comunicación Rs232 con SC10"
    Else
        Write #Fichero, True, "Comunicación Rs232 con SC10"
    End If
    Write #Fichero, MiTienda.SC10_RS232, "Parametros del Puerto Serie RS-232 SC10"
    Write #Fichero, False, "Activación de envíos directos"
    Write #Fichero, 25, "Hora de Cierre"
    Write #Fichero, 61, "Minuto de Cierre"
    Write #Fichero, 25, "Hora de Recogida"
    Write #Fichero, 61, "Minuto de Recogida"
    Write #Fichero, "127.000.000.001", "Host - Enlace Socket"
    Write #Fichero, "5000", "Puerto TCP - Enlace Socket"
    Write #Fichero, "5001", "Puerto UDP - Enlace Socket"
    Write #Fichero, 0, "Enlace Socket"
    Write #Fichero, 0, "Enlace ASCII"
    Write #Fichero, "balglo.dat", "Enlace ASCII - fichero de importación contínua"
    Write #Fichero, 1, "Enlace ASCII - Dejar ficheros de tiquets"
    Write #Fichero, 0, "Iconificar en barra de estado"
    Write #Fichero, 0, "Iconificar al inicio"
    Write #Fichero, 0, "Log de Artículos"
    Write #Fichero, 1, "Log del Sistema"
    Write #Fichero, "", "Comando Libre - Nombre"
    Write #Fichero, "", "Comando Libre - Linea de comandos"
    Write #Fichero, "", "Comando Libre - Nombre"
    Write #Fichero, "", "Comando Libre - Linea de comandos"
    Write #Fichero, "", "Comando Libre - Nombre"
    Write #Fichero, "", "Comando Libre - Linea de comandos"
    Write #Fichero, "", "Comando Libre - Nombre"
    Write #Fichero, "", "Comando Libre - Linea de comandos"
    Write #Fichero, "", "Comando Libre - Nombre"
    Write #Fichero, "", "Comando Libre - Linea de comandos"
    Write #Fichero, 0, "Activar Pantallas Remotas"
    Write #Fichero, "32335", "Puerto de Pantallas Remotas"
    Write #Fichero, False, "Utilizar Moneda distinta del Euro"
    Write #Fichero, 10, "Borrar totales más antiguos de ... días"
    Write #Fichero, "", "Borrar copias de seguridad más antíguas de ... días"
    Write #Fichero, 0, "Exportación - Formato de Fecha"
    Write #Fichero, 0, "Exportación - Formato de Hora"
    Write #Fichero, "", "Exportación - Signo positivo"
    Write #Fichero, "", "Exportación - Signo negativo"
    Write #Fichero, 1, "Exportación - Utilizar valores absolutos"
    Write #Fichero, App.Path & "\t" & Format(MiTienda.Numero, "00"), "Exportación - Path"
    Write #Fichero, 1, "Exportación - Sólo Diario"
    Write #Fichero, "BBTTTT", "Nombre Fichero Tiquets"
    Write #Fichero, "1", "ASCII : envío directo"
    Write #Fichero, "0", "Vacuno versión D o posterior"
    Write #Fichero, "0", "Ticket ASCII en fichero único"
    Write #Fichero, "", "Nombre Fichero Tiquets Único"
    Write #Fichero, 1, "Reiniciar Numeración de Tiquets Euroscale"
    If MiTienda.TipoC9 = True Then
        Write #Fichero, True, "Concentrador C9 o C10"
    Else
        Write #Fichero, False, "Concentrador C9 o C10"
    End If
    Write #Fichero, MiTienda.modoc9, "Número de Secciones C9 o C10"
    Write #Fichero, 0, "Chequeo Rígido de Tiquets"
    Write #Fichero, 0, "Versión D : clientes"
    Write #Fichero, MiTienda.Modem_Puerto, "Modem --> Puerto"
    Write #Fichero, MiTienda.Modem_Velocidad, "Modem --> Velocidad"
    Write #Fichero, MiTienda.Modem_Telefono, "Modem --> Teléfono"
    Write #Fichero, MiTienda.Modem_Cadena, "Modem --> Cadena"
    Write #Fichero, "", "Modem --> Cadena"
    Write #Fichero, "", "Modem --> Cadena"
    Write #Fichero, 1, "Modem --> Colgar Automáticamente"
    If MiTienda.UsaModem Then
        Write #Fichero, 1, "Modem --> Sí(1) o No(2)"
    Else
        Write #Fichero, 0, "Modem --> Sí(1) o No(2)"
    End If
    Write #Fichero, 0, "Usa Pasarela Remota"
    Write #Fichero, "127.000.000.001", "Host Pasarela"
    Write #Fichero, "32340", "Puerto Pasarela"
    Write #Fichero, 0, "Absorver Cambios de Precio de Balanzas"
    Write #Fichero, 0, "Rechazar artículos con Precio cero"
    Write #Fichero, 0, "Solicitar pedido en recogida contínua"
    Write #Fichero, 0, "borrar pedido tras cada recogida"
    Write #Fichero, 1, "Borrar pedido a fin de día"
    Write #Fichero, 0, "Solicitar inventario en recogida contínua"
    Write #Fichero, 0, "borrar inventario tras cada recogida"
    Write #Fichero, 1, "borrar inventario a fin de día"
    Write #Fichero, 0, "Solicitar pedido a fin de día"
    Write #Fichero, 0, "solicitar inventario a fin de día"
    Write #Fichero, MiTienda.UsaEco1, "Equipos EcoNet / V8"
    Write #Fichero, True, "Tipo EcoNet (true) o V8 (false)"
    Write #Fichero, MiTienda.UsaEco2, "Equipos EcoLabel"
    Write #Fichero, 1, "Número equipo V8/EcoNet conectado"
    Write #Fichero, 1, "Número equipo EcoLabel conectado"
    Write #Fichero, MiTienda.PuertoECO, "puerto RS-232 EcoNet/V8"
    Write #Fichero, MiTienda.PuertoECO, "puerto RS-232 EcoLabel"
    Write #Fichero, MiTienda.UsaEcoPrint, "Equipos Ecoprint / IV4"
    Write #Fichero, 0, "Tipo EcoPrint"
    Write #Fichero, MiTienda.PuertoECO - 1, "puerto RS-232 EcoLabel"
    Write #Fichero, 0, "Estilo de la interfaz"
    Write #Fichero, 0, "ECO/IV4 - Comunicar sólo con una máquina"
    Write #Fichero, 0, "IV4 TQ - Solicitar pedido/inventario al recibir totales"
    Close #Fichero
    

End Sub
Public Function Encripta(texto As String) As String
    Dim Cl1 As String
    Dim Cl2 As String
    Dim MiSalida As String
    Dim Bucle As Integer
    For Bucle = 0 To 255
        Cl2 = Cl2 & Chr(Bucle)
    Next Bucle
    For Bucle = 255 To 133 Step -1
        Cl1 = Cl1 & Chr(Bucle)
    Next Bucle
    For Bucle = 0 To 132
        Cl1 = Cl1 & Chr(Bucle)
    Next Bucle
    For Bucle = 1 To Len(texto)
        MiSalida = MiSalida & Mid(Cl1, InStr(Cl2, Mid(texto, Bucle, 1)), 1)
    Next Bucle
    Encripta = MiSalida
End Function
Public Function DesEncripta(texto As String) As String
    Dim Cl1 As String
    Dim Cl2 As String
    Dim MiSalida As String
    Dim Bucle As Integer
    For Bucle = 0 To 255
        Cl1 = Cl1 & Chr(Bucle)
    Next Bucle
    For Bucle = 255 To 133 Step -1
        Cl2 = Cl2 & Chr(Bucle)
    Next Bucle
    For Bucle = 0 To 132
        Cl2 = Cl2 & Chr(Bucle)
    Next Bucle
    For Bucle = 1 To Len(texto)
        MiSalida = MiSalida & Mid(Cl1, InStr(Cl2, Mid(texto, Bucle, 1)), 1)
    Next Bucle
    DesEncripta = MiSalida
End Function
Public Function CargaPermisos(usuarios As Tipo_User, Buffer As String) As Boolean
    Dim MiCheck As String
    Dim MiCheck2 As String
    Dim Mipermiso As Byte
    Dim Bucle As Integer
    MiCheck = Right(Buffer, 1)
    calcheck Left(Buffer, Len(Buffer) - 1), MiCheck2
    If Asc(MiCheck) <> MiCheck2 Then
        CargaPermisos = False
        
        Exit Function
    End If
    usuarios.Nombre = StrConv(Left(Buffer, 8), vbLowerCase)
    usuarios.Asociado = Val(Mid(Buffer, 25, 3))
    usuarios.Password = Mid(Buffer, 9, 8)
    '*****************
    'obtiene Permisos
    '*****************
    For Bucle = 1 To 8
        Mipermiso = Asc(Mid(Buffer, 16 + Bucle, 1))
        usuarios.Permisos(Bucle - 1) = Mipermiso
    Next Bucle
End Function
Public Function Secure_Login(Mi_Usuario As Tipo_User) As Boolean
    Dim Bucle As Integer
    Dim Archivo As Integer
    Dim Buffer As String
    Dim Nulos As String * 29
    Dim BufByte(29) As Byte
    Dim Encontrado As Boolean
    Dim MiP As Byte
    Dim MiCheck
    Archivo = FreeFile()
    '1.8.9
    'Open MiRuta & "\pass.sec" For Binary As Archivo
    Open App.Path & "\pass.sec" For Binary As Archivo
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Do Until EOF(Archivo) Or Encontrado
        For Bucle = 0 To 28
            Get #Archivo, , BufByte(Bucle)
        Next Bucle
        'Buffer = Mid(Buffer, 2, Len(Buffer) - 1)
        Buffer = ""
        For Bucle = 0 To 28
            Buffer = Buffer & Chr(BufByte(Bucle))
        Next Bucle
            
        If Buffer <> Nulos Then Buffer = DesEncripta(Buffer)
        If Buffer <> Nulos Then
            If Trim(Left(Buffer, 8)) = Trim(Mi_Usuario.Nombre) Then
                Encontrado = True
            End If
        End If
    Loop
    Close #Archivo
    
    If Encontrado Then
        
        calcheck Left(Buffer, Len(Buffer) - 1), MiCheck
        If Val(MiCheck) <> Asc(Right(Buffer, 1)) Then
            'CadenadeLog "violación de seguridad : fichero pass.sec modificado"
            
            End
        End If
    
        Mi_Usuario.Permisos(0) = 0
        If Trim(Mi_Usuario.Password) <> Trim(Mid(Buffer, 9, 8)) Then
            Secure_Login = False
            
            Exit Function
        End If
        Mi_Usuario.Asociado = Val(Mid(Buffer, 25, 8))
        MiP = Asc(Mid(Buffer, 17, 1))
        For Bucle = 0 To 7
            If MiP And (2 ^ Bucle) Then
                Mi_Usuario.Permisos(0) = Mi_Usuario.Permisos(0) + (2 ^ Bucle)
            End If
        Next Bucle
        Mi_Usuario.Permisos(1) = 0
        MiP = Asc(Mid(Buffer, 18, 1))
        For Bucle = 0 To 7
            If MiP And (2 ^ Bucle) Then
                Mi_Usuario.Permisos(1) = Mi_Usuario.Permisos(1) + (2 ^ Bucle)
            End If
        Next Bucle
        Mi_Usuario.Permisos(2) = 0
        MiP = Asc(Mid(Buffer, 19, 1))
        For Bucle = 0 To 7
            If MiP And (2 ^ Bucle) Then
                Mi_Usuario.Permisos(2) = Mi_Usuario.Permisos(2) + (2 ^ Bucle)
            End If
        Next Bucle
        Mi_Usuario.Permisos(3) = 0
        MiP = Asc(Mid(Buffer, 20, 1))
        For Bucle = 0 To 7
            If MiP And (2 ^ Bucle) Then
                Mi_Usuario.Permisos(3) = Mi_Usuario.Permisos(3) + (2 ^ Bucle)
            End If
        Next Bucle
        Mi_Usuario.Permisos(4) = 0
        MiP = Asc(Mid(Buffer, 21, 1))
        For Bucle = 0 To 7
            If MiP And (2 ^ Bucle) Then
                Mi_Usuario.Permisos(4) = Mi_Usuario.Permisos(4) + (2 ^ Bucle)
            End If
        Next Bucle
        '*********************
        ' tienda
        '*********************
        Mi_Usuario.Permisos(7) = Asc(Mid(Buffer, 24, 1))
        If Mi_Usuario.Permisos(7) <> 255 Then
            
            Mi_Usuario.Permisos(7) = Mi_Usuario.Permisos(7) + 1
        Else
            Mi_Usuario.Permisos(7) = 0
        End If
        Secure_Login = True
    Else
        Secure_Login = False
    End If
    
End Function

'Public Function EncontrarUsuario() As Boolean
'    Dim MiOK As Boolean
'    Do Until MiOK
'        If frmusuario.Visible = False Then
'            Load frmusuario
'            frmusuario.Show
'            Do While frmusuario.Visible = True
'                Do_Events
'            Loop
'            Unload frmusuario
'        Else
'            Exit Function
'        End If
'        If UsuarioActual.Nombre = "NONE" Then
'            MiOK = False
'            EncontrarUsuario = False
'            Exit Function
'        Else
'            MiOK = Secure_Login(UsuarioActual)
'            If MiOK = False Then
'                MsgBox CargaCadena(769), vbCritical
'                CadenadeLog CargaCadena(863) & " : LOGIN = " & UsuarioActual.Nombre
'            Else
'                If Not StoreFlow Then grventas
'            End If
'        End If
'    Loop
'    EncontrarUsuario = MiOK
'End Function
Public Function ConsultarPermiso(ByVal Tipo As Integer) As Boolean
    Dim MiIndice As Integer
    '**********
    If Tipo < 8 Then
        MiIndice = 0
    Else
        If Tipo < 16 Then
            MiIndice = 1
            Tipo = Tipo - 8
        Else
            If Tipo < 24 Then
                MiIndice = 2
                Tipo = Tipo - 16
            Else
                If Tipo < 32 Then
                    MiIndice = 3
                    Tipo = Tipo - 24
                Else
                    MiIndice = 4
                    Tipo = Tipo - 32
                End If
            End If
        End If
    End If
    '**********
    If UsuarioActual.Permisos(MiIndice) And (2 ^ Tipo) Then
        ConsultarPermiso = True
    Else
        ConsultarPermiso = False
    End If
End Function
Public Sub CrearPassSec(PathSec As String)
    Dim Fichero As Integer
    Dim Buffer As String
    Dim BufByte As Byte
    Dim Checksum
    Dim Bucle As Integer
    Fichero = FreeFile()
    '1.8.9
    'Open PathSec & "\pass.sec" For Output As #Fichero
    Open App.Path & "\pass.sec" For Output As #Fichero
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Close #Fichero
    Fichero = FreeFile()
    '1.8.9
    'Open PathSec & "\pass.sec" For Binary As #Fichero
    Open App.Path & "\pass.sec" For Binary As #Fichero
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Buffer = "admin   " & "drake   "
    For Bucle = 1 To 8
        Buffer = Buffer & Chr(255)
    Next Bucle
    Buffer = Buffer & "000"
    Buffer = Buffer & Chr(1)
    calcheck Buffer, Checksum
    Buffer = Buffer & Chr(Checksum)
    Buffer = Encripta(Buffer)
    For Bucle = 1 To 29
        BufByte = Asc(Mid(Buffer, Bucle, 1))
        Put #Fichero, , BufByte
    Next Bucle
    Buffer = "cmdline " & "cmdline "
    For Bucle = 1 To 8
        Buffer = Buffer & Chr(255)
    Next Bucle
    Buffer = Buffer & "000"
    Buffer = Buffer & Chr(2)
    calcheck Buffer, Checksum
    Buffer = Buffer & Chr(Checksum)
    Buffer = Encripta(Buffer)
    For Bucle = 1 To 29
        BufByte = Asc(Mid(Buffer, Bucle, 1))
        Put #Fichero, , BufByte
    Next Bucle
    Close #Fichero
End Sub
Public Function calcheck(ByVal micade As String, check)

Dim i
Dim buffercheck(255)
Dim longitud_cade

longitud_cade = Len(micade)

check = 0

For i = 1 To longitud_cade Step 1
    buffercheck(i) = Mid(micade, i, 1)

    If buffercheck(i) = Chr$(126) Then
        buffercheck(i) = Chr$(0)
    End If
    check = check Xor Asc(buffercheck(i))
Next i

check = check Xor (longitud_cade + 2)
check = ((check And 63) Or 64)
End Function
Public Sub ActualizaSeguridad()
    Dim Bucle As Integer
    Dim Bucle2 As Integer
    Dim Arch As Integer
    Dim bUF() As String
    Dim NBuf As String
    Dim MiValor As Integer
    Dim MiValor2 As Integer
    FrmTools.Dir1.Path = App.Path
    FrmTools.Dir1.Refresh
    For Bucle = 0 To FrmTools.Dir1.ListCount - 1
        If Dir(FrmTools.Dir1.List(Bucle) & "\hydratouch.ini") <> "" Then
            NBuf = 0
            Arch = FreeFile()
            Open FrmTools.Dir1.List(Bucle) & "\hydratouch.ini" For Input As #Arch
            Do Until EOF(Arch)
                NBuf = NBuf + 1
                ReDim Preserve bUF(NBuf)
                Line Input #Arch, bUF(NBuf - 1)
            Loop
            Close #Arch
            Arch = FreeFile()
            Open FrmTools.Dir1.List(Bucle) & "\hydratouch.ini" For Output As #Arch
            Randomize Time
            MiValor = Int((4500 * Rnd) + 1)
            If HaySeguridad Then
                MiValor2 = MiValor + 1
            Else
                MiValor2 = MiValor + 3
            End If
            Write #Arch, Encripta(Format(MiValor, "0000") & Format(MiValor2, "0000")), "NEVER EDIT NOR REMOVE THIS LINE"
            For Bucle2 = 1 To NBuf - 1
                Print #Arch, bUF(Bucle2)
            Next Bucle2
            Close #Arch
        End If
    Next Bucle
End Sub

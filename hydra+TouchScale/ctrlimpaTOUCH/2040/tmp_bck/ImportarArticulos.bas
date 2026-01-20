Attribute VB_Name = "importararticulos"
Option Explicit
Public AHOLD As Boolean
Public RechazarPlu0 As Boolean
Public TodoModificaciones As Boolean
Private Total_lineas As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400
Public Type T_DataCaducidad
    Numero As Integer
    Seccion() As Integer
    Subseccion() As Integer
    Familia() As Integer
    Caducidad() As Integer
End Type
Public DataCaducidad As T_DataCaducidad
Sub ImportaElFichero(ByVal Fichero As String)
Dim Balenv As Boolean
Dim tipo() As String
Dim Inicio() As Long
Dim Final() As Long
Dim posicion() As Integer
Dim contador As Long
'
Dim SabecoRetry As Boolean
Dim SabecoCodigos As Boolean
Dim RespAlta As Integer
'
Dim EsBalenv As String
Dim SeccionGestion As Long
Dim dato As String
Dim buffer() As String
Dim Buffer2 As String
Dim fileConfig As Integer
Dim fileImport As Integer
Dim IncrementoBarra As Double
Dim bucle As Long
Dim Base As Database
Dim Registro As Recordset
Dim MiTecla As DB_Teclas
Dim MiContador As Long
Dim MiBucle As Integer
Dim CtEvents As Integer
Dim MiProceso As Long
Dim MiResultado As Long
Dim sRuta As String

    
    TodoModificaciones = False
    If Dir(App.Path & "\allmodify.txt") <> "" Then TodoModificaciones = True
    '
    Buscado = False
    Data_Caducidad
    frmimportar.Lista.Clear
    CtEvents = 0
    '***********************
    ' activa de momento para
    ' SABECO el 'response'
    '***********************
    Dim MISalida As Integer
    If Dir(PathLocal & "\sabeco.sbc") <> "" Then
        SabecoRetry = True
        RESPONSE = True
        MISalida = FreeFile()
        Open PathLocal & "\_articus.tmp" For Output As MISalida
    End If
    If Dir(PathLocal & "\codigos.sbc") <> "" Then
        SabecoCodigos = True
    End If
    lEsAho = False
    If Dir(PathLocal & "\ahorramas.cfg") <> "" Then
       lEsAho = True
    End If
    
    '**************************
    ' si existe un ejecutable
    ' XMSB200.EXE, lo llama
    ' primero
    '***************************
    If LCase(Trim(FicheroImport)) = "orden.dat" Then
        NoCaducidad = True
    Else
        NoCaducidad = False
    End If
    If Dir(PathImport & "\xmsb200.exe") <> "" Then
        If LogActivo Then CadenadeLog "xmsb200.exe found. Executing"
        If Dir(PathImport & "\" & Fichero) <> "" Then
            Kill PathImport & "\" & Fichero
        End If
        ChDir PathImport
        MiResultado = STILL_ACTIVE
        frmimportar.LblProgreso.Visible = True
        frmimportar.LblProgreso.Caption = "|"
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("xmsb200.EXE", vbHide))
        Do While MiResultado = STILL_ACTIVE
            Select Case frmimportar.LblProgreso.Caption
                Case "\"
                    frmimportar.LblProgreso.Caption = "|"
                Case "|"
                    frmimportar.LblProgreso.Caption = "/"
                Case "/"
                    frmimportar.LblProgreso.Caption = "-"
                Case "-"
                    frmimportar.LblProgreso.Caption = "\"
            End Select
            GetExitCodeProcess MiProceso, MiResultado
            DoEvents
            Sleep (1)
        Loop
        If LogActivo Then CadenadeLog "XMSB200.EXE finished"
    End If
    '************************************
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from articulo")
    inicia
    fileConfig = FreeFile
    ' Abre fichero de configuración y lo lee
    DoEvents
    Sleep (1)
    On Error GoTo noexiste1
    Open PathLocal & "\etc\" & FicheroImport & ".fia" For Input As fileConfig
    On Error GoTo 0
    contador = 0
    Input #fileConfig, EsBalenv
    If EsBalenv = "1" Then
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
        contador = contador + 1
    Loop
    Close fileConfig
     ' Abre fichero a importar
    On Error GoTo noexiste2
    fileImport = FreeFile
    Open PathImport & "\" & FicheroImport For Input As fileImport
    '***************************
    ' Cuenta el nro de líneas del fichero
    ' Para la barra de progreso
    '***************************
    Do Until EOF(fileImport)
        IncrementoBarra = IncrementoBarra + 1
        Line Input #fileImport, Buffer2
    Loop
    IncrementoBarra = 100 / IncrementoBarra
    frmimportar.BarraProgreso.Value = 0
    '*********************
    Close fileImport
    fileImport = FreeFile
    sRuta = PathImport & "\" & FicheroImport
    Open sRuta For Input As fileImport
    On Error GoTo 0
    '*****************************
    ' número de líneas del fichero
    '*****************************
    Total_lineas = 1
    frmimportar.LblProgreso.Visible = True
    ' Lee las líneas
    Do While (Not EOF(fileImport)) And Cancelar = False
        MiContador = 0
        MiBucle = 0
        Do Until MiBucle = 100 Or EOF(fileImport)
            ReDim Preserve buffer(MiContador + 1)
            Line Input #fileImport, buffer(MiContador)
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
                DoEvents
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
                For bucle = 0 To contador - 1
                    dato = Mid(buffer(MiBucle), Inicio(bucle), Final(bucle) - Inicio(bucle) + 1)
                    Select Case tratadato(tipo(bucle), dato, posicion(bucle))
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
                                MsgBox "Datos incorrectos en fichero de artículos"
                            End If
                            Success = 18
                        End If
                        Base.Close
                        Exit Sub
                    End Select
                Next bucle
                ' Aquí se introduce el registro en la base
                If (articulo.sec = 999 And articulo.cod = 999999) Then
                   'makro se dan de baja todos los artículos.
                   'Borrar_todos_Articulo Base
                   'Previamente se lanzan todos los existentes a fichero.
                    If Dir(App.Path & "\crfpvp.exe") <> "" Then
                        MiResultado = STILL_ACTIVE
                        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell("crfpvp.exe", vbHide))
                        Do While MiResultado = STILL_ACTIVE
                            GetExitCodeProcess MiProceso, MiResultado
                            DoEvents
                            Sleep (1)
                        Loop
                    End If
                    Borrar_todos_Articulo Base
                   
                Else 'se procesan los registros siguientes...
                
                    If (articulo.prc = 0 And Trim(articulo.txt(0)) = "") Or articulo.baja > 1 Then
                        If RESPONSE = False Then
                            Borrar_Articulo Balenv, Base
                        Else
                            Borrar_Articulo Balenv, Base, MISalida
                        End If
                        Buscado = False
                    Else
                        'Agregar_Articulo Balenv, Base, Registro
                        If RESPONSE = False Then
                            RespAlta = Alta_Articulo(Balenv, Base, Registro)
                            If RespAlta = 2 And (SabecoRetry Or lEsAho) Then
                                Borrar_Articulo Balenv, Base, , True
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro
                            End If
                            If RespAlta = 21 And (SabecoRetry Or lEsAho) Then
                                If SabecoRetry Then
                                   Borrar_Articulo Balenv, Base, , True, True
                                Else
                                   Borrar_Articulo Balenv, Base, , True, False
                                End If
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro
                            End If
                            If RespAlta = 69 And lEsAho Then
                                Borrar_Articulo Balenv, Base, MISalida, True, False
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro, MISalida
                            End If
                        
                        Else
                            RespAlta = Alta_Articulo(Balenv, Base, Registro, MISalida)
                            If RespAlta = 2 And (SabecoRetry Or lEsAho) Then
                                Borrar_Articulo Balenv, Base, MISalida, True
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro, MISalida
                            End If
                            If RespAlta = 21 And (SabecoCodigos Or lEsAho) Then
                                If SabecoRetry Then
                                   Borrar_Articulo Balenv, Base, MISalida, True, True
                                Else
                                   Borrar_Articulo Balenv, Base, MISalida, True, False
                                End If
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro, MISalida
                            End If
                            If RespAlta = 69 And lEsAho Then
                                Borrar_Articulo Balenv, Base, MISalida, True, False
                                Buscado = False
                                Alta_Articulo Balenv, Base, Registro, MISalida
                            End If
                            
                        End If
                        If articulo.tecla > 0 And articulo.JuegoTecla > -1 Then
                            MiTecla.Destino_NBalanza = 0
                            MiTecla.Destino_NMostrador = articulo.sec
                            MiTecla.Destino_CodiTec = articulo.tecla
                            MiTecla.Destino_Tabla = articulo.JuegoTecla
                            MiTecla.CodArticulo = articulo.cod
                            MiTecla.SecArticulo = 0
                            MiTecla.PluArticulo = 0
                            Alta_Teclas MiTecla, Base
                        End If
                        If Dir(App.Path & "\pathmaj.txt") <> "" Then
                           If Dir(App.Path & "\clkeymaj.txt") <> "" Then
                                If articulo.tecla = 0 And articulo.JuegoTecla = 0 Then
                                    MiTecla.Destino_NBalanza = 0
                                    MiTecla.Destino_NMostrador = articulo.sec
                                    MiTecla.Destino_CodiTec = articulo.tecla
                                    MiTecla.Destino_Tabla = articulo.JuegoTecla
                                    MiTecla.CodArticulo = articulo.cod
                                    MiTecla.SecArticulo = 0
                                    MiTecla.PluArticulo = 0
                                    Baja_Teclas MiTecla, Base
                                End If
                           
                           End If
                        Else
                        If articulo.tecla = 0 And articulo.JuegoTecla = 0 Then
                            MiTecla.Destino_NBalanza = 0
                            MiTecla.Destino_NMostrador = articulo.sec
                            MiTecla.Destino_CodiTec = articulo.tecla
                            MiTecla.Destino_Tabla = articulo.JuegoTecla
                            MiTecla.CodArticulo = articulo.cod
                            MiTecla.SecArticulo = 0
                            MiTecla.PluArticulo = 0
                            Baja_Teclas MiTecla, Base
                        End If
                        End If
                    End If
                
                End If
                
            'TratarRegistro Base, Balenv, FamiliaVacuno
            End If
        Next MiBucle
    Loop
    Close fileImport
    Base.Close
    If RESPONSE Then Close MISalida
    If Cancelar = False Then
        Success = 0
    Else
        Success = 1
    End If
    Exit Sub
noexiste1:
    Success = 10
    Base.Close
    Exit Sub
noexiste2:
    Success = 11
    Base.Close
    Exit Sub
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
    Dim bucle As Integer
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
    For bucle = 0 To 10
        .tlt(bucle) = 0
    Next bucle
    For bucle = 0 To 20
        .txt(bucle) = ""
    Next bucle
    .typ = -1
    .gpc = -1
    .etq = -1
    .prc1 = -1
    .prc2 = -1
    .trm1 = -1
    .trm2 = -1
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
    For bucle = 0 To 10
        .m_tlt(bucle) = False
    Next bucle
    For bucle = 0 To 20
        .m_txt(bucle) = False
    Next bucle
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
    End With
End Sub
Public Sub Data_Caducidad()
    Dim ArchCad As Integer
    Dim BufCad As String
    Dim MiniBufCad As String
    Dim ContCad As Integer
    Dim MySecCad As Integer
    Dim MySubSecCad As Integer
    Dim MyFamCad As Integer
    Dim MyCadCad As Integer
    Dim Errlinea As Boolean
    Dim bucle As Integer
    DataCaducidad.Numero = 0
    '************************************
    ' caducidad por sección, subsección, familia
    '************************************
    If Dir(PathLocal & "\caddata.txt") <> "" Then
        ArchCad = FreeFile()
        Open PathLocal & "\caddata.txt" For Input As #ArchCad
        Do Until EOF(ArchCad)
            Line Input #ArchCad, BufCad
            BufCad = Trim(BufCad)
            If BufCad <> "" Then
                Errlinea = False
                ContCad = 0
                MiniBufCad = ""
                For bucle = 1 To Len(BufCad)
                    If IsNumeric(Mid(BufCad, bucle, 1)) Then
                        MiniBufCad = MiniBufCad & Mid(BufCad, bucle, 1)
                    Else
                        If Mid(BufCad, bucle, 1) = "," Then
                            ContCad = ContCad + 1
                            Select Case ContCad
                                Case 1
                                    MySecCad = Val(MiniBufCad)
                                    MiniBufCad = ""
                                Case 2
                                    MySubSecCad = Val(MiniBufCad)
                                    MiniBufCad = ""
                                Case 3
                                    MyFamCad = Val(MiniBufCad)
                                    MiniBufCad = ""
                                Case Else
                                    Errlinea = True
                            End Select
                            
                        Else
                            Errlinea = True
                        End If
                    End If
                Next bucle
                If Not IsNumeric(MiniBufCad) Then Errlinea = True
                If Errlinea = False And ContCad = 3 Then
                    MyCadCad = Val(MiniBufCad)
                    DataCaducidad.Numero = DataCaducidad.Numero + 1
                    ReDim Preserve DataCaducidad.Familia(DataCaducidad.Numero)
                    ReDim Preserve DataCaducidad.Subseccion(DataCaducidad.Numero)
                    ReDim Preserve DataCaducidad.Seccion(DataCaducidad.Numero)
                    ReDim Preserve DataCaducidad.Caducidad(DataCaducidad.Numero)
                    DataCaducidad.Familia(DataCaducidad.Numero - 1) = MyFamCad
                    DataCaducidad.Subseccion(DataCaducidad.Numero - 1) = MySubSecCad
                    DataCaducidad.Seccion(DataCaducidad.Numero - 1) = MySecCad
                    DataCaducidad.Caducidad(DataCaducidad.Numero - 1) = MyCadCad
                End If
            End If
        Loop
        Close #ArchCad
    End If
    '************************************
End Sub

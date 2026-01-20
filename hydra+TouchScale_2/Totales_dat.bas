Attribute VB_Name = "Totales_dat"
Option Explicit
Public Function ExportaTotalesDat(MostrarMensajes As Boolean, Optional Fecha As String, Optional sSel As String) As Boolean
'************************
' esporta el fichero
' totales.dat, y
' si procede, genera msb100
'************************
Dim aviso As String
Dim nID As Integer
Dim nLen As Double
Dim Registro As dao.Recordset
Dim RegAux As dao.Recordset
Dim Base As dao.Database
Dim Archivo1 As Integer
Dim Archivo2 As Integer
Dim ArchLocal As Integer
Dim MiProceso As Long
Dim MiResultado As Long
Dim sBackName As String
Dim sBackLocal As String
Dim sCadena As String
Dim MiPath As String
Dim MiFecha As String
Dim MiFechaMSB As String
Dim MiFactor1 As Double
Dim MiFactor2 As Double
Dim MyPlu As String
Dim MyVend As String
Dim STPETIC As Integer
Dim CuentaEvents As Integer
Dim Monetario As Double
'Dim RegBal as dao.recordset
'Dim NumRel As Integer
Dim MiCad1 As String, MiCad4 As String
Dim sJuliano As String
Dim nFich As Integer
Dim sCadMaj As String
Dim lValFrancos As Double
Dim sEan As String
Dim lTotalFrancos As Double

lValFrancos = CDbl("6,55957")

'c2f para exportar todos los sistemas si hay multitienda y se llama a exportatotalesdat con true
'    , es decir, cuando se valla a utilizar como copia de seguridad.
Dim ncont As Integer
Dim nBucle As Integer

ncont = 1

If Fecha = "" Then
    MiFecha = Format(Date, "mm/dd/yy")
    MiFechaMSB = Format(Date, "dd/mm/yy")
    sBackName = Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.dat"
Else
    MiFecha = Fecha
    MiFechaMSB = Mid(Fecha, 4, 2) & "/" & left(Fecha, 2) & "/" & Right(Fecha, 2)
    sBackName = Right(Fecha, 2) & left(Fecha, 2) & Mid(Fecha, 4, 2) & "BK.dat"
End If
sBackLocal = Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.loc.tmp"
MiPath = Miruta
lCogeTiquet = False
'***************************
' pregunta antes de exportar
'***************************
nID = vbOK
'If xlcampo And MostrarMensajes Then
'    nId = MsgBox("¿Realizar Exportación?", vbOKCancel)
'    If nId <> vbOK Then
'        lCogeTiquet = True
'        Exit Sub
'    End If
'End If
If Dir(MiPath, vbDirectory) = "" Then
    CadenadeLog "Path de exportación e Totales.dat no válido"
    ExportaTotalesDat = False
    Exit Function
End If
Do_Events
If Dir(MiPath & "\" & "totales.dat") <> "" Then
    FileCopy MiPath & "\" & "totales.dat", MiPath & "\" & "totales.ant"
End If
Do_Events

'c2f copia de seguridad totales que van a ser borrados...
If MostrarMensajes Then
    If Dir(App.Path & "\hydramulti.exe") <> "" Then
       ncont = 3
    End If
    Archivo1 = FreeFile()
    Open MiPath & "\" & sBackName For Output As #Archivo1
    sBackName = ""
    sBackLocal = ""
Else
    Archivo1 = FreeFile()
    Open MiPath & "\" & "totales.dat" For Output As #Archivo1
    Archivo2 = FreeFile()
    If sBackName <> "" Then Open MiPath & "\" & sBackName For Output As #Archivo2
    ArchLocal = FreeFile()
    If sBackLocal <> "" Then Open MiPath & "\" & sBackLocal For Output As #ArchLocal
End If

CuentaEvents = 0
Set Base = AbrirBase
Set RegAux = Base.OpenRecordset("select * from articulo where borrado=false")
'comienza exportación desde tablas totales .
'log
If MostrarMensajes Then
    Set Registro = Base.OpenRecordset("select * from log")
Else
    Set Registro = Base.OpenRecordset("select * from log where cstr(cdate(fecha)) = " & Chr(34) & CStr(CDate(MiFecha)) & Chr(34))
End If
    With Registro
    If Not .EOF Then
    .MoveFirst
        Do Until .EOF
            '***************
            ' do_events para master/slave
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 4 Then
                CuentaEvents = 0
                Do_Events
            End If
            '***************
            sCadena = Trim(!texto)
            Print #Archivo1, sCadena
            If sBackName <> "" Then Print #Archivo2, sCadena
            .Movenext
        Loop
    End If
End With
'*******************
' cabeceras de tiquets
'*******************
For nBucle = 1 To ncont

If MostrarMensajes Then
    'If nBucle = 1 Then
        Set Registro = Base.OpenRecordset("select * from cabecera")
    'Else
    '    If nBucle = 2 Then
    '        Set Registro = Base.OpenRecordset("select * from cabeceraSC10")
    '    Else
    '        Set Registro = Base.OpenRecordset("select * from cabeceraTQ")
    '    End If
    'End If

Else
    If sSel = "" Then
        '2.0.a era mifechamsb en lugar de mifecha
        Set Registro = Base.OpenRecordset _
        ("select * from cabecera where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") And (cstr(cdate(Fecha)) = " & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & " or reserv2<>" & Chr(34) & "*" & Chr(34) & ")")
    Else
        '2.0.a era mifechamsb en lugar de mifecha
        If sSel = "*" Then
        Set Registro = Base.OpenRecordset _
        ("select * from cabecera where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") And (cstr(cdate(Fecha)) = " & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & ")")
        Else
        '2.0.a mifechamsb en lugar de mifecha
        Set Registro = Base.OpenRecordset _
        ("select * from cabecera where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") And (cstr(cdate(Fecha)) = " & Chr(34) & CStr(CDate(MiFecha)) & Chr(34) & " and reserv2<>" & Chr(34) & "*" & Chr(34) & ")")
        
        End If
    End If
End If
With Registro
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            '***************
            ' do_events para master/slave
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 4 Then
                CuentaEvents = 0
                Do_Events
            End If
            '***************
            sCadena = "CTQ "
            STPETIC = Val(!STPETIC)
            sCadena = sCadena & left(!Fecha, 2) & Mid(!Fecha, 4, 2) & Right(!Fecha, 2) & ","
            If Mid(!hora, 2, 1) = ":" Then
                sCadena = sCadena & "0" & left(!hora, 1) & Mid(!hora, 3, 2) & ","
            Else
                sCadena = sCadena & left(!hora, 2) & Mid(!hora, 4, 2) & ","
            End If
            sCadena = sCadena & Format(Val(!Balanza), "00") & ","
            '*********************************
            ' nro. de tiquet
            If Dir(App.Path & "\fede386.exe") <> "" Or Dir(App.Path & "\fedeiv8.exe") <> "" Then
                sCadena = sCadena & Format(Val(!nume), "00000") & ","
            Else
                If (STPETIC And 1) <> 0 Then
                    'local
                    sCadena = sCadena & Format(99 - Val(!Balanza), "00")
                Else
                    'red
                    sCadena = sCadena & Format(Val(!Balanza), "00")
                End If
                sCadena = sCadena & Format(Val(!nume), "000") & ","
            End If
            '***********************************
            ' vendedor
            sCadena = sCadena & Format(Val(!vended), "0000") & ","
            '***********************************
            ' 12 primeros dígitos del EAN
            sCadena = sCadena & "2" & Format(Val(!seccion), "0")
            If (STPETIC And 1) <> 0 Then
                'local
                sCadena = sCadena & Format(99 - Val(!Balanza), "00")
            Else
                'red
                sCadena = sCadena & Format(Val(!Balanza), "00")
            End If
            sCadena = sCadena & Format(Val(!nume), "000")
            If Val(!Importe) < 0 Then
                sCadena = sCadena & "00000"
            Else
                Monetario = !Importe
                If UsaEuro Then
                    sCadena = sCadena & Format(Monetario * 100, "00000")
                Else
                    sCadena = sCadena & Format(Monetario * (10 ^ decimales), "00000")
                End If
            End If
            '****************************
            ' dentro y fuera de fecha
            '****************************
            If !Fecha = MiFechaMSB Then
                Print #Archivo1, sCadena
                If sBackName <> "" Then Print #Archivo2, sCadena
                Edit_Record Registro
                .Fields("reserv2") = "*"
                .Update
            Else
                If DateSerial(Val(Right(MiFechaMSB, 2)), Val(Mid(MiFechaMSB, 4, 2)), Val(left(MiFechaMSB, 2))) > DateSerial(Val(Right(!Fecha, 2)), Val(Mid(!Fecha, 4, 2)), Val(left(!Fecha, 2))) Then
                    If sBackLocal <> "" Then Print #ArchLocal, sCadena
                    Edit_Record Registro
                    .Fields("reserv2") = "*"
                    .Update
                End If
            End If
            
            .Movenext
        Loop
    End If
End With

Next nBucle


'lineas
For nBucle = 1 To ncont

If MostrarMensajes Then
    'If nBucle = 1 Then
        Set Registro = Base.OpenRecordset("select * from tickets")
    'Else
    '    If nBucle = 2 Then
    '        Set Registro = Base.OpenRecordset("select * from ticketsSC10")
    '    Else
    '        Set Registro = Base.OpenRecordset("select * from ticketsTQ")
    '    End If
    'End If
Else
    If sSel = "" Then
        Set Registro = Base.OpenRecordset _
        ("select * from tickets where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") and (date = " & Chr(34) & MiFechaMSB & Chr(34) & " or reserv2<>" & Chr(34) & "*" & Chr(34) & ")")
    Else
        If sSel = "*" Then
        Set Registro = Base.OpenRecordset _
        ("select * from tickets where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") and (date = " & Chr(34) & MiFechaMSB & Chr(34) & ")")
        Else
        Set Registro = Base.OpenRecordset _
        ("select * from tickets where (typtic=" & Chr(34) & "1" & Chr(34) & " or typtic=" & Chr(34) & "5" & Chr(34) & _
        ") and (date = " & Chr(34) & MiFechaMSB & Chr(34) & " and reserv2<>" & Chr(34) & "*" & Chr(34) & ")")
        
        End If
    End If
End If

With Registro
    If Not .EOF Then
    .MoveFirst
        Do Until .EOF
            '***************
            ' do_events para master/slave
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 4 Then
                CuentaEvents = 0
                Do_Events
            End If
            '***************
            If Val(!lincan) = 1 Then
                MiFactor1 = -1
            Else
                MiFactor1 = 1
            End If
            If (Val(!STPETIC) And 64) <> 0 Or (Val(!STPETIC) And 128) <> 0 Then
                MiFactor2 = -1
            Else
                MiFactor2 = 1
            End If
            '*********************
            ' micad1 --> 2, no pesado
            '            0, pesado
            '            1, cancelado
            '*********************
            If !units <> 0 Then
                MiCad1 = "2"
            Else
                MiCad1 = "0"
            End If
            '*********************
            ' micad4=PLU
            '*********************
            If Val(!Code) = 0 Then
                MiCad4 = "0000"
            Else
                If Val(!Plu) = 0 Then
                    RegAux.FindFirst "codigo=" & Val(!Code)
                    If RegAux.EOF Or RegAux.NoMatch Then
                        MiCad4 = "0000"
                    Else
                        MiCad4 = Format(RegAux!Plu, "0000")
                    End If
                Else
                    MiCad4 = Format(!Plu, "0000")
                End If
            End If
            STPETIC = Val(!STPETIC)
            sCadena = "LTQ "
            sCadena = sCadena & Format(left(!Date, 2), "00") & Format(Mid(!Date, 4, 2), "00") & Format(Right(!Date, 2), "00")
            sCadena = sCadena & ","
            sCadena = sCadena & Format(left(!Hour, 2), "00") & Format(Mid(!Hour, 4, 2), "00")
            sCadena = sCadena & "," & Format(!numbal, "00") & ","
            '**********************
            ' busca el nro. de balanza en la sección
            '**********************
            'NumRel = 1
            '
            ' local
            'Set RegBal = Base.OpenRecordset("select * from equipos where secc_maqui=" & !SECC & " order by numero_eqp")
            'If Not RegBal.EOF Then
            '    RegBal.MoveFirst
            '    Do Until RegBal.EOF
            '        If Val(!numbal) <> !numero_eqp Then
            '            NumRel = NumRel + 1
            '        Else
            '            Exit Do
            '        End If
            '        RegBal.MoveNext
            '    Loop
            'End If
            '
            'If NumRel > 9 Then NumRel = 0
            '**********************
            'If (STPETIC And 17) <> 0 Then
            '    NumRel = 99 - NumRel
            '    sCadena = sCadena & Format(NumRel, "00")
            '    sCadena = sCadena & Format(Val(!nume), "000")
            'Else
            '    sCadena = sCadena & Format(NumRel, "0")
            '    sCadena = sCadena & Format(Val(!nume), "0000")
            '    ' red
            'End If
            '**********************
            If Dir(App.Path & "\fede386.exe") <> "" Or Dir(App.Path & "\fedeiv8.exe") <> "" Then
                sCadena = sCadena & Format(Val(!nume), "00000")
            Else
                If (STPETIC And 1) <> 0 Then
                    'local
                    sCadena = sCadena & Format(99 - Val(!numbal), "00")
                Else
                    'red
                    sCadena = sCadena & Format(Val(!numbal), "00")
                End If
                sCadena = sCadena & Format(Val(!nume), "000")
            End If
            sCadena = sCadena & ","
            sCadena = sCadena & Format(Val(!vendor), "0000") & ","
            If (MiFactor1 * MiFactor2) = -1 Then
                sCadena = sCadena & "0" & ","
            Else
                sCadena = sCadena & MiCad1 & ","
            End If
            sCadena = sCadena & MiCad4 & ","
            sCadena = sCadena & Format(Val(!Code), "000000") & ","
            sCadena = sCadena & Format(!Weight * 1000, "000000000") & ","
            sCadena = sCadena & Format(!units, "000000000") & ","
            Monetario = !amount
            If UsaEuro Then
                sCadena = sCadena & Format(Monetario * 100, "000000000")
            Else
                sCadena = sCadena & Format(Monetario * (10 ^ decimales), "000000000")
            End If
            
            '****************************
            ' fuera de fecha
            '****************************
            
            If !Date = MiFechaMSB Then
                Print #Archivo1, sCadena
                If sBackName <> "" Then Print #Archivo2, sCadena
                Edit_Record Registro
                !reserv2 = "*"
                .Update
            Else
                If DateSerial(Val(Right(MiFechaMSB, 2)), Val(Mid(MiFechaMSB, 4, 2)), Val(left(MiFechaMSB, 2))) > DateSerial(Val(Right(!Date, 2)), Val(Mid(!Date, 4, 2)), Val(left(!Date, 2))) Then
                    If sBackLocal <> "" Then Print #ArchLocal, sCadena
                    Edit_Record Registro
                    !reserv2 = "*"
                    .Update
                End If
            End If
            
            .Movenext
        Loop
    End If
End With
On Error Resume Next
Close #ArchLocal
On Error GoTo 0

Next nBucle

If MostrarMensajes Then
    Set Registro = Base.OpenRecordset("select * from gtsecs")
Else
    Set Registro = Base.OpenRecordset("select * from gtsecs where fecha = #" & MiFecha & "#")
End If
With Registro
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            '***************
            ' do_events para master/slave
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 4 Then
                CuentaEvents = 0
                Do_Events
            End If
            '***************
            If UsaEuro Then
                sCadena = "SEC " & Format(!codi_ident, "000") & "," & Format(!codi_sub, "000") & "," & Format(!codi_fam, "0000") & "," & _
                Format(!tiquets, "000000000") & "," & Format(!operacion, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format((!pesetas * 100), "000000000") & "," & Format(!secc_Maqui, "00")
            Else
                sCadena = "SEC " & Format(!codi_ident, "000") & "," & Format(!codi_sub, "000") & "," & Format(!codi_fam, "0000") & "," & _
                Format(!tiquets, "000000000") & "," & Format(!operacion, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format(!pesetas * (10 ^ decimales), "000000000") & "," & Format(!secc_Maqui, "00")
            End If
            Print #Archivo1, sCadena
            If sBackName <> "" Then Print #Archivo2, sCadena
            .Movenext
        Loop
    End If
End With

If lClR And sPathMaj <> "" Then
    sJuliano = Calcula_Juliano1()
    nFich = FreeFile()
    Open sPathMaj & "\BAL_VTE.LCK" For Output As #nFich
    Close #nFich
    nFich = FreeFile()
    Open sPathMaj & "\bal_vte" & sJuliano For Output As #nFich
End If

If MostrarMensajes Then
    Set Registro = Base.OpenRecordset("select * from gtarti")
Else
    Set Registro = Base.OpenRecordset("select * from gtarti where fecha = #" & MiFecha & "#")
End If

With Registro
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
        '***************
        ' do_events para master/slave
        CuentaEvents = CuentaEvents + 1
        If CuentaEvents >= 4 Then
            CuentaEvents = 0
            Do_Events
        End If
        '***************
        If lClR And sPathMaj <> "" Then
            lTotalFrancos = Round((((lValFrancos * !pesetas) * 100) * 10) / 10)
            '1.7.19
            'sCadMaj = Format(!codi_ident, "00000") & "#" & Format(!plu, "000000") & "#" & Mid(Trim(!txt_sub), 1, 13) & "#" & _
            '          Format(lValFrancos, "0000000000") & "#" & Format((!pesetas * 100), "000000000") & "#" & _
            '          Format((!Peso * 1000), "00000000") & "#" & Format(!Unidades, "00000")
            If Dir(App.Path & "\j0.txt") <> "" Or lTotalFrancos > 99999 Then
                sEan = Mid(Trim(!txt_sub), 1, 7) & "00000"
            Else
                sEan = Mid(Trim(!txt_sub), 1, 7) & Format(lTotalFrancos, "00000")
            End If
            If Len(sEan) = 12 Then
                sEan = sEan & Format(DigitoControl(sEan), "0")
            Else
                If Len(sEan) <> 13 Then
                    sEan = Space(13)
                End If
            End If
            sCadMaj = Format(!codi_ident, "00000") & "#" & Format(!Plu, "000000") & "#" & sEan & "#" & _
                      Format(lTotalFrancos, "0000000000") & "#" & Format((!pesetas * 100), "000000000") & "#" & _
                      Format((!Peso * 1000), "00000000") & "#" & Format(!Unidades, "00000")
            
            Print #nFich, sCadMaj
        End If
        '***************
        If UsaEuro Then
            MyPlu = Trim(!des_plu1)
            If Len(MyPlu) < 25 Then
                MyPlu = MyPlu & Space(25 - Len(MyPlu))
            Else
                MyPlu = Mid(MyPlu, 1, 25)
            End If
            sCadena = "ART " & Format(Date, "ddmmyy") & "," & Format(!Plu, "000000") & "," & MyPlu & "," & Format(!codi_ident, "000") & "," & _
            Format(!codi_sub, "0000") & "," & Format(!codi_fam, "000") & "," & Format(!codigo, "000000") & "," & Format((!precio * 100), "0000000") & "," & _
            Format(!Unidades, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format((!pesetas * 100), "000000000") & "," & Format(!tara, "0000000") & "," & _
            Format(!caducidad, "0000000") & "," & Format(!secc_Maqui, "00")
        Else
            sCadena = "ART " & Format(Date, "ddmmyy") & "," & Format(!Plu, "000000") & "," & MyPlu & "," & Format(!codi_ident, "000") & "," & _
            Format(!codi_sub, "0000") & "," & Format(!codi_fam, "000") & "," & Format(!codigo, "000000") & "," & Format(!precio * (10 ^ decimales), "0000000") & "," & _
            Format(!Unidades, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format(!pesetas * (10 ^ decimales), "000000000") & "," & Format(!tara, "0000000") & "," & _
            Format(!caducidad, "0000000") & "," & Format(!secc_Maqui, "00")
        End If
        Print #Archivo1, sCadena
        If sBackName <> "" Then Print #Archivo2, sCadena
        .Movenext
        Loop
    End If
End With

If lClR And sPathMaj <> "" Then
    Close #nFich
    Kill sPathMaj & "\BAL_VTE.LCK"
End If

If MostrarMensajes Then
    Set Registro = Base.OpenRecordset("select * from gtvend")
Else
    Set Registro = Base.OpenRecordset("select * from gtvend where fecha = #" & MiFecha & "#")
End If
With Registro
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            '***************
            ' do_events para master/slave
            CuentaEvents = CuentaEvents + 1
            If CuentaEvents >= 4 Then
                CuentaEvents = 0
                Do_Events
            End If
            '***************
            MyVend = Trim(!Nombre)
            If Len(MyVend) < 25 Then
                MyVend = MyVend & Space(25 - Len(MyVend))
            End If
            If UsaEuro Then
                sCadena = "VND " & Format(!ident_vend, "0000") & "," & MyVend & "," & Format((!Supertot * 100), "000000000") & "," & Format((!totcance * 100), "000000000") & "," & Format((!Total * 100), "000000000") & "," & _
                Format(!tiquets, "000000000") & "," & Format(!operacion, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format(!cancela, "000000000") & "," & "000000000" & "," & Format(!codi_ident, "00")
            Else
                sCadena = "VND " & Format(!ident_vend, "0000") & "," & MyVend & "," & Format(!Supertot * (10 ^ decimales), "000000000") & "," & Format(!totcance * (10 ^ decimales), "000000000") & "," & Format(!Total * (10 ^ decimales), "000000000") & "," & _
                Format(!tiquets, "000000000") & "," & Format(!operacion, "000000000") & "," & Format((!Peso * 1000), "000000000") & "," & Format(!cancela, "000000000") & "," & "000000000" & "," & Format(!codi_ident, "00")
            End If
            Print #Archivo1, sCadena
            If sBackName <> "" Then Print #Archivo2, sCadena
            .Movenext
        Loop
    End If
End With
Base.Close
Close #Archivo1
If sBackName <> "" Then Close #Archivo2
If sBackLocal <> "" Then
    Close #ArchLocal
    If FileLen(MiPath & "\" & Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.loc.tmp") > 0 Then
        Do_Events
        FileCopy MiPath & "\" & Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.loc.tmp", MiPath & "\" & Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.loc"
    Else
        Kill MiPath & "\" & Mid(Format(Year(Date), "0000"), 3, 2) + Format(Month(Date), "00") + Format(Day(Date), "00") + "BK.loc.tmp"
    End If
End If
'*******************************

'*******************************
' tras generar el archivo
' llama a msb100
'*******************************
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
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiPath & "\" & "fedeiv8.exe " & MiFechaMSB, vbHide))
        Do While MiResultado = STILL_ACTIVE
            GetExitCodeProcess MiProceso, MiResultado
            Do_Events
        Loop
    End If
    If Dir(MiPath & "\fede386.exe") <> "" Then
        MiResultado = STILL_ACTIVE
        MiProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(MiPath & "\" & "fede386.exe " & MiFechaMSB, vbHide))
        Do While MiResultado = STILL_ACTIVE
            GetExitCodeProcess MiProceso, MiResultado
            Do_Events
        Loop
    End If
    
    If Dir(MiPath & "\msb100") = "" And Dir(MiPath & "\xmsb100.exe") <> "" Then
        'If MostrarMensajes Then MsgBox (CargaCadena( 861))
        If DebugActivo Then CadenadeLog CargaCadena(861)
        lCogeTiquet = True
        ExportaTotalesDat = False
        Exit Function
    End If
   
    If DebugActivo Then CadenadeLog CargaCadena(862)
    
    lCogeTiquet = True
    ExportaTotalesDat = True
End If

End Function


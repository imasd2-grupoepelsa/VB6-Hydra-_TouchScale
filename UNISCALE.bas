Attribute VB_Name = "UNISCALE"
Option Explicit
'*********************************
' Euroscale=true // SC10=false
'*********************************
Public Sub export_CABLIN(lTodos As Boolean, Euroscale_o_SC10 As Boolean)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    '**************************
    ' Cabecera
    Dim B1 As String
    Dim B2 As String
    Dim Codigo_mos As String
    Dim Codigo_bal As String
    Dim Codigo_ven As String
    'Dim Cod_Bar As String
    Dim Lineas As String
    Dim precio As String
    Dim Tipo As String
    Dim Fecha As String
    Dim hora As String
    Dim LineaAnulada As String
    '********
    ' Linea
    Dim Num_lin As String
    Dim Cod_Plu As String
    Dim Cod_Cod As String
    Dim Precio_Uni As String
    Dim Pes_Uni As String
    Dim Total As String
    Dim Signo As String
    Dim Tipo_Linea As String
    Dim Anulado As String
    '************
    Dim MiEnvio As String
    Dim nintentos As Integer
    'Dim Bucle As Integer

    Dim MiTabla1 As String
    Dim MiTabla2 As String
    Dim Fichero1 As Integer
    Dim Fichero2 As Integer
    'Dim NombreTiquet As String
    'Dim BufferLineas() As String
    'Dim MiCuentaLineas As Integer
    'Dim BucleLineas As Integer
    'Dim HayCabecera As Boolean
    'Dim PrecioSinIVA As Double
    'Dim PrecioConIVA As Double
    'Dim BufferIVA As String
    'Dim BufferIVA2 As String
    Dim sLinCan As String
    Dim sCodCli As String
    Dim nSysToca As Integer

    'Dim NombreF As Tipo_NombreTiquet
    '****************************************
    Dim sSQL As String
    Dim sruta As String

    If Dir(App.Path & "\cabtqt.sdf") <> "" Then
        If Dir(App.Path & "\cabtqt.bak") <> "" Then
            If Dir(App.Path & "\cabtqt.old") <> "" Then
                Kill App.Path & "\cabtqt.old"
            End If
            FileCopy App.Path & "\cabtqt.bak", App.Path & "\cabtqt.old"
            Kill App.Path & "\cabtqt.bak"
        End If

        FileCopy App.Path & "\cabtqt.sdf", App.Path & "\cabtqt.bak"

    End If
    If Dir(App.Path & "\tiquet.sdf") <> "" Then
        If Dir(App.Path & "\tiquet.bak") <> "" Then
            If Dir(App.Path & "\tiquet.old") <> "" Then
                Kill App.Path & "\tiquet.old"
            End If
            FileCopy App.Path & "\tiquet.bak", App.Path & "\tiquet.old"
            Kill App.Path & "\tiquet.bak"
        End If

        FileCopy App.Path & "\tiquet.sdf", App.Path & "\tiquet.bak"

    End If


    Fichero1 = FreeFile()
    Open App.Path & "\cabtqt.sdf" For Append As #Fichero1
    Fichero2 = FreeFile()
    Open App.Path & "\tiquet.sdf" For Append As #Fichero2

    '1.7.14
    'If Euroscale_o_SC10 Then
    '    MiTabla1 = "Cabecera"
    '    MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If


    Set Base = OpenDatabase(Base_General)

    'For nSysToca = 0 To 1

    'If nSysToca = 1 Then
    MiTabla1 = "Cabecera"
    MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If

    sSQL = "SELECT * " & _
           "FROM " & MiTabla1 & " " & _
           "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
           ") order by val(nume)"

    CadenadeLog sSQL

    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then


        'With Registro

        ' si no está enviada la cabecera, se envía
        'Do Until Registro.EOF Or CI_Salir = True
        'With Registro


        'With Registro
        Registro.MoveFirst


        Do While Not Registro.EOF And CI_Salir = False

            With Registro

                Codigo_mos = Format(Val(!seccion), "00")
                Codigo_bal = Format(Val(!Balanza), "00")
                Codigo_ven = Format(Val(!vended), "0000")
                sLinCan = Format(Val(!lincan), "00")

                Lineas = Format(Val(!numlin), "00")
                If !Importe >= 0 Then
                    precio = !Importe
                    precio = Format(100 * (precio), "0000000")
                Else
                    precio = (-1) * !Importe
                    precio = "-" & Format(100 * (precio), "000000")
                End If
                Fecha = Format(!Fecha, "ddmmyy") & "  "
                hora = Format(!hora, "hhmm") & "    "
                If lUpper Then
                    sCodCli = "000000"
                Else
                    sCodCli = Format(Val(!codcli), "000000")
                End If
                ' Enviar la cabecera
                MiEnvio = Format(Val(!nume), "00000") & "        "
                MiEnvio = MiEnvio & Codigo_mos & Codigo_ven & Codigo_bal & Lineas & sLinCan
                MiEnvio = MiEnvio & precio & hora & "0" & "0" & sCodCli & Fecha
                Print #Fichero1, MiEnvio

                Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
                                              Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                                            " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                                            & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                                            & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & ") order by val(numlin)")

                If Not Reg2.EOF Then
                    Reg2.MoveFirst


                    Do While Not Reg2.EOF
                        With Reg2
                            Num_lin = Format(!numlin, "00")
                            ''Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                            If lUpper Then
                                Cod_Plu = CodSecPlu(!Code)
                            Else
                                Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                            End If
                            Cod_Cod = Format(!Code, "000000")

                            If !Price >= 0 Then
                                Precio_Uni = !Price
                                Precio_Uni = Format(100 * (Precio_Uni), "0000000")
                            Else
                                Precio_Uni = (-1) * !Price
                                Precio_Uni = "-" & Format(100 * (Precio_Uni), "000000")
                            End If

                            If !units = 0 Then
                                If !Weight >= 0 Then
                                    Pes_Uni = !Weight
                                    Pes_Uni = Format(1000 * (Pes_Uni), "000000000")
                                Else
                                    Pes_Uni = (-1) * !Weight
                                    Pes_Uni = "-" & Format(1000 * (Pes_Uni), "00000000")
                                End If
                                Tipo_Linea = "0"
                            Else
                                If !units > 0 Then
                                    Pes_Uni = !units
                                    Pes_Uni = Format(Pes_Uni, "000000000")
                                Else
                                    Pes_Uni = (-1) * !units
                                    Pes_Uni = "-" & Format(Pes_Uni, "00000000")
                                End If
                                Tipo_Linea = "1"
                            End If
                            If !amount < 0 Then
                                Total = (-1) * !amount
                                Total = "-" & Format(100 * (Total), "000000")
                                Signo = "1"
                            Else
                                Total = !amount
                                Total = Format(100 * (Total), "0000000")
                                Signo = "0"
                            End If
                            If !lincan = 0 Then
                                LineaAnulada = "0"
                            Else
                                LineaAnulada = "1"
                            End If

                            MiEnvio = "00" & Fecha & Mid(hora, 1, 4) & Codigo_ven & Format(Val(!nume), "00000")
                            'If lUpper Then
                            '    MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Codigo_mos & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                            'Else
                            '    MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Cod_Cod & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                            'End If
                            If lUpper Then
                                'MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Format(!secc, "00") & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                                MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                            Else
                                MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Cod_Cod & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                            End If

                            'MiEnvio = MiEnvio & LineaAnulada & Total & "1"
                            Print #Fichero2, MiEnvio

                            Reg2.Edit
                            Reg2.Fields("reserv3") = "SI"
                            Reg2.Update
                            Reg2.Movenext

                        End With


                    Loop

                End If

                Registro.Edit
                Registro.Fields("reserv3") = "SI"
                Registro.Update

                Registro.Movenext

            End With

        Loop    'loop registros cabecera

        'End With

    End If    'si hay cabeceras pendientes

    'Next nSysToca

    Close #Fichero1
    Close #Fichero2

    CerrarBase Base
    lCogeTiquet = True

End Sub



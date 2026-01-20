Attribute VB_Name = "coms"
'*********************************
' Euroscale=true // SC10=false
'*********************************
Public Sub export_CABLIN(ByVal lTodos As Boolean)
Dim Base As DAO.Database
Dim Registro As DAO.Recordset
Dim Reg2 As DAO.Recordset
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
Dim lEX As Boolean

'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
    
    If Dir(var.PathExpor & "\cabtqt.sdf") <> "" Then
        If Dir(var.PathExpor & "\cabtqt.bak") <> "" Then
            If Dir(var.PathExpor & "\cabtqt.old") <> "" Then
                Kill var.PathExpor & "\cabtqt.old"
            End If
            FileCopy var.PathExpor & "\cabtqt.bak", var.PathExpor & "\cabtqt.old"
            Kill var.PathExpor & "\cabtqt.bak"
        End If
        
        FileCopy var.PathExpor & "\cabtqt.sdf", var.PathExpor & "\cabtqt.bak"
        
    End If
    If Dir(var.PathExpor & "\tiquet.sdf") <> "" Then
        If Dir(var.PathExpor & "\tiquet.bak") <> "" Then
            If Dir(var.PathExpor & "\tiquet.old") <> "" Then
                Kill var.PathExpor & "\tiquet.old"
            End If
            FileCopy var.PathExpor & "\tiquet.bak", var.PathExpor & "\tiquet.old"
            Kill var.PathExpor & "\tiquet.bak"
        End If
        
        FileCopy var.PathExpor & "\tiquet.sdf", var.PathExpor & "\tiquet.bak"
        
    End If
    
    
    Fichero1 = FreeFile()
    'Open App.Path & "\cabtqt.sdf" For Output As #Fichero1
    Open var.PathExpor & "\cabtqt.sdf" For Append As #Fichero1
    Fichero2 = FreeFile()
    'Open App.Path & "\tiquet.sdf" For Output As #Fichero2
    Open var.PathExpor & "\tiquet.sdf" For Append As #Fichero2
    
    '1.7.14
    'If Euroscale_o_SC10 Then
    '    MiTabla1 = "Cabecera"
    '    MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If
    
    
    Set Base = AbrirBase()
    
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
            sCodCli = Format(Val(!codcli), "000000")
            ' Enviar la cabecera
            MiEnvio = Format(Val(!nume), "00000") & "        "
            MiEnvio = MiEnvio & Codigo_mos & Codigo_ven & Codigo_bal & Lineas & sLinCan
            MiEnvio = MiEnvio & precio & hora & "0" & "0" & sCodCli & Fecha
            
            lEX = True
            If (!STPETIC And 64) = 0 Then '2.0.a no exportar cancelados...
                Print #Fichero1, MiEnvio
            Else
                lEX = False
            End If
            
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
                        Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        Cod_Cod = Format(!Code, "000000")
                        
                        If !price >= 0 Then
                            Precio_Uni = !price
                            Precio_Uni = Format(100 * (Precio_Uni), "0000000")
                        Else
                            Precio_Uni = (-1) * !price
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
                        If Dir(App.Path & "\upper.txt") <> "" Then
                            MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Format(!secc, "00") & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                        Else
                            MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Cod_Cod & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                        End If
                        'MiEnvio = MiEnvio & LineaAnulada & Total & "1"
                        
                        If (lEX) Then
                            Print #Fichero2, MiEnvio
                        End If
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.MoveNext
                    
                    End With
                    
                
                Loop
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Registro.MoveNext
        
            End With
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    Close #Fichero1
    Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    lCogeTiquet = True

End Sub
Public Sub export_BL(ByVal lTodos As Boolean)
Dim Base As DAO.Database
Dim Registro As DAO.Recordset
Dim Reg2 As DAO.Recordset
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
Dim MiEnvio, Mienvio2 As String
Dim nintentos As Integer
'Dim Bucle As Integer

Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim Fichero2 As Integer
Dim sLinCan As String
Dim sCodCli As String
Dim nSysToca As Integer
Dim sFich As String

'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
    
    
    Set Base = AbrirBase()
    'If lTodos Then 'euroscale
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "CabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
                  ") order by val(nume)"
    
    CadenadeLog sSQL
    
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        
        
        Registro.MoveFirst
        
        
        Do While Not Registro.EOF And CI_Salir = False
        
            With Registro
            
            Codigo_mos = Format(Val(!seccion), "00")
            Codigo_bal = Format(Val(!Balanza), "00")
            Codigo_ven = Mid(Format(Val(!vended), "0000"), 3, 2)
            Lineas = Format(Val(!numlin) - Val(!lincan), "000") ' No se exportan las canceladas...
            Fecha = Format(!Fecha, "yyyymmdd")
            hora = Format(!hora, "hhmm")
            'código de Cliente se pasa a 4 de longitud según Informático.
            'sCodCli = Format(Val(!codcli), "00000")
            sCodCli = Format(Val(!codcli), "0000")
            
            MiEnvio = Codigo_mos & Format(Val(!nume), "00000")
            
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & ") order by val(numlin)")

            If Not Reg2.EOF Then
                Reg2.MoveFirst
                
               
                Do While Not Reg2.EOF
                    With Reg2
                        Num_lin = Format(!numlin, "000")
                        'Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        Cod_Cod = Format(!Code, "000000")
                        
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                Pes_Uni = Format(1000 * (Pes_Uni), "000000")
                            Else
                                Pes_Uni = (-1) * !Weight
                                Pes_Uni = "-" & Format(1000 * (Pes_Uni), "00000")
                            End If
                            Tipo_Linea = "0"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                Pes_Uni = Format((Pes_Uni) * 1000, "000000")
                            Else
                                Pes_Uni = (-1) * !units
                                Pes_Uni = "-" & Format((Pes_Uni) * 1000, "00000")
                            End If
                            Tipo_Linea = "1"
                        End If
                        
                        If !amount < 0 Then
                            Total = (-1) * !amount
                            Total = Format(100 * (Total), "000000000")
                            Signo = "1"
                        Else
                            Total = !amount
                            Total = Format(100 * (Total), "000000000")
                            Signo = "0"
                        End If
                        If !lincan = 0 Then
                            LineaAnulada = "0"
                        Else
                            LineaAnulada = "1"
                        End If
                        
                        If !lincan = 0 Then
                            
                            Mienvio2 = ""
                            
                            If Val(Mid(Fecha, 5, 2)) < 10 Then
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Format(Val(Mid(Fecha, 5, 2)), "0") & ".TOT"
                            Else
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Chr(55 + Val(Mid(Fecha, 5, 2))) & ".TOT"
                            End If
                            
                            Mienvio2 = MiEnvio & Num_lin & Cod_Cod & Pes_Uni & Total & Codigo_ven & Fecha & hora & sCodCli
                            
                            Fichero1 = FreeFile()
                            Open var.PathExpor & "\" & sFich For Append Shared As #Fichero1
                            Print #Fichero1, Mienvio2
                            Close #Fichero1
                            
                        End If
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.MoveNext
                    
                    End With
                    
                
                Loop
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Registro.MoveNext
        
            End With
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    Close #Fichero1
    Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    lCogeTiquet = True

End Sub
Public Sub export_BDP(ByVal lTodos As Boolean)
Dim Base As Database
Dim Registro As Recordset
Dim Reg2 As Recordset
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
Dim MiEnvio, Mienvio2 As String
Dim nintentos As Integer
'Dim Bucle As Integer

Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim Fichero2 As Integer
Dim sLinCan As String
Dim sCodCli As String
Dim snumtiq As String
Dim sImporte6 As String
Dim sIporte5 As String
Dim sPre As String
Dim sCab As String

Dim nSysToca As Integer
Dim sFich As String

'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
'c2f 1.4.4 no se exportan lineas canceladas...
    
    Set Base = AbrirBase()
    
    MiTabla1 = "Cabecera"
    MiTabla2 = "tickets"
    
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
                  ") order by val(nume)"
    
    CadenadeLog sSQL
    
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        
        
        Registro.MoveFirst
        
        
        Do While Not Registro.EOF And CI_Salir = False
        
            With Registro
            
            snumtiq = Format(!nume, "0000")
            sImporte6 = Format((!Importe) * 100, "000000")
            'simporte5 = Format((!Importe) * 100, "00000")
            Codigo_mos = Format(Val(!seccion), "00")
            'Codigo_bal = Format(Val(!Balanza), "00")
            Codigo_ven = Format(Val(!vended), "00000")
            'Lineas = Format(Val(!numlin) - Val(!lincan), "000") ' No se exportan las canceladas...
            Fecha = Format(!Fecha, "ddmmyy")
            hora = Format(!hora, "hhmm")
            'código de Cliente se pasa a 4 de longitud según Informático.
            'sCodCli = Format(Val(!codcli), "00000")
            'sCodCli = Format(Val(!codcli), "0000")
            

            sFich = (!codbar) & ".txt"
            
            'sCab = sImporte6 & Fecha & hora & Codigo_ven & Codigo_mos & snumtiq
            sCab = Fecha & hora & Codigo_ven & Codigo_mos & snumtiq
            
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & " and lincan=" & Chr(34) & "0" & Chr(34) & ") order by val(numlin)")

            If Not Reg2.EOF Then
                Reg2.MoveFirst
                
                
                Fichero1 = FreeFile()
                Open var.PathExpor & "\" & sFich For Output As #Fichero1
          
                Do While Not Reg2.EOF
                    With Reg2
                        MiEnvio = ""
                        'Num_lin = Format(!numlin, "000")
                        'Cod_Plu = Format(CodigoToPlu(!Code), "0000")
                        Cod_Cod = Format(!Code, "00000")
                        sPre = Format((!price) * 100, "00000")
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                Pes_Uni = Format(1000 * (Pes_Uni), "000000")
                            Else
                                Pes_Uni = (-1) * !Weight
                                Pes_Uni = "-" & Format(1000 * (Pes_Uni), "00000")
                            End If
                            'Tipo_Linea = "0"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                'Fanjul 21/02/2013
                                'Pes_Uni = Format((Pes_Uni), "000000")
                                Pes_Uni = Format((Pes_Uni) * 1000, "000000")
                                ''''''''''''''''''''''''''''''''''''''
                            Else
                                Pes_Uni = (-1) * !units
                                'Fanjul 21/02/2013
                                'Pes_Uni = "-" & Format((Pes_Uni), "00000")
                                Pes_Uni = "-" & Format((Pes_Uni) * 1000, "00000")
                            End If
                            'Tipo_Linea = "1"
                        End If
                        
                        If !amount < 0 Then
                            Total = (-1) * !amount
                            Total = "-" & Format(100 * (Total), "00000") 'c2f 1.4.4 se deja negativo.
                        '    Signo = "1"
                        Else
                            Total = !amount
                            Total = Format(100 * (Total), "000000")
                        '    Signo = "0"
                        End If
                        'If !lincan = 0 Then
                        '    LineaAnulada = "0"
                        'Else
                        '    LineaAnulada = "1"
                        'End If
                        
                        'If !lincan = 0 Then
                            
                        '    Mienvio2 = ""
                            
                        '    If Val(Mid(Fecha, 5, 2)) < 10 Then
                        '        sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Format(Val(Mid(Fecha, 5, 2)), "0") & ".TOT"
                        '    Else
                        '        sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Chr(55 + Val(Mid(Fecha, 5, 2))) & ".TOT"
                        '    End If
                            
                        '    Mienvio2 = MiEnvio & Num_lin & Cod_Cod & Pes_Uni & Total & Codigo_ven & Fecha & hora & sCodCli
                            
                        '    Fichero1 = FreeFile()
                        '    Open var.PathExpor & "\" & sFich For Append Shared As #Fichero1
                        '    Print #Fichero1, Mienvio2
                        '    Close #Fichero1
                            
                        'End If
                        
                        MiEnvio = Cod_Cod & Pes_Uni & sPre & Total & sCab
                        
                        Print #Fichero1, MiEnvio
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.MoveNext
                    
                    End With
                    
                
                Loop
                
                Close #Fichero1
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Registro.MoveNext
        
            End With
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    'Close #Fichero1
    'Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    lCogeTiquet = True

End Sub

Public Sub export_MT(ByVal lTodos As Boolean)
Dim Base As DAO.Database
Dim Registro As DAO.Recordset
Dim Reg2 As DAO.Recordset
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

Dim sPrice As String

Dim Signo As String
Dim Tipo_Linea As String
Dim Anulado As String
'************
Dim MiEnvio, Mienvio2 As String
Dim nintentos As Integer
'Dim Bucle As Integer

Dim MiTabla1 As String
Dim MiTabla2 As String
Dim Fichero1 As Integer
Dim Fichero2 As Integer
Dim sLinCan As String
Dim sCodCli As String
Dim snumtiq As String
Dim sImporte6 As String
Dim sPeso As String
Dim sImporte5 As String
Dim sPre As String
Dim sCab As String

Dim nSysToca As Integer
Dim sFich As String
Dim sMT As String
Dim nPeso As Double
Dim nUnidades As Long
Dim nContLin As Integer
Dim sVen As String

'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
    
Dim nFichMT As Integer

CadenadeLog "export_MT"

    If Dir(App.Path & "\tiendamt.txt") <> "" Then
        nFichMT = FreeFile()
        Open App.Path & "\tiendamt.txt" For Input As #nFichMT
        Line Input #nFichMT, sMT
        Close #nFichMT
        If Val(sMT) > 0 And Val(sMT) < 100 Then
        Else
            sMT = "01"
        End If
    Else
        sMT = "01"
    End If

    Set Base = AbrirBase()
    
    MiTabla1 = "Cabecera"
    MiTabla2 = "tickets"
    
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
                  ") order by val(nume)"
    
    CadenadeLog sSQL
    
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        
        
        Registro.MoveFirst
        
        
        Do While Not Registro.EOF And CI_Salir = False
        
            With Registro
            '"40" & "01" & sMT & "0" & --> inicio encabezado cabecera tique
            snumtiq = Format(!nume, "000000")
            If (!Importe) >= 0 Then
                sImporte6 = Format((!Importe) * 100, "0000000000")
            Else
                sImporte6 = Format((!Importe) * 100, "000000000")
                sImporte6 = "-" & sImporte6
            End If
            sImporte5 = Format(Abs((!Importe) * 100), "00000")
            sImporte6 = Mid(sImporte6, 1, 8) & "," & Mid(sImporte6, 9)
            Codigo_mos = Format(Val(!seccion), "0000")
            Codigo_bal = Format(Val(!Balanza), "00")
            If (!TYPTIC) = 1 Or (!TYPTIC) = 5 Or (!TYPTIC) = 2 Or (!TYPTIC) = 6 Then
                If (!vended) < 99 Then
                    Codigo_ven = Format(Val(!vended), "00")
                Else
                    Codigo_ven = Mid((!vended), Len(!vended) - 1)
                End If
            Else
                If (!TYPTIC) = 4 Or (!TYPTIC) = 8 Then
                    Codigo_ven = "SS"
                Else
                    Codigo_ven = "PP"
                End If
            End If
            sVen = Codigo_ven
            If (Val(!STPETIC) And 64) Or (Val(!STPETIC) And 128) Then
                Codigo_ven = Codigo_ven & "1" ' se le añade cancelado o no
            Else
                Codigo_ven = Codigo_ven & "0" ' se le añade cancelado o no
            End If
            Codigo_ven = Codigo_ven & "0" ' tipo de venta se coloca "0"
            Lineas = Format(Val(!numlin) - Val(!lincan), "00000000000") ' No se exportan las canceladas...
            Fecha = CStr(Format(!Fecha, "dd-mm-yy"))
            hora = CStr(Format(!hora, "hh:mm"))
            'código de Cliente se pasa a 11 de longitud
            If (!codcli) = 0 Then
                sCodCli = Space(11)
            Else
                sCodCli = Format(Val(!codcli), "###########")
            End If

            

            sFich = snumtiq & sImporte5
            sFich = sFich & "." & Mid(Fecha, 1, 2)
            Select Case Val(Mid(Fecha, 4, 2))
                Case 1 To 9
                    sFich = sFich & Format(Val(Mid(Fecha, 4, 2)), "0")
                Case 10 To 12
                    sFich = sFich & Chr(65 + Val(Mid(Fecha, 4, 2)) - 10)
                Case Else
                    sFich = sFich & "0"
            End Select
            
            
            'sCab = sImporte6 & Fecha & hora & Codigo_ven & Codigo_mos & snumtiq
            'sCab = Fecha & hora & Codigo_ven & Codigo_mos & snumtiq
            
            nPeso = 0
            nUnidades = 0
            
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & ") order by val(numlin)")

            If Not Reg2.EOF Then
                Reg2.MoveFirst
                
                Fichero1 = FreeFile()
                Open var.PathExpor & "\" & sFich For Output As #Fichero1
                nContLin = 0
                Do While Not Reg2.EOF
                    With Reg2

                        'Cod_Cod = Format(!Code, "000000")
                        Cod_Plu = Format(CodigoToPlu(!Code), "000000")
                        MiEnvio = "4101" & sMT & "0" & snumtiq & Fecha & Format(!numlin, "00") & Codigo_mos & Codigo_bal & Space(15) & Space(20) & Cod_Plu
                        
                        If !lincan = 0 Then
                            LineaAnulada = "0"
                        Else
                            LineaAnulada = "1"
                        End If
                        

                        'sPre = Format((!price) * 100, "00000")

                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                Pes_Uni = Format(1000 * (Pes_Uni), "0000000000")
                            Else
                                Pes_Uni = (-1) * !Weight
                                Pes_Uni = "-" & Format(1000 * (Pes_Uni), "000000000")
                            End If
                            Pes_Uni = Mid(Pes_Uni, 1, 7) & "," & Mid(Pes_Uni, 8)
                            If LineaAnulada = "0" Then
                                nPeso = nPeso + ((!Weight) * 1000)
                            End If
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                Pes_Uni = Format((Pes_Uni), "00000000000")
                            Else
                                Pes_Uni = (-1) * !units
                                Pes_Uni = "-" & Format((Pes_Uni), "0000000000")
                            End If
                            If LineaAnulada = "0" Then
                                nUnidades = nUnidades + (!units)
                            End If
                        End If

                        If !amount < 0 Then
                            Total = (-1) * !amount
                            Total = "-" & Format(100 * (Total), "000000000")
                        Else
                            Total = !amount
                            Total = Format(100 * (Total), "0000000000")
                        End If
                        Total = Mid(Total, 1, 8) & "," & Mid(Total, 9)
                        
                        'cambios _MT
                        If !price < 0 Then
                            sPrice = (-1) * !price
                            sPrice = "-" & Format(100 * (sPrice), "000000")
                        Else
                            sPrice = !price
                            sPrice = Format(100 * (sPrice), "0000000")
                        End If
                        sPrice = Mid(sPrice, 1, 5) & "," & Mid(sPrice, 6)
                        
                        
                        ''''''''''''
                        
                        MiEnvio = MiEnvio & Pes_Uni & Total & Space(11) & LineaAnulada & sVen & Mid(Fecha, 1, 2) & _
                                  Mid(Fecha, 4, 2) & Mid(Fecha, 7, 2) & Mid(hora, 1, 2) & Mid(hora, 4, 2) & sPrice
                        
                        Print #Fichero1, MiEnvio
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.MoveNext
                    
                    End With
                    
                
                Loop
                
                If nPeso > 0 Then
                    sPeso = Format(nPeso, "0000000000")
                Else
                    sPeso = "-" & Format((-1) * nPeso, "000000000")
                End If
                sPeso = Mid(sPeso, 1, 7) & "," & Mid(sPeso, 8)
                sCab = "40" & "01" & sMT & "0" & snumtiq & Fecha & hora & Codigo_mos & Codigo_bal & Codigo_ven & _
                       Lineas & sPeso & sImporte6 & sCodCli
                       'Format(nUnidades, "00000000000") & sPeso & sImporte6 & sCodCli
                Print #Fichero1, sCab
                Close #Fichero1
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Registro.MoveNext
        
            End With
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    'Close #Fichero1
    'Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    lCogeTiquet = True

End Sub

Public Sub CG_PideTicket_ah(MachineID As Integer, aModTrab As Integer)
'//////////////////////////
'// PETICIÓN DE TIQUETS  //
'// PARA EUROSCALE,      //
'// SC10 (NO PATH RED)   //
'// Y C9/C10             //
'//////////////////////////
' aModTrab => modo de trabajo
' 0 --> Venta
' 1 --> Super
' 2 --> Envasado
' 3 --> AutoServicio
' 4 --> Local
'///////////////////////////
' aTipoTqt => tipo de tiquet
' 1 --> 386
' 2 --> C9/C10
' 3 --> SC10
' 4 --> IV4TQ
'\\\\\\\\\\\\
Dim aTipoTqt As Integer
Dim MiContador As Integer
Dim Bucle As Long
Dim MiEnvio
Dim MiRecibe As Long
Dim PedirTiquet As String
Dim Cabecera As tipo_cabecera
Dim Linea() As Tipo_Linea
Dim Repetir As Boolean
Dim CuentaLineas As Integer
Dim NroLineas As Integer
Dim LineaInicial As Integer
Dim LineaComun As String
Dim ToTalLineas As Integer
Dim ErrorLineas As Boolean
Dim Buffer As String
Dim MiTipo As String
Dim Base As DAO.Database
Dim RegArt As DAO.Recordset
Dim BalTick As Epel_CabTick
Dim baseID As String
Dim RegTramo As DAO.Recordset
Dim HayTramos As Boolean
Dim MyArt As DB_Articulo
Dim ScaleArt As DB_Articulo
Dim OKCambiar As Boolean
Dim CSiempre As Boolean
Dim ContadorTQ As Long
Dim SeccionTQ As Integer
Dim MaquinaTQ As Integer
Dim RegTQ As DAO.Recordset
Dim tp As Variant
Dim ce As Variant
'*********************cas.v1.1.6
Dim lclImporte As Long
Dim cnt As Long
'*********************cas.v1.1.8
Dim sSQL As String
Dim codArt As Long
'/////////////////

    aTipoTqt = 0
    
    ContadorTQ = 1
    If Dir(App.Path & "\csiempre.txt") <> "" Then
        CSiempre = True
    Else
        CSiempre = False
    End If
    
    Switch_mode nMaqType
    
    'If MachineID >= 700 And MachineID < 800 Then MachineID = 0
    'MachineID = 0
    'Select Case MachineID
    '    Case 0
    '        baseID = "SC10"
    '        MiTipo = aModTrab
    '    'Case 1 To 2
    '    '    baseID = "SC10"
    '    '    If TipoC9 Then
    '    '        frmControl.Balanza.C9_TickRewind
    '    '    End If
    '    '    Select Case aModTrab
    '    '        Case 0
    '    '            MiTipo = 0
    '    '        Case 1
    '    '            MiTipo = 1
    '    '        Case 2
    '    '            MiTipo = 4
    '    '        Case 3
    '    '            MiTipo = 2
    '    '        Case 4
    '    '            MiTipo = 3
    '    '    End Select
    '    'Case 501 To 599
    '    '    baseID = "TQ"
    '    '    MiTipo = 0
    'End Select
    Cabecera.Ntiquet = 0
    Repetir = True
    
    Set Base = AbrirBase()
    'If MachineID > 500 And MachineID < 600 Then 'cas.n->si es gama baja
    '    MaquinaTQ = MachineID - 500
    '    Set RegTQ = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & MaquinaTQ)
    '    SeccionTQ = RegTQ.Fields("codi_ident")
    'End If
    Do Until (Repetir = False) Or (ErrorLineas = True)
        'frmEpelsa.CmdComunicaciones(3).Enabled = False
        'frmEpelsa.CmdComunicaciones(4).Enabled = False
        'If MachineID > 500 And MachineID < 600 Then ' IV4TQ
        '    MiRecibe = frmControl.Balanza.Query_One_Ticket(Val(MiTipo), 0, 0, ContadorTQ, BalTick)
        '    ContadorTQ = ContadorTQ + 1
        'Else
            ' Resto de balanzas
            MiRecibe = frmControl.Balanza.Query_Ticket(Val(MiTipo), BalTick)
        'End If
        'frmEpelsa.CmdComunicaciones(3).Enabled = True
        'frmEpelsa.CmdComunicaciones(4).Enabled = True
        If MiRecibe >= 1000 Or MiRecibe = 4 Then
            
            Exit Do
        End If
        If MiRecibe = 4 Then
            Repetir = False
        Else
            '***
            ' si STS=0, lo procesa (tiquet cerrado)
            '***
            If MiRecibe = 0 Then
                Repetir = True
                '********************
                ' Toma los datos de Cabecera
                '********************
                
                Cabecera.NLineas = BalTick.nLines
                ReDim Linea(Cabecera.NLineas + 1)
                Cabecera.LineasMensaje = BalTick.nLines
                Cabecera.LineaFinal = BalTick.nLines
                Cabecera.LineaInicial = 1
                Cabecera.LineasCanceladas = 0
                Cabecera.Ntiquet = BalTick.NTicket
                Cabecera.NMostrador = BalTick.Section
                Cabecera.NVendedor = BalTick.Vendor
                Cabecera.NBalanza = BalTick.Machine
                Cabecera.NCliente = BalTick.Client
                If BalTick.EAN13 = "(null)" Then BalTick.EAN13 = Space(13)
                '////////////////////////////////////////
                ' IV4TQ, recompone el mostrador y balanza
                'If MachineID > 500 And MachineID < 600 Then
                '    Cabecera.NBalanza = MaquinaTQ
                '    Cabecera.NMostrador = SeccionTQ
                'End If
                '////////////////////////////////////////
                Buffer = BalTick.Hour
                Cabecera.hora = Buffer 'TimeSerial(Val(Left(Buffer, 2)), Val(Mid(Buffer, 3, 2)), Val(Right(Buffer, 2)))
                Buffer = BalTick.Date
                Cabecera.Fecha = Buffer 'DateSerial(Val(Right(Buffer, 2)), Val(Mid(Buffer, 3, 2)), Val(Left(Buffer, 2)))
                If BalTick.positive = 1 Then
                    Cabecera.ImporteTotal = BalTick.amount
                Else
                    Cabecera.ImporteTotal = (-1) * BalTick.amount
                End If
                If MachineID <> 0 Then
                    Cabecera.ImporteBruto = Cabecera.ImporteTotal
                    Cabecera.descuento = 0
                Else
                    Cabecera.ImporteBruto = BalTick.P1
                    Cabecera.descuento = BalTick.P2
                End If
                '///cas.v117
                Cabecera.tipoPago = BalTick.tipoPago
                Cabecera.cantidadEntre = BalTick.cantidadEntre
                Cabecera.Cambio = Cabecera.cantidadEntre - Cabecera.ImporteBruto
                If (Cabecera.tipoPago = 255) Or (Cabecera.cantidadEntre = 0) Then Cabecera.Cambio = 0
                '///cas.v117
                If Not (IsNumeric(MiTipo)) Then MiTipo = 0
                If aModTrab > 3 Then
                    'Cabecera.EstadoTiquet = 17
                    Cabecera.CadenaEstadoTiquet = "Local"
                    'Cabecera.tipoTiquet = MiTipo + 1
                Else
                    'Cabecera.EstadoTiquet = 16
                    Cabecera.CadenaEstadoTiquet = "Red"
                    'Cabecera.tipoTiquet = MiTipo + 1
                End If
                If BalTick.Cancelled = 1 Then
                    'Cabecera.EstadoTiquet = Cabecera.EstadoTiquet + 64
                    Cabecera.CadenaEstadoTiquet = Cabecera.CadenaEstadoTiquet & " Cancelado"
                End If
                
                Cabecera.EstadoTiquet = BalTick.EstadoTiquet
                
                Cabecera.lote = BalTick.lote
                
                If IsNull(Cabecera.lote) Then
                    Cabecera.lote = Space(20)
                Else
                    If Len(Cabecera.lote) = 0 Then
                        Cabecera.lote = Space(20)
                    End If
                End If
                
                '*******************
                ' Pide las lineas
                '*******************
                CuentaLineas = 1
                LineaInicial = 1
                ToTalLineas = Cabecera.NLineas
                For Bucle = 1 To Cabecera.NLineas
                    'STV (positivo o negativo)
                    Linea(Bucle).ModificadoTemporal = False
                    If BalTick.Lines(Bucle - 1).positive = 1 Then
                         Linea(Bucle).FactorPeso = 1
                         Linea(Bucle).TipoLinea = 0
                    Else
                         Linea(Bucle).FactorPeso = -1
                         Linea(Bucle).TipoLinea = 2
                    End If
                    'STP (pesado o no pesado)
                    If BalTick.Lines(Bucle - 1).Weighted = 1 Then
                         Linea(Bucle).EsPesado = True
                    Else
                         Linea(Bucle).EsPesado = False
                         Linea(Bucle).TipoLinea = Linea(Bucle).TipoLinea + 1
                    End If
                    ' STC (cancelado o no cancelado)
                    If BalTick.Lines(Bucle - 1).Cancelled = 0 Then
                         Linea(Bucle).LineaCancelada = False
                    Else
                         Cabecera.LineasCanceladas = Cabecera.LineasCanceladas + 1
                         Linea(Bucle).LineaCancelada = True
                         Linea(Bucle).TipoLinea = Linea(Bucle).TipoLinea + 16
                    End If
                    Linea(Bucle).FichaVacuno = BalTick.Lines(Bucle - 1).BovineCard
                    Linea(Bucle).NumeroLinea = BalTick.Lines(Bucle - 1).Number
                    Linea(Bucle).DescuentoLinea = BalTick.Lines(Bucle - 1).discount
                    Linea(Bucle).CodigoArticulo = BalTick.Lines(Bucle - 1).item_code
                    '**************************************
                    ' IV4TQ recompone el código de artículo
                    '**************************************
                    Linea(Bucle).precio = BalTick.Lines(Bucle - 1).Item_Price
                    If BalTick.Lines(Bucle - 1).positive = 1 Then
                        Linea(Bucle).ImporteLinea = BalTick.Lines(Bucle - 1).amount
                        Linea(Bucle).cantidad = BalTick.Lines(Bucle - 1).Item_Pes
                    Else
                        Linea(Bucle).ImporteLinea = BalTick.Lines(Bucle - 1).amount
                        Linea(Bucle).cantidad = BalTick.Lines(Bucle - 1).Item_Pes
                    End If
                    '*****************************
                    ''''VER If Linea(Bucle).EsPesado = True Then Linea(Bucle).cantidad = (Linea(Bucle).cantidad / 1000)
                    '**************************
                    ' ML100 --> Si el peso está
                    ' entre 15 y 30 kg, sólo
                    ' proporciona dos decimales
                    ' de peso, y hay que corregirlo
                    '******************************
                    '******************************cas.v1.1.6
                    'If ML100 And aTipoTqt <> 2 And (MachineID = 0 Or MachineID = 1) And Linea(Bucle).EsPesado Then
                    '    If Abs(Linea(Bucle).ImporteLinea) >= 999 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                    '        Linea(Bucle).cantidad = Linea(Bucle).cantidad * 1000
                    '    Else
                    '       If Abs(Linea(Bucle).ImporteLinea) >= 99 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                    '           Linea(Bucle).cantidad = Linea(Bucle).cantidad * 100
                    '       Else
                    '           If Abs(Linea(Bucle).ImporteLinea) >= 9 * (Abs(Linea(Bucle).cantidad * Linea(Bucle).precio)) Then
                    '               Linea(Bucle).cantidad = Linea(Bucle).cantidad * 10
                    '           End If
                    '       End If
                    '    End If
                    'End If
                    '***************************************cas.v1.1.6
                    ' verifica si se ha modificado el precio
                    '***************************************
                    If Linea(Bucle).CodigoArticulo <> 0 Then
                        If aTipoTqt = 2 Then
                            sSQL = "SELECT codigo " & _
                                   "FROM articulo " & _
                                   "WHERE plu=" & Linea(Bucle).CodigoArticulo & _
                                   " AND codi_ident=" & Cabecera.NMostrador & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If Not RegArt.EOF Then
                                codArt = RegArt.Fields("codigo")
                                Linea(Bucle).CodigoArticulo = codArt
                            End If
                            RegArt.Close
                            Set RegArt = Nothing
                        Else
                            codArt = Linea(Bucle).CodigoArticulo
                        End If
                        If Not (TipoC9) And Not (MachineID > 500 And MachineID < 600) Then
                            sSQL = "SELECT * " & _
                                   "FROM articulo " & _
                                   "WHERE codigo=" & CStr(codArt) & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If RegArt.EOF Then
                                 Linea(Bucle).SehaModificado = False
                                 Linea(Bucle).IVA = 0
                            Else
                                ScaleArt.codigo = 0
                                Linea(Bucle).IVA = RegArt!tipo_iva
                                sSQL = "SELECT * " & _
                                       "FROM tramos " & _
                                       "WHERE borrado=false and codigo=" & CStr(codArt)
                                Set RegTramo = Base.OpenRecordset(sSQL)
                                If RegTramo.EOF Then
                                   HayTramos = False
                                Else
                                   HayTramos = True
                                End If
                                RegTramo.Close
                                If HayTramos = False Then
                                   If UsaEuro Then
                                       If (Linea(Bucle).precio / 100) <> RegArt!Euros Then
                                           Linea(Bucle).SehaModificado = True
                                       Else
                                           Linea(Bucle).SehaModificado = False
                                       End If
                                   Else
                                       If (Linea(Bucle).precio / (10 ^ decimales)) <> RegArt!precio Then
                                           Linea(Bucle).SehaModificado = True
                                       Else
                                           Linea(Bucle).SehaModificado = False
                                       End If
                                   End If
                                   'If Linea(Bucle).SehaModificado And (MachineID = 0 Or MachineID = 1) Then
                                   'If Linea(Bucle).SehaModificado And Not (MachineID > 500 And MachineID < 600) And Not (TipoC9) Then
                                   '    '*****************************************
                                   '    '* consulta el precio real del artículo  *
                                   '    '*****************************************
                                   '    If MachineID = 0 Then
                                   '        ScaleArt = Binario_PideArticulo(RegArt!codigo, False)
                                   '    Else
                                   '        If TipoC9 Then
                                   '            ScaleArt = FrmSC10_PideArticulo((10000 * RegArt!secc_Maqui) + RegArt!plu, False, True)
                                   '        Else
                                   '            ScaleArt = FrmSC10_PideArticulo(RegArt!codigo, False, False)
                                   '        End If
                                   '    End If
                                   '    If UsaEuro Then
                                   '        ScaleArt.precio = ScaleArt.precio / 100
                                   '    Else
                                   '        ScaleArt.precio = ScaleArt.precio / (10 ^ decimales)
                                   '    End If
                                   '    If (UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!Euros)) Or _
                                   '    (Not UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!precio)) Then
                                   '        If AbsorverCambios Then
                                   '            MyArt.codigo = Linea(Bucle).CodigoArticulo
                                   '            MyArt.plu = 0
                                   '            MyArt.Mostrador = 0
                                   '            MyArt.precio = ScaleArt.precio
                                   '            If CSiempre Then
                                   '                OKCambiar = True
                                   '            Else
                                   '                OKCambiar = False
                                   '                If IsNull(RegArt!prc3) Then
                                   '                    OKCambiar = True
                                   '                Else
                                   '                    If RegArt!prc3 < (Val(Format(Cabecera.Fecha, "yymmdd")) * 10000 + Val(Format(Cabecera.hora, "hhmm"))) Then OKCambiar = True
                                   '                End If
                                   '            End If
                                   '            If OKCambiar Then
                                   '                Modificacion_Articulo_Precio MyArt, "Balanzas", True
                                   '            End If
                                   '        End If
                                   '    Else
                                   '        Linea(Bucle).ModificadoTemporal = True
                                   '        Linea(Bucle).SehaModificado = False
                                   '    End If
                                   'End If
                                End If
                            End If
                            RegArt.Close
                            Set RegArt = Nothing
                        End If
                    End If
                    MiContador = MiContador + 1
                Next Bucle
                '********************************
                ' Actualiza las tablas de totales
                '********************************
                Cabecera.Incluir_en_Base = True
                Actualiza_Tabla_Cabecera Base, Cabecera
                If Cabecera.Incluir_en_Base = True Then
                    Actualiza_Tabla_Tickets Base, Cabecera
                    CadenadeLog "Actualizar Total artículos..."
                    Actualiza_Tabla_Gtarti Base, Cabecera
                    CadenadeLog "Actualizar Total secciones..."
                    Actualiza_Tabla_Gtsecs Base, Cabecera
                    CadenadeLog "Actualizar Total vendedores..."
                    Actualiza_Tabla_GtVend Base, Cabecera
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v118
                    CadenadeLog "Actualizar Total info Vende...."
                    actualiza_tabla_gtInfVen Base, Cabecera
                    CadenadeLog "Actualizar Total Horario..."
                    actualiza_tabla_gtHora Base, Cabecera '\\\cas.v170
                    
                    '///////////////////////////////////////////////////
                    CadenadeLog "Actualizar Log..."
                    Actualiza_Tabla_Log Base, Cabecera
                    CadenadeLog "Fin actualización totales..."
                End If
            End If
        End If
        If cancelar = 1 Then Repetir = False
    Loop
    Base.Close
End Sub

Public Sub CG_PideTicket(MachineID As Integer, aModTrab As Integer)
'//////////////////////////
'// PETICIÓN DE TIQUETS  //
'// PARA EUROSCALE,      //
'// SC10 (NO PATH RED)   //
'// Y C9/C10             //
'//////////////////////////
'************
' aModTrab
' 0 --> Venta
' 1 --> Super
' 2 --> Envasado
' 3 --> Autoservicio
' 4 --> Venta Local
' 5 --> Super Local
' 6 --> Envasado Local
' 7 --> Autoservicio Local
' 9 --> anulados....
'*************************
'///////////////////////////
' aTipoTqt => tipo de tiquet
' 1 --> 386
' 2 --> C9/C10
' 3 --> SC10
' 4 --> IV4TQ
'\\\\\\\\\\\\
Dim aTipoTqt As Integer
Dim MiContador As Integer
Dim Bucle As Long
Dim MiEnvio
Dim MiRecibe As Long
Dim PedirTiquet As String
Dim Cabecera As tipo_cabecera
'Dim Linea() As Tipo_Linea
Dim Repetir As Boolean
Dim CuentaLineas As Integer
Dim NroLineas As Integer
Dim LineaInicial As Integer
Dim LineaComun As String
Dim ToTalLineas As Integer
Dim ErrorLineas As Boolean
Dim Buffer As String
Dim MiTipo As String
Dim Base As DAO.Database
Dim RegArt As DAO.Recordset
Dim BalTick As Epel_CabTick
Dim baseID As String
Dim RegTramo As DAO.Recordset
Dim HayTramos As Boolean
Dim MyArt As DB_Articulo
Dim ScaleArt As DB_Articulo
Dim OKCambiar As Boolean
Dim CSiempre As Boolean
Dim ContadorTQ As Long
Dim SeccionTQ As Integer
Dim MaquinaTQ As Integer
Dim RegTQ As DAO.Recordset
Dim tp As Variant
Dim ce As Variant
'*********************cas.v1.1.6
Dim lclImporte As Long
Dim cnt As Long
'*********************cas.v1.1.8
Dim sSQL As String
Dim codArt As Long
'/////////////////

    aTipoTqt = 0
    
    ContadorTQ = 1
    If Dir(App.Path & "\csiempre.txt") <> "" Then
        CSiempre = True
    Else
        CSiempre = False
    End If
    
    
    '1.4.1 upper
    Switch_mode nMaqType
    ''''''
    'If nMaqType = 0 Then
    '    baseID = "GA"
    'Else
    '    baseID = "SC10"
    'End If
    MiTipo = aModTrab
    ''''''''''''
    
    Cabecera.Ntiquet = 0
    Repetir = True
    
    '1.4.1
    'Set Base = AbrirBase()
    'If MachineID > 500 And MachineID < 600 Then 'cas.n->si es gama baja
    '    MaquinaTQ = MachineID - 500
    '    Set RegTQ = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & MaquinaTQ)
    '    SeccionTQ = RegTQ.Fields("codi_ident")
    'End If
    '''''''
    Do Until (Repetir = False) Or (ErrorLineas = True)
        
        If (Dir(App.Path & "\hydrarun.txt") <> "") Then
            CadenadeLog "Detectado  HydraRun.txt en petición Tique desde TIQAH."
            '''''''''''''''''''''''''
            frmControl.Balanza.Reset
            '''''''''''''''''''''''''
           Repetir = False
           Exit Sub
        End If
        
        'frmEpelsa.CmdComunicaciones(3).Enabled = False
        'frmEpelsa.CmdComunicaciones(4).Enabled = False
        'If MachineID > 500 And MachineID < 600 Then ' IV4TQ
        '    MiRecibe = frmControl.Balanza.Query_One_Ticket(Val(MiTipo), 0, 0, ContadorTQ, BalTick)
        '    ContadorTQ = ContadorTQ + 1
        'Else
            ' Resto de balanzas
            
            
            If nMaqType = 0 Then
                'compatibilidad...TiqahTouch
                'MiRecibe = frmControl.Balanza.Query_TicketH(Val(MiTipo), BalTick)
                ''''''''''''''''''
            Else
                MiRecibe = frmControl.Balanza.Query_Ticket(Val(MiTipo), BalTick)
            End If
        'End If
        'frmEpelsa.CmdComunicaciones(3).Enabled = True
        'frmEpelsa.CmdComunicaciones(4).Enabled = True
        If MiRecibe >= 1000 Or MiRecibe = 4 Then
            CadenadeLog "Salida por inexistencia de tique..."
            lNoConnect = True
            Exit Do
        End If
        If MiRecibe = 4 Then
            lNoConnect = False
            Repetir = False
        Else
            '***
            ' si STS=0, lo procesa (tiquet cerrado)
            '***
            If MiRecibe = 0 Then
                
                '1.4.1
                Set Base = AbrirBase()
                ''''''
                
                Repetir = True
                '********************
                ' Toma los datos de Cabecera
                '********************
                
                Cabecera.NLineas = BalTick.nLines
                ReDim Cabecera.Lines(Cabecera.NLineas + 1)
                Cabecera.LineasMensaje = BalTick.nLines
                Cabecera.LineaFinal = BalTick.nLines
                Cabecera.LineaInicial = 1
                Cabecera.LineasCanceladas = BalTick.nLinCanceled
                Cabecera.Ntiquet = BalTick.NTicket
                Cabecera.NMostrador = BalTick.Section
                Cabecera.NVendedor = BalTick.Vendor
                Cabecera.NBalanza = BalTick.Machine
                Cabecera.NCliente = BalTick.Client
                Cabecera.lote = BalTick.lote
                If BalTick.EAN13 = "(null)" Then BalTick.EAN13 = Space(13)
                Cabecera.EAN13 = BalTick.EAN13
                'Cabecera.codigo
                If IsNull(Cabecera.lote) Then
                    Cabecera.lote = Space(20)
                Else
                    If Len(Cabecera.lote) = 0 Then
                        Cabecera.lote = Space(20)
                    End If
                End If
                Cabecera.tipoTiquet = BalTick.tipoTiquet
                
                '////////////////////////////////////////
                ' IV4TQ, recompone el mostrador y balanza
                'If MachineID > 500 And MachineID < 600 Then
                '    Cabecera.NBalanza = MaquinaTQ
                '    Cabecera.NMostrador = SeccionTQ
                'End If
                '////////////////////////////////////////
                Buffer = BalTick.Hour
                Cabecera.hora = Buffer 'Left(Buffer, 2) & ":" & Mid(Buffer, 3, 2) & ":" & Right(Buffer, 2)
                Buffer = BalTick.Date
                Cabecera.Fecha = Buffer 'DateSerial(Val(Right(Buffer, 2)), Val(Mid(Buffer, 3, 2)), Val(Left(Buffer, 2)))
                'If BalTick.positive = 1 Then
                    Cabecera.ImporteTotal = BalTick.amount
                'Else
                '    Cabecera.ImporteTotal = (-1) * BalTick.amount
                'End If
                'If MachineID <> 0 Then
                '    Cabecera.ImporteBruto = Cabecera.ImporteTotal
                '    Cabecera.descuento = 0
                'Else
                    Cabecera.ImporteBruto = BalTick.P1
                    Cabecera.descuento = BalTick.P2
                'End If
                '///cas.v117
                Cabecera.tipoPago = BalTick.tipoPago
                Cabecera.cantidadEntre = BalTick.cantidadEntre
                Cabecera.Cambio = Cabecera.cantidadEntre - Cabecera.ImporteTotal
                If (Cabecera.tipoPago = 255) Or (Cabecera.cantidadEntre = 0) Then Cabecera.Cambio = 0
                '///cas.v117
                If Not (IsNumeric(MiTipo)) Then MiTipo = 0
                If aModTrab > 3 Then
                    'Cabecera.EstadoTiquet = 17
                    Cabecera.CadenaEstadoTiquet = "Local"
                    'Cabecera.tipoTiquet = MiTipo + 1
                Else
                    'Cabecera.EstadoTiquet = 16
                    Cabecera.CadenaEstadoTiquet = "Red"
                    'Cabecera.tipoTiquet = MiTipo + 1
                End If
                If BalTick.Cancelled = 1 Then
                    'Cabecera.EstadoTiquet = Cabecera.EstadoTiquet + 64
                    Cabecera.CadenaEstadoTiquet = Cabecera.CadenaEstadoTiquet & " Cancelado"
                End If
                
                Cabecera.EstadoTiquet = BalTick.EstadoTiquet
                
                Cabecera.lote = BalTick.lote
                If IsNull(Cabecera.lote) Then
                    Cabecera.lote = Space(20)
                Else
                    If Len(Cabecera.lote) = 0 Then
                        Cabecera.lote = Space(20)
                    End If
                End If
                
                
                '*******************
                ' Pide las lineas
                '*******************
                CuentaLineas = 1
                LineaInicial = 1
                ToTalLineas = Cabecera.NLineas
                For Bucle = 1 To Cabecera.NLineas
                    'STV (positivo o negativo)
                    Cabecera.Lines(Bucle).ModificadoTemporal = False
                    If BalTick.Lines(Bucle).positive = 1 Then
                         Cabecera.Lines(Bucle).FactorPeso = 1
                         Cabecera.Lines(Bucle).TipoLinea = 0
                    Else
                         Cabecera.Lines(Bucle).FactorPeso = -1
                         Cabecera.Lines(Bucle).TipoLinea = 2
                    End If
                    'STP (pesado o no pesado)
                    If BalTick.Lines(Bucle).Weighted = True Then
                         Cabecera.Lines(Bucle).EsPesado = True
                    Else
                         Cabecera.Lines(Bucle).EsPesado = False
                         Cabecera.Lines(Bucle).TipoLinea = Cabecera.Lines(Bucle).TipoLinea + 1
                    End If
                    ' STC (cancelado o no cancelado)
                    If BalTick.Lines(Bucle).Cancelled = 0 Then
                         Cabecera.Lines(Bucle).LineaCancelada = False
                    Else
                         Cabecera.LineasCanceladas = Cabecera.LineasCanceladas + 1
                         Cabecera.Lines(Bucle).LineaCancelada = True
                         Cabecera.Lines(Bucle).TipoLinea = Cabecera.Lines(Bucle).TipoLinea + 16
                    End If
                    Cabecera.Lines(Bucle).FichaVacuno = BalTick.Lines(Bucle).BovineCard
                    Cabecera.Lines(Bucle).NumeroLinea = BalTick.Lines(Bucle).Number
                    Cabecera.Lines(Bucle).DescuentoLinea = BalTick.Lines(Bucle).discount
                    Cabecera.Lines(Bucle).CodigoArticulo = BalTick.Lines(Bucle).item_code
                    '**************************************
                    ' IV4TQ recompone el código de artículo
                    '**************************************
                    Cabecera.Lines(Bucle).precio = BalTick.Lines(Bucle).Item_Price
                    'If BalTick.Lines(Bucle - 1).positive = 1 Then
                        Cabecera.Lines(Bucle).ImporteLinea = BalTick.Lines(Bucle).amount
                        Cabecera.Lines(Bucle).cantidad = BalTick.Lines(Bucle).Item_Pes
                    'Else
                    '    Cabecera.Lines(Bucle).ImporteLinea = BalTick.Lines(Bucle - 1).amount
                    '    Cabecera.Lines(Bucle).cantidad = BalTick.Lines(Bucle - 1).Item_Pes
                    'End If
                    '*****************************
                    'If Cabecera.Lines(Bucle).EsPesado = True Then Cabecera.Lines(Bucle).cantidad = (Cabecera.Lines(Bucle).cantidad / 1000)
                    '**************************
                    ' ML100 --> Si el peso está
                    ' entre 15 y 30 kg, sólo
                    ' proporciona dos decimales
                    ' de peso, y hay que corregirlo
                    '******************************
                    '******************************cas.v1.1.6
                    'If ML100 And aTipoTqt <> 2 And (MachineID = 0 Or MachineID = 1) And Cabecera.Lines(Bucle).EsPesado Then
                    '    If Abs(Cabecera.Lines(Bucle).ImporteLinea) >= 999 * (Abs(Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).precio)) Then
                    '        Cabecera.Lines(Bucle).cantidad = Cabecera.Lines(Bucle).cantidad * 1000
                    '    Else
                    '       If Abs(Cabecera.Lines(Bucle).ImporteLinea) >= 99 * (Abs(Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).precio)) Then
                    '           Cabecera.Lines(Bucle).cantidad = Cabecera.Lines(Bucle).cantidad * 100
                    '       Else
                    '           If Abs(Cabecera.Lines(Bucle).ImporteLinea) >= 9 * (Abs(Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).precio)) Then
                    '               Cabecera.Lines(Bucle).cantidad = Cabecera.Lines(Bucle).cantidad * 10
                    '           End If
                    '       End If
                    '    End If
                    'End If
                    '***************************************cas.v1.1.6
                    ' verifica si se ha modificado el precio
                    '***************************************
                    If Cabecera.Lines(Bucle).CodigoArticulo <> 0 Then
                        If aTipoTqt = 2 Then
                            sSQL = "SELECT codigo " & _
                                   "FROM articulo " & _
                                   "WHERE plu=" & Cabecera.Lines(Bucle).CodigoArticulo & _
                                   " AND codi_ident=" & Cabecera.NMostrador & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If Not RegArt.EOF Then
                                codArt = RegArt.Fields("codigo")
                                Cabecera.Lines(Bucle).CodigoArticulo = codArt
                            End If
                            RegArt.Close
                            Set RegArt = Nothing
                        Else
                            codArt = Cabecera.Lines(Bucle).CodigoArticulo
                        End If
                        If Not (TipoC9) And Not (MachineID > 500 And MachineID < 600) Then
                            sSQL = "SELECT * " & _
                                   "FROM articulo " & _
                                   "WHERE codigo=" & CStr(codArt) & _
                                   " AND borrado=false"
                            Set RegArt = Base.OpenRecordset(sSQL)
                            If RegArt.EOF Then
                                 Cabecera.Lines(Bucle).SehaModificado = False
                                 Cabecera.Lines(Bucle).IVA = 0
                            Else
                                ScaleArt.codigo = 0
                                Cabecera.Lines(Bucle).IVA = RegArt!tipo_iva
                                sSQL = "SELECT * " & _
                                       "FROM tramos " & _
                                       "WHERE borrado=false and codigo=" & CStr(codArt)
                                Set RegTramo = Base.OpenRecordset(sSQL)
                                If RegTramo.EOF Then
                                   HayTramos = False
                                Else
                                   HayTramos = True
                                End If
                                RegTramo.Close
                                If HayTramos = False Then
                                   If UsaEuro Then
                                       If (Cabecera.Lines(Bucle).precio / 100) <> RegArt!Euros Then
                                           Cabecera.Lines(Bucle).SehaModificado = True
                                       Else
                                           Cabecera.Lines(Bucle).SehaModificado = False
                                       End If
                                   Else
                                       If (Cabecera.Lines(Bucle).precio / (10 ^ decimales)) <> RegArt!precio Then
                                           Cabecera.Lines(Bucle).SehaModificado = True
                                       Else
                                           Cabecera.Lines(Bucle).SehaModificado = False
                                       End If
                                   End If
                                   'If Cabecera.Lines(Bucle).SehaModificado And (MachineID = 0 Or MachineID = 1) Then
                                   'If Cabecera.Lines(Bucle).SehaModificado And Not (MachineID > 500 And MachineID < 600) And Not (TipoC9) Then
                                   '    '*****************************************
                                   '    '* consulta el precio real del artículo  *
                                   '    '*****************************************
                                   '    If MachineID = 0 Then
                                   '        ScaleArt = Binario_PideArticulo(RegArt!codigo, False)
                                   '    Else
                                   '        If TipoC9 Then
                                   '            ScaleArt = FrmSC10_PideArticulo((10000 * RegArt!secc_Maqui) + RegArt!plu, False, True)
                                   '        Else
                                   '            ScaleArt = FrmSC10_PideArticulo(RegArt!codigo, False, False)
                                   '        End If
                                   '    End If
                                   '    If UsaEuro Then
                                   '        ScaleArt.precio = ScaleArt.precio / 100
                                   '    Else
                                   '        ScaleArt.precio = ScaleArt.precio / (10 ^ decimales)
                                   '    End If
                                   '    If (UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!Euros)) Or _
                                   '    (Not UsaEuro And (ScaleArt.codigo = RegArt!codigo) And (ScaleArt.precio <> RegArt!precio)) Then
                                   '        If AbsorverCambios Then
                                   '            MyArt.codigo = Cabecera.Lines(Bucle).CodigoArticulo
                                   '            MyArt.plu = 0
                                   '            MyArt.Mostrador = 0
                                   '            MyArt.precio = ScaleArt.precio
                                   '            If CSiempre Then
                                   '                OKCambiar = True
                                   '            Else
                                   '                OKCambiar = False
                                   '                If IsNull(RegArt!prc3) Then
                                   '                    OKCambiar = True
                                   '                Else
                                   '                    If RegArt!prc3 < (Val(Format(Cabecera.Fecha, "yymmdd")) * 10000 + Val(Format(Cabecera.hora, "hhmm"))) Then OKCambiar = True
                                   '                End If
                                   '            End If
                                   '            If OKCambiar Then
                                   '                Modificacion_Articulo_Precio MyArt, "Balanzas", True
                                   '            End If
                                   '        End If
                                   '    Else
                                   '        Cabecera.Lines(Bucle).ModificadoTemporal = True
                                   '        Cabecera.Lines(Bucle).SehaModificado = False
                                   '    End If
                                   'End If
                                End If
                            End If
                            RegArt.Close
                            Set RegArt = Nothing
                        End If
                    End If
                    MiContador = MiContador + 1
                Next Bucle
                '********************************
                ' Actualiza las tablas de totales
                '********************************
                Cabecera.Incluir_en_Base = True
                Actualiza_Tabla_Cabecera Base, Cabecera
                If Cabecera.Incluir_en_Base = True Then
                                  
                    Actualiza_Tabla_Tickets Base, Cabecera
                    CadenadeLog "Actualiza Total Artículos..."
                    Actualiza_Tabla_Gtarti Base, Cabecera
                    CadenadeLog "Actualiza Total Secciones..."
                    Actualiza_Tabla_Gtsecs Base, Cabecera
                    CadenadeLog "Actualiza Total Vendedores..."
                    Actualiza_Tabla_GtVend Base, Cabecera
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\cas.v118
                    CadenadeLog "Actualiza Total Info. Vend. ..."
                    actualiza_tabla_gtInfVen Base, Cabecera
                    CadenadeLog "Actualiza Total Horario..."
                    actualiza_tabla_gtHora Base, Cabecera '\\\cas.v170
                    
                    '///////////////////////////////////////////////////
                    CadenadeLog "Actualiza Log ..."
                    Actualiza_Tabla_Log Base, Cabecera
                End If
            
                Base.Close
                
            End If
        End If
        If cancelar = 1 Then Repetir = False
    Loop
    
    '1.4.1 Base.Close
    
End Sub

Public Sub Switch_mode(ByVal Scaletype As Integer, Optional PortGB As Integer, Optional MaqGB As Integer)
    Dim MyC9 As String
    Dim MyResp As Long
    Dim UsuarioOK As Long
    Dim MyNum As String
    Dim BufDetiene As Boolean
    Dim MyNumMaq As Integer
    Dim MyIP As String
    Dim valorip As String
    Dim nST As Long
    
    'If lSerial Then
    '    Scaletype = 0 ''' cambios upper ... tener en cuenta seleccionar SC10...
    'Else
    '    Scaletype = 0
    'End If
    
    If Dir(App.Path & "\cancel.tqt") <> "" Then Kill App.Path & "\cancel.tqt"
    cancelar = 0
    ' 1.4.1
    'frmControl.Balanza.Set_Timeout (5)
    'frmControl.Balanza.Set_Retries 3
    'frmControl.Balanza.Set_Timeout (1)
    'frmControl.Balanza.Set_Retries 1
    

    Select Case Scaletype
        Case 10
            On Error Resume Next
            nST = frmControl.Balanza.Get_Status()
            If nST <> 1 Or (Err.Number <> 0) Then
                If (Err.Number <> 0) Then
                    On Error GoTo 0
                End If
                MyResp = frmControl.Balanza.Configure(10, 0, sAddIP)
            End If
            On Error GoTo 0
            If MyResp = 0 Then
               estadobal = 10
            Else
               estadobal = 1000
            End If
        
        Case 11
            On Error Resume Next
            nST = frmControl.Balanza.Get_Status()
            If nST <> 1 Or (Err.Number <> 0) Then
                If (Err.Number <> 0) Then
                    On Error GoTo 0
                End If
                MyResp = frmControl.Balanza.Configure(11, 0, sAddIP)
            End If
            On Error GoTo 0
            If MyResp = 0 Then
               estadobal = 11
            Else
               estadobal = 1000
            End If
        
        Case 0
            On Error Resume Next
            If lSerial Then
                MyResp = frmControl.Balanza.Configure(0, 1, sCfgSerial)
            Else
                nST = frmControl.Balanza.Get_Status()
                If nST <> 1 Or (Err.Number <> 0) Then
                    If (Err.Number <> 0) Then
                        On Error GoTo 0
                    End If
                    MyResp = frmControl.Balanza.Configure(0, 0, sAddIP)
                End If
            End If
            On Error GoTo 0
            If MyResp = 0 Then
               estadobal = 0
            Else
               estadobal = 1000
            End If
        Case 1
            On Error Resume Next
            If lSerial Then
                MyResp = frmControl.Balanza.Configure(1, 1, sCfgSerial)
            Else
                'c2f 1.4.2 Ahorramas SC10 TCP MyResp = 1010
                nST = frmControl.Balanza.Get_Status()
                If nST <> 1 Or (Err.Number <> 0) Then
                    If (Err.Number <> 0) Then
                        On Error GoTo 0
                    End If
                    MyResp = frmControl.Balanza.Configure(1, 0, sAddIP)
                    
                End If
                ''''''''''''''''''''''''''''''''''''''''''''
            End If
            On Error GoTo 0
            If MyResp = 0 Then
               estadobal = 1
            Else
               estadobal = 1000
            End If
            
    End Select
    'Select Case estadobal
    '    Case 0
    '        frmControl.Balanza.ScaleVersion (1)
    '        If Dir(App.Path & "\noask13") <> "" Then
    '            'compatibilidad...TiqAhTouch
    '            'frmControl.Balanza.NoAsk13 True
    '            ''''''''''''''''''''''''''''
    '        End If
    '
    'End Select
    If MyResp <> 0 And MyResp <> 1 Then
        cancelar = 1

        If MyResp <> 7 Then
            frmControl.Balanza.Reset
            estadobal = 1000
        Else
            estadobal = Scaletype
        End If

    Else
        cancelar = 0
    End If
End Sub



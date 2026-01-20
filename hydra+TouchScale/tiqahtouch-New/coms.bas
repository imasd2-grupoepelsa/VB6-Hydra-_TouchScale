Attribute VB_Name = "coms"
Option Explicit
Public EstadoBal As Integer
Public CI_Salir As Boolean
Public lCogeTiquet As Boolean

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
Dim sSecPlu As String
Dim sFich1 As String
Dim sFich2 As String

On Error GoTo noEsPosible
If (Dir(App.Path & "\hydrarun.txt") <> "") And lUpper Then
    GoTo noEsPosible
End If
If Dir(App.Path & "\tiqah.ini") <> "" Then
    sFich1 = "cabtqttouch"
    sFich2 = "tiquettouch"
Else
    sFich1 = "cabtqt"
    sFich2 = "tiquet"
End If
    'If lUpper = False Then
    
        If Dir(var.PathExpor & "\" & sFich1 & ".sdf") <> "" Then
            If Dir(var.PathExpor & "\" & sFich1 & ".bak") <> "" Then
                If Dir(var.PathExpor & "\" & sFich1 & ".old") <> "" Then
                    Kill var.PathExpor & "\" & sFich1 & ".old"
                End If
                FileCopy var.PathExpor & "\" & sFich1 & ".bak", var.PathExpor & "\" & sFich1 & ".old"
                Kill var.PathExpor & "\" & sFich1 & ".bak"
            End If
            
            FileCopy var.PathExpor & "\" & sFich1 & ".sdf", var.PathExpor & "\" & sFich1 & ".bak"
            
        End If
        'If Dir(var.PathExpor & "\tiquettouch.sdf") <> "" Then
        '    If Dir(var.PathExpor & "\tiquettouch.bak") <> "" Then
        '        If Dir(var.PathExpor & "\tiquettouch.old") <> "" Then
        '            Kill var.PathExpor & "\tiquettouch.old"
        '        End If
        '        FileCopy var.PathExpor & "\tiquettouch.bak", var.PathExpor & "\tiquettouch.old"
        '        Kill var.PathExpor & "\tiquettouch.bak"
        '    End If
        '
        '    FileCopy var.PathExpor & "\tiquettouch.sdf", var.PathExpor & "\tiquettouch.bak"
        '
        'End If
        If Dir(var.PathExpor & "\" & sFich2 & ".sdf") <> "" Then
            If Dir(var.PathExpor & "\" & sFich2 & ".bak") <> "" Then
                If Dir(var.PathExpor & "\" & sFich2 & ".old") <> "" Then
                    Kill var.PathExpor & "\" & sFich2 & ".old"
                End If
                FileCopy var.PathExpor & "\" & sFich2 & ".bak", var.PathExpor & "\" & sFich2 & ".old"
                Kill var.PathExpor & "\" & sFich2 & ".bak"
            End If
            
            FileCopy var.PathExpor & "\" & sFich2 & ".sdf", var.PathExpor & "\" & sFich2 & ".bak"
            
        End If
    
    'End If
    
    'If FileDateTime(var.PathExpor & "\cabtqt.sdf") = Now() Then GoTo noEsPosible
    'If FileDateTime(var.PathExpor & "\tiquet.sdf") = Now() Then GoTo noEsPosible
    
    Fichero1 = FreeFile()
    'Open App.Path & "\cabtqt.sdf" For Output As #Fichero1
    Open var.PathExpor & "\" & sFich1 & ".sdf" For Append As #Fichero1
    Fichero2 = FreeFile()
    'Open App.Path & "\tiquet.sdf" For Output As #Fichero2
    Open var.PathExpor & "\" & sFich2 & ".sdf" For Append As #Fichero2
    
    '1.7.14
    'If Euroscale_o_SC10 Then
    '    MiTabla1 = "Cabecera"
    '    MiTabla2 = "tickets"
    'Else
    '    MiTabla1 = "cabeceraSC10"
    '    MiTabla2 = "ticketsSC10"
    'End If
    
    
    Set Base = AbrirBase()
    On Error GoTo 0
    On Error GoTo finCabLin
    
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
                        If lUpper Then
                            'MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Format(!secc, "00") & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
                            MiEnvio = MiEnvio & "0" & Codigo_bal & Num_lin & Cod_Plu & Precio_Uni & Pes_Uni & Signo & Tipo_Linea & LineaAnulada & Total & "1"
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
    GoTo noEsPosible
    
finCabLin:
    If Err.Number <> 0 Then
        CadenadeLog "export_CABLIN, Err.:" & CStr(Err.Number) & " Desc.:" & Err.Description
        On Error GoTo 0
    End If
    On Error Resume Next
    
    Close #Fichero1
    Close #Fichero2
    
    Base.Close
    Set Base = Nothing

noEsPosible:
    On Error GoTo 0
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
                        sPre = Format((!Price) * 100, "00000")
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
                        If !Price < 0 Then
                            sPrice = (-1) * !Price
                            sPrice = "-" & Format(100 * (sPrice), "000000")
                        Else
                            sPrice = !Price
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


Public Sub Switch_M(ByVal MiP As String)

Dim MyFich As Integer
Dim s As String
Dim BufDetiene As Variant
Dim valorip As String
Dim sTi As String
Dim nResp As Integer

    On Error Resume Next
    MyConnObj.Close
    Set MyConnObj = Nothing
    EstadoBal = 1000
    On Error GoTo 0
    
    MiP = Val(Left(MiP, 3)) & "." & Val(Mid(MiP, 5, 3)) & "." & Val(Mid(MiP, 9, 3)) & "." & Val(Right(MiP, 3))
    

    '''''''''''''''''''''''''''''''
    'Conectar con sAddIp y sPortBal
    '''''''''''''''''''''''''''''''
    On Error Resume Next

    CadenadeLog "Cadena de Conexión" & "-->" & MiP & "," & sPortBal
    Do_Events

    nResp = ElPing(MiP)
    If nResp = 1 Then
    
        If Dir(App.Path & "\odbc53.txt") <> "" Then
            MyConnObj.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        Else
            MyConnObj.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=" & MiP & "; Port=" & sPortBal & "; DATABASE=pcscale; UID=pcscale; PWD=epelsa; OPTION=3"
        End If
        
        CadenadeLog "Estado Conexión:" & CStr(MyConnObj.State)
        CadenadeLog "Error:" & CStr(Err.Number)
        
        If Err.Number <> 0 Then

            On Error GoTo 0
            EstadoBal = 1000
            GoTo finMSWITCH
        Else
            EstadoBal = 0
        End If
    
    Else
        'frmmonitor.MostrarDato "No Ping:" & MiP
        EstadoBal = 1000
    End If
                

finMSWITCH:
    If EstadoBal <> 0 Then
        cancelar = 1
    Else
        cancelar = 0
    End If

End Sub

Public Sub Switch_mode(ByVal Scaletype As Integer)
    Dim MyC9 As String
    Dim MyResp As Long
    Dim UsuarioOK As Long
    Dim MyNum As String
    Dim BufDetiene As Boolean
    Dim MyNumMaq As Integer
    Dim MyIP As String
    Dim valorip As String
    Dim nST As Long
    Dim nResp As Integer
    
    On Error GoTo sigueSw
    MyResp = 1000
    nResp = ElPing(sAddIP)
    If nResp <> 1 Then
        If MyConnObj.State <> 0 Or MyConnObj.Errors.Count <> 0 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
        End If
        CadenadeLog "sin ping..."
        GoTo finSWITCH
    End If

    'If MyConnObj.State <> 0 Or MyConnObj.Errors.Count <> 0 Then
    '    MyConnObj.Close
    '    Set MyConnObj = Nothing
    'End If
    'Do_Events
    'nutricion
    'carga_data_Vitamin
    ''''''''''
    Do_Events
    
sigueSw:
    On Error GoTo 0
    
    On Error GoTo finSWITCH
    
    
    If Dir(App.Path & "\cancel.tqt") <> "" Then Kill App.Path & "\cancel.tqt"
    cancelar = 0

    On Error Resume Next
    MyResp = 0
    nST = Get_Status() '.balanza
    If nST <> 1 Or (Err.Number <> 0) Then
        If (Err.Number <> 0) Then
            On Error GoTo 0
        End If
        MyResp = Configure(11, 0, sAddIP) '.balanza
    End If
finSWITCH:
    On Error GoTo 0
    If MyResp = 0 Then
       EstadoBal = 11
       cancelar = 0
    Else
       reset
       EstadoBal = 1000
       cancelar = 1
    End If
End Sub


Public Sub LeeTickets(Optional Tabla As Integer)
    
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecSetTmp As New ADODB.Recordset
    Dim MyRecImp As New ADODB.Recordset
    
    Dim Mybase As DAO.Database
    Dim myRec As DAO.Recordset
    Dim MisEqp As DAO.Recordset
    Dim nLR As Integer
    
    Dim cm As String
    Dim Ffecha As Variant
    Dim Fhora As Variant
    Dim IdDoc As Variant
    Dim nNumLin As Variant
    Dim Cabecera As tipo_cabecera
    Dim ncont As Long
    Dim dLImporte As Double
    Dim nLocal As Integer
    Dim sCap As String
    
    'If Dir(App.Path & "\hydramulti.exe") <> "" Or Dir(App.Path & "\noslaves.txt") <> "" Then
    '    nLocal = 0
    'Else
        nLocal = 1
    'End If
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'Call capturaClientes
    
    On Error GoTo fin
    Set Mybase = DAO.OpenDatabase(Base_General)
    'CadenadeLog "paso 1"
    'If nLocal <> 0 Then
    If Dir(App.Path & "\tqslaves.txt") <> "" Then
        Set MisEqp = Mybase.OpenRecordset("select * from equipos order by modelo asc")
    Else
        Set MisEqp = Mybase.OpenRecordset("select * from equipos where modelo=1")
    End If
    'CadenadeLog "paso 2"
    If MisEqp.EOF Then
        GoTo SalirLeeTique
    Else
        MisEqp.MoveFirst
        Do While Not MisEqp.EOF
            
            'If Dir(App.Path & "\tqslaves.txt") <> "" Then
            '    frmControl.Switch_M MisEqp.Fields("prog_eqp")
            'Else
            '    frmControl.Switch_mode 0
            'End If
            If MisEqp.Fields("modelo") = 1 Then
                Switch_mode 11
            Else
                Switch_M MisEqp.Fields("prog_eqp")
            End If
            'CadenadeLog "paso 3"
            If EstadoBal = 11 Then
            
            For nLR = 0 To 1 ' Red/Local
            
            If MisEqp.Fields("modelo") <> 1 And nLR = 0 Then GoTo SigueNlr
            'CadenadeLog "paso 4"
            If nLR = 0 Then
            'MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat<>'L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic

                If IsNull(Tabla) Or Tabla = 0 Then
                    If Dir(App.Path & "\noencargo.txt") <> "" Then
                        'MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or isnull(captured)) and (CapturedDateTime='0000-00-00 00:00:00' or  CapturedDateTime='0000-00-00 11:00:00') and TckType<>'O' and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        If lMarcaTime Then
                            MyRecSetTmp.Open "select * from pcscale.htickets where TckType<>'O' and (CapturedDateTime='0000-00-00 00:00:00' or  CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        Else
                            MyRecSetTmp.Open "select * from pcscale.htickets where TckType<>'O' and (captured='' or isnull(captured) or captured<>'*') and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        End If
                    Else
                        'MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or isnull(captured)) and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00') and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        If lMarcaTime Then
                            MyRecSetTmp.Open "select * from pcscale.htickets where (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        Else
                            MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or isnull(captured) or captured<>'*') and (status='C' or status='R' or status='X') and NetStat='N' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        End If
                    End If
                    'CadenadeLog "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='N' order by id"
                Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.packetlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.packetlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.sservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.sservicelabels where customer<>'9999999999' order by id"
                    End If
                End If
            Else
            'MyRecSetTmp.Open "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
            
                If IsNull(Tabla) Or Tabla = 0 Then
                    'MyRecSetTmp.Open "select * from pcscale.localhtickets where captured='' and status='C' and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    If Dir(App.Path & "\noencargo.txt") <> "" Then
                        MyRecSetTmp.Open "select * from pcscale.htickets where TckType<>'O' and (captured='' or captured<>'*') and TckType<>'O' and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    Else
                        MyRecSetTmp.Open "select * from pcscale.htickets where (captured='' or captured<>'*') and (CapturedDateTime='0000-00-00 00:00:00' or CapturedDateTime='0000-00-00 11:00:00' or CapturedDateTime='1970-01-01 00:00:00') and (status='C' or status='R' or status='X' or status='W' or status='V' or status='w') and NetStat='L' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                    End If
                    'CadenadeLog "select * from pcscale.htickets where captured='' and (status='C' or status='R' or status='X') and NetStat='L' order by id"
                Else
                    If Tabla = 1 Then
                        MyRecSetTmp.Open "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localpacketlabels where customer<>'9999999999' order by id"
                    Else
                        MyRecSetTmp.Open "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id", MyConnObj, adOpenKeyset, adLockOptimistic
                        'CadenadeLog "select * from pcscale.localsservicelabels where customer<>'9999999999' order by id"
                    End If
                End If
            End If
            'CadenadeLog "paso 5"
            If Not MyRecSetTmp.EOF Then
                Do While Not MyRecSetTmp.EOF
                    'CadenadeLog "paso 6"
                    Ffecha = Format(MyRecSetTmp.Fields("dattim"), "dd/mm/yyyy hh:mm:ss")
                    Fhora = Mid(Ffecha, 12, 8)
                    Ffecha = Mid(Ffecha, 1, 10)
                    'CadenadeLog "paso 7"
                    IdDoc = MyRecSetTmp.Fields("id")
                    If IsNull(MyRecSetTmp.Fields("CapturedDateTime")) Then
                        sCap = "0000-00-00 00:00:00"
                    Else
                        sCap = Format(MyRecSetTmp.Fields("CapturedDateTime"), "yyyy-mm-dd hh:mm:ss")
                    End If
                    sCap = Mid(sCap, 1, 17) & "11"
                    'CadenadeLog "paso 8"
                    ''''medio de pago
                    '1=efectivo
                    '2=t.crédito
                    '3=cheque
                    Cabecera.tipoPago = 1
                    Cabecera.cantidadEntre = 0
                    
                    If IsNull(Tabla) Or Tabla = 0 Then
'Select Case !TYPTIC --> Cabecera.tipoTiquet
'    Case 1
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(336) '"Venta"
'    Case 2
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(195) '"Envasado" 'CargaCadena(741) '"Super"
'    Case 3
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(196) '"Autoservicio" 'CargaCadena(195) '"Envasado"
'    Case 4
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(1213) '"Venta Local" 'CargaCadena(196) '"Autoservicio"
'    Case 5
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(199) '"Envasado Local" 'CargaCadena(1213) '"Venta Local"
'    Case 6
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = CargaCadena(200) '"Autoservicio Local" 'CargaCadena(198) '"Super Local"
'    Case 7
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Net" '"Envasado Local"
'    Case 8
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Fact. Local" '"Autoservicio Local"
'    Case 12
'        Lista.ListItems(Lista.ListItems.Count).SubItems(11) = "Descart."
'End Select
                    
                        'If nLR = 0 Then
                            MyRecAux.Open "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "select * from pcscale.paytickets where idhticket =" & CStr(IdDoc)
                            If Not MyRecAux.EOF Then
                                'CadenadeLog "paso 9"
                                Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                                Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                                'CadenadeLog "paso 10"
                            End If
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'CadenadeLog "select * from pcscale.localpaytickets where idhticket =" & CStr(IdDoc)
                        '    If Not MyRecAux.EOF Then
                        '        'CadenadeLog "paso 11"
                        '        Cabecera.tipoPago = MyRecAux.Fields("paymethod")
                        '        Cabecera.cantidadEntre = MyRecAux.Fields("delivered")
                        '        'CadenadeLog "paso 12"
                        '    End If
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                    End If
        '''''''
                    'Impuestos
                    Cabecera.base0 = 0
                    Cabecera.por0 = 0
                    Cabecera.imp0 = 0
                    Cabecera.base1 = 0
                    Cabecera.por1 = 0
                    Cabecera.imp1 = 0
                    Cabecera.base2 = 0
                    Cabecera.por2 = 0
                    Cabecera.imp2 = 0
                    Cabecera.base3 = 0
                    Cabecera.por3 = 0
                    Cabecera.imp3 = 0
                    Cabecera.base4 = 0
                    Cabecera.por4 = 0
                    Cabecera.imp4 = 0
        ''''''''''
                    
                    '''''''''''''''''
                    If IsNull(Tabla) Or Tabla = 0 Then
                        
                        'If nLR = 0 Then
                            MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                            'CadenadeLog "paso 13"
                            nNumLin = MyRecAux.Fields("lineas")
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            ReDim Cabecera.Lines(nNumLin)
                            'MyRecAux.Open "select sum(amount) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                            'On Error Resume Next
                            ''CadenadeLog "paso 14"
                            'dLImporte = MyRecAux.Fields("importe")
                            ''CadenadeLog "paso 15"
                            'If Err.Number <> 0 Then
                            '    On Error GoTo 0
                            '    On Error GoTo FIN
                            '    dLImporte = 0
                            'End If
                            'MyRecAux.Close
                            'Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    'CadenadeLog "paso 16"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '    ReDim Cabecera.Lines(nNumLin)
                        '    'MyRecAux.Open "select sum(amount) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc) & " and linetype=1", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    'On Error Resume Next
                        '    'dLImporte = MyRecAux.Fields("importe")
                        '    ''CadenadeLog "paso 17"
                        '    'If Err.Number <> 0 Then
                        '    '    On Error GoTo 0
                        '    '    On Error GoTo FIN
                        '    '    dLImporte = 0
                        '    'End If
                        '    'MyRecAux.Close
                        '    'Set MyRecAux = Nothing
                        '
                        'End If
                    Else
                        nNumLin = 1
                        'CadenadeLog "paso 18"
                        ReDim Cabecera.Lines(nNumLin)
                        'dLImporte = MyRecSetTmp.Fields("amount")
                        'CadenadeLog "paso 19"
                    End If
                    
                    'recoger datos
                    Cabecera.Ntiquet = MyRecSetTmp.Fields("number")
                    Cabecera.NVendedor = MyRecSetTmp.Fields("vendor")
                    Cabecera.NMostrador = MyRecSetTmp.Fields("counter")
                    'CadenadeLog "paso 20"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        If IsNull(MyRecSetTmp.Fields("invoice")) Or Len(MyRecSetTmp.Fields("invoice")) = 0 Then
                            Cabecera.factura = Space(15)
                        Else
                            Cabecera.factura = Trim(MyRecSetTmp.Fields("invoice"))
                            'CadenadeLog "paso 21"
                        End If
                        
                        Cabecera.ean13 = Trim(MyRecSetTmp.Fields("ean13"))
                        'CadenadeLog "paso 22"
                        If MyRecSetTmp.Fields("customer") <> "" Then
                            If IsNull(MyRecSetTmp.Fields("customer")) Then
                                Cabecera.NCliente = 0
                            Else
                                On Error Resume Next
                                Cabecera.NCliente = MyRecSetTmp.Fields("customer")
                                If Err.Number <> 0 Then
                                    On Error GoTo 0
                                    Cabecera.NCliente = 0
                                End If
                                On Error GoTo fin
                            End If
                        Else
                            Cabecera.NCliente = 0
                        End If
                        'CadenadeLog "paso 23"
                    Else
                        Cabecera.factura = Space(15)
                        Cabecera.NCliente = 0
                    End If
                    'CadenadeLog "paso 24"
                    
                    Cabecera.Fecha = Format(Ffecha, "dd/mm/yyyy")
                    Cabecera.hora = Fhora
                    Cabecera.NBalanza = MyRecSetTmp.Fields("scale")
                    If Tabla <> 0 Then
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("amount")
                    Else
                        Cabecera.ImporteTotal = MyRecSetTmp.Fields("netamount")
                    End If
                    'CadenadeLog "paso 25"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        Cabecera.nLineas = nNumLin
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = nNumLin
                        Cabecera.LineasMensaje = nNumLin
                    Else
                        Cabecera.nLineas = 1
                        Cabecera.LineaInicial = 1
                        Cabecera.LineaFinal = 1
                        Cabecera.LineasMensaje = 1
                    End If
                    'cabecera.codigo ???
                    
                    Cabecera.EstadoTiquet = 16
                    'CadenadeLog "paso 26"
                    If IsNull(Tabla) Or Tabla = 0 Then
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If MyRecSetTmp.Fields("NetStat") = "L" Then
                            Cabecera.CadenaEstadoTiquet = "Local"
                            Cabecera.tipoTiquet = 4
                        Else
                            Cabecera.CadenaEstadoTiquet = "Red"
                            Cabecera.tipoTiquet = 1
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'CadenadeLog "paso 27"
                        'If nLR = 0 Then
                        '    'CadenadeLog "paso 28"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.ltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 29"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    'CadenadeLog "paso 30"
                        '    MyRecAux.Open "select sum(price*weight) as importe from pcscale.localltickets where Idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        '    dLImporte = Round(MyRecAux.Fields("importe"), 2)
                        '    'CadenadeLog "paso 31"
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'End If
                        'CadenadeLog "paso 32"
                        'Cabecera.Trainning = False
                        'If MyRecSetTmp.Fields("tcktype") = "U" Then
                        '    Cabecera.tipoTiquet = 12
                        '    Cabecera.Trainning = True
                        'End If
                        'If MyRecSetTmp.Fields("status") = "R" Then
                        '    Cabecera.tipoTiquet = 13
                        '    Cabecera.EstadoTiquet = 64
'
                        'End If
                        ''CadenadeLog "paso 33"
                        'If MyRecSetTmp.Fields("status") = "X" Then
                        '    Cabecera.tipoTiquet = 14
                        '    Cabecera.EstadoTiquet = 64
                        '
                        'End If
''V': Tique reembolso (cantidades vendidas negativas). (SAFT_PT) (+DB.135)
''w: Tique reembolsado (cancelado) pero no recargado. (SAFT_PT) (+DB.135)
''W': Tique reembolsado (cancelado) y recargado. (SAFT_PT) (+DB.135) .

''X': Cancelado por el operador. Sin imprimir.
''R': Tique recargado, Tique “Cerrado” copiado sobre un nuevo tique para añadir nuevas operaciones.

                        Cabecera.Trainning = False
                        If MyRecSetTmp.Fields("tcktype") = "U" Then
                            Cabecera.tipoTiquet = 12
                            Cabecera.Trainning = True
                        End If
                        If MyRecSetTmp.Fields("status") = "R" Or MyRecSetTmp.Fields("status") = "V" Then
                            Cabecera.tipoTiquet = 13
                            Cabecera.EstadoTiquet = 64

                        End If
                        'CadenadeLog "paso 33"
                        If MyRecSetTmp.Fields("status") = "X" Or MyRecSetTmp.Fields("status") = "W" Or MyRecSetTmp.Fields("status") = "w" Then
                            Cabecera.tipoTiquet = 14
                            Cabecera.EstadoTiquet = 64
                            
                        End If
                        'CadenadeLog "paso 34"
                        'If Tabla <> 0 Then
                        '    Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                        'Else
                            Cabecera.ImporteBruto = MyRecSetTmp.Fields("grossamount")
                        'End If
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Round(Cabecera.ImporteBruto - Cabecera.ImporteTotal, 2) 'más abajo se hace como porcentaje...
                        Cabecera.documentoabono = 0
                    
                    Else
                        'CadenadeLog "paso 35"
                        If nLR = 0 Then
                            Cabecera.CadenaEstadoTiquet = "Red"
                        Else
                            Cabecera.CadenaEstadoTiquet = "Local"
                        End If
                        
                        ' en "OnOff" tabla cabecera se indicará red o local... (como reconocer en tabla sqlserver?)
                        '    se guarda también (en cadena) si es cancelado o no...por ejemplo "Red cancelado".
                        If Tabla = 1 Then
                            Cabecera.tipoTiquet = 2
                        Else
                            Cabecera.tipoTiquet = 3
                        End If
                        
                        'Cabecera.CadenaEstadoTiquet = "Red"
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        'Cabecera.ImporteBruto = dLImporte
                        Cabecera.ImporteBruto = MyRecSetTmp.Fields("amount")
                        Cabecera.Impuestos = 0
                        'Cabecera.descuento = Cabecera.ImporteBruto - Cabecera.ImporteTotal
                        Cabecera.documentoabono = 0
                        'Cabecera.Cambio = 0
                        'Cabecera.descuento = 0
                        'Cabecera.LineasCanceladas = 0
                    End If
                    'CadenadeLog "paso 36"
                    If nLR = 0 Then
                        If Tabla = 0 Or IsNull(Tabla) Then
                            ''frmmonitor.MostrarDato CargaCadena(868) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            'frmmonitor.MostrarDato CargaCadena(1412) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            'frmmonitor.MostrarDato CargaCadena(1413) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    Else
                        If Tabla = 0 Or IsNull(Tabla) Then
                            'frmmonitor.MostrarDato CargaCadena(1414) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Ticket Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                        Else
                            If Tabla = 1 Then
                            'frmmonitor.MostrarDato CargaCadena(1415) & " " & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Envasado Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            Else
                            'frmmonitor.MostrarDato CargaCadena(1416) & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            CadenadeLog "HydraTouch:" & "Etiqueta Autoservicio Local:" & Cabecera.Ntiquet & "Scale:" & Cabecera.NBalanza & Cabecera.CadenaEstadoTiquet
                            End If
                        End If
                    
                    End If
                    'CadenadeLog "paso 37"
                    If Tabla = 0 Or IsNull(Tabla) Then
                        'If nLR = 0 Then
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.ltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        'Else
                        '    MyRecAux.Open "select COUNT(*) as lineas from pcscale.localltickets where (linetype=0 and idhticket=" & CStr(IdDoc) & ")", MyConnObj, adOpenKeyset, adLockOptimistic
                        '    nNumLin = MyRecAux.Fields("lineas")
                        '    MyRecAux.Close
                        '    Set MyRecAux = Nothing
                        '
                        'End If
                        'CadenadeLog "paso 38"
                        'CadenadeLog "7"
                        
                        Cabecera.LineasCanceladas = MyRecSetTmp.Fields("canceledlines")
                        
                        'también se vé afectado si es cancelado o no pero de mamento cerrado...
                        'Cabecera.EstadoTiquet = 16
                        '********************
                        ' este dato se introduce
                        ' tras pasar por actualiza_tabla_cabecera
                        'Incluir_en_Base As Boolean
                        'Cabecera.tipoPago = myrecsettmp.Fields("idmoneda")
                        'Cabecera.cantidadEntre = myrecsettmp.Fields("totalneto") + myrecsettmp.Fields("importecambio")
                        
                        Cabecera.Cambio = Round(Cabecera.ImporteTotal - Cabecera.cantidadEntre, 2) ' myrecsettmp.Fields("importecambio")
                        Cabecera.descuento = MyRecSetTmp.Fields("tckdiscount")
                        'CadenadeLog "paso 39"
                        'lote As String
                        'lEncargo As Boolean
            
            
                        'LINEAS
                        If lFornes Then
                            If MyRecSetTmp.Fields("usesectionitem") = 1 Then
                                Cabecera.NMostrador = MyRecSetTmp.Fields("sectionitem")
                            End If
                        End If
            
                        nNumLin = 0
                        'If nLR = 0 Then
                            MyRecAux.Open "select * from pcscale.ltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'Else
                        '    MyRecAux.Open "select * from pcscale.localltickets where idhticket =" & CStr(IdDoc), MyConnObj, adOpenKeyset, adLockOptimistic
                        'End If
                        'CadenadeLog "8"
                        'CadenadeLog "paso 40"
                        If Not MyRecAux.EOF Then
                            
                            'CadenadeLog "8.1"
                            
                            MyRecAux.MoveFirst
                            
                            'CadenadeLog "8.2"
                            
                            Do While Not MyRecAux.EOF
                                
                                'CadenadeLog "8.3"
            'DROP TABLE IF EXISTS `pcscale`.`ltickets`;
            'CREATE TABLE  `pcscale`.`ltickets` (
            '  `Id` int(11) NOT NULL AUTO_INCREMENT,
            '  `IdHTicket` int(11) DEFAULT NULL,
            '  `lineType` smallint(6) DEFAULT NULL,
            '  `SaleForm` smallint(6) DEFAULT NULL,
            '  `Item` int(11) DEFAULT NULL,
            '  `Weight` decimal(5,3) DEFAULT NULL,
            '  `Tare` decimal(5,3) DEFAULT '0.000',
            '  `TTare` char(1) DEFAULT '',
            '  `lineDiscount` decimal(4,2) DEFAULT NULL,
            '  `Price` decimal(6,2) DEFAULT NULL,
            '  `Amount` decimal(8,2) DEFAULT NULL,
            '  `BovineCard` smallint(3) DEFAULT NULL,
            '  `EarTag` varchar(30) DEFAULT NULL,
            '  `BovineLot` varchar(20) DEFAULT NULL,
            '  PRIMARY KEY (`Id`),
            '  KEY `IdHTicket` (`IdHTicket`)
            ') ENGINE=InnoDB DEFAULT CHARSET=latin1;
                                'CadenadeLog "paso 41"
                                nNumLin = nNumLin + 1
                                If MyRecAux.Fields("saleform") = 1 Then
                                    Cabecera.Lines(nNumLin).EsPesado = True
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                Else
                                    Cabecera.Lines(nNumLin).EsPesado = False
                                    Cabecera.Lines(nNumLin).cantidad = MyRecAux.Fields("weight")
                                End If
                                'CadenadeLog "paso 42"
                                'CadenadeLog "8.4"
                                
                                Cabecera.Lines(nNumLin).FactorPeso = 1
                                Cabecera.Lines(nNumLin).precio = MyRecAux.Fields("price")
                                'CadenadeLog "paso 43"
                                'CadenadeLog "8.5"
                                
                                Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                                Cabecera.Lines(nNumLin).CodigoArticulo = MyRecAux.Fields("item")
                                
                                'CadenadeLog "paso 44"
                                'ver qué se hace IVA As Integer
                                Cabecera.Lines(nNumLin).IVA = MyRecAux.Fields("vatcode")
                                Cabecera.Lines(nNumLin).porcentaje = MyRecAux.Fields("vatpercent")
                                'CadenadeLog "paso 45"
                                'CadenadeLog "8.6"
                                
                                Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("ticketdiscount") + MyRecAux.Fields("linediscount") 'porcentaje debido a la cabecera + porcentaje en linea
                                'CadenadeLog "paso 46"
                                Cabecera.Lines(nNumLin).ImporteBruto = MyRecAux.Fields("amount") 'Round(MyRecAux.Fields("weight") * MyRecAux.Fields("price"), 2) 'MyRecAux.Fields("importe") - MyRecAux.Fields("dtoscabecera")
                                'CadenadeLog "paso 47"
                                Cabecera.Lines(nNumLin).ImporteLinea = Round(MyRecAux.Fields("amount") - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                'Cabecera.Lines(nNumLin).ImporteLinea = Round(Cabecera.Lines(nNumLin).ImporteLinea - Round(((MyRecAux.Fields("ticketdiscount") * MyRecAux.Fields("amount")) / 100), 3), 2)
                                
                                'CadenadeLog "paso 48"
                                'CadenadeLog "8.7"
                                
                                'If MyRecAux.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                                If MyRecAux.Fields("linetype") = 0 Then
                                    Cabecera.Lines(nNumLin).LineaCancelada = True
                                Else
                                    Cabecera.Lines(nNumLin).LineaCancelada = False
                                End If
                                'CadenadeLog "paso 49"
                                'ver qué se hace con TipoLinea As Byte
                                'STV (positivo o negativo)
                                'Linea(Bucle).ModificadoTemporal = False
                                If MyRecAux.Fields("amount") > 0 Then
                                    'Cabecera.Lines(nNumLin).positive = 1
                                     Cabecera.Lines(nNumLin).FactorPeso = 1
                                     Cabecera.Lines(nNumLin).TipoLinea = 0
                                
                                Else
                                    'Cabecera.Lines(nNumLin).positive = 0
                                     Cabecera.Lines(nNumLin).FactorPeso = -1
                                     Cabecera.Lines(nNumLin).TipoLinea = 2
                                    
                                End If
                                'CadenadeLog "paso 50"
            '                    If Cabecera.Lines(nNumLin).positive = 1 Then
            '                         Cabecera.Lines(nNumLin).FactorPeso = 1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 0
            '                    Else
            '                         Cabecera.Lines(nNumLin).FactorPeso = -1
            '                         Cabecera.Lines(nNumLin).TipoLinea = 2
            '                    End If
                                'STP (pesado o no pesado)
                                If Cabecera.Lines(nNumLin).EsPesado Then
                                     'Cabecera.Lines(nNumLin).Weighted = 1
                                Else
                                     'Cabecera.Lines(nNumLin).Weighted = 0
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                                End If
                                
                                'CadenadeLog "paso 51"
                                ' STC (cancelado o no cancelado)
                                If Cabecera.Lines(nNumLin).LineaCancelada Then
                                     Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                                End If
                                'CadenadeLog "paso 52"
                                'If MyRecAux.Fields("precio") <> MyRecAux.Fields("preciodefecto") Then
                                '    Cabecera.Lines(nNumLin).SehaModificado = True
                                'Else
                                    Cabecera.Lines(nNumLin).SehaModificado = False
                                'End If
                                '¿? ModificadoTemporal As Boolean
                                'CadenadeLog "paso 53"
                                'CadenadeLog "8.8"
                                'c2f 1.1.1 aclarar ficha Vacuno...
                                'If IsNull(MyRecAux.Fields("eartag")) Or MyRecAux.Fields("eartag") = "" Then
                                    Cabecera.Lines(nNumLin).FichaVacuno = 0
                                'Else
                                '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecAux.Fields("eartag")
                                'End If
                                'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                                'Cabecera.descuento = Cabecera.descuento + MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecAux.Fields("dtoslineaconimpuestos")
                                
                                'CadenadeLog "8.9"
                                If lFornes Then Cabecera.NMostrador = MyRecAux.Fields("section")
                                
                                MyRecAux.MoveNext
                            Loop
                            MyRecAux.Close
                            Set MyRecAux = Nothing
                            'CadenadeLog "paso 54"
                            
                            Actualiza_Tabla_Cabecera Mybase, Cabecera
                            
                            'CadenadeLog "paso 55"
                            Do_Events
                            If Cabecera.Incluir_en_Base = True Then
                                Actualiza_Tabla_Tickets Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtarti Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_GtVend Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtHora Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Gtsecs Mybase, Cabecera
                                Do_Events
                                actualiza_tabla_gtInfVen Mybase, Cabecera
                                Do_Events
                                Actualiza_Tabla_Log Mybase, Cabecera
                            End If
                            
                            'CadenadeLog "paso 56"
                        End If
            
            
            
                        '''''''''''''''
                        'marcar...
                        'cm = "update dbo.documento set fecha='" & Format(Ffecha, "yyyy-dd-mm") & " 00:00:02.000" & "'" & " where iddocumento=" & CStr(IdDoc)
                        If lMarcaTime Then
                            'cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
                            cm = "update pcscale.htickets set CapturedDateTime='" & sCap & "' where id=" & CStr(IdDoc)
                        Else
                            'cm = "update pcscale.localhtickets set captured='*' where id=" & CStr(IdDoc)
                            cm = "update pcscale.htickets set captured='*' where id=" & CStr(IdDoc)
                        End If
                        CadenadeLogTotales "Marcando:" & "update mbd.hTickets set Captured='*' where id=" & CStr(IdDoc)
                        
                        MyConnObj.Execute cm
                        ''''''''''
                        'CadenadeLog "paso 57"
                    Else 'REVISAR...
                    
                    
                            'CadenadeLog "paso 58"
                            nNumLin = 1
                            If MyRecSetTmp.Fields("saleform") = 1 Then
                                Cabecera.Lines(nNumLin).EsPesado = True
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            Else
                                Cabecera.Lines(nNumLin).EsPesado = False
                                Cabecera.Lines(nNumLin).cantidad = MyRecSetTmp.Fields("weight")
                            End If
                            'CadenadeLog "paso 59"
                            'CadenadeLog "8.4"
                            
                            Cabecera.Lines(nNumLin).FactorPeso = 1
                            Cabecera.Lines(nNumLin).precio = MyRecSetTmp.Fields("price")
                            
                            'CadenadeLog "8.5"
                            'CadenadeLog "paso 60"
                            Cabecera.Lines(nNumLin).NumeroLinea = nNumLin
                            Cabecera.Lines(nNumLin).CodigoArticulo = MyRecSetTmp.Fields("item")
                            'ver qué se hace IVA As Integer
                            Cabecera.Lines(nNumLin).IVA = 0
        
                            'CadenadeLog "paso 61"
                            'CadenadeLog "8.6"
                            
                            Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("linediscount") 'porcentaje
                            Cabecera.Lines(nNumLin).ImporteLinea = MyRecSetTmp.Fields("amount") 'MyRecSetTmp.Fields("importeconimpuestos") - MyRecSetTmp.Fields("dtoscabeceraconimpuestos")
                            Cabecera.Lines(nNumLin).ImporteBruto = Round(MyRecSetTmp.Fields("weight") * MyRecSetTmp.Fields("price"), 2) 'MyRecSetTmp.Fields("importe") - MyRecSetTmp.Fields("dtoscabecera")
                            'CadenadeLog "paso 62"
                            'CadenadeLog "8.7"
                            
                            'If MyRecSetTmp.Fields("idlineaabono") <> 0 Then '-1 Then 'era <>0
                            If MyRecSetTmp.Fields("linetype") = 0 Then
                                Cabecera.Lines(nNumLin).LineaCancelada = True
                            Else
                                Cabecera.Lines(nNumLin).LineaCancelada = False
                            End If
                            'CadenadeLog "paso 63"
                            'ver qué se hace con TipoLinea As Byte
                            'STV (positivo o negativo)
                            'Linea(Bucle).ModificadoTemporal = False
                            If MyRecSetTmp.Fields("amount") > 0 Then
                                'Cabecera.Lines(nNumLin).positive = 1
                                 Cabecera.Lines(nNumLin).FactorPeso = 1
                                 Cabecera.Lines(nNumLin).TipoLinea = 0
                            
                            Else
                                'Cabecera.Lines(nNumLin).positive = 0
                                 Cabecera.Lines(nNumLin).FactorPeso = -1
                                 Cabecera.Lines(nNumLin).TipoLinea = 2
                                
                            End If
                            'CadenadeLog "paso 64"
        '                    If Cabecera.Lines(nNumLin).positive = 1 Then
        '                         Cabecera.Lines(nNumLin).FactorPeso = 1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 0
        '                    Else
        '                         Cabecera.Lines(nNumLin).FactorPeso = -1
        '                         Cabecera.Lines(nNumLin).TipoLinea = 2
        '                    End If
                            'STP (pesado o no pesado)
                            If Cabecera.Lines(nNumLin).EsPesado Then
                                 'Cabecera.Lines(nNumLin).Weighted = 1
                            Else
                                 'Cabecera.Lines(nNumLin).Weighted = 0
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 1
                            End If
                            'CadenadeLog "paso 65"
                            
                            ' STC (cancelado o no cancelado)
                            If Cabecera.Lines(nNumLin).LineaCancelada Then
                                 Cabecera.Lines(nNumLin).TipoLinea = Cabecera.Lines(nNumLin).TipoLinea + 16
                            End If
                            'CadenadeLog "paso 66"
                            'If MyRecSetTmp.Fields("precio") <> MyRecSetTmp.Fields("preciodefecto") Then
                            '    Cabecera.Lines(nNumLin).SehaModificado = True
                            'Else
                                Cabecera.Lines(nNumLin).SehaModificado = False
                            'End If
                            '¿? ModificadoTemporal As Boolean
                            'CadenadeLog "paso 67"
                            'CadenadeLog "8.8"
                            'c2f 1.1.1 aclarar ficha Vacuno...
                            'If IsNull(MyRecSetTmp.Fields("eartag")) Or MyRecSetTmp.Fields("eartag") = "" Then
                                Cabecera.Lines(nNumLin).FichaVacuno = 0
                            'Else
                            '    Cabecera.Lines(nNumLin).FichaVacuno = MyRecSetTmp.Fields("eartag")
                            'End If
                            'descuentos de cada linea se añaden al descuento de cabecera para la tabla "DESCUENTOS")
                            'Cabecera.descuento = Cabecera.descuento + MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'Cabecera.Lines(nNumLin).DescuentoLinea = MyRecSetTmp.Fields("dtoslineaconimpuestos")
                            
                            'CadenadeLog "8.9"
                    
                        'CadenadeLog "paso 68"
                        Actualiza_Tabla_Cabecera Mybase, Cabecera
                        Do_Events
                        If Cabecera.Incluir_en_Base = True Then
                            Actualiza_Tabla_Tickets Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Gtarti Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_GtVend Mybase, Cabecera
                            Do_Events
                            actualiza_tabla_gtHora Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Gtsecs Mybase, Cabecera
                            Do_Events
                            actualiza_tabla_gtInfVen Mybase, Cabecera
                            Do_Events
                            Actualiza_Tabla_Log Mybase, Cabecera
                        End If
                        'CadenadeLog "paso 69"
                        'marcar...
                        If nLR = 0 Then
                            If Tabla = 1 Then
                                cm = "update pcscale.packetlabels set customer='9999999999' where id=" & CStr(IdDoc)
                            Else
                                cm = "update pcscale.sservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                            End If
                        Else
                            If Tabla = 1 Then
                                cm = "update pcscale.localpacketlabels set customer='9999999999' where id=" & CStr(IdDoc)
                            Else
                                cm = "update pcscale.localsservicelabels set customer='9999999999' where id=" & CStr(IdDoc)
                            End If
                        End If
                        MyConnObj.Execute cm
                        ''''''''''
                        'CadenadeLog "paso 70"
                    End If
                    
                    MyRecSetTmp.Update
        
                    MyRecSetTmp.MoveNext
                
                    
                    If cancelar = 1 Then GoTo SalirLeeTique
                    Do_Events
                    
                Loop
            End If
            'CadenadeLog "paso 71"
            'On Error Resume Next
            If MyRecSetTmp.State <> 0 Then
                MyRecSetTmp.Close
                Set MyRecSetTmp = Nothing
            End If
            'On Error GoTo 0
SigueNlr:
            'CadenadeLog "paso 72"
            Next nLR
            
            End If
            'CadenadeLog "paso 73"
            MisEqp.MoveNext
            
        Loop
        
        MisEqp.Close
        Set MisEqp = Nothing
        
    End If
    
SalirLeeTique:
''''''''
    On Error GoTo 0
    On Error GoTo fin
    'myRec.Close
    'Set myRec = Nothing
    If MyRecSetTmp.State <> 0 Then
        MyRecSetTmp.Close
        Set MyRecSetTmp = Nothing
    End If
    If Mybase.Name <> "" Then
    Mybase.Close
    Set Mybase = Nothing
    End If
fin:
    If Err.Number <> 0 Then
        CadenadeLog "Error LeeTickets:" & CStr(Err.Number) & " --> " & Err.Description
        On Error GoTo 0
        On Error GoTo FIN1
        If MyConnObj.State <> 0 Then
            MyConnObj.Close
            Set MyConnObj = Nothing
            Do_Events
            Sleep (200)
            Switch_mode 11
            Do_Events
            Sleep (200)
            EstadoBal = 1000
        End If
    End If
FIN1:
    On Error GoTo 0

End Sub

Public Function Erase_Totals(ByVal No_Reinit As Long) As Long
    '///////////////////////////////////////
    '// Borrado de Totales para TouchScale  //
    '///////////////////////////////////////
    Dim MyRecAux As New ADODB.Recordset
    Dim MyRecAux1 As New ADODB.Recordset
    Dim cm As String
    Dim nf As Integer
    Dim nC As Long
    Dim ncont As Integer
    Dim Resp As Long
    Dim IdEqp As Variant
   
    On Error Resume Next
    MyRecSet.Close
    Set MyRecSet = Nothing
    On Error GoTo 0

    CadenadeLog "Capa1A_GA_Erase_totals"
'*******
    
    
    CadenadeLog "Iniciar Numeración si Procede..."
    If No_Reinit = 0 Or lBoka Then 'Iniciar numeración tiques...
        'On Error Resume Next
        'cm = "delete from pcscale.localnumtickets"
        'MyConnObj.Execute cm
        'On Error GoTo 0
        'On Error Resume Next
        'cm = "delete from pcscale.numtickets"
        'MyConnObj.Execute cm
        'On Error GoTo 0
        'On Error Resume Next
        'cm = "delete from pcscale.erasetickets"
        'MyConnObj.Execute cm
        'On Error GoTo 0
        
        On Error Resume Next
        cm = "update pcscale.numberingtickets set resetnumber=1"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "update pcscale.erasetickets set dattim=now() where type='T'"
        MyConnObj.Execute cm
        On Error GoTo 0
        
        On Error Resume Next
        cm = "delete from pcscale.localnumtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        On Error Resume Next
        cm = "delete from pcscale.numtickets"
        MyConnObj.Execute cm
        On Error GoTo 0
        'On Error Resume Next
        'cm = "delete from pcscale.erasetickets"
        'MyConnObj.Execute cm
        'On Error GoTo 0
        
    End If
    CadenadeLog "Borrado Documentos y Lineas Documento..."
    On Error Resume Next
    cm = "delete from pcscale.HTickets where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.sservicelabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    cm = "delete from pcscale.packetlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next
    
    'Añadididas 1.0.4
    cm = "delete from pcscale.localerasetickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localhtickets where invoice=''"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketlabels"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpacketpresets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localpaytickets"
    MyConnObj.Execute cm
    On Error GoTo 0
    On Error Resume Next

    cm = "delete from pcscale.localsservicelabels"
    MyConnObj.Execute cm
    ''''''''''''''''''''''''''''''''
    On Error GoTo 0
    On Error Resume Next
    
    Resp = 0
    
FIN_Capa1A_GA_Erase_Totals:
    If Err.Number <> 0 Then
        CadenadeLog "Error en Capa1A_GA_Erase_Totals. Error:" & CStr(Err.Number) & " Descr.:" & Err.Description
        On Error GoTo 0
        Resp = 1004
    End If

    Erase_Totals = Resp
End Function

'''''''''''''''''''''''''''''''''
Public Sub export_BL_Fornes(ByVal lTodos As Boolean)
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
Dim sPes As String
Dim Total As String
Dim Signo As String
Dim Tipo_Linea As String
Dim Anulado As String
'************
Dim MiEnvio, Mienvio2, Mienvio3, sParte As String
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
Dim nLinea As Integer
Dim PathExpor As String
Dim PathExporRX As String
Dim lFirst As Boolean
Dim nLinLa As Integer
Dim nC As Integer
Dim nMess As Integer
'Dim NombreF As Tipo_NombreTiquet
'****************************************
Dim sSQL As String
Dim sruta As String
    'If lFornes Then
        If Dir("c:\telecom\epelsa", vbDirectory) = "" Then
            MkDir (App.Path & "\exportabl")
        End If
        'If Dir(App.Path & "\exportarx", vbDirectory) = "" Then
        '    MkDir (App.Path & "\exportarx")
        'End If
    'End If
    'If Val(TiendaActual) < 100 Then
    '    If TiendaActual = 0 And lFornes Then
            PathExpor = "c:\telecom\epelsa\" 'App.Path & "\exportaBL\"
            PathExporRX = "c:\telecom\epelsa\" 'App.Path & "\exportarx\"
    '    Else
    '        PathExpor = App.Path & "\t" & Format(TiendaActual, "00") & "\"
    '    End If
    'Else
    '    PathExpor = App.Path & "\t" & Format(TiendaActual, "000") & "\"
    'End If
    
    Set Base = DAO.OpenDatabase(App.Path & "\dbasetouch.mdb")
    
    If lTodos Then 'euroscale
        MiTabla1 = "Cabecera"
        MiTabla2 = "tickets"
    Else
        MiTabla1 = "CabeceraSC10"
        MiTabla2 = "ticketsSC10"
    End If
    sSQL = "SELECT * " & _
                  "FROM " & MiTabla1 & " " & _
                  "WHERE (reserv3=" & Chr(34) & "NO" & Chr(34) & " OR reserv3=" & Chr(34) & "--" & Chr(34) & _
                  ") order by val(nume)"
    
    'CadenadeLog sSQL
    
    Set Registro = Base.OpenRecordset(sSQL)
    If Not Registro.EOF Then
        
        
        Registro.MoveFirst
        
        
        Do While Not Registro.EOF And CI_Salir = False
        
            If lFornes Then
                Fichero2 = FreeFile()
                If Dir(PathExporRX & "rx.txt") <> "" Then
                    Open PathExporRX & "rx.txt" For Append Shared As #Fichero2
                Else
                    Open PathExporRX & "rx.txt" For Output As #Fichero2
                End If
                lFirst = True
                nLinLa = 0
                Mienvio3 = ""
                nMess = 0
            End If
            
            With Registro
            
            Codigo_mos = Format(Val(!seccion), "00")
            Codigo_bal = Format(Val(!Balanza), "00")
            Codigo_ven = Mid(Format(Val(!vended), "0000"), 3, 2)
            Lineas = Format(Val(!numlin) - Val(!lincan), "000") ' No se exportan las canceladas...
            Fecha = Format(!Fecha, "yyyymmdd")
            hora = Format(!hora, "hhmm")
            'código de Cliente se pasa a 4 de longitud según Informático.
            'sCodCli = Format(Val(!codcli), "00000")
            sCodCli = Format(Val(!codcli), "00000")
            
            MiEnvio = Codigo_mos & Format(Val(!nume), "00000")
            
            If lFornes Then
'CLAVE  "LA"
'GRUPO  3        4           2                 VARIOS/SEVERAL
'CLAVE  5        5           1                 VARIOS/SEVERAL
'   (A). Si el registro no es la última parte del ticket.
'   (B). Si este registro es la última parte del ticket. El importe corresponderá al total con descuento de todos los tramos que componen
'       el ticket.
'   (P): Es un Pedido. No se acumula en listados y no se incrementa el contador de ticket.
'   (Q):Si es la última parte de un pedido.
'   (R). Etiquetas
'descripción            Inicio Fin   Tamaño   Formato
'NUMERO DE TICKET       6      10    5        VARIOS/SEVERAL
'CODIGO DE CLIENTE      11     15    5        VARIOS/SEVERAL --> 0 SI NO HAY CLIENTE
'NUMERO DE TRAMO        16     17    2        VARIOS/SEVERAL --> '1,2,3..
'NUMERO MENSAJE         18     19    2        VARIOS/SEVERAL --> Número mensaje dentro del tramo: 0, 1, 2..
'SIGNO                  20     20    1        VARIOS/SEVERAL
'IMPORTE                21     29    9        VARIOS/SEVERAL
'VENDEDOR               30     34    5        VARIOS/SEVERAL
                Mienvio3 = Codigo_mos & "LA" & Codigo_mos & "B" & Format(Val(!nume), "00000") & sCodCli & "00"
            End If
            
            Set Reg2 = Base.OpenRecordset("select * from " & MiTabla2 & " where (reserv3=" & _
            Chr(34) & "NO" & Chr(34) & " and nume=" & Chr(34) & Registro!nume & Chr(34) & _
                     " and numbal=" & Registro!Balanza & " and date=" & Chr(34) & Registro!Fecha & Chr(34) _
                     & " and hour=" & Chr(34) & Registro!hora & Chr(34) _
                     & " and typtic=" & Chr(34) & Registro!TYPTIC & Chr(34) & ") order by val(numlin)")
            
            nLinea = 0
            
            If Not Reg2.EOF Then
                Reg2.MoveFirst
                                
                Do While Not Reg2.EOF
                    With Reg2
                        
                        Cod_Cod = Format(!Code, "000000")
                        
                        If !units = 0 Then
                            If !Weight >= 0 Then
                                Pes_Uni = !Weight
                                sPes = Format(1000 * (Pes_Uni), "00000")
                                Pes_Uni = Format(1000 * (Pes_Uni), "000000")
                            Else
                                Pes_Uni = (-1) * !Weight
                                sPes = Format(1000 * (Pes_Uni), "00000")
                                Pes_Uni = "-" & Format(1000 * (Pes_Uni), "00000")
                            End If
                            
                            Tipo_Linea = "0"
                        Else
                            If !units > 0 Then
                                Pes_Uni = !units
                                sPes = Format((Pes_Uni), "00000")
                                Pes_Uni = Format((Pes_Uni) * 1000, "000000")
                            Else
                                Pes_Uni = (-1) * !units
                                sPes = Format((Pes_Uni), "00000")
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
                            
                            nLinea = nLinea + 1
                            Num_lin = Format(nLinea, "000")
                            
                            Mienvio2 = ""
                            
                            If Val(Mid(Fecha, 5, 2)) < 10 Then
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Format(Val(Mid(Fecha, 5, 2)), "0") & ".TOT"
                            Else
                                sFich = "BL0" & Codigo_mos & Mid(Fecha, 7, 2) & Chr(55 + Val(Mid(Fecha, 5, 2))) & ".TOT"
                            End If
                            
                            Mienvio2 = MiEnvio & Num_lin & Cod_Cod & Pes_Uni & Total & Codigo_ven & Fecha & hora & sCodCli
                            
                            Fichero1 = FreeFile()
                            If Dir(PathExpor & sFich) <> "" Then
                                Open PathExpor & sFich For Append Shared As #Fichero1
                            Else
                                Open PathExpor & sFich For Output As #Fichero1
                            End If
                            Print #Fichero1, Mienvio2
                            Close #Fichero1
                            
                            If lFornes Then
'CODIGO       35      40           6                 VARIOS/SEVERAL
'PESO 2 + 3   41      45           5                 VARIOS/SEVERAL
'SIGNO        46      46           1                 VARIOS/SEVERAL
'IMPORTE      47      55           9                 VARIOS/SEVERAL
'CODIGO       56      61           6                 VARIOS/SEVERAL
'PESO 2 + 3   62      66           5                 VARIOS/SEVERAL
'SIGNO        67      67           1                 VARIOS/SEVERAL
'IMPORTE      68      76           9                 VARIOS/SEVERAL
'CODIGO       77      82           6                 VARIOS/SEVERAL
'PESO 2 + 3   83      87           5                 VARIOS/SEVERAL
'SIGNO        88      88           1                 VARIOS/SEVERAL
'IMPORTE      89      97           9                 VARIOS/SEVERAL
'CODIGO       98    103            6                 VARIOS/SEVERAL
'PESO 2 + 3   104   108            5                 VARIOS/SEVERAL
'SIGNO        109    109           1                 VARIOS/SEVERAL
'IMPORTE      110    118           9                 VARIOS
'HORA Y FECHA 119    128          10  (HHMMDD VARIOS/SEVERAL
                                If lFirst = True Then
                                    lFirst = False
                                    sParte = ""
                                End If
                                nLinLa = nLinLa + 1
                                sParte = sParte & Cod_Cod & sPes & Signo & Total
                                
                                If nLinLa = 4 Then 'escribir registro
                                    nLinLa = 0
                                    If Registro!Importe > 0 Then 'añadir número de mensaje.
                                        sPes = Mienvio3 & Format(nMess, "00") & "0" & Format(Registro!Importe * 100, "000000000")
                                    Else
                                        sPes = Mienvio3 & Format(nMess, "00") & "1" & Format(Registro!Importe * (-100), "000000000")
                                    End If
                                    'Mienvio3 = Mienvio3 & Format(Registro!vended, "00000")
                                    nMess = nMess + 1
                                    Print #Fichero2, sPes & Format(Registro!vended, "00000") & sParte & hora & Format(CDate(Registro!Fecha), "ddmmyy")
                                    lFirst = True
                                End If
                            End If
                            
                        End If
                        
                        
                        Reg2.Edit
                        Reg2.Fields("reserv3") = "SI"
                        Reg2.Update
                        Reg2.MoveNext
                        
                        
                    
                    End With
                    
                    
                    
                Loop
                
                If lFornes Then
                    If nLinLa <> 0 Then 'completar registro
                        For nC = nLinLa + 1 To 4
                            sParte = sParte & "000000000000000000000"
                        Next nC
                        If Registro!Importe > 0 Then 'añadir número de mensaje.
                            sPes = Mienvio3 & Format(nMess, "00") & "0" & Format(Registro!Importe * 100, "000000000")
                        Else
                            sPes = Mienvio3 & Format(nMess, "00") & "1" & Format(Registro!Importe * (-100), "000000000")
                        End If
                        'Mienvio3 = Mienvio3 & Format(Registro!vended, "00000")
                        Print #Fichero2, sPes & Format(Registro!vended, "00000") & sParte & hora & Format(CDate(Registro!Fecha), "ddmmyy")
                    End If
                End If
                
                
            End If
            
            Registro.Edit
            Registro.Fields("reserv3") = "SI"
            Registro.Update
            
            Close #Fichero2
            
            Registro.MoveNext
        
            End With
            
            If lFornes Then Close #Fichero2
            
        Loop 'loop registros cabecera
    
        'End With
        
    End If 'si hay cabeceras pendientes
    
    'Next nSysToca
    
    'On Error Resume Next
    'Close #Fichero1
    'On Error GoTo 0
    
    'Close #Fichero2
    
    Base.Close
    Set Base = Nothing
    'lCogeTiquet = True

End Sub


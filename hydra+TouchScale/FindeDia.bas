Attribute VB_Name = "s"
Option Explicit
'//////////////
'//////////////
'//////////////
Public Sub Trata_Fin_Dia()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Dim sZip As String
    Dim nResp As Integer
    Dim siP As String
    Dim susR As String
    Dim spsS As String
    Dim Ret, nFich
    Dim sext As String
    Dim dB1 As dao.Database
    Dim RstdB1 As dao.Recordset
    Dim nf As Integer
    Dim cm As String
    
    Contador = 0
    If Dir(Miruta & "\end_day.ini") <> "" Then
        'exportar formas de pago
        'Base_General = App.Path & "\dbase.mdb"
        Set dB1 = AbrirBase
        Set RstdB1 = dB1.OpenRecordset("select d_fecha,tipopago,sum(importe) as imp,count(*) as cuantos from cabecera  where reserv3<>'SI' and tipopago<5 group by d_fecha,tipopago")
        If Not RstdB1.EOF Then
            RstdB1.MoveFirst
            nf = FreeFile()
            Open App.Path & "\mpagos.dat" For Output As #nf
        
            Do While Not RstdB1.EOF
                'MsgBox "Importe:" & CStr(RstdB1.Fields("imp")) & " , " & "Fecha:" & RstdB1.Fields("d_fecha") & " , " & "Tipo pago:" & RstdB1.Fields("tipopago") & " , " & "Cuantos:" & RstdB1.Fields("cuantos")
                cm = Format(RstdB1.Fields("d_fecha"), "dd/mm/yy") & Format(RstdB1.Fields("cuantos"), "000000") & CStr(RstdB1.Fields("tipopago"))
                Select Case RstdB1.Fields("tipopago")
                    Case 0
                    cm = cm & "Moneda       "
                    Case 1
                    cm = cm & "Tarjeta      "
                    Case 2
                    cm = cm & "Cheque       "
                    Case 3
                    cm = cm & "T.Restaurante"
                    Case 4
                    cm = cm & "Genérico     "
                End Select
                cm = cm & Format(RstdB1.Fields("imp") * 100, "000000000")
                Print #nf, cm
                RstdB1.Movenext
            Loop '
        
            Close #nf
        
        End If
        RstdB1.Close
        dB1.Close
        '''''''''''''''''''''''''''''''''''''
    
        Archivo = FreeFile()
        Open Miruta & "\end_day.ini" For Input As Archivo
        Do Until EOF(Archivo)
            Contador = Contador + 1
            ReDim Preserve Buffer(Contador + 1)
            Line Input #Archivo, Buffer(Contador)
        Loop
        Close #Archivo
        Load FrmExportar
        FrmExportar.Hide
        For bucle = 1 To Contador
            Buffer(bucle) = Trim(Buffer(bucle))
            If Buffer(bucle) <> "" Then
               FrmExportar.MostrarMensajes = False
               FrmExportar.ExportarFichero 0, Buffer(bucle), CDate(Trim(Mid(Now(), 1, Len(Now()) - 8))), False '2.0.24 (no estaba ,true de continua)
               '2.0.2 frmMonitor.MostrarDato CargaCadena(984) & " " & Buffer(Bucle)
                If Dir(App.Path & "\getputfileftp.dat") <> "" And Dir(Exp_Path & "\" & Buffer(bucle)) <> "" Then
                    nFich = FreeFile()
                    Open App.Path & "\getputfileftp.dat" For Input As #nFich
                    'Do While Not EOF(nFich)
                        Line Input #nFich, siP
                        Line Input #nFich, susR
                        Line Input #nFich, spsS
                    'Loop
                    Close #nFich
                    
                    nResp = ElPing(siP)
                    If nResp = 1 Then
                    
                        If nRetConnFTP = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            If Ret = 1 Then nRetConnFTP = 1
                        Else
                            Ret = 1
                        End If
                        If Ret = 1 Then 'ha sido posible conectar
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                                
                                
                            End If
                        End If
                        If Ret = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            'Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                                
                            End If
                        
                        End If
                    
                    End If
                    
                    If Ret = 0 Or nResp = 0 Then
                        nRetConnFTP = 0
                    End If
                    
                End If
            
            End If
        Next bucle
        
        Set FrmExportar = Nothing
    End If
End Sub

Public Sub Trata_Fin_Dia_1()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Dim MyProc As Long
    Dim MyResul As Long
    Dim sZip As String
    Dim nResp As Integer
    Dim siP As String
    Dim susR As String
    Dim spsS As String
    Dim Ret, nFich
    Dim dB1 As dao.Database
    Dim RstdB1 As dao.Recordset
    Dim nf As Integer
    Dim cm As String
    
    Contador = 0
    If Dir(Miruta & "\end_day.ini") <> "" Then
        'exportar formas de pago
        'Base_General = App.Path & "\dbase.mdb"
        Set dB1 = AbrirBase
        Set RstdB1 = dB1.OpenRecordset("select d_fecha,tipopago,sum(importe) as imp,count(*) as cuantos from cabecera  where reserv3<>'SI' and tipopago<5 group by d_fecha,tipopago")
        If Not RstdB1.EOF Then
            RstdB1.MoveFirst
            nf = FreeFile()
            Open App.Path & "\mpagos.dat" For Output As #nf
        
            Do While Not RstdB1.EOF
                'MsgBox "Importe:" & CStr(RstdB1.Fields("imp")) & " , " & "Fecha:" & RstdB1.Fields("d_fecha") & " , " & "Tipo pago:" & RstdB1.Fields("tipopago") & " , " & "Cuantos:" & RstdB1.Fields("cuantos")
                cm = Format(RstdB1.Fields("d_fecha"), "dd/mm/yy") & Format(RstdB1.Fields("cuantos"), "000000") & CStr(RstdB1.Fields("tipopago"))
                Select Case RstdB1.Fields("tipopago")
                    Case 0
                    cm = cm & "Moneda       "
                    Case 1
                    cm = cm & "Tarjeta      "
                    Case 2
                    cm = cm & "Cheque       "
                    Case 3
                    cm = cm & "T.Restaurante"
                    Case 4
                    cm = cm & "Genérico     "
                End Select
                cm = cm & Format(RstdB1.Fields("imp") * 100, "000000000")
                Print #nf, cm
                RstdB1.Movenext
            Loop '
        
            Close #nf
        
        End If
        RstdB1.Close
        dB1.Close
        '''''''''''''''''''''''''''''''''''''
        Archivo = FreeFile()
        Open Miruta & "\end_day.ini" For Input As Archivo
        Do Until EOF(Archivo)
            Contador = Contador + 1
            ReDim Preserve Buffer(Contador + 1)
            Line Input #Archivo, Buffer(Contador)
        Loop
        Close #Archivo
        Load FrmExportar
        FrmExportar.Hide
        For bucle = 1 To Contador
            Buffer(bucle) = Trim(Buffer(bucle))
            If Buffer(bucle) <> "" Then
               FrmExportar.MostrarMensajes = False
               FrmExportar.ExportarFichero 0, Buffer(bucle), CDate(Trim(Mid(Now(), 1, Len(Now()) - 8))), True '2.0.24 (no estaba ,true de continua)
               '2.0.2 frmMonitor.MostrarDato CargaCadena(984) & " " & Buffer(Bucle)
            
                If Dir(App.Path & "\getputfileftp.dat") <> "" And Dir(Exp_Path & "\" & Buffer(bucle)) <> "" Then
                    nFich = FreeFile()
                    Open App.Path & "\getputfileftp.dat" For Input As #nFich
                    'Do While Not EOF(nFich)
                        Line Input #nFich, siP
                        Line Input #nFich, susR
                        Line Input #nFich, spsS
                    'Loop
                    Close #nFich
                    
                    nResp = ElPing(siP)
                    If nResp = 1 Then
                    
                        If nRetConnFTP = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            If Ret = 1 Then nRetConnFTP = 1
                        Else
                            Ret = 1
                        End If
    
                        'If Ret = 1 Then 'ha sido posible conectar
                        '    Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                        'End If
                        ''Ret = Desconectar_FTP(0)
                        ''nRetConnFTP = 0
                        'If Ret = 0 Then
                        '    Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                        '    Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                        'End If
                        
                        If Ret = 1 Then 'ha sido posible conectar
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                                
                            End If
                        End If
                        If Ret = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            'Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                                
                            End If
                        
                        End If
                    End If
                    
                    If Ret = 0 Or nResp = 0 Then
                        nRetConnFTP = 0
                    End If
                End If
            
            End If
        Next bucle
        Set FrmExportar = Nothing
    End If
End Sub

Public Sub Trata_Fin_Dia_2()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Dim MyProc As Long
    Dim MyResul As Long
    Dim sZip As String
    Dim nResp As Integer
    Dim siP As String
    Dim susR As String
    Dim spsS As String
    Dim Ret, nFich
    Dim dB1 As dao.Database
    Dim RstdB1 As dao.Recordset
    Dim nf As Integer
    Dim cm As String
    
    Contador = 0
    If Dir(Miruta & "\end_day.ini") <> "" Then
        'exportar formas de pago
        'Base_General = App.Path & "\dbase.mdb"
        Set dB1 = AbrirBase
        Set RstdB1 = dB1.OpenRecordset("select d_fecha,tipopago,sum(importe) as imp,count(*) as cuantos from cabecera  where reserv3<>'SI' and tipopago<5 group by d_fecha,tipopago")
        If Not RstdB1.EOF Then
            RstdB1.MoveFirst
            nf = FreeFile()
            Open App.Path & "\mpagos.dat" For Output As #nf
        
            Do While Not RstdB1.EOF
                'MsgBox "Importe:" & CStr(RstdB1.Fields("imp")) & " , " & "Fecha:" & RstdB1.Fields("d_fecha") & " , " & "Tipo pago:" & RstdB1.Fields("tipopago") & " , " & "Cuantos:" & RstdB1.Fields("cuantos")
                cm = Format(RstdB1.Fields("d_fecha"), "dd/mm/yy") & Format(RstdB1.Fields("cuantos"), "000000") & CStr(RstdB1.Fields("tipopago"))
                Select Case RstdB1.Fields("tipopago")
                    Case 0
                    cm = cm & "Moneda       "
                    Case 1
                    cm = cm & "Tarjeta      "
                    Case 2
                    cm = cm & "Cheque       "
                    Case 3
                    cm = cm & "T.Restaurante"
                    Case 4
                    cm = cm & "Genérico     "
                End Select
                cm = cm & Format(RstdB1.Fields("imp") * 100, "000000000")
                Print #nf, cm
                RstdB1.Movenext
            Loop '
        
            Close #nf
        
        End If
        RstdB1.Close
        dB1.Close
        '''''''''''''''''''''''''''''''''''''
        Archivo = FreeFile()
        Open Miruta & "\end_day.ini" For Input As Archivo
        Do Until EOF(Archivo)
            Contador = Contador + 1
            ReDim Preserve Buffer(Contador + 1)
            Line Input #Archivo, Buffer(Contador)
        Loop
        Close #Archivo
        Load FrmExportar
        FrmExportar.Hide
        For bucle = 1 To Contador
            Buffer(bucle) = Trim(Buffer(bucle))
            If Buffer(bucle) <> "" Then
               FrmExportar.MostrarMensajes = False
               FrmExportar.ExportarFichero 0, Buffer(bucle), CDate(Trim(Mid(Now(), 1, Len(Now()) - 8))), False '2.0.24 (no estaba ,true de continua)
               '2.0.2 frmMonitor.MostrarDato CargaCadena(984) & " " & Buffer(Bucle)
            
                If Dir(App.Path & "\getputfileftp.dat") <> "" And Dir(Exp_Path & "\" & Buffer(bucle)) <> "" Then
                    nFich = FreeFile()
                    Open App.Path & "\getputfileftp.dat" For Input As #nFich
                    'Do While Not EOF(nFich)
                        Line Input #nFich, siP
                        Line Input #nFich, susR
                        Line Input #nFich, spsS
                    'Loop
                    Close #nFich
                    
                    nResp = ElPing(siP)
                    If nResp = 1 Then
                    
                        If nRetConnFTP = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            If Ret = 1 Then nRetConnFTP = 1
                        Else
                            Ret = 1
                        End If
    
                        'If Ret = 1 Then 'ha sido posible conectar
                        '    Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                        'End If
                        ''Ret = Desconectar_FTP(0)
                        ''nRetConnFTP = 0
                        'If Ret = 0 Then
                        '    Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                        '    Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                        'End If
                        If Ret = 1 Then 'ha sido posible conectar
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                            
                            End If
                        End If
                        If Ret = 0 Then
                            Ret = Conectar_FTP(siP, susR, spsS, 0) '0=no ha podido conectar, 1=se ha conectado
                            'Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(Bucle), Buffer(Bucle), 0)
                            If Dir(App.Path & "\nozip.txt") <> "" Then
                                Ret = Enviar_Fichero_FTP(Exp_Path & "\" & Buffer(bucle), Buffer(bucle), 0)
                            Else
                                If LCase(Exp_Path) <> LCase(App.Path) Then
                                    FileCopy Exp_Path & "\" & Buffer(bucle), App.Path & "\" & Buffer(bucle)
                                End If
                                
                                Ret = Zipear(Miruta, Buffer(bucle))
                            
                            End If
                        
                        End If
                    
                    End If
                    If Ret = 0 Or nResp = 0 Then
                        nRetConnFTP = 0
                    End If
                End If
            
            End If
        Next bucle
        Set FrmExportar = Nothing
    End If
End Sub

Public Sub Trata_Boka()
    Dim Archivo As Integer
    Dim bucle As Integer
    Dim Buffer() As String
    Dim Contador As String
    Contador = 0
    Load FrmExportar
    FrmExportar.Hide

    CadenadeLog "Se exporta:" & "boka.txt"
    'FrmExportar.MostrarMensajes = False
    FrmExportKuups.ExportarFichero_BOKA 0, "boka.txt", 0, True, True
    
    Set FrmExportKuups = Nothing
    
End Sub

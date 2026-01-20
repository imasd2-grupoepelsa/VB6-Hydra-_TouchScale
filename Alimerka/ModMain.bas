Attribute VB_Name = "ModMain"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type MyType
    MyS() As String
End Type
Public myRec(10000) As MyType
Public nEnvios As Integer
Public nRechazos As Integer
Public sArrIP(10) As String
Public nServidoras As Integer
Public sArrSTATUS(10) As String
Public nSerEnCur As Integer

'Formato de fichero de importación
'El nombre del fichero será:
'Se busca en c:\program files (x86)\storeline\office\export\
'Nombre de fichero: DIBAL_MART_YYYYMMDD_HHMMSS.TXT
'una vez procesado se mueve con el mismo nombre a :
'c:\program files (x86)\storeline\office\export\processed\
'Siendo:
'-  YYYYMMDD la fecha de generación del fichero
'-  HHMMSS la hora de generación del fichero
'Se usará el separador |
'descripción Contenido
'Operación   M-Alta o modificación, B-Borrado
'PLU Código de artículo (cuatrodígitos)
'Familia Existe una tabla de conversion de cuatro a dos dígitos
'Precio  Siete digitos completados con ceros a la izquierdo y sin punto decimal.
'1ª Línea descripción    LABEL_DESC1 (Items.Name y Items.Shortname)
'2ª Línea descripción    LABEL_DESC2 (Items.Text1)
'3ª Línea descripción    LABEL_DESC3 (Items.Text2)
'Tipo de venta   Pesado=1, No Pesado=0 (Nota: Al contrario que en tiques)
'Formato de etiqueta Numero de etiqueta, si no viene nada, etiqueta por defecto.
'Días de caducidad   Si tiene cad. Número de días, si no, 0
'Ingredientes    Si lleva ingredientes, (Items.longtext) Este campo puede no estar programado.

'"   La tabla de relación de departamentos (familias):
'Campo en
'fichero Familia a programar
'0201    21
'0202    22
'0203    23
'0204    24
'0205    25
'0206    26
'0207    27
'0208    28
'0209    29
'7002    30
Public Sub Main()
Dim nF As Integer
Dim nC As Integer
Dim sS As String
Dim sFile As String
Dim sIp As String
Dim sPathORG As String
Dim sPathDES As String
Dim Resp As Long
Dim lBorTot As Boolean
Dim lReinit As Long
Dim Bucle As Integer
    If App.PrevInstance Then
        LogingE Now() & " -->Instancia previa en funcionamiento...", "", ""
        End
    End If
    LogingE Now() & " --> Arranque.", "", ""
    Load frmSysTray
    If Dir("C:\tiquetbal", vbDirectory) = "" Then
        MkDir "C:\tiquetbal\"
    End If
    If Dir(App.Path & "\procesados", vbDirectory) = "" Then
        MkDir App.Path & "\procesados"
    End If
    nServidoras = 0
    If Dir(App.Path & "\servidoras.txt") <> "" Then
        nF = FreeFile()
        Open App.Path & "\servidoras.txt" For Input As #nF
        Do While Not EOF(nF)
            Line Input #nF, sIp
            nServidoras = nServidoras + 1
            sArrIP(nServidoras) = sIp
        Loop
        Close nF
    Else
        If Dir(App.Path & "\serip.txt") <> "" Then
            nF = FreeFile()
            Open App.Path & "\serip.txt" For Input As #nF
            Line Input #nF, sIp
            Close nF
            nServidoras = 1
            sArrIP(1) = sIp
        Else
            MsgBox "Dirección IP Servidora no definida...", vbCritical
            End
        End If
    End If
    
    sPathORG = App.Path & "\"
    sPathDES = App.Path & "\procesados\"
    On Error GoTo sigueEntrada
    If Dir(App.Path & "\paths.txt") <> "" Then
        nF = FreeFile()
        Open App.Path & "\paths.txt" For Input As #nF
        Line Input #nF, sPathORG
        Line Input #nF, sPathDES
        Close #nF
        If Right(sPathORG, 1) = "\" Then
            sPathORG = Mid(sPathORG, 1, Len(sPathORG) - 1)
        End If
        If Right(sPathDES, 1) = "\" Then
            sPathDES = Mid(sPathDES, 1, Len(sPathDES) - 1)
        End If
        'If Dir("c:\program files (x86)\storeline", vbDirectory) = "" Then
        '          'C:\Program Files (x86)
        '    MkDir "c:\program files (x86)\storeline"
        'End If
        'If Dir(sPathORG, vbDirectory) = "" Then
        '    MkDir sPathORG
        'End If
        'If Dir(sPathDES, vbDirectory) = "" Then
        '    MkDir sPathDES
        'End If
        sPathORG = sPathORG & "\"
        sPathDES = sPathDES & "\"
    End If
sigueEntrada:
    If Err.Number <> 0 Then
        CadenadeLog "Error en arranque:" & CStr(Err.Number) & " -->  " & Err.Description
        MsgBox "Error en arranque. Por favor verifique: 'paths.txt'y 'serip.txt'", vbCritical
        End
    End If
    On Error GoTo 0
    On Error GoTo sigueERROR
    
    For Bucle = 1 To nServidoras
        sAddIP = sArrIP(Bucle)
        Resp = Capa1B_SendFullItem_TouchScale(True)
    Next Bucle
    
    lBorTot = False
    lReinit = 1
    If UCase(Trim(Command)) = "BORINI" Then
        lBorTot = True
        lReinit = 0
    End If
    If UCase(Trim(Command)) = "BORNOINI" Then
        lBorTot = True
        lReinit = 1
    End If
    
    Do While 1
        nF = FreeFile()
        sFile = Format(Now, "yyyymmdd")
        'sFile = "20180806"
        sFile = Dir(sPathORG & "DIBAL_MART_" & sFile & "_*.txt")
        If sFile <> "" Then
            Open sPathORG & sFile For Input As #nF
            CadenadeLog "Fichero detectado:" & sPathORG & sFile
            nC = 0
            nEnvios = 0
            nRechazos = 0
            Do While (Not EOF(nF)) And (nC < 10000)
                Line Input #nF, sS
                If Len(sS) > 10 Then
                    myRec(nC).MyS() = Split(sS, "|")
                    'Select Case Val(myRec(nC).MyS(2))
                    'Case 201
                    'myRec(nC).MyS(2) = "21"
                    'Case 202
                    'myRec(nC).MyS(2) = "22"
                    'Case 203
                    'myRec(nC).MyS(2) = "23"
                    'Case 204
                    'myRec(nC).MyS(2) = "24"
                    'Case 205
                    'myRec(nC).MyS(2) = "25"
                    'Case 206
                    'myRec(nC).MyS(2) = "26"
                    'Case 207
                    'myRec(nC).MyS(2) = "27"
                    'Case 208
                    'myRec(nC).MyS(2) = "28"
                    'Case 209
                    'myRec(nC).MyS(2) = "29"
                    'Case 7002
                    'myRec(nC).MyS(2) = "30"
                    'End Select
                    If Val(myRec(nC).MyS(1)) > 9999 Then
                        nRechazos = nRechazos + 1
                        CadenadeLog "Registro rechazado (Item nr.):" & myRec(nC).MyS(1)
                    Else
                        nC = nC + 1
                    End If
                Else
                    nRechazos = nRechazos + 1
                    CadenadeLog "Registro con longitud incorrecta..."
                End If
            Loop
            Close #nF
            
            Sleep (200)
            FileCopy sPathORG & sFile, sPathDES & sFile
            Sleep (200)
            Kill sPathORG & sFile
            'sAddIP = "10.8.1.185:3306"
            nEnvios = nC
            CadenadeLog "Fichero procesado:" & sFile & " --> Registros procesados:" & CStr(nEnvios) & " // Registros rechazados:" & CStr(nRechazos)
            For Bucle = 1 To nServidoras
                nSerEnCur = Bucle
                sAddIP = sArrIP(Bucle)
                CadenadeLog "Envío datos a Servidora:" & sAddIP
                Resp = Capa1B_SendFullItem_TouchScale(False)
            Next Bucle
            'CadenadeLog "Fichero procesado:" & sFile & " --> Registros procesados:" & CStr(nEnvios) & " // Registros rechazados:" & CStr(nRechazos)
        End If
        
        Do_Events
        'CadenadeLog "Captura de tiques..."
        For Bucle = 1 To nServidoras
            nSerEnCur = Bucle
            sAddIP = sArrIP(Bucle)
            Call LeeTickets
        Next Bucle
        Do_Events
        
        If Dir(App.Path & "\fin.dat") <> "" Then
            Kill App.Path & "\fin.dat"
            'Sleep (200)
            End
        End If
sigueERROR:
        If Err.Number <> 0 Then
            CadenadeLog "Error en bucle:" & CStr(Err.Number) & " --> " & Err.Description
        End If
        On Error GoTo 0
        For nF = 1 To 5
            Do_Events
        Next nF
        If lBorTot Then
            For Bucle = 1 To nServidoras
                sAddIP = sArrIP(Bucle)
                CadenadeLog "Borrado de Totales..."
                Switch_mode 11
                Resp = Erase_Totals(lReinit)
            Next Bucle
        End If
    Loop
End Sub

Public Sub CadenadeLog(MiCadena As String)
    Dim Fichero As Integer
    On Error GoTo FIN
    
    Fichero = FreeFile
    If Dir(App.Path & "\log.txt") <> "" Then
        If FileLen(App.Path & "\log.txt") > 3000000 Then
            If Dir(App.Path & "\log.bck") <> "" Then
                Kill App.Path & "\log.bck"
            End If
            Sleep (200)
            FileCopy App.Path & "\log.txt", App.Path & "\log.bck"
            Sleep (200)
            Kill App.Path & "\log.txt"
            Sleep (200)
        End If
        Open App.Path & "\log.txt" For Append As Fichero
    Else
        Open App.Path & "\log.txt" For Output As Fichero
    End If
    On Error Resume Next
    Print #Fichero, Now() & "-->" & MiCadena
    Close #Fichero
    Exit Sub
FIN:
    On Error GoTo 0
End Sub

Public Sub Do_Events()
    Sleep 2
    DoEvents
End Sub

Public Function DigitoControl(ByVal Valor As String) As Integer
   '**************************
   ' Calcula el Checksum de
   ' un código de barras
   '***************************
   Dim Longitud As Integer
   Dim sumapar As Integer
   Dim sumaimpar As Integer
   Dim suma As Integer
   Dim digito, i As Integer
   sumapar = 0
   sumaimpar = 0
   Longitud = Len(Valor)
   For i = 1 To Longitud
      If i Mod 2 = 0 Then
         sumapar = sumapar + Val(Mid(Valor, i, 1))
      Else
         sumaimpar = sumaimpar + Val(Mid(Valor, i, 1))
      End If
   Next
   If Longitud Mod 2 = 0 Then
      suma = sumapar * 3 + sumaimpar
   Else
      suma = sumapar + sumaimpar * 3
   End If

   digito = 10 - suma Mod 10
   If digito = 10 Then digito = 0
   DigitoControl = digito
End Function

Public Sub LogingE(ByVal sInfo As String, ByVal sError As String, ByVal sWarning As String)
    '-- Start Event Logging
    Call App.StartLogging("", vbLogToNT)
    
    '-- Log Events to NT
    If sInfo <> "" Then Call App.LogEvent(sInfo, vbLogEventTypeInformation)
    If sError <> "" Then Call App.LogEvent(sError, vbLogEventTypeError)
    If sWarning <> "" Then Call App.LogEvent(sWarning, vbLogEventTypeWarning)
End Sub

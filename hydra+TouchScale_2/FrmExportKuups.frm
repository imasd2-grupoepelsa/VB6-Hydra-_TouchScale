VERSION 5.00
Begin VB.Form FrmExportKuups 
   ClientHeight    =   900
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   1560
End
Attribute VB_Name = "FrmExportKuups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TipoFragmento
    id As Integer
    inicio As Long
    fin As Long
    Adicional As String
End Type
Private Fragmento() As TipoFragmento
Private fragmentos As Long

Public Sub ExportarFichero_BOKA(ElBoton As Integer, Optional MiArchivo As String, Optional MiFecha As Date, Optional Continua As Boolean, Optional NoExportados As Boolean)
    Dim Fichero As Integer
    Dim Buffer As String
    Dim FaEx As String
    
    NoExportados = True
    If MiArchivo <> "" Then
        FaEx = MiArchivo
    Else
        FaEx = "boka.txt"
    End If
    Fichero = FreeFile()
    Open App.Path & "\etc\" & FaEx & ".fet" For Input As Fichero
    Line Input #Fichero, Buffer
    Close #Fichero
    If InStr(1, Buffer, "+") Then
        Buffer = Mid(Buffer, InStr(1, Buffer, "+") - 1)
    End If
    
    ExportaLinea FaEx, MiFecha, Continua, NoExportados
End Sub

Private Sub ExportaLinea(Elfichero As String, Optional MiFecha As Date, Optional lCont As Boolean, Optional NoExportados As Boolean, Optional ByVal mIfECHA2 As Date)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Archivo As Integer
    Dim bucle As Long
    Dim Buffer As String
    Dim Contbuffer As Long
    Dim Cadenilla As String
    Dim CMax As Integer
    Dim CMin As Integer
    Dim BucleGrande As Long
    Dim Condicion As String
    Dim ContReg As Long
    Dim BGrande As Long
    Dim NSistema As Integer 'c2f/caspiunza
    Dim nSisLoop As Integer 'c2f/caspiunza
    Dim sSs As String
    Dim sPar As String
    Dim nFichHys As Integer
    Dim sBas As String
    
    If MiFecha = "0" Then
        Condicion = " where (reserv3 <>'SI')" & " order by val(nume),val(numlin)"
    Else
        'Condicion = " where cdate(d_fecha)=cdate(" & Chr(34) & MiFecha & Chr(34) & ") order by val(nume),val(numlin)"
        Condicion = " where d_fecha=#" & Format(MiFecha, "dd/mm/yyyy") & "# order by val(nume),val(numlin)"
    End If
    Set Base = AbrirBase 'abrirbase 'OpenDatabase(Base_General)
   
    LeeFragmentos (Elfichero)
    
    Set Registro = Base.OpenRecordset("select * from tickets" & Condicion)
    With Registro
        If Not .EOF Then
            .MoveFirst
            .MoveLast
            'Barra.Max = .RecordCount + 1
            ContReg = .Recordcount
            'Barra.Value = 0
            Archivo = FreeFile
            On Error Resume Next
            If MiFecha = "0" And InStr(1, UCase(Elfichero), "BOKA") <> 0 Then
                If Dir(Exp_Path & "\" & Elfichero) <> "" Then
                    Open Exp_Path & "\" & Elfichero For Append As #Archivo
                Else
                    Open Exp_Path & "\" & Elfichero For Output As #Archivo
                End If
            Else
                Open Exp_Path & "\" & Elfichero For Output As #Archivo
            End If
            If Err.Number <> 0 Then
                CadenadeLog "Error. File or Folder not Available : " & Exp_Path & "\" & Elfichero
                Exit Sub
            End If
            On Error GoTo 0
            .MoveFirst
            For bucle = 1 To ContReg
                Buffer = ""
                Contbuffer = 1
                'Barra.Value = Barra.Value + 1
                
                ExportaunaLinea Buffer, Contbuffer, Registro, BucleGrande, "", lCont '2.0.24
                
                Print #Archivo, Buffer
                
                .Movenext
            Next bucle
        End If
    End With
    
    Close #Archivo
    Base.Close
    
    Unload Me
End Sub

Private Sub ExportaunaLinea(Buffer As String, ByRef Contbuffer As Long, Registro As dao.Recordset, ByRef BucleGrande As Long, BufCab As String, Optional Continua As Boolean, Optional lEspecial As Boolean)
    Dim bucle As Long
    Dim laLongitud As Integer
    Dim MiLen As Integer
    Dim sM As String
    Dim Exp_Mas, Exp_Menos As String
    Dim MiValor As Long
    
    Exp_Mas = "0"
    Exp_Menos = "-"
    
    If IsNull(lEspecial) Then lEspecial = False
        
    With Registro
    For bucle = 1 To fragmentos
        Do While Contbuffer < Fragmento(bucle).inicio
            Contbuffer = Contbuffer + 1
            'Buffer = Buffer & " "
        Loop
        laLongitud = Fragmento(bucle).fin - Fragmento(bucle).inicio + 1
        Select Case Val(Fragmento(bucle).id)
            Case 0
                Buffer = Buffer & Formatea(Fragmento(bucle).Adicional, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 1
                Buffer = Buffer & Formatea(FormFecha(!Date), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 2
                Buffer = Buffer & Formatea(FormHora(!Hour), laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 3
                Buffer = Buffer & Formatea(Val(!Vendor), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 4
                'If TipoExport <> 3 Then
                    Buffer = Buffer & Formatea(Val(!nume), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'Else
                '    If BucleGrande = 1 Then
                '        Buffer = Buffer & GA
                '    Else
                '        Buffer = Buffer & SC10
                '    End If
                '    Buffer = Buffer & Formatea(Val(!nume), laLongitud - 1, True)
                '    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 5
                'If Sabeco And !lincan = 1 Then
                '    Buffer = Buffer & Formatea(SabLineas, laLongitud, True)
                '    Contbuffer = Contbuffer + laLongitud
                '    SabLineas = SabLineas + 1
                'Else
                    Buffer = Buffer & Formatea(Val(!numlin), laLongitud, True)
                    Contbuffer = Contbuffer + laLongitud
                'End If
            Case 6
                Buffer = Buffer & Formatea(Val(!secc), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 7
                If Dir(App.Path & "\decepal.txt") <> "" Then
                    Buffer = Buffer & FormateaR(!Code, laLongitud, False) '  Formatea(Val(!Code), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(Val(!Code), laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 8
                If Not lEspecial Then
                If UsaEuro Then
                    Buffer = Buffer & Formatea(!Price * 100, laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(!Price * (10 ^ decimales), laLongitud, True)
                End If
                Else
                    sM = CStr(!Price)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic_B(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 9
                If Not lEspecial Then
                If UsaEuro Then
                    Buffer = Buffer & Formatea(!amount * 100, laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(!amount * (10 ^ decimales), laLongitud, True)
                End If
                Else
                    sM = CStr(!amount)
                    If IsNull(sM) Or sM = "0" Then sM = "0,00"
                    sM = formatic_B(sM, 2, laLongitud)
                    Buffer = Buffer & sM
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 10
                If Not lEspecial Then
                Buffer = Buffer & Formatea(!Weight * 1000, laLongitud, True)
                Else
                    sM = CStr(!Weight)
                    If IsNull(sM) Or sM = "0" Then sM = "0,000"
                    sM = formatic_B(sM, 3, laLongitud)
                    Buffer = Buffer & sM
                                     
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 11
                Buffer = Buffer & Formatea(!units, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 12
                If !units = 0 Then
                    If Not lEspecial Then
                        Buffer = Buffer & Formatea(!Weight * 1000, laLongitud, True)
                    Else
                        sM = CStr(!Weight)
                        If IsNull(sM) Or sM = "0" Then sM = "0,000"
                        sM = formatic_B(sM, 3, laLongitud)
                        Buffer = Buffer & sM
                    
                    End If
                Else
                    'If Not Sabeco Then
                        Buffer = Buffer & Formatea(!units, laLongitud, True)
                    'Else
                    '    Buffer = Buffer & Formatea(!units * 1000, laLongitud, True)
                    'End If
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 13
                Buffer = Buffer & Formatea(!TYPTIC, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 14
                Buffer = Buffer & Formatea(Val(!STPETIC), laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 15
                If !amount >= 0 Then
                    Buffer = Buffer & Formatea(Exp_Mas, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Exp_Menos, laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 16
                MiLen = 1 + Val(Right(Fragmento(bucle).Adicional, 2)) - Val(left(Fragmento(bucle).Adicional, 2))
                If Val(Right(Fragmento(bucle).Adicional, 2)) <= Len(BufCab) Then
                    Buffer = Buffer & Formatea(Mid(BufCab, Val(left(Fragmento(bucle).Adicional, 2)), MiLen), laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Space(MiLen), laLongitud, False)
                End If
                Contbuffer = Contbuffer & laLongitud
            Case 17
                If Not IsNull(!Item) Then
                    Buffer = Buffer & Formatea(!Item, laLongitud, False)
                Else
                    Buffer = Buffer & Formatea(Space(25), laLongitud, False)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 18
                If !units <> 0 Then
                    Buffer = Buffer & Formatea("1", laLongitud, False) '¡¡ era 2 C2F
                Else
                    If !Weight <> 0 Then
                        Buffer = Buffer & Formatea("0", laLongitud, False) '¡¡ era 1 C2F
                    Else
                        Buffer = Buffer & Formatea("1", laLongitud, False) '¡¡ era 0 C2F
                    End If
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 19
                Buffer = Buffer & Formatea(!lincan, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 20
                If Not IsNull(!Plu) Then
                    Buffer = Buffer & Formatea(Val(!Plu), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 21
                If Not IsNull(!numbal) Then
                    Buffer = Buffer & Formatea(!numbal, laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 22
                'tiendaactual=1
                'Buffer = Buffer & Formatea(TiendaActual, laLongitud, True)
                Buffer = Buffer & Formatea(1, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 23 'c2f invicta --> Departamento (familia)
                If Not IsNull(!DPT) Then
                    Buffer = Buffer & Formatea(Val(!DPT), laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(0, laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 24 'c2f invicta importe bruto
                If UsaEuro Then
                    Buffer = Buffer & Formatea(!importe_bruto * 100, laLongitud, True)
                Else
                    Buffer = Buffer & Formatea(!importe_bruto * (10 ^ decimales), laLongitud, True)
                End If
                Contbuffer = Contbuffer + laLongitud
            Case 25 'c2f lote
                Buffer = Buffer & Formatea(!numlote, laLongitud, False)
                Contbuffer = Contbuffer + laLongitud
            Case 26 'c2f cliente
                Buffer = Buffer & Formatea(!cliente, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            'Case 27 'c2f factura
            '    Buffer = Buffer & Formatea(!factura, laLongitud, False)
            '    Contbuffer = Contbuffer + laLongitud
            Case 27 '% 0
                MiValor = !porcentaje * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            Case 28 'c2f factura Simpli.
                'Buffer = Buffer & Formatea(!factura, laLongitud, False)
                'Contbuffer = Contbuffer + laLongitud
                On Error Resume Next
                Buffer = Buffer & Formatea(!factura, laLongitud, False)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Buffer = Buffer & Formatea(" ", laLongitud, False)
                End If
                On Error GoTo 0
                Contbuffer = Contbuffer + laLongitud
            Case 29
                MiValor = !porcentaje * 100
                Buffer = Buffer & Formatea(MiValor, laLongitud, True)
                Contbuffer = Contbuffer + laLongitud
            
        End Select
    Next bucle
    If Continua Then
        Registro.Edit
        !reserv3 = "SI"
        .Update
    End If
    End With
    
End Sub

Private Sub LeeFragmentos(Elfichero As String)
    Dim bucle As Long
    Dim Bucle2 As Long
    Dim Archivo As Integer
    Dim Buffer As String
    Dim Buf2 As String
    Dim Buf3 As String
    Dim FragmentoAux As TipoFragmento
    Dim C1 As String
    Dim C2 As String
    Dim C3 As String
    Dim C4 As String
    Dim Contador As Long
    Buf2 = ""
    For bucle = Len(Elfichero) To 1 Step -1
        If Mid(Elfichero, bucle, 1) <> "\" Then
            Buf2 = Buf2 & Mid(Elfichero, bucle, 1)
        Else
            Exit For
        End If
    Next bucle
    Buf3 = ""
    For bucle = Len(Buf2) To 1 Step -1
        Buf3 = Buf3 & Mid(Buf2, bucle, 1)
    Next bucle
    Archivo = FreeFile
    Open App.Path & "\etc\" & Buf3 & ".fet" For Input As Archivo
    Line Input #Archivo, Buffer
    If Val(Buffer) = 4 Or Val(Buffer) = 5 Then
        Input #Archivo, C1, C2, C3
        'TipoExport = Val(C1)
        'GA = C2
        'SC10 = C3
    End If
    Contador = 0
    Do Until EOF(Archivo)
        Contador = Contador + 1
        ReDim Preserve Fragmento(Contador + 1)
        Input #Archivo, C1, C2, C3, C4
        Fragmento(Contador).inicio = Val(C1)
        Fragmento(Contador).fin = Val(C2)
        Fragmento(Contador).id = Val(left(C3, 3))
        Fragmento(Contador).Adicional = C4
    Loop
    Close #Archivo
    fragmentos = Contador
    ' ordena
    For bucle = 1 To fragmentos - 1
        For Bucle2 = bucle + 1 To fragmentos
            If Fragmento(bucle).inicio > Fragmento(Bucle2).inicio Then
                FragmentoAux = Fragmento(bucle)
                Fragmento(bucle) = Fragmento(Bucle2)
                Fragmento(Bucle2) = FragmentoAux
            End If
        Next Bucle2
    Next bucle
  
End Sub
Private Function Formatea(cadena, Longitud As Integer, EsNumerico As Boolean) As String
    Dim Buffer As String
    Dim BufferFormato As String
    Dim bucle As Long
    Dim Exp_Absoluto As Boolean
    
    Exp_Absoluto = True
    
    If EsNumerico Then
        For bucle = 1 To Longitud
            BufferFormato = BufferFormato & "0"
        Next bucle
        Buffer = Format(cadena, BufferFormato)
    Else
        Buffer = cadena
        If Len(Buffer) < Longitud Then
            Buffer = Buffer & Space(Longitud - Len(Buffer))
        End If
        Buffer = left(Buffer, Longitud)
    End If
    If Exp_Absoluto Then
        If EsNumerico Then
            If Val(Buffer) < 0 Then
                Buffer = Right(Buffer, Len(Buffer) - 1)
                Buffer = "0" & Buffer
            End If
        End If
    End If
    If Len(Buffer) > Longitud Then
        'Buffer = Right(Buffer, Longitud)
        If left(Buffer, 1) = "-" Then
            Buffer = "-" & Mid(Buffer, 3)
        Else
            Buffer = Right(Buffer, Longitud)
        End If
    End If
    Formatea = Buffer
End Function
Private Function FormateaP(cadena, Longitud As Integer, EsNumerico As Boolean) As String
    Dim Buffer As String
    Dim BufferFormato As String
    Dim bucle As Long
    Dim Exp_Absoluto As Boolean
    
    Exp_Absoluto = True
    
    If EsNumerico Then
        For bucle = 1 To Longitud
            BufferFormato = BufferFormato & "0"
        Next bucle
        Buffer = Format(cadena, BufferFormato)
    Else
        Buffer = cadena
        If Len(Buffer) < Longitud Then
            Buffer = Buffer & Space(Longitud - Len(Buffer))
        End If
        Buffer = left(Buffer, Longitud)
    End If
    If Exp_Absoluto Then
        If EsNumerico Then
            If Val(Buffer) < 0 Then
                Buffer = Right(Buffer, Len(Buffer) - 1)
                Buffer = "0" & Buffer
            End If
        End If
    End If
    If Len(Buffer) > Longitud Then
        Buffer = Right(Buffer, Longitud)
    End If
    Buffer = Mid(Buffer, 2)
    Buffer = Mid(Buffer, 1, Len(Buffer) - 3) & "." & Right(Buffer, 3)
    FormateaP = Buffer
End Function

Private Function FormateaR(cadena, Longitud As Integer, EsNumerico As Boolean) As String
    Dim Buffer As String
    Dim BufferFormato As String
    Dim bucle As Long
    Dim Exp_Absoluto As Boolean
    Exp_Absoluto = True
    If EsNumerico Then
        For bucle = 1 To Longitud
            BufferFormato = BufferFormato & "0"
        Next bucle
        Buffer = Format(cadena, BufferFormato)
    Else
        Buffer = cadena
        If Len(Buffer) < Longitud Then
            Buffer = Space(Longitud - Len(Buffer)) & Buffer
        End If
        Buffer = left(Buffer, Longitud)
    End If
    If Exp_Absoluto Then
        If EsNumerico Then
            If Val(Buffer) < 0 Then
                Buffer = Right(Buffer, Len(Buffer) - 1)
                Buffer = "0" & Buffer
            End If
        End If
    End If
    If Len(Buffer) > Longitud Then
        Buffer = Right(Buffer, Longitud)
    End If
    FormateaR = Buffer
End Function

Private Function FormFecha(MiFecha As String) As String
    Dim Retorno As String
    Dim Dia As String
    Dim Mes As String
    Dim Anyo As String
    Dia = left(MiFecha, 2)
    Mes = Mid(MiFecha, 4, 2)
    Anyo = Right(MiFecha, 2)
    '**************
    ' mifecha --> ddmmyy
    '**************
    'Select Case Exp_Fecha
    '    Case 0
    '        Retorno = Dia & "/" & Mes & "/" & Anyo
    '    Case 1
    '        Retorno = Mes & "/" & Dia & "/" & Anyo
    '    Case 2
            Retorno = Dia & "/" & Mes & "/" & "20" & Anyo
    '    Case 3
    '        Retorno = Mes & "/" & Dia & "/" & "20" & Anyo
    '    Case 4
    '        Retorno = Dia & Mes & Anyo
    '    Case 5
    '        Retorno = Mes & Dia & Anyo
    '    Case 6
    '        Retorno = Dia & Mes & "20" & Anyo
    '    Case 7
    '        Retorno = Mes & Dia & "20" & Anyo
    '    Case 8
    '        Retorno = Dia & "-" & Mes & "-" & Anyo
    'End Select
    FormFecha = Retorno
End Function
Private Function FormHora(MiHora As String) As String
    
    Dim Mhora As String
    Dim MMinuto As String
    Dim Retorno As String
    '****************
    ' mihora --> hh:mm
    '****************
    Mhora = left(MiHora, 2)
    MMinuto = Mid(MiHora, 4, 2)
    'Select Case Exp_hora
    '    Case 0
            Retorno = Mhora & ":" & MMinuto
    '    Case 1
    '        Retorno = Mhora & MMinuto
    '    Case 2
    '        Retorno = Mhora & "-" & MMinuto
    '    Case 3
    '        Retorno = Mhora
    'End Select
    FormHora = Retorno
End Function

Private Function formatic_B(sM As String, nDec As Integer, laLo As Integer) As String
   Dim nL As Integer
   Dim nB As Integer
   Dim nD
   nL = Len(sM)
   nD = 0
   For nB = 1 To nL
          'MsgBox Format(Asc(Mid(sM, nB, 1)), "000")

       If Mid(sM, nB, 1) = "," Then
       
          nD = nB
          nB = nL + 1
       End If
   
   Next nB
   If nD <> 0 Then
       If (nL - nD) < nDec Then
          
          sM = sM & Mid("0000000000", 1, nDec - (nL - nD))
           
       End If
   
   Else
       sM = sM & "," & Mid("0000000000", 1, nDec)
   End If
   sM = Mid("00000000000000000000", 1, laLo - Len(sM)) & sM
   formatic_B = sM
End Function


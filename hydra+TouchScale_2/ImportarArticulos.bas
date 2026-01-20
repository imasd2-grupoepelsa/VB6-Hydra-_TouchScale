Attribute VB_Name = "importararticulos"
Option Explicit
Sub ImportaElFichero(ByVal fichero As String)
    Dim tipo(DimensionArticulo) As String
    Dim Inicio(DimensionArticulo) As Long
    Dim Final(DimensionArticulo) As Long
    Dim posicion(DimensionArticulo) As Integer
    Dim contador As Long
        
    Dim dato As String
    Dim buffer As String
    Dim fileConfig As Integer
    Dim fileImport As Integer
    
    Dim bucle As Long
    inicia
    fileConfig = FreeFile
    fileImport = FreeFile
    ' Abre fichero de configuración y lo lee
    PathImport = fichero
    On Error GoTo noexiste1
    Open PathImport & "\" & fichero & ".cfg" For Input As fileConfig
    On Error GoTo 0
    contador = 0
    Do While Not EOF(fileConfig)
        Input #fileConfig, tipo(contador), buffer, Inicio(contador), Final(contador), posicion(contador)
        contador = contador + 1
    Loop
    Close fileConfig
    
     
     ' Abre fichero a importar
    On Error GoTo noexiste2
    Open PathImport & "\" & fichero For Input As fileImport
    On Error GoTo 0
    
    
    ' Lee las líneas
    Do While Not EOF(fileConfig)
        Line Input #fileImport, buffer
        ' Almacena la información de la línea y comprueba que es correcta
        
        If Val(buffer) <> 0 Then
            BorrarArticulo
            For bucle = 0 To contador - 1
                dato = Mid(buffer, Inicio(bucle), Final(bucle) - Inicio(bucle) + 1)
                If tratadato(tipo(bucle), dato, posicion(bucle)) = False Then
                    MsgBox "Datos incorrectos en fichero de configuración"
                    Exit Sub
                End If
            Next bucle
    ' Aquí se introduce el registro en la base
            TratarRegistro
        End If
    Loop
    MsgBox "Importación finalizada. Envíe modificaciones."
    
    Exit Sub
    
noexiste1:

    MsgBox "No se pudo abrir el fichero de configuración de importación"
    Exit Sub
    
noexiste2:

    MsgBox "No se pudo abrir el fichero de importación"
    Exit Sub
    
End Sub

Private Sub TratarRegistro()
' ***************************************************************
' Daniel Campos Fernández
' 25-7-2001
' Esta función toma un registro a importar, que se halla en la
' variable "articulo", y estudia si se ha de insertar un registro
' en la base, modificar uno existente, o marcar uno para borrar.
' La función es llamada desde cmdimportar_click()
' ***************************************************************
' Aviso :
' ¿y si varias líneas tras borrado?
'
    Dim a As Long
    Dim bucle As Long
    Dim cadenaVacia As Boolean
    Dim varmarcador As Variant
    Dim Modificar As Boolean
    Dim ExisteCodigo As Boolean
    Dim ExistePluySeccion As Boolean
    Dim valido As Byte
    Dim base As Database
    Dim registro As Recordset
    Dim Equipos As Recordset
    Dim strpeticion
    Dim ModificacionBasica As Boolean
    Dim ModificacionTexto0 As Boolean
    Dim ModificacionTexto5 As Boolean
    Dim PluEsCero As Boolean
    
    Set base = OpenDatabase(PathImport & "\dbase.mdb")
    Set registro = base.OpenRecordset("Articulo", dbOpenDynaset)
    Set Equipos = base.OpenRecordset("Equipos", dbOpenDynaset)
    ' Comprueba si no existe
    If registro.RecordCount <> 0 Then
    
       ExisteCodigo = True
       ExistePluySeccion = True
       Modificar = False
    
       strpeticion = "codigo = " & arTIculo.cod
        ' ¿Existe el código?
       With registro
             .MoveLast
             .FindFirst strpeticion
             If .NoMatch Then
                ExisteCodigo = False
                ' Si existe, pero está marcado para borrar, se considera que
                ' no existe para añadir una nueva línea
             Else
                If !borrado = True Then
                ' se asegura de que no se haya añadido ya otro igual
                     ExisteCodigo = False
                     Do While (Not .NoMatch) And ExisteCodigo = False
                          .FindNext strpeticion
                          
                     If !borrado = True Then
                           ExisteCodigo = False
                     Else
                            ExisteCodigo = True
                     End If
                     Loop
                End If
            End If
       End With
    '********************************************************************************
    ' Si el plu está definido se pide sección y plu, si no, sólo la sección
    '********************************************************************************
       If arTIculo.plu = 0 Then
            PluEsCero = True
            arTIculo.plu = registro!plu
       Else
            PluEsCero = False
       End If
     ' ¿Existe Plu y sección?
       strpeticion = "secc_maqui = " & arTIculo.sec & "and plu = " & arTIculo.plu
       With registro
            .MoveLast
            .FindFirst strpeticion
            If .NoMatch Then
                ExistePluySeccion = False
            Else
                ' Si existe, pero está marcado para borrar, se considera que
                ' no existe para añadir una nueva línea
                ' se asegura de que no se haya añadido ya otro igual
                If !borrado = True Then
                    ExistePluySeccion = False
                    Do While (Not .NoMatch) And ExisteCodigo = False
                        .FindNext strpeticion
                        If !borrado = True Then
                          ExistePluySeccion = False
                        Else
                         ExistePluySeccion = True
                   End If
                   Loop
                End If
            End If
       End With
        ' Si existe seccion+plu o codigo, tienen que coincidir
        ' y además ser iguales la familia y la subsección
        
       If ExistePluySeccion Or ExisteCodigo Then
            strpeticion = "secc_maqui = " & arTIculo.sec & "and plu = " & arTIculo.plu & "and codigo = " & arTIculo.cod & "and codi_sub = " & arTIculo.sub & "and codi_fam = " & arTIculo.fam
          
          With registro
            .MoveLast
            .FindFirst strpeticion
            If .NoMatch Then ' Si no coinciden, sale sin hacer nada
                base.Close
                Exit Sub
            Else ' Si coincide, modificará en lugar de insertar
                Modificar = True
                varmarcador = .Bookmark
            End If
          End With
       End If
    Else
        If arTIculo.plu = 0 Then arTIculo.plu = 1
    End If
    ' ***********************************************************************************
    ' ***********************************************************************************
    ' ¿Marcado para destruir?
    
    ' comprueba que txt(0) está en blanco
    
    a = Len(arTIculo.txt(0))
    cadenaVacia = True
    For bucle = 1 To a
        If Mid(arTIculo.txt(0), bucle, 1) <> " " Then
            cadenaVacia = False
            bucle = a + 1
        End If
    Next bucle
    ' Comprueba que se dan las condiciones para borrar
    
    If arTIculo.prc = 0 And cadenaVacia = True Then
      With registro
            .Edit
            !borrado = True
            !tran_plu = ""
            .Update
            .Bookmark = .LastModified
      End With
      Exit Sub
    End If
    
    ' Si no existe el registro, se inserta
    ' Si existe, se modifica
    '
    ' 1º Si plu es 0 y hay que añadir registro, hay que buscar un plu libre
    If PluEsCero = True And Modificar = False Then
        For bucle = 1 To 9999
            strpeticion = "secc_maqui = " & arTIculo.sec & "and plu = " & bucle
           With registro
                .MoveLast
                .FindFirst strpeticion
                If .NoMatch Then
                    arTIculo.plu = bucle
                    bucle = 10000
                End If
           End With
        Next bucle
    End If
    '
    ' 2º Modifica o añade
    ' En primer lugar comprueba que marcas hay que cambiar
    If Modificar = True Then
        ModificacionBasica = False
        ModificacionTexto0 = False
        ModificacionTexto5 = False
        With registro
            .Bookmark = varmarcador
            If (!tara <> arTIculo.tar) Or (!caducidad <> arTIculo.cad) Then ModificacionBasica = True
            If (!precio <> arTIculo.prc) Or (!codi_pes <> arTIculo.wgh) Then ModificacionBasica = True
            For bucle = 25 To 35
                If .Fields(bucle) <> arTIculo.tlt(bucle - 25) Then ModificacionBasica = True
            Next bucle
            For bucle = 36 To 40
                If .Fields(bucle) <> arTIculo.txt(bucle - 36) Then
                    ModificacionBasica = True
                    ModificacionTexto0 = True
                End If
            Next bucle
             For bucle = 41 To 56
                If .Fields(bucle) <> arTIculo.txt(bucle - 36) Then
                    ModificacionBasica = True
                    ModificacionTexto5 = True
                    
                End If
            Next bucle
        End With
    Else
        ModificacionBasica = True
        ModificacionTexto0 = True
        ModificacionTexto5 = True
    End If
    With registro
        If Modificar = False Then
           .AddNew
        Else
            .Bookmark = varmarcador
            .Edit
        End If
        
        !codigo = arTIculo.cod
        ' Corrección 26-07-2001
        ' Para completar el campo codi_ident (código de tres cifras)
        ' Busca correspondencia entre secc_maqui (2digitos) y codi_ident (3 digitos))
        strpeticion = "secc_maqui =" & arTIculo.sec
        Equipos.MoveLast
        Equipos.FindFirst strpeticion
            registro!codi_ident = Equipos!codi_ident
        
        ' e introduce la sección
        !secc_maqui = arTIculo.sec
        !plu = arTIculo.plu
        !codi_sub = arTIculo.sub
        !codi_fam = arTIculo.fam
        !precio = arTIculo.prc
        !caducidad = arTIculo.cad
        !tara = arTIculo.tar
        !ean8 = arTIculo.cba
        !tipo_desc = arTIculo.typ
        !etq = arTIculo.etq
        !pref = arTIculo.pre
        !codi_pes = arTIculo.wgh
        ' textos
        For bucle = 36 To 56
            .Fields(bucle) = arTIculo.txt(bucle - 36)
        Next bucle
        ' tipos de letra
        For bucle = 25 To 35
            .Fields(bucle) = arTIculo.tlt(bucle - 25)
        Next bucle
                    
        ' se se añaden las marcas que correspondan
        If ModificacionBasica = True Then !tran_plu = ""
        If ModificacionTexto0 = True Then !tran_texto = ""
        If ModificacionTexto5 = True Then !tran_tx1 = ""
            
        .Update
        .Bookmark = .LastModified
    End With
    base.Close
      
    
    
End Sub
Private Sub BorrarArticulo()
' ***************************************************************
' Daniel Campos Fernández
' 25-7-2001
' Esta función se utiliza para poner a cero todos los valores
' de la estructura que almacena temporalmente los datos de
' cada línea del fichero de configuración leída.
' Es necesario poner todos los valores a cero cada vez que
' Se lee una nueva línea para no mezclar datos de las anteriores.
' ***************************************************************
    Dim bucle As Integer
    With arTIculo
    .cod = 0
    .plu = 0
    .sec = 0
    .sub = 0
    .fam = 0
    .prc = 1
    
    .wgh = ""
    .cad = 0
    .pre = 0
    .tar = 0
    .etq = 0
    .cba = 0
    For bucle = 0 To 10
        .tlt(bucle) = 0
    Next bucle
    For bucle = 0 To 20
        .txt(bucle) = ""
    Next bucle
    .typ = 0
    
    End With
End Sub

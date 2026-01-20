Attribute VB_Name = "db"
Option Explicit
'''''''''''''''
'''''''''''''''
'''''''''''''''
'¡ATENCIÓN A LA CONEXIÓN...
' Por defecto en la Balanza el puerto para conectar será: 3306
'''''''''''''''
#Const BALSAM = False
'***********************
' Valores que devuelven
' las funciones
' 0 --> Alta realizada, o borrado
' 1 --> Modificación realizada
' 2 --> El dato (equipo, plu, código...) no corresponde con la sección. Rechazo
' 3 --> No Existe la sección. Rechazo
' 4 --> No se ha modificado nada (los datos eran iguales a los existentes)
' 5 --> Código de Barras, dato con menos de 12 caracteres.Rechazo
' 6 --> Longitud de publicidad u oferta excesiva.Rechazo
' 7 --> Número de Mostrador fuera de rango.Rechazo
' 8 --> Número de Sección fuera de rango.Rechazo
' 9 --> Texto de descripción (21 caracteres) demasiado largo.Rechazo
' 10 --> Artículos asociados al mostrador a borrar.Rechazo
' 11 --> Equipos asociados al mostrador a borrar.Rechazo
' 12 --> Número de Balanza fuera de Rango.Rechazo
' 13 --> Se necesita el nro. IP de la balanza.Rechazo
' 14 --> Falta código del artículo.Rechazo
' 15 --> Falta el número de mostrador.Rechazo
' 16 --> Error en Precios por tramos.Rechazo
' 17 --> Base de precios por tramos Llena.Rechazo
' 18 --> Tres datos simultáneos en función que no los admite (sec+plu) y codigo.Rechazo
' 19 --> Vendedor, falta tecla (Rechazo)
' 20 --> Vendedor, tecla ya asignada a otro vendedor en esa sección
' 21 --> Articulo, PLU repetido
' 22 --> Nº decódigo de cliente fuera de rango
' 50 --> No puede haber más de un mostrador de ese tipo (ECONET/V8/ECOLABEL)
' 51 --> Sección máquina incompatible (ECO/V8)
' 52 --> Balanza incompatible con el tipo de sección indicado
' 53 --> Ya existe una Balanza EcoLabel/net/V8 en el mostrador
' 54 --> Grupo de Conservacion incorrecto
' 60 --> El dato buscado ya no existe
' 61 --> No hay secciones para la familia
Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
'***********************
' Tipos de las diferentes tablas de la base
'***********************
Public Type DB_Vendedor
    NVendedor As Long
    Descripcion As String
    tecla As Integer
    seccion As Integer
    estado As Boolean
End Type
Public Type DB_Balanzas
    NBalanza As Integer
    NMostrador As Integer
    Descripcion As String
    DireccionIP As String
    ' Tipo --> 0 --> Euroscale no servidora, 1 --> Euroscale Servidora, 100 --> V-12
    Tipo As Integer
    lss165 As Boolean
    
End Type
Public Type DB_Cabeceras
    Cabecera(5) As String
    leyenda(5) As String
    TLCabecera(5) As String
    TLLeyenda(5) As String
    Mostrador As Integer
    Equipo As Integer
    Reservado As Integer
End Type
Public Type DB_CodigoBarras
    ' datos : Red 0-->cb_vnt,1-->cb_sup,2-->cb_mix,3-->cb_ev1,4-->cb_ev2
    '         Local 5-->cb_vnt, 6 --> cb_sup,7-->cb_mix,8-->cb_ev1,9 --> cb_ev2
    Mostrador As Integer
    Equipo As Integer
    Datos(10) As String
End Type
Public Type DB_Publicidad
    ' Activación : 0 --> Desactivar todo, 1 --> Activar Contínua, 2 --> Activar Ofertas
    Activacion As Integer
    OFERTAS(10) As String
    Publicidad As String
    Equipo As Integer
    Mostrador As Integer
End Type
Public Type DB_mostrador
    ' Tipo : 0 --> Artículos a Euroscale, 1 --> Artículos a SC10, 2 --> A Ambos sistemas
    Tipo As Integer
    NMostrador As Integer
    NSeccion As Integer
    name As String
    Descripcion As String
    PuertoCOM As Long ' puerto COM para ecolabel, econet y v8
    DireccionIP As String ' dirección IP Euroscale para Sistemas con múltiples NetID
End Type
Public Type DB_Articulo
    codigo As Long
    Mostrador As Integer
    Plu As Integer
    subsec As Integer
    familia As Integer
    precio As Double
    caducidad As Integer
    tara As Double
    WGH As Boolean
    Balenv As Boolean
    GrpConserv As Integer
    Descriptivos(21) As String
    TipoLetra(10) As Integer
    Preferente As Integer
    Etiqueta As Integer
    ean13 As String
    TRM1 As Double
    TRM2 As Double
    PRC1 As Double
    PRC2 As Double
    IVA As Integer
    Merma As Integer
    Presel As Integer
    PRC100G As Boolean ' EcoLabel, precio por 100 gramos
    
    descriptivos_2040(20) As String 'cas->v171
    lMix As Boolean
    nPoid As Long
    
    PRC3 As Double
    Imagen As String
    posicion As Integer
    Label1 As String
    Label2 As String
    ean14 As String
    
    tarifa(10) As Double
    onkey As Integer
    
End Type
Public Type DB_Teclas
    Destino_NMostrador As Integer
    Destino_NBalanza As Integer
    Destino_Tabla As Integer
    Destino_CodiTec As Integer
    CodArticulo As Long
    PluArticulo As Long
    SecArticulo As Integer
    lEs165 As Boolean
    Section As Integer
    Tipo As String
    Descripcion As String
End Type
Public Type DB_Cliente
    codigo As Long
    Nombre As String
    Datos(4) As String
    factura As Boolean
    discount As Double
End Type
Public Type DB_Ficha
    codigo As Integer
    Identificador As String
    RSI_Sacrificio As String
    RSI_Despiece As String
    texto As String
    lote As String
    Fecha_Nacimiento As String
    Fecha_Sacrificio As String
    Fecha_Despiece As String
    Categoria As Integer
    Raza As Integer
    Cod_nacimiento As Integer
    Cod_Crianza As Integer
    Cod_Sacrificio As Integer
    Cod_Despiece As Integer
    Cod_Produccion As Integer
    Edad As Integer
    Sexo As String
    Peso As Double
End Type

Public Type Tipo_Linea
    EsPesado As Boolean
    FactorPeso As Integer
    precio As Double
    cantidad As Double
    NumeroLinea As Byte
    CodigoArticulo As Long
    IVA As Integer
    TipoLinea As Byte
    DescuentoLinea As Double
    ImporteLinea As Double
    LineaCancelada As Boolean
    SehaModificado As Boolean
    ModificadoTemporal As Boolean
    FichaVacuno As Long
    ImporteBruto As Double
     descriptivo As String
    porcentaje As Double
     Plu As Long
     desc As String
     familia As Long
End Type


Public Type tipo_cabecera
     Ntiquet As Long
     NVendedor As Long
     NMostrador As Long
     NCliente As String
     Fecha As Variant
     hora As Variant
     NBalanza As Long
     ImporteTotal As Double
     nLineas As Long
     LineaInicial As Long
     LineaFinal As Long
     codigo As Long
     CadenaEstadoTiquet As String
     ImporteBruto As Double
     Impuestos As Double
     descuento As Double
     LineasCanceladas As Long
     tipoTiquet As Long
     EstadoTiquet As Long
     LineasMensaje As Long
     '********************
     ' este dato se introduce
     ' tras pasar por actualiza_tabla_cabecera
     Incluir_en_Base As Boolean
     tipoPago As Long
     cantidadEntre As Double
     Cambio As Double
     lote As String
     lEncargo As Boolean
     documentoabono As Long
     Trainning As Boolean
     Lines() As Tipo_Linea
     base0 As Double
     por0 As Double
     imp0 As Double
     base1 As Double
     por1 As Double
     imp1 As Double
     base2 As Double
     por2 As Double
     imp2 As Double
     base3 As Double
     por3 As Double
     imp3 As Double
     base4 As Double
     por4 As Double
     imp4 As Double
     factura As String
     ean13 As String
     TotalWeight As Double
     TotalUnits As Double
End Type

'\\\\\\\
'///////
Public Sub CargaIdiomas()
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Long
    Dim MyID As Integer
    MyID = (id / 1000)
    MyID = MyID + 1
    Set Base = dao.OpenDatabase(App.Path & "\langtouch\langtouch.mdb")
    Set Registro = Base.OpenRecordset("select max(id) from texts")
    With Registro
        nIdioma = Val(.Fields(0))
        ReDim Preserve sIdioma(nIdioma + 1)
    End With
    Set Registro = Base.OpenRecordset("select * from texts order by id")
    With Registro
        .MoveFirst
        Do Until .EOF
            If .Fields("id") = 1373 Then
            bucle = bucle
            End If
            If Not IsNull(.Fields("text" & Format(MyID, "0"))) Then
                sIdioma(.Fields("id")) = .Fields("text" & Format(MyID, "0"))
            End If
            .Movenext
        Loop
    End With
    Base.Close
End Sub
Public Function CargaCadena(ByVal NumCad As Integer) As String
    If NumCad <= nIdioma Then
        CargaCadena = sIdioma(NumCad)
    Else
        CargaCadena = "Cad : " & NumCad
    End If
End Function
'\\\\\\\\\\\
'///////////
Private Function Record_Action(Registro As dao.Recordset, lAction As Long) As Boolean
Dim RetVal As Boolean
Dim ErrorVal As Long
Dim Continuar As Boolean
    Continuar = True
    RetVal = False
    Do While Continuar
        On Error Resume Next
        If lAction = 0 Then
            Registro.Edit
            ErrorVal = Err.Number
        End If
        If lAction = 1 Then
            Registro.Delete
            ErrorVal = Err.Number
        End If
        On Error GoTo 0
        Select Case ErrorVal
            Case 0
                RetVal = True
                Continuar = False
            Case 3260 ' otro está bloqueando
                Sleep (10)
            Case 3197 ' el registro ha sido borrado
                RetVal = False
                Continuar = False
            Case 3027   'cas.n->registro de solo lectura
                RetVal = False
                Continuar = False
            Case Else ' base dañada
                If lAction = 0 And (ErrorVal = 3021 Or ErrorVal = 91) Then 'c2f jordi
                    RetVal = True
                    Continuar = False
                
                Else
                
                    CadenadeLog "Critical error while updating recordset"
                    End
                
                End If
        End Select
    Loop
    Record_Action = RetVal
End Function
Public Function Edit_Record(Registro As dao.Recordset) As Boolean
    Edit_Record = Record_Action(Registro, 0)
End Function
Public Function Delete_Record(Registro As dao.Recordset) As Boolean
    Delete_Record = Record_Action(Registro, 1)
End Function
Public Sub Corregir_Nulos()
'/////////////////////////////////////
'// esta función es llamada al inicio
'// del programa para eliminar los posibles
'// valores nulos en campos 'dbText' de las tablas, ya que
'// algunos usuarios pican artículos
'// directamente en la base, dejando
'// los campos de texto como nulos.
'// Se podría implementar también este
'// proceso para valores numéricos, pero
'// hay que tener mucho cuidado con los
'// valores por defecto que se asignan.
'//////////////////////////////////////
Dim Base As dao.Database
'Dim rst As DAO.Recordset
Dim rst As dao.Recordset
Dim cntTab As Long
Dim cntReg As Long
Dim cntCam As Long
Dim StrSQL As String
    Do_Events
    On Error GoTo Ret
    Set Base = OpenDatabase(Base_General) 'abrirbase 'OpenDatabase(Base_General)
    For cntTab = 0 To Base.TableDefs.Count - 1
        If left(UCase(Base.TableDefs(cntTab).name), 4) <> "MSYS" Then
            StrSQL = "SELECT * FROM " & Base.TableDefs(cntTab).name
            '2.0.65
            'If UCase(Mid(Base.TableDefs(cntTab).Name, 1, 8)) <> "CABECERA" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 7)) <> "TICKETS" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 3)) <> "LOG" And _
            '   UCase(Mid(Base.TableDefs(cntTab).Name, 1, 6)) <> "TECLAS" Then
            '   'UCase(Mid(Base.TableDefs(cntTab).Name, 1, 8)) <> "ARTICULO" And

            If UCase(Mid(Base.TableDefs(cntTab).name, 1, 3)) <> "LOG" And _
               UCase(Mid(Base.TableDefs(cntTab).name, 1, 6)) <> "TECLAS" Then
            
If UCase(Mid(Base.TableDefs(cntTab).name, 1, 3)) = "tickets" Then
    cntTab = cntTab
End If

            Set rst = Base.OpenRecordset(StrSQL)
            If Not rst.BOF Then
                rst.MoveFirst
                Do Until rst.EOF
                    For cntCam = 0 To rst.Fields.Count - 1
                    
                    'If UCase(Mid(Base.TableDefs(cntTab).Name, 1, 8)) = "ARTICULO" And rst.Fields(cntCam).Name = "posicion" Then
                    '    cntCam = cntCam
                    'End If
                        If rst.Fields(cntCam).type = dbText Then
                            If IsNull(rst.Fields(cntCam)) Then
                                'Edit_Record rst
                                rst.Edit
                                rst.Fields(cntCam) = " "
                                rst.Update
                            End If
                        Else
                            If rst.Fields(cntCam).type = dbNumeric Or rst.Fields(cntCam).type = 4 Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    'Edit_Record rst
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).type = dbBoolean Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    'Edit_Record rst
                                    rst.Edit
                                    rst.Fields(cntCam) = False
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).type = dbDouble Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    'Edit_Record rst
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            If rst.Fields(cntCam).type = dbBigInt Then
                                If IsNull(rst.Fields(cntCam)) Then
                                    'Edit_Record rst
                                    rst.Edit
                                    rst.Fields(cntCam) = 0
                                    rst.Update
                                End If
                            End If
                            
                        
                        End If
                    Next cntCam
                    'Do_Events
                    rst.Movenext
                Loop
            End If
            rst.Close
            Set rst = Nothing
            End If
        End If
    Next cntTab
    
    Base.Execute "update cabecera set imp0=0 where base0=0"
    Base.Execute "update cabecera set imp1=0 where base1=0"
    Base.Execute "update cabecera set imp2=0 where base2=0"
    Base.Execute "update cabecera set imp3=0 where base3=0"
    Base.Execute "update cabecera set imp4=0 where base4=0"
    
    
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    Corregir_Articulos_Nulos
    On Error GoTo 0
Exit Sub
Ret:
    On Error Resume Next
    '1.7.2 If TypeName(Base) <> "Nothing" Then Base.Close
    'rst.Close
    'Set rst = Nothing
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
    On Error GoTo 0
End Sub

'adaptar TouchScale Barna
Private Sub Corregir_Articulos_Nulos()
    Dim Base As dao.Database
    Dim bucle As Integer
    Dim Registro As dao.Recordset
    Dim Campos(27) As String
    Do_Events
    Campos(0) = "codi_ident"
    Campos(1) = "plu"
    Campos(2) = "codigo"
    Campos(3) = "secc_maqui"
    Campos(4) = "codi_sub"
    Campos(5) = "codi_fam"
    Campos(6) = "tipo_iva"
    Campos(7) = "caducidad"
    Campos(8) = "grupo_conserv"
    Campos(9) = "tara"
    Campos(10) = "pref"
    Campos(11) = "fcb"
    Campos(12) = "ning"
    Campos(13) = "precio"
    Campos(14) = "euros"
    Campos(15) = "usatramos"
    Campos(16) = "borrado"
    Campos(17) = "balenv"
    Campos(18) = "prc100g"
    Campos(19) = "etq"
    Campos(20) = "tip_let0"
    Campos(21) = "poid"
    Campos(22) = "tar1"
    Campos(23) = "tar2"
    Campos(24) = "tar3"
    Campos(25) = "tar4"
    Campos(26) = "tar5"
    Campos(26) = "tar6"
    Campos(26) = "tar7"
    Campos(26) = "tar8"
    Campos(26) = "tar9"
    Campos(26) = "tar10"
    
    Set Base = OpenDatabase(Base_General)
    For bucle = 0 To 26
        Set Registro = Base.OpenRecordset("select * from articulo where isnull(" & Campos(bucle) & ")")
        With Registro
            If Not .EOF Then
                .MoveFirst
                If bucle < 4 Then
                    Do Until .EOF
                        If Not .EOF Then .Delete
                        If Not .EOF Then .Movenext
                    Loop
                Else
                    Do Until .EOF
                        If Not .EOF Then .Edit
                        Select Case bucle
                            Case 13 To 14
                                .Fields(Campos(bucle)) = 1
                            Case 15 To 18
                                .Fields(Campos(bucle)) = False
                            Case 19
                                .Fields(Campos(bucle)) = 0 ' ETQ
                            Case 21 To 26
                                .Fields(Campos(bucle)) = 0
                            Case Else
                                .Fields(Campos(bucle)) = 0
                        End Select
                        .Update
                        If Not .EOF Then .Movenext
                    Loop
                End If
            End If
        End With
    Next bucle
    
    Registro.Close
    Set Registro = Nothing
    Base.Close
    Set Base = Nothing
    
End Sub
Public Sub Crear_Base_Maestra()
'///////////////////////////////////////////
' Función llamada cuando no existe la
' base de datos, para crearla.
'///////////////////////////////////////////
    Dim Base As dao.Database
    Set Base = CreateDatabase(Base_General, dbLangGeneral, dbVersion40)
    Base.Close
    crea_bases
    crea_cabtiquets
    'crea_cabtiquetsTQ
    tot_table
    Crea_IVA
    crea_cliente
    'Crea_TextLibre
    Crea_Vacuno
    'Crea_ECO
    Crea_101
    'Crea_Familias
    Crea_1_1_0
    Crea_1_1_7
    Crea_1_1_8
    Crea_1_4_0
    Crea_1_4_1
    Crea_1_7_0
    Crea_1_7_1
    Crea_1_7_1_1
    crea_Gen_sam
    Crea_1_1_7_T
    'Crea_Tara_Leclerc
    'Crea_Ingre_Leclerc
    FrmCambiaFormato.CambiarFormatoBase

End Sub
Public Sub crea_Gen_sam()
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    
    Set Base = OpenDatabase(Base_General)
    Do_Events
    
    Set Tabla = Base.CreateTableDef("gen_sam")
    With Tabla
        .Fields.Append .CreateField("tipo", dbText, 10)
        .Fields.Append .CreateField("destino", dbInteger)
        .Fields.Append .CreateField("numero", dbInteger)
        .Fields.Append .CreateField("redlocal", dbBoolean)
        .Fields.Append .CreateField("baja", dbBoolean)
        .Fields.Append .CreateField("txt01", dbText, 40)
        .Fields.Append .CreateField("txt02", dbText, 40)
        .Fields.Append .CreateField("txt03", dbText, 40)
        .Fields.Append .CreateField("txt04", dbText, 40)
        .Fields.Append .CreateField("txt05", dbText, 40)
        .Fields.Append .CreateField("txt06", dbText, 40)
        .Fields.Append .CreateField("txt07", dbText, 40)
        .Fields.Append .CreateField("txt08", dbText, 40)
        .Fields.Append .CreateField("txt09", dbText, 40)
        .Fields.Append .CreateField("txt10", dbText, 40)
        .Fields.Append .CreateField("txt11", dbText, 40)
        .Fields.Append .CreateField("txt12", dbText, 40)
        .Fields.Append .CreateField("txt13", dbText, 40)
        .Fields.Append .CreateField("txt14", dbText, 40)
        .Fields.Append .CreateField("txt15", dbText, 40)
        .Fields.Append .CreateField("txt16", dbText, 40)
        .Fields.Append .CreateField("txt17", dbText, 40)
        .Fields.Append .CreateField("txt18", dbText, 40)
        .Fields.Append .CreateField("txt19", dbText, 40)
        .Fields.Append .CreateField("txt20", dbText, 40)
        .Fields.Append .CreateField("txt21", dbText, 40)
        .Fields.Append .CreateField("txt22", dbText, 40)
        .Fields.Append .CreateField("txt23", dbText, 40)
        .Fields.Append .CreateField("txt24", dbText, 40)
        .Fields.Append .CreateField("txt25", dbText, 40)
        .Fields.Append .CreateField("txt26", dbText, 40)
        .Fields.Append .CreateField("txt27", dbText, 40)
        .Fields.Append .CreateField("txt28", dbText, 40)
        .Fields.Append .CreateField("txt29", dbText, 40)
        .Fields.Append .CreateField("txt30", dbText, 40)
        .Fields.Append .CreateField("txt31", dbText, 40)
        .Fields.Append .CreateField("txt32", dbText, 40)
        .Fields.Append .CreateField("txt33", dbText, 40)
        .Fields.Append .CreateField("txt34", dbText, 40)
        .Fields.Append .CreateField("txt35", dbText, 40)
        .Fields.Append .CreateField("txt36", dbText, 40)
        .Fields.Append .CreateField("txt37", dbText, 40)
        .Fields.Append .CreateField("txt38", dbText, 40)
        .Fields.Append .CreateField("txt39", dbText, 40)
        .Fields.Append .CreateField("txt40", dbText, 40)
        .Fields.Append .CreateField("tran_txt", dbText, 1)
        For bucle = 0 To .Fields.Count - 1
            .Fields(bucle).AllowZeroLength = True
        Next bucle
        Base.TableDefs.Append Tabla
    End With
    '
    Base.Close

End Sub

Public Sub Crea_101()
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim scad As String
    '
    Do_Events
    Set Base = OpenDatabase(Base_General)
    '
    'For Bucle = 1 To 3
        'Select Case Bucle
        '    Case 1
                scad = ""
        '    Case 2
        '        scad = "sc10"
        '    Case 3
        '        scad = "TQ"
        'End Select
        On Error Resume Next
        Set Tabla = Base.TableDefs("cabecera" & scad)
        Tabla.Fields.Append Tabla.CreateField("d_fecha", dbDate)
        Tabla.Fields.Append Tabla.CreateField("d_hora", dbDate)
        Set Tabla = Base.TableDefs("tickets" & scad)
        Tabla.Fields.Append Tabla.CreateField("d_fecha", dbDate)
        Tabla.Fields.Append Tabla.CreateField("d_hora", dbDate)
    'Next Bucle
    '
    Set Tabla = Base.TableDefs("gtarti")
    Tabla.Fields.Append Tabla.CreateField("txt_fam", dbText, 30)
    Tabla.Fields("txt_fam").AllowZeroLength = True
    Set Tabla = Base.TableDefs("gtsecs")
    Tabla.Fields.Append Tabla.CreateField("txt_fam", dbText, 30)
    Tabla.Fields("txt_fam").AllowZeroLength = True
    '
    Set Tabla = Base.CreateTableDef("fam_code")
    With Tabla
        .Fields.Append .CreateField("codi_fam", dbLong)
        .Fields.Append .CreateField("secc_maqui", dbLong)
        .Fields.Append .CreateField("codi_ident", dbLong)
        .Fields.Append .CreateField("posicion", dbLong)
        .Fields.Append .CreateField("txt_fam", dbText, 35)
        .Fields.Append .CreateField("imagen", dbText, 250)
        .Fields.Append .CreateField("etiqueta", dbText, 250)
        .Fields.Append .CreateField("Index", dbInteger)
        For bucle = 0 To .Fields.Count - 1
            .Fields(bucle).AllowZeroLength = True
        Next bucle
        Base.TableDefs.Append Tabla
    End With
    '
    Base.Close
End Sub

Public Sub Crea_Familias()
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    Dim Resp As Long
    '
    Do_Events
    Set Base = OpenDatabase(Base_General)
    '
' With mifam'
'
'      .Fields.Append .CreateField("CODI_IDENT", dbInteger)
'      .Fields.Append .CreateField("SECC_MAQUI", dbInteger)
'      .Fields.Append .CreateField("codi_fam", dbInteger)
'      .Fields.Append .CreateField("TRAN_FAM", dbText, 1)
'      .Fields.Append .CreateField("INDEX", dbInteger)
'      .Fields.Append .CreateField("BORRADO", dbBoolean)
'
'      For i = 0 To .Fields.Count - 1
'      .Fields(i).AllowZeroLength = True
'      Next i
'
'      Base.TableDefs.Append mifam
' End With
    
    Set Tabla = Base.CreateTableDef("familias")
    With Tabla
        .Fields.Append .CreateField("CODI_IDENT", dbInteger)
        .Fields.Append .CreateField("SECC_MAQUI", dbInteger)
        .Fields.Append .CreateField("CODI_FAM", dbInteger)
        .Fields.Append .CreateField("TRAN_FAM", dbText, 1)
        .Fields.Append .CreateField("INDEX", dbInteger)
        .Fields.Append .CreateField("BORRADO", dbBoolean)
    
        For bucle = 0 To .Fields.Count - 1
            .Fields(bucle).AllowZeroLength = True
        Next bucle
        Base.TableDefs.Append Tabla
    End With
    '
    Base.Close
End Sub

Public Sub Crea_1_6_5() 'c2f invicta
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim scad As String
    '
    Do_Events
    Set Base = OpenDatabase(Base_General)
    '
    'For Bucle = 1 To 3
        'Select Case Bucle
        '    Case 1
                scad = ""
        '    Case 2
        '        scad = "sc10"
        '    Case 3
        '        scad = "TQ"
        'End Select
        On Error Resume Next
        Set Tabla = Base.TableDefs("tickets" & scad)
        Tabla.Fields.Append Tabla.CreateField("DPT", dbText, 4)
        Tabla.Fields.Append Tabla.CreateField("importe_bruto", dbDouble)
    
        Set Tabla = Nothing
        scad = "SELECT importe_bruto,DPT,amount " & _
               "FROM " & "tickets" & scad & " " & _
               "ORDER BY nume"
        Set Registro = Base.OpenRecordset(scad)
        With Registro
            If Not .EOF Then
                .MoveFirst
                For Resp = 0 To .Recordcount - 1
                    .Edit
                    !importe_bruto = !amount
                    !DPT = 0
                    .Update
                    .Movenext
                Next Resp
            End If
        End With
        Registro.Close
        Set Registro = Nothing
    
    'Next Bucle
    '
    '
    Base.Close
End Sub
Public Sub adapta_Cobro() 'c2f se añade campo cobrado...
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim scad As String
    '
    Do_Events
    Set Base = OpenDatabase(Base_General)
    '
    'For Bucle = 1 To 3
        'Select Case Bucle
        '    Case 1
                scad = ""
        '    Case 2
        '        scad = "sc10"
        '    Case 3
        '        scad = "TQ"
        'End Select
        On Error Resume Next
        Set Tabla = Base.TableDefs("Cabecera" & scad)
        Tabla.Fields.Append Tabla.CreateField("Cobrado", dbBoolean, 1)
                            
        Set Tabla = Nothing
        'sCad = "SELECT Cobrado " & _
        '       "FROM " & "tickets" & sCad & " " & _
        '       "ORDER BY nume"
        scad = "SELECT Cobrado " & _
               "FROM " & "cabecera" & scad & " " & _
               "ORDER BY nume"
        
        Set Registro = Base.OpenRecordset(scad)
        With Registro
            If Not .EOF Then
                .MoveFirst
                For Resp = 0 To .Recordcount - 1
                    .Edit
                    !cobrado = False
                    .Update
                    .Movenext
                Next Resp
            End If
        End With
        Registro.Close
        Set Registro = Nothing
    
    'Next Bucle
    '
    '
    Base.Close
End Sub
Public Sub Crea_1_1_0()
'**************************************
' peticiones de INVICTA retail systems
' a partir de la versión 1.1.0 --> Código de subsección
'**************************************
    Dim Base As dao.Database
    Dim bucle As Long
    Dim Tabla As TableDef
    Dim Registro As dao.Recordset
    Do_Events
    Set Base = OpenDatabase(Base_General)
    Set Tabla = Base.TableDefs("gtarti")
    Tabla.Fields.Append Tabla.CreateField("txt_sub", dbText, 30)
    Tabla.Fields("txt_sub").AllowZeroLength = True
    
    Set Tabla = Base.TableDefs("gtsecs")
    Tabla.Fields.Append Tabla.CreateField("txt_sub", dbText, 30)
    Tabla.Fields("txt_sub").AllowZeroLength = True
    Set Tabla = Base.CreateTableDef("sub_code")
    With Tabla
        .Fields.Append .CreateField("codi_sub", dbLong)
        .Fields.Append .CreateField("codi_ident", dbLong)
        .Fields.Append .CreateField("txt_sub", dbText, 35)
        For bucle = 0 To .Fields.Count - 1
            .Fields(bucle).AllowZeroLength = True
        Next bucle
        Base.TableDefs.Append Tabla
    End With
    '
    Base.Close
End Sub
'******
'******
'******
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'cas.n->peticiones de INVICTA retail systems a partir de la versión 1.1.7 --> se introducen
'cambios en la base para almacenar las funciones de cambio
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_1_7()
Dim db As dao.Database
Dim tdf As dao.TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = dao.OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
            sTabla = "cabecera"
        'endif
        'If cnt0 = 1 Then sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("tipoPago", dbLong)
            .Fields.Append .CreateField("cantidadEntre", dbDouble, 15)
            .Fields.Append .CreateField("cambio", dbDouble, 15)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,tipoPago,cantidadEntre,cambio " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("tipoPago") = CLng(0)
                    .Fields("cantidadEntre") = CDbl(0)
                    .Fields("cambio") = CDbl(0)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_1_7_T()
Dim db As dao.Database
Dim tdf As dao.TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = dao.OpenDatabase(Base_General)
        sTabla = "cabecera"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("totalweight", dbDouble)
            .Fields.Append .CreateField("totalunits", dbLong)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,totalweight,totalunits " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("totalunits") = CLng(0)
                    .Fields("totalweight") = CDbl(0)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_1_7_T_1()
Dim db As dao.Database
Dim tdf As dao.TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = dao.OpenDatabase(Base_General)
        sTabla = "tickets"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("ivapercent", dbDouble)
        End With
        Set tdf = Nothing
        sSQL = "SELECT ivapercent " & _
               "FROM " & sTabla
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("ivapercent") = CDbl(0)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    db.Close
    Set db = Nothing
End Function

'se añade numero de lote a la cabecera
Public Function Crea_1_6_6()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
            sTabla = "cabecera"
        'End If
        'If cnt0 = 1 Then sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("numlote", dbText)
            .Fields.Append .CreateField("abonado", dbBoolean)
            .Fields.Append .CreateField("ticketabono", dbLong)
            .Fields.Append .CreateField("ticketorigen", dbLong)
            .Fields.Append .CreateField("abono", dbBoolean)
            .Fields.Append .CreateField("descuento", dbDouble)
        End With
        
        Set tdf = Nothing
        sSQL = "SELECT nume,numlote " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("numlote") = " "
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function
'se añade numero de lote a las tablas de lineas
Public Function Crea_1_6_6_lin()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
            sTabla = "tickets"
        'End If
        'If cnt0 = 1 Then sTabla = "ticketsSC10"
        'If cnt0 = 2 Then sTabla = "ticketsTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("numlote", dbText)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,numlote " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("numlote") = " "
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_6_6_2()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
    sTabla = "cabecera"
    Set tdf = db.TableDefs(sTabla)
    With tdf
        .Fields.Append .CreateField("factura", dbText, 25)
    End With
    Set tdf = Nothing
    sSQL = "SELECT factura " & _
           "FROM " & sTabla
    Set rst = db.OpenRecordset(sSQL)
    With rst
        If Not .EOF Then
            For cnt1 = 0 To .Recordcount - 1
                .Edit
                .Fields("factura") = Space(15)
                .Update
                .Movenext
            Next cnt1
        End If
    End With
    Set rst = Nothing
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_6_6_3()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
    sTabla = "tickets"
    Set tdf = db.TableDefs(sTabla)
    With tdf
        .Fields.Append .CreateField("factura", dbText, 25)
    End With
    Set tdf = Nothing
    sSQL = "SELECT factura " & _
           "FROM " & sTabla
    Set rst = db.OpenRecordset(sSQL)
    With rst
        If Not .EOF Then
            For cnt1 = 0 To .Recordcount - 1
                .Edit
                .Fields("factura") = Space(15)
                .Update
                .Movenext
            Next cnt1
        End If
    End With
    Set rst = Nothing
    db.Close
    Set db = Nothing
End Function

'se añaden impuestos a la cabecera
Public Function Crea_1_6_6_1()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = OpenDatabase(Base_General)
    sTabla = "cabecera"
    Set tdf = db.TableDefs(sTabla)
    With tdf
        .Fields.Append .CreateField("base0", dbDouble)
        .Fields.Append .CreateField("por0", dbDouble)
        .Fields.Append .CreateField("imp0", dbDouble)
        .Fields.Append .CreateField("base1", dbDouble)
        .Fields.Append .CreateField("por1", dbDouble)
        .Fields.Append .CreateField("imp1", dbDouble)
        .Fields.Append .CreateField("base2", dbDouble)
        .Fields.Append .CreateField("por2", dbDouble)
        .Fields.Append .CreateField("imp2", dbDouble)
        .Fields.Append .CreateField("base3", dbDouble)
        .Fields.Append .CreateField("por3", dbDouble)
        .Fields.Append .CreateField("imp3", dbDouble)
        .Fields.Append .CreateField("base4", dbDouble)
        .Fields.Append .CreateField("por4", dbDouble)
        .Fields.Append .CreateField("imp4", dbDouble)
    
    End With
    Set tdf = Nothing
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_6_6_lin_1()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
        sTabla = "tickets"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("porcentaje", dbDouble)
        End With
        Set tdf = Nothing
    db.Close
    Set db = Nothing
End Function

Public Function Crea_1_6_6_Descuento()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
        sTabla = "tickets"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("descuento", dbDouble)
        End With
        Set tdf = Nothing
    db.Close
    Set db = Nothing
End Function

'se añade numero de cliente a las tablas de lineas
Public Function Crea_1_6_6_Cli()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String
    
    Do_Events
    Set db = OpenDatabase(Base_General)
        sTabla = "tickets"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            On Error Resume Next
            .Fields.Delete ("cliente")
            On Error GoTo 0
            .Fields.Append .CreateField("cliente", dbText, 15)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,cliente " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("cliente") = Space(15)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    db.Close
    Set db = Nothing
End Function

Public Function VerTipoCliente()
Dim db As dao.Database
Dim tdf As dao.TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String
    Do_Events
    Set db = OpenDatabase(Base_General)
        sTabla = "tickets"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            On Error Resume Next
            .Fields.Delete ("clienteT")
            On Error GoTo 0
            .Fields.Append .CreateField("clienteT", dbText, 15)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,clienteT " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("clienteT") = Space(15)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    db.Close
    Set db = Nothing

End Function


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'c2f ahorramás parece que no se actualizan bien...
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_1_7_ahSc10()
Dim db As dao.Database
Dim tdf As dao.TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = dao.OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then sTabla = "cabecera"
        'If cnt0 = 1 Then
        sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("tipoPago", dbLong)
            .Fields.Append .CreateField("cantidadEntre", dbDouble, 15)
            .Fields.Append .CreateField("cambio", dbDouble, 15)
            
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,tipoPago,cantidadEntre,cambio " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("tipoPago") = CLng(0)
                    .Fields("cantidadEntre") = CDbl(0)
                    .Fields("cambio") = CDbl(0)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'c2f ahorramás parece que no se actualizan bien...
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_1_7_ahTq()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then sTabla = "cabecera"
        'If cnt0 = 1 Then sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then
        sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("tipoPago", dbLong)
            .Fields.Append .CreateField("cantidadEntre", dbDouble, 15)
            .Fields.Append .CreateField("cambio", dbDouble, 15)
        End With
        Set tdf = Nothing
        sSQL = "SELECT nume,tipoPago,cantidadEntre,cambio " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("tipoPago") = CLng(0)
                    .Fields("cantidadEntre") = CDbl(0)
                    .Fields("cambio") = CDbl(0)
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function
'******
'******
'******
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'cas.n->peticiones de INVICTA retail systems a partir de la versión 1.1.7 --> se introducen
'cambios en la base para almacenar las funciones de cambio
'c2f invicta
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_6_5_C()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
            sTabla = "cabecera"
        'End If
        'If cnt0 = 1 Then sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("importe_bruto", dbDouble)
        End With
        Set tdf = Nothing
        sSQL = "SELECT importe_bruto,IMPORTE " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                .MoveFirst
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    !importe_bruto = !Importe
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'c2f ahorramás parece que no se actualizan bien...
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_6_5_ahSc10()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
        sTabla = "cabecera"
        'endif
        'If cnt0 = 1 Then
        'sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("importe_bruto", dbLong)
        End With
        Set tdf = Nothing
        sSQL = "SELECT importe_bruto " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("importe_bruto") = .Fields("importe")
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'c2f ahorramás parece que no se actualizan bien...
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_6_5_ahTq()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim sSQL As String
Dim cnt0 As Long
Dim cnt1 As Long
Dim sTabla As String

    Do_Events
    Set db = OpenDatabase(Base_General)
    'For cnt0 = 0 To 2
        'If cnt0 = 0 Then
        sTabla = "cabecera"
        'endif
        'If cnt0 = 1 Then sTabla = "cabeceraSC10"
        'If cnt0 = 2 Then
        'sTabla = "cabeceraTQ"
        Set tdf = db.TableDefs(sTabla)
        With tdf
            .Fields.Append .CreateField("importe_bruto", dbLong)
        End With
        Set tdf = Nothing
        sSQL = "SELECT importe_bruto " & _
               "FROM " & sTabla & " " & _
               "ORDER BY nume"
        Set rst = db.OpenRecordset(sSQL)
        With rst
            If Not .EOF Then
                For cnt1 = 0 To .Recordcount - 1
                    .Edit
                    .Fields("importe_bruto") = .Fields("importe")
                    .Update
                    .Movenext
                Next cnt1
            End If
        End With
        Set rst = Nothing
    'Next cnt0
    db.Close
    Set db = Nothing
End Function

'***********
'***********
'***********
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'//cas.n->peticion de AHORRAMAS recopila informacion sobre la forma de operar de los vendedores
'//////////////////////////////////////////////////////////////////////////////////////////////
Public Function Crea_1_1_8() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.CreateTableDef("gtInfVen")
    With tdf
        .Fields.Append .CreateField("ident_vend", dbLong)
        .Fields.Append .CreateField("codi_ident", dbLong)
        .Fields.Append .CreateField("secc_maqui", dbLong)
        .Fields.Append .CreateField("imporCod", dbDouble, 9) 'cas.n->importe codificado (con PLU)
        .Fields.Append .CreateField("operCod", dbDouble, 9) 'cas.n->operaciones codificadas (con PLU)
        .Fields.Append .CreateField("imporDir", dbDouble, 9) 'cas.n->importe directo (con PLU 0)
        .Fields.Append .CreateField("operDir", dbDouble, 9) 'cas.n->operaciones directas (con PLU 0)
        .Fields.Append .CreateField("imporCan", dbDouble, 9) 'cas.n->importe anulado
        .Fields.Append .CreateField("operCan", dbDouble, 9) 'cas.n->operaciones anuladas
        .Fields.Append .CreateField("imporNeg", dbDouble, 9) 'cas.n->importe negativo
        .Fields.Append .CreateField("operNeg", dbDouble, 9) 'cas.n->operaciones negativas
        .Fields.Append .CreateField("modo", dbLong)
        .Fields.Append .CreateField("fecha", dbDate)
        
        
        db.TableDefs.Append tdf
    End With
    db.Close
End Function
Public Function Crea_1_1_8_punto2() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("gtInfVen")
    With tdf
        .Fields.Append .CreateField("imporNeg", dbDouble, 9) 'cas.n->importe negativo
        .Fields.Append .CreateField("operNeg", dbDouble, 9) 'cas.n->operaciones negativas
    End With
    Set tdf = Nothing
    db.Close
End Function
'\\\\\\\\\\\
Public Function Crea_1_1_8_punto3() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("gtInfVen")
    With tdf
        .Fields.Append .CreateField("operuni", dbLong) 'cas.n->operaciones venta por unidades
        .Fields.Append .CreateField("operpeso", dbLong) 'cas.n->operaciones venta por peso
        .Fields.Append .CreateField("imporuni", dbDouble, 9) 'importe venta por unidades
        .Fields.Append .CreateField("imporpeso", dbDouble, 9) 'cas.n->importe venta por peso
        .Fields.Append .CreateField("ntotal", dbLong) 'cas.n->importe directo (con PLU 0)
        .Fields.Append .CreateField("exported", dbBoolean) 'cas.n->importe directo (con PLU 0)
    End With
    Set tdf = Nothing
    db.Close
    
End Function

Public Function Crea_1_1_8_punto4() As Integer
Dim db As dao.Database
Dim MiTotPend As TableDef
Dim i As Integer
    Set db = OpenDatabase(Base_General)

    Set MiTotPend = db.CreateTableDef("gtpend")

    'totales pendientes de exportar --> Belros
    With MiTotPend
        .Fields.Append .CreateField("id", dbLong)
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("ntotal", dbLong)
        .Fields.Append .CreateField("exported", dbBoolean, 1)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        db.TableDefs.Append MiTotPend
    End With

    Do_Events
    Set MiTotPend = Nothing
    db.Close
End Function


'///////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'cas.n->recopila informacion de los descuentos realizados
'////////////////////////////////////////////////////////
Public Function Crea_1_4_0() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long
    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.CreateTableDef("descuentos")
    With tdf
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("hora", dbDate)
        .Fields.Append .CreateField("nume", dbLong)
        .Fields.Append .CreateField("seccion", dbLong)
        .Fields.Append .CreateField("balanza", dbLong)
        .Fields.Append .CreateField("vended", dbLong)
        .Fields.Append .CreateField("typTic", dbByte)
        .Fields.Append .CreateField("importeBruto", dbDouble, 9)
        .Fields.Append .CreateField("importeNeto", dbDouble, 9)
        db.TableDefs.Append tdf
    End With
    db.Close
End Function
Public Function Crea_1_4_1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long
    Do_Events
    Set db = OpenDatabase(Base_General)
    
    On Error Resume Next '2.0.31
    db.TableDefs.Delete "descuentos"
    On Error GoTo 0
    
    Set tdf = db.CreateTableDef("descuentos")
    With tdf
        .Fields.Append .CreateField("fecha", dbDate)
        .Fields.Append .CreateField("hora", dbDate)
        .Fields.Append .CreateField("nume", dbLong)
        .Fields.Append .CreateField("seccion", dbLong)
        .Fields.Append .CreateField("balanza", dbLong)
        .Fields.Append .CreateField("vended", dbLong)
        .Fields.Append .CreateField("typTic", dbByte)
        .Fields.Append .CreateField("importeBruto", dbDouble, 9)
        .Fields.Append .CreateField("importeNeto", dbDouble, 9)
        .Fields.Append .CreateField("descuentototal", dbDouble, 9)
        db.TableDefs.Append tdf
    End With
    db.Close
End Function

'\\\\\\\\\\\
'///////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'cas.n->recopila informacion de los totales horarios
'///////////////////////////////////////////////////
Public Function Crea_1_7_0() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    
    On Error Resume Next '2.0.31
    db.TableDefs.Delete "gthora"
    On Error GoTo 0
    
    Set tdf = db.CreateTableDef("gthora")
    With tdf
        .Fields.Append .CreateField("FECHA", dbDate)
        '.Fields.Append .CreateField("HORA", dbText)
        .Fields.Append .CreateField("HORA", dbDate)
        .Fields.Append .CreateField("SECCION", dbLong)
        .Fields.Append .CreateField("EQUIPO", dbLong)
        .Fields.Append .CreateField("VENDEDOR", dbLong)
        .Fields.Append .CreateField("TIQUET", dbLong)
        .Fields.Append .CreateField("OPERACIONES", dbLong)
        .Fields.Append .CreateField("OPERACIONES_CAN", dbLong)
        .Fields.Append .CreateField("OPERACIONES_NEG", dbLong)
        .Fields.Append .CreateField("IMPORTE", dbDouble, 9)
        .Fields.Append .CreateField("IMPORTE_CAN", dbDouble, 9)
        .Fields.Append .CreateField("IMPORTE_NEG", dbDouble, 9)
        .Fields.Append .CreateField("PESO", dbDouble)
        .Fields.Append .CreateField("PESO_CAN", dbDouble)
        .Fields.Append .CreateField("PESO_NEG", dbDouble)
        .Fields.Append .CreateField("UNIDADES", dbLong)
        .Fields.Append .CreateField("UNIDADES_CAN", dbLong)
        .Fields.Append .CreateField("UNIDADES_NEG", dbLong)
        db.TableDefs.Append tdf
    End With
    db.Close
End Function
'\\\\\\\\\\\
'///////////
Public Function Crea_1_7_2() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
        .Fields.Append .CreateField("lMix", dbBoolean)
        .Fields.Append .CreateField("Poid", dbLong)
    End With
    Set tdf = Nothing
    Set rst = db.OpenRecordset("select lmix,poid from articulo")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            rst.Edit
            rst.Fields("lmix") = False
            rst.Fields("poid") = 0
            rst.Update
            rst.Movenext
        Loop
    
    End If
    rst.Close
    Set rst = Nothing
    db.Close

End Function
'CreaOnKey
Public Function creaOnKey() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
        .Fields.Append .CreateField("onkey", dbBoolean)
    End With
    Set tdf = Nothing
    Set rst = db.OpenRecordset("select onkey from articulo")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            rst.Edit
            rst.Fields("onkey") = True
            rst.Update
            rst.Movenext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    db.Close

End Function

Public Function Crea_1_9_0() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("equipos")
    With tdf
        .Fields.Append .CreateField("ss165", dbBoolean)
    End With
    Set tdf = Nothing
    Set rst = db.OpenRecordset("select ss165 from equipos")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            rst.Edit
            rst.Fields("ss165") = False
            rst.Update
            rst.Movenext
        Loop
    
    End If
    rst.Close
    Set rst = Nothing
    db.Close

End Function
      
'      .Fields.Append .CreateField("ART_CB14", dbText, 14)
'      .Fields.Append .CreateField("label2", dbText, 60)
'Crea_art_sam1
Public Function Crea_art_sam1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
        .Fields.Append .CreateField("art_cb14", dbText, 14)
        .Fields.Append .CreateField("label2", dbText, 60)
        .Fields("art_cb14").AllowZeroLength = True
        .Fields("label2").AllowZeroLength = True
        
    End With
    Set tdf = Nothing
    db.Close

End Function

Public Function Crea_art_label1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
        .Fields.Append .CreateField("label1", dbText, 60)
        .Fields("label1").AllowZeroLength = True
    
    End With
    Set tdf = Nothing
    db.Close

End Function

Public Function Crea_fam_label1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("fam_code")
    With tdf
        .Fields.Append .CreateField("label1", dbText, 60)
        .Fields.Append .CreateField("label2", dbText, 60)
        .Fields("label1").AllowZeroLength = True
        .Fields("label2").AllowZeroLength = True
    
    End With
    Set tdf = Nothing
    db.Close

End Function

Public Function Crea_art_Tar() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
    
      .Fields.Append .CreateField("tar1", dbDouble, 7)
      .Fields.Append .CreateField("tar2", dbDouble, 7)
      .Fields.Append .CreateField("tar3", dbDouble, 7)
      .Fields.Append .CreateField("tar4", dbDouble, 7)
      .Fields.Append .CreateField("tar5", dbDouble, 7)
            
    End With
    Set tdf = Nothing
    db.Close

End Function

Public Function Crea_art_Tar10() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("articulo")
    With tdf
    
      .Fields.Append .CreateField("tar6", dbDouble, 7)
      .Fields.Append .CreateField("tar7", dbDouble, 7)
      .Fields.Append .CreateField("tar8", dbDouble, 7)
      .Fields.Append .CreateField("tar9", dbDouble, 7)
      .Fields.Append .CreateField("tar10", dbDouble, 7)
            
    End With
    Set tdf = Nothing
    db.Close

End Function

'cliente
'        .Fields.Append .CreateField("factura", dbBoolean)
'        .Fields.Append .CreateField("discount", dbDouble)
Public Function Crea_Cliente_72() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("cliente")
    With tdf
    
      .Fields.Append .CreateField("factura", dbBoolean)
      .Fields.Append .CreateField("discount", dbDouble)
            
    End With
    Set tdf = Nothing
    db.Close

End Function

Public Function Crea_Traza_fam_Sam() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("fam_code")
    With tdf
        .Fields.Append .CreateField("istraza", dbBoolean)
    End With
    Set tdf = Nothing
    
    db.Close

End Function
    
Public Function Crea_Tax_fam() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("fam_code")
    With tdf
        .Fields.Append .CreateField("tax", dbInteger)
    End With
    Set tdf = Nothing
    Set rst = db.OpenRecordset("select tax from fam_code")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            rst.Edit
            rst.Fields("tax") = 0
            rst.Update
            rst.Movenext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    db.Close

End Function
Public Function Crea_2_0_0() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("fam_code")
    With tdf
        'On Error Resume Next
        .Fields.Append .CreateField("Index", dbInteger)
        .Fields.Append .CreateField("etiqueta", dbText, 250)
        For cnt = 0 To .Fields.Count - 1
            If (.Fields(cnt).type = dbText) Then .Fields(cnt).AllowZeroLength = True
        Next cnt
        
        'On Error GoTo 0
        
    End With
    Set tdf = Nothing
    Set rst = db.OpenRecordset("select * from fam_code")
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            rst.Edit
            rst.Fields("index") = 0
            'If Not IsNull(rst.Fields("imagen")) And rst.Fields("imagen") <> "" Then
                rst.Fields("etiqueta") = rst.Fields("imagen")
            'Else
            '    rst.Fields("etiqueta") = "-"
            'End If
            rst.Fields("imagen") = ""

            rst.Update
            
            rst.Movenext
        Loop
    
    End If
    rst.Close
    Set rst = Nothing
    db.Close

End Function


Public Function RevisaText15()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.TableDefs("Text15")
    With tdf
        For cnt = 0 To .Fields.Count - 1
            If (.Fields(cnt).type = dbText) Or (.Fields(cnt).type = dbMemo) Then .Fields(cnt).AllowZeroLength = True
        Next cnt
        db.TableDefs.Append tdf
    End With
    db.Close

End Function

Public Function CreaText15()
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.CreateTableDef("Text15")
    With tdf
        .Fields.Append .CreateField("codigo", dbLong)
        .Fields.Append .CreateField("plu", dbLong)
        .Fields.Append .CreateField("mostrador", dbLong)

#If BALSAM = True Then
        .Fields.Append .CreateField("text1", dbMemo, 1000)
#Else
        .Fields.Append .CreateField("text1", dbMemo, 1500)
        .Fields.Append .CreateField("text2", dbMemo, 1500)
        .Fields.Append .CreateField("text3", dbMemo, 1500)
        .Fields.Append .CreateField("text4", dbMemo, 1500)
        .Fields.Append .CreateField("text5", dbMemo, 1500)
#End If

        .Fields.Append .CreateField("tran_txt15", dbText, 1)
        For cnt = 0 To .Fields.Count - 1
            If (.Fields(cnt).type = dbText) Or (.Fields(cnt).type = dbMemo) Then .Fields(cnt).AllowZeroLength = True
        Next cnt
        db.TableDefs.Append tdf
    End With
    db.Close
End Function

Public Function Crea_1_7_1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim rst As dao.Recordset
Dim cnt As Long

    Do_Events
    Set db = OpenDatabase(Base_General)
    Set tdf = db.CreateTableDef("linTxt2040")
    With tdf
        .Fields.Append .CreateField("codigo", dbLong)
        .Fields.Append .CreateField("plu", dbLong)
        .Fields.Append .CreateField("mostrador", dbLong)
        .Fields.Append .CreateField("txt_21", dbText, 80)
        .Fields.Append .CreateField("txt_22", dbText, 80)
        .Fields.Append .CreateField("txt_23", dbText, 80)
        .Fields.Append .CreateField("txt_24", dbText, 80)
        .Fields.Append .CreateField("txt_25", dbText, 80)
        .Fields.Append .CreateField("txt_26", dbText, 80)
        .Fields.Append .CreateField("txt_27", dbText, 80)
        .Fields.Append .CreateField("txt_28", dbText, 80)
        .Fields.Append .CreateField("txt_29", dbText, 80)
        .Fields.Append .CreateField("txt_30", dbText, 80)
        .Fields.Append .CreateField("txt_31", dbText, 80)
        .Fields.Append .CreateField("txt_32", dbText, 80)
        .Fields.Append .CreateField("txt_33", dbText, 80)
        .Fields.Append .CreateField("txt_34", dbText, 80)
        .Fields.Append .CreateField("txt_35", dbText, 80)
        .Fields.Append .CreateField("txt_36", dbText, 80)
        .Fields.Append .CreateField("txt_37", dbText, 80)
        .Fields.Append .CreateField("txt_38", dbText, 80)
        .Fields.Append .CreateField("txt_39", dbText, 80)
        .Fields.Append .CreateField("txt_40", dbText, 80)
        .Fields.Append .CreateField("tran_txt2", dbText, 1)
        For cnt = 0 To .Fields.Count - 1
            If (.Fields(cnt).type = dbText) Then .Fields(cnt).AllowZeroLength = True
        Next cnt
        db.TableDefs.Append tdf
    End With
    db.Close
End Function
''''''''''''
''''''''''''

''''''''''''
''''''''''''
Sub crea_bases()
'//////////////////////////////////////////
' Aquí se crean las tablas principales de
' las bases de datos, pero no todas, ya que
' algunas se han ido añadiendo más tarde,
' y es más cómodo tenerlas separadas para
' llamarlas por separado cuando el programa
' se encuentra con bases obsoletas (ver
' FrmCambiaFormato) en instalaciones para
' actualización. Estas tablas provienen del programa MNG
'//////////////////////////////////////////
Dim Base As dao.Database
Dim mitabledef As TableDef
Dim MiSeccion As TableDef
Dim Trade As TableDef
Dim misubsec As TableDef
Dim mifam As TableDef
Dim micodbar As TableDef
Dim micabley As TableDef
Dim MiVendedor As TableDef
Dim miequip As TableDef
Dim MiTecla As TableDef
Dim mipubli As TableDef
Dim miGTVacuno As TableDef
Dim miTramo As TableDef
'**********
' MNG 2.0.7
Dim MiGrpConservacion As TableDef
Dim MiPais As TableDef
Dim MiCategoria As TableDef
Dim MiRaza As TableDef
Dim MiSexo As TableDef
Dim MiTipoEtq As TableDef
Dim MiTotPend As TableDef

'Dim MiTecla As TableDef
'*********************
Dim i As Integer
Dim bucle As Integer
Dim prpBucle As Property
'Set base = OpenDatabase(miruta, False, 0, "DBASE IV")
Set Base = OpenDatabase(Base_General)
Set mitabledef = Base.CreateTableDef("Articulo")
Set MiSeccion = Base.CreateTableDef("Seccion")
Set Trade = Base.CreateTableDef("Trade")
'Set misubsec = Base.CreateTableDef("subsec")
Set mifam = Base.CreateTableDef("familias")
Set miequip = Base.CreateTableDef("equipos")
Set micodbar = Base.CreateTableDef("codbar")
'Set micabley = Base.CreateTableDef("cabley")
Set MiTecla = Base.CreateTableDef("teclas")
Set MiVendedor = Base.CreateTableDef("vendedor")
'Set mipubli = Base.CreateTableDef("publi")
'Set MiGrpConservacion = Base.CreateTableDef("GrpConserv")
Set miGTVacuno = Base.CreateTableDef("GTVacuno")
Set MiPais = Base.CreateTableDef("Paises")
Set miTramo = Base.CreateTableDef("Tramos")
Set MiTipoEtq = Base.CreateTableDef("tipoetiqueta")
Set MiTotPend = Base.CreateTableDef("gtpend")

'totales pendientes de exportar --> Belros
With MiTotPend
    .Fields.Append .CreateField("fecha", dbDate)
    .Fields.Append .CreateField("ntotal", dbLong)
    .Fields.Append .CreateField("exported", dbBoolean, 1)
    For i = 0 To .Fields.Count - 1
        .Fields(i).AllowZeroLength = True
    Next i
    Base.TableDefs.Append MiTotPend
End With

'//////////////////////////////////////////////////
' Tabla de precios por tramo : Euroscale admite
' precios escalonados por peso o número de unidades
'//////////////////////////////////////////////////
With miTramo
    .Fields.Append .CreateField("Codigo", dbDouble, 6)
    .Fields.Append .CreateField("Tramo1", dbDouble, 9)
    .Fields.Append .CreateField("Precio1", dbDouble, 9)
    .Fields.Append .CreateField("Euros1", dbDouble, 9)
    .Fields.Append .CreateField("Tramo2", dbDouble, 9)
    .Fields.Append .CreateField("Precio2", dbDouble, 9)
    .Fields.Append .CreateField("Euros2", dbDouble, 9)
    .Fields.Append .CreateField("tran_tramo", dbText, 1)
    .Fields.Append .CreateField("Borrado", dbBoolean, 1)
    For i = 0 To .Fields.Count - 1
        .Fields(i).AllowZeroLength = True
    Next i
    Base.TableDefs.Append miTramo
End With
'///////////////////////////////////////////
' Tabla de totales para vacuno : se creó muy
' deprisa para Ahorramas, cuando en la propia
' balanza Euroscale no estaba demasiado
' claro el modo de implementarlo. Cuando sea
' posible sería mejor cambiar el formato
' a uno similar al de GTARTI, por ejemplo
'////////////////////////////////////////
' buen comentario que nunca será atendido y pasará al olvido
'///////////////////////////////////////////////////////////
    With miGTVacuno
        On Error GoTo 0
        .Fields.Append .CreateField("Ficha", dbDouble, 2)
        .Fields.Append .CreateField("LNumOperaciones", dbDouble, 2)
        .Fields.Append .CreateField("LNumUnidades", dbDouble, 2)
        .Fields.Append .CreateField("LImporte", dbDouble, 2)
        .Fields.Append .CreateField("LPeso", dbDouble, 2)
        .Fields.Append .CreateField("VNumOperaciones", dbDouble, 2)
        .Fields.Append .CreateField("VNumUnidades", dbDouble, 2)
        .Fields.Append .CreateField("VImporte", dbDouble, 2)
        .Fields.Append .CreateField("VPeso", dbDouble, 2)
        .Fields.Append .CreateField("Bloqueo", dbBoolean, 2)
        .Fields.Append .CreateField("Total", dbBoolean, 2)
        .Fields.Append .CreateField("Fecha", dbText, 10)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append miGTVacuno
    End With
    '/////////////////////////////
    ' Tabla de países para fichas de Vacuno
    With MiPais
        On Error GoTo 0
        .Fields.Append .CreateField("CODIGO", dbDouble, 3)
        .Fields.Append .CreateField("valor", dbText, 100)
        .Fields.Append .CreateField("NOMBRE", dbText, 100)
        .Fields.Append .CreateField("TRAN_PAIS", dbText, 1)
        .Fields.Append .CreateField("BORRADO", dbBoolean, 1)
        .Fields.Append .CreateField("MODIFICADO", dbBoolean, 1)
        .Fields.Append .CreateField("Abrv_3", dbText, 3)
        .Fields.Append .CreateField("Abrv_2", dbText, 2)
        .Fields.Append .CreateField("tipo", dbInteger)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append MiPais
    End With
    With MiTipoEtq
        On Error GoTo 0
        .Fields.Append .CreateField("CODIGO", dbDouble, 3)
        .Fields.Append .CreateField("tipo", dbInteger)
        .Fields.Append .CreateField("descripcion", dbText, 25)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append MiTipoEtq
    End With
    
    '//////////////////////////////////////
    ' Grupos de conservación : textos
    ' para poner a cañón en en las líneas
    ' 10 y 11 de Euroscale, o la línea
    ' 10 de SC10 (y EcoLabel?)
    '//////////////////////////////////////
    'With MiGrpConservacion
    '    On Error GoTo 0
    '    .Fields.Append .CreateField("GRUPO", dbDouble, 2)
    '    .Fields.Append .CreateField("TXT1", dbText, 25)
    '    .Fields.Append .CreateField("TXT2", dbText, 25)
    '    .Fields.Append .CreateField("TXTSC10", dbText, 25)
    '    For i = 0 To .Fields.Count - 1
    '        .Fields(i).AllowZeroLength = True
    '    Next i
    '    Base.TableDefs.Append MiGrpConservacion
    'End With
    
'End If
'*********************
'/////////////////////////////////////
' tabla de artículos
'/////////////////////////////////////
With mitabledef
      On Error Resume Next
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("CODI_SUB", dbDouble, 3)
      .Fields.Append .CreateField("CODI_FAM", dbDouble, 4)
      .Fields.Append .CreateField("PLU", dbDouble, 4)
      .Fields.Append .CreateField("CODIGO", dbDouble, 6)
      .Fields.Append .CreateField("CODI_PES", dbText, 1)
      .Fields.Append .CreateField("TIPO_IVA", dbDouble, 1)
      .Fields.Append .CreateField("PRECIO", dbDouble, 7)
      .Fields.Append .CreateField("USATRAMOS", dbBoolean, 1)
      .Fields.Append .CreateField("PRC100G", dbBoolean, 1)
      .Fields.Append .CreateField("EUROS", dbDouble, 7)
      .Fields.Append .CreateField("PRC3", dbDouble, 7)
      .Fields.Append .CreateField("ETQ", dbDouble, 2)
      .Fields.Append .CreateField("PREF", dbDouble, 3)
      .Fields.Append .CreateField("FCB", dbDouble, 3)
      .Fields.Append .CreateField("NING", dbDouble, 3)
      .Fields.Append .CreateField("CADUCIDAD", dbDouble, 3)
      .Fields.Append .CreateField("TARA", dbDouble, 5)
      .Fields.Append .CreateField("TIP_LET1", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET2", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET3", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET4", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET5", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET6", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET7", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET8", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET9", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LET0", dbDouble, 1)
      .Fields.Append .CreateField("TIP_LETX", dbDouble, 1)
      .Fields.Append .CreateField("DES_PLU1", dbText, 40)
      .Fields.Append .CreateField("DES_PLU2", dbText, 40)
      .Fields.Append .CreateField("DES_PLU3", dbText, 40)
      .Fields.Append .CreateField("DES_PLU4", dbText, 40)
      .Fields.Append .CreateField("DES_PLU5", dbText, 40)
      .Fields.Append .CreateField("DES_PLU6", dbText, 40)
      .Fields.Append .CreateField("DES_PLU7", dbText, 40)
      .Fields.Append .CreateField("DES_PLU8", dbText, 40)
      .Fields.Append .CreateField("DES_PLU9", dbText, 40)
      .Fields.Append .CreateField("DES_PLU0", dbText, 40)
      .Fields.Append .CreateField("DES_PLUX", dbText, 40)
      .Fields.Append .CreateField("DES_PLU11", dbText, 40)
      .Fields.Append .CreateField("DES_PLU12", dbText, 40)
      .Fields.Append .CreateField("DES_PLU13", dbText, 40)
      .Fields.Append .CreateField("DES_PLU14", dbText, 40)
      .Fields.Append .CreateField("DES_PLU15", dbText, 40)
      .Fields.Append .CreateField("DES_PLU16", dbText, 40)
      .Fields.Append .CreateField("DES_PLU17", dbText, 40)
      .Fields.Append .CreateField("DES_PLU18", dbText, 40)
      .Fields.Append .CreateField("DES_PLU19", dbText, 40)
      .Fields.Append .CreateField("DES_PLU20", dbText, 40)
      .Fields.Append .CreateField("TRAN_PLU", dbText, 1)
      .Fields.Append .CreateField("TRAN_PLUSC10", dbText, 1)
      .Fields.Append .CreateField("TRAN_TEXTO", dbText, 1)
      .Fields.Append .CreateField("TRAN_TEXTOSC10", dbText, 1)
      .Fields.Append .CreateField("TRAN_TX1", dbText, 1)
      .Fields.Append .CreateField("TRAN_TX1SC10", dbText, 1)
      .Fields.Append .CreateField("ART_CB", dbText, 13)
      
      .Fields.Append .CreateField("ART_CB14", dbText, 14)
      .Fields.Append .CreateField("label2", dbText, 60)
      
      .Fields.Append .CreateField("TRAN_CB", dbText, 1)
      .Fields.Append .CreateField("TRAN_CBSC10", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      .Fields.Append .CreateField("TRAN_EN", dbText, 1)
      .Fields.Append .CreateField("TRAN_EL", dbText, 1)
      .Fields.Append .CreateField("TRAN_EP", dbText, 60)

      .Fields.Append .CreateField("TRAN_TEXTOEL", dbText, 1)
      '************************
      ' Ahorramas
      .Fields.Append .CreateField("GRUPO_CONSERV", dbDouble, 2)
      .Fields.Append .CreateField("BALENV", dbBoolean, 1)
      .Fields.Append .CreateField("lMix", dbBoolean)
      .Fields.Append .CreateField("Poid", dbLong)
      .Fields.Append .CreateField("imagen", dbText, 250)
      .Fields.Append .CreateField("posicion", dbLong)
      .Fields.Append .CreateField("tara_envasado", dbDouble, 5)
      .Fields.Append .CreateField("tipo_ning", dbInteger)
      .Fields.Append .CreateField("tar1", dbDouble, 7)
      .Fields.Append .CreateField("tar2", dbDouble, 7)
      .Fields.Append .CreateField("tar3", dbDouble, 7)
      .Fields.Append .CreateField("tar4", dbDouble, 7)
      .Fields.Append .CreateField("tar5", dbDouble, 7)
      .Fields.Append .CreateField("tar6", dbDouble, 7)
      .Fields.Append .CreateField("tar7", dbDouble, 7)
      .Fields.Append .CreateField("tar8", dbDouble, 7)
      .Fields.Append .CreateField("tar9", dbDouble, 7)
      .Fields.Append .CreateField("tar10", dbDouble, 7)
      .Fields.Append .CreateField("onkey", dbBoolean)
      ' ***************************
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
'     For Each prpBucle In .Properties
'            On Error Resume Next
'            If prpBucle <> "" Then Debug.Print "    " & prpBucle.Name & " = " & prpBucle
'     Next prpBucle

      Base.TableDefs.Append mitabledef
End With
'///////////////////////
' secciones
'///////////////////////
With MiSeccion
      .Fields.Append .CreateField("COM", dbDouble, 2)
      .Fields.Append .CreateField("SEC_IP", dbText, 15)
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("DESCRIPCIO", dbText, 50)
      .Fields.Append .CreateField("NOMBRE", dbText, 50)
      .Fields.Append .CreateField("TRAN_SEC", dbText, 1)
      .Fields.Append .CreateField("TRAN_SECSC10", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      .Fields.Append .CreateField("ENVIARDATOS", dbText, 5)
      .Fields.Append .CreateField("MULTIPLE", dbDouble, 2)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
      Base.TableDefs.Append MiSeccion
    
End With
'//////////////////////////
'///////////////////////
' Trade
'///////////////////////
With Trade
      .Fields.Append .CreateField("company", dbText, 50)
      .Fields.Append .CreateField("name", dbText, 50)
      .Fields.Append .CreateField("address", dbText, 50)
      .Fields.Append .CreateField("quarter", dbText, 50)
      .Fields.Append .CreateField("web", dbText, 50)
      .Fields.Append .CreateField("shopweb", dbText, 50)
      .Fields.Append .CreateField("description", dbText, 50)
      .Fields.Append .CreateField("slogan", dbText, 50)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
      Base.TableDefs.Append Trade
    
End With

' subsecciones
With misubsec
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("CODI_SUB", dbDouble, 3)
      .Fields.Append .CreateField("TRAN_SUB", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
      Base.TableDefs.Append misubsec
 End With
 '////////////////////////////
 ' Familias
 With mifam

      .Fields.Append .CreateField("CODI_IDENT", dbInteger)
      .Fields.Append .CreateField("SECC_MAQUI", dbInteger)
      .Fields.Append .CreateField("codi_fam", dbInteger)
      .Fields.Append .CreateField("TRAN_FAM", dbText, 1)
      .Fields.Append .CreateField("INDEX", dbInteger)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
      Base.TableDefs.Append mifam
 End With
'/////////////////////
' equipos
With miequip
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
      .Fields.Append .CreateField("PROG_EQP", dbText, 20)
      .Fields.Append .CreateField("MODELO", dbDouble, 1)
      .Fields.Append .CreateField("DESCRIPCIO", dbText, 21)
      .Fields.Append .CreateField("TRAN_EQP", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      .Fields.Append .CreateField("ss165", dbBoolean)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
            
      Base.TableDefs.Append miequip
 End With


 '///////////////////////
 ' códigos de barras
 With micodbar
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
      .Fields.Append .CreateField("SECC_EQP", dbDouble, 1)
      .Fields.Append .CreateField("CB_VNT", dbText, 12)
      .Fields.Append .CreateField("CB_SUP", dbText, 12)
      .Fields.Append .CreateField("CB_MIX", dbText, 12)
      .Fields.Append .CreateField("CB_EV1", dbText, 12)
      .Fields.Append .CreateField("CB_EV2", dbText, 12)
      
      .Fields.Append .CreateField("TRAN_BAR", dbText, 1)
      .Fields.Append .CreateField("TRAN_BARSC10", dbText, 1)
      .Fields.Append .CreateField("TRAN_EN", dbText, 1)
      .Fields.Append .CreateField("TRAN_EL", dbText, 1)
      .Fields.Append .CreateField("TRAN_EP", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
            
      Base.TableDefs.Append micodbar
 End With

'//////////////////////////////
' cabeceras y leyendas
With micabley
      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
      .Fields.Append .CreateField("SECC_EQP", dbBoolean, 1)
      .Fields.Append .CreateField("CABECERA01", dbText, 25)
      .Fields.Append .CreateField("CABECERA02", dbText, 25)
      .Fields.Append .CreateField("CABECERA03", dbText, 25)
      .Fields.Append .CreateField("CABECERA04", dbText, 25)
      .Fields.Append .CreateField("CABECERA05", dbText, 25)
      .Fields.Append .CreateField("LEYENDA01", dbText, 25)
      .Fields.Append .CreateField("LEYENDA02", dbText, 25)
      .Fields.Append .CreateField("LEYENDA03", dbText, 25)
      .Fields.Append .CreateField("LEYENDA04", dbText, 25)
      .Fields.Append .CreateField("LEYENDA05", dbText, 25)
      .Fields.Append .CreateField("TLETRACAB1", dbText, 1)
      .Fields.Append .CreateField("TLETRACAB2", dbText, 1)
      .Fields.Append .CreateField("TLETRACAB3", dbText, 1)
      .Fields.Append .CreateField("TLETRACAB4", dbText, 1)
      .Fields.Append .CreateField("TLETRACAB5", dbText, 1)
      .Fields.Append .CreateField("TLETRALEY1", dbText, 1)
      .Fields.Append .CreateField("TLETRALEY2", dbText, 1)
      .Fields.Append .CreateField("TLETRALEY3", dbText, 1)
      .Fields.Append .CreateField("TLETRALEY4", dbText, 1)
      .Fields.Append .CreateField("TLETRALEY5", dbText, 1)
      .Fields.Append .CreateField("TRAN_CAB", dbText, 1)
      .Fields.Append .CreateField("TRAN_CABSC10", dbText, 1)
      .Fields.Append .CreateField("TRAN_EN", dbText, 1)
      .Fields.Append .CreateField("TRAN_EL", dbText, 1)
      .Fields.Append .CreateField("TRAN_EP", dbText, 1)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      
      
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
            
      Base.TableDefs.Append micabley
 End With

'/////////////////////////////
' teclas de PLUs
With MiTecla
      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("CODIGO", dbDouble, 6)
      .Fields.Append .CreateField("PLU", dbDouble, 4)
      .Fields.Append .CreateField("DESC_PLU", dbText, 25)
      .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
      .Fields.Append .CreateField("SECC_EQP", dbBoolean, 1)
      .Fields.Append .CreateField("TABLA", dbDouble, 2)
      .Fields.Append .CreateField("CODI_TEC", dbDouble, 3)
      .Fields.Append .CreateField("PROG_TEC", dbText, 5)
      .Fields.Append .CreateField("TRAN_TEC", dbText, 1)
      .Fields.Append .CreateField("TRAN_TECSC10", dbText, 2)
      .Fields.Append .CreateField("BORRADO", dbBoolean)

      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i

      Base.TableDefs.Append MiTecla
 End With
'///////////////////////
' vendedores
With MiVendedor
      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
      .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
      .Fields.Append .CreateField("IDENT_VEND", dbDouble, 4)
      .Fields.Append .CreateField("CODI_VEND", dbDouble, 6)
      .Fields.Append .CreateField("TEC_VEND", dbDouble, 2)
      .Fields.Append .CreateField("NOMBRE", dbText, 30)
      .Fields.Append .CreateField("PROG_VEND", dbText, 45)
      .Fields.Append .CreateField("TRAN_VEND", dbText, 1)
      .Fields.Append .CreateField("TRAN_VENDSC10", dbText, 1)
      .Fields.Append .CreateField("TRAN_ASIGN", dbText, 1)
      .Fields.Append .CreateField("TRAN_ASIGNSC10", dbText, 1)
      .Fields.Append .CreateField("FECHA", dbDate, 8)
      .Fields.Append .CreateField("BORRADO", dbBoolean)
      .Fields.Append .CreateField("imagen", dbText, 250)
      For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
      Next i
      
      Base.TableDefs.Append MiVendedor
 End With

'//////////////////////
' publicidad y ofertas
'With mipubli
'      .Fields.Append .CreateField("TIENDA", dbDouble, 2)
''      .Fields.Append .CreateField("CODI_IDENT", dbDouble, 2)
 '     .Fields.Append .CreateField("SECC_MAQUI", dbDouble, 2)
 '     .Fields.Append .CreateField("NUMERO_EQP", dbDouble, 2)
 '     .Fields.Append .CreateField("PUBLICIDAD", dbText, 100)
 '     .Fields.Append .CreateField("OFERTA01", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA02", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA03", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA04", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA05", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA06", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA07", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA08", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA09", dbText, 16)
 '     .Fields.Append .CreateField("OFERTA10", dbText, 16)
 '     .Fields.Append .CreateField("PROG_PUB", dbText, 113)
 '     .Fields.Append .CreateField("TRAN_PUB", dbText, 1)
 '     .Fields.Append .CreateField("TRAN_PUBSC10", dbText, 1)
 '     .Fields.Append .CreateField("TRAN_OFE", dbText, 1)
 '     .Fields.Append .CreateField("TRAN_OFESC10", dbText, 1)
 '     .Fields.Append .CreateField("TRAN_ACT", dbText, 1)
 '     .Fields.Append .CreateField("TRAN_ACTSC10", dbText, 1)
 '      .Fields.Append .CreateField("BORRADO", dbBoolean)
 '
 '     For i = 0 To .Fields.Count - 1
 '     .Fields(i).AllowZeroLength = True
 '     Next i
 '
 '     Base.TableDefs.Append mipubli
 'End With


CerrarBase Base




End Sub
Public Sub Crea_IVA()
    '////////////////////////
    ' Tabla de tipos de IVA
    ' (Euroscale, y Gama Baja, ya
    ' que las V-12 ignoran este
    ' dato, si bien se recompone
    ' al tratar totales en Hydra +
    Dim Base As dao.Database
    Dim MiIva As TableDef
    Dim i As Integer
    Set Base = OpenDatabase(Base_General)
    Set MiIva = Base.CreateTableDef("IVA")
    With MiIva
        .Fields.Append .CreateField("Tipo", dbDouble, 6)
        .Fields.Append .CreateField("Valor", dbDouble, 6)
        .Fields.Append .CreateField("Tran_IVA", dbText, 1)
        .Fields.Append .CreateField("Tran_en", dbText, 1)
        .Fields.Append .CreateField("Tran_el", dbText, 1)
        .Fields.Append .CreateField("interno", dbLong)
        .Fields.Append .CreateField("descripcion", dbText, 50)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append MiIva
    End With
    
    Base.Close
    Set Base = Nothing
    'Workspaces(0).close
    
End Sub
Sub tot_table()
   Dim tdftot As TableDef
   Dim dbftot As dao.Database
   Dim i As Integer
   
   Set dbftot = OpenDatabase(Base_General)
   
   
   '//////////////////////////////////
   ' LOG de cambios de precio y rechazos
   Set tdftot = dbftot.CreateTableDef("log")
   With tdftot
      .Fields.Append .CreateField("codi_ident", dbInteger)
      .Fields.Append .CreateField("codi_sub", dbInteger)
      .Fields.Append .CreateField("codigo", dbDouble)
      .Fields.Append .CreateField("plu", dbInteger)
      .Fields.Append .CreateField("precio", dbDouble)
      .Fields.Append .CreateField("tipo_ven", dbText) 'tipo de venta
      .Fields.Append .CreateField("codi_fam", dbInteger)
      .Fields.Append .CreateField("caducidad", dbInteger)
      .Fields.Append .CreateField("tara", dbDouble)
      .Fields.Append .CreateField("des_plu1", dbText, 40)
      .Fields.Append .CreateField("comentario", dbText)
      .Fields.Append .CreateField("fecha", dbDate)
      .Fields.Append .CreateField("texto", dbText, 255)
      
      On Error Resume Next

      For i = 0 To .Fields.Count - 1
        .Fields(i).AllowZeroLength = True
      Next i
        
      dbftot.TableDefs.Append tdftot
  
   End With
   
   Set tdftot = dbftot.CreateTableDef("rechazos")
   With tdftot
      .Fields.Append .CreateField("codi_ident", dbInteger)
      .Fields.Append .CreateField("codi_sub", dbInteger)
      .Fields.Append .CreateField("codigo", dbDouble)
      .Fields.Append .CreateField("plu", dbInteger)
      .Fields.Append .CreateField("precio", dbDouble)
      .Fields.Append .CreateField("tipo_ven", dbText) 'tipo de venta
      .Fields.Append .CreateField("codi_fam", dbInteger)
      .Fields.Append .CreateField("caducidad", dbInteger)
      .Fields.Append .CreateField("tara", dbDouble)
      .Fields.Append .CreateField("des_plu1", dbText, 40)
      .Fields.Append .CreateField("comentario", dbText)
      On Error Resume Next

      For i = 0 To .Fields.Count - 1
        .Fields(i).AllowZeroLength = True
      Next i
        
      dbftot.TableDefs.Append tdftot
  
   End With
   
   
   
   
   dbftot.Close
   

End Sub
Public Sub Crea_Vacuno()
    '////////////////////////////////////////////
    ' Tablas creadas para almacenar las
    ' fichas de vacuno, se modificó ya
    ' que las balanzas Euroacale de versión C
    ' tenían menos datos de trazabilidad
    ' que las D y posteriores
    '/////////////////////////////////////////////
    Dim Base As dao.Database
    Dim MiFichaVacuno As TableDef
    Dim i As Integer
    Set Base = OpenDatabase(Base_General)
    Set MiFichaVacuno = Base.CreateTableDef("FichaVacuno")
    With MiFichaVacuno
        On Error GoTo 0
        .Fields.Append .CreateField("Codigo", dbText, 6)
        .Fields.Append .CreateField("Identificador", dbText, 255)
        .Fields.Append .CreateField("CodNacimiento", dbDouble, 2)
        .Fields.Append .CreateField("CodCrianza", dbDouble, 2)
        .Fields.Append .CreateField("CodSacrificio", dbDouble, 2)
        .Fields.Append .CreateField("CodDespiece", dbDouble, 2)
        .Fields.Append .CreateField("CodProduccion", dbDouble, 2)
        .Fields.Append .CreateField("RSISacrificio", dbText, 255)
        .Fields.Append .CreateField("RSIdespiece", dbText, 255)
        .Fields.Append .CreateField("Categoria", dbDouble, 2)
        .Fields.Append .CreateField("Raza", dbDouble, 2)
        .Fields.Append .CreateField("Fecha_Sac", dbDate)
        .Fields.Append .CreateField("Edad", dbDouble)
        .Fields.Append .CreateField("Sexo", dbDouble, 2)
        .Fields.Append .CreateField("Peso", dbDouble)
        .Fields.Append .CreateField("Texto", dbText, 255)
        .Fields.Append .CreateField("Fecha_Nac", dbDate)
        .Fields.Append .CreateField("Fecha_Des", dbDate)
        .Fields.Append .CreateField("Lote", dbText, 255)
        .Fields.Append .CreateField("TranFicha", dbText, 1)
        .Fields.Append .CreateField("Modificado", dbBoolean, 1)
        .Fields.Append .CreateField("Borrado", dbBoolean, 1)
        .Fields.Append .CreateField("valorpproduccion", dbText, 100)
        .Fields.Append .CreateField("valorpnacimiento", dbText, 100)
        .Fields.Append .CreateField("valorpcrianza", dbText, 100)
        .Fields.Append .CreateField("valorpsacrificio", dbText, 100)
        .Fields.Append .CreateField("valorpdespiece", dbText, 100)
        .Fields.Append .CreateField("valorsexo", dbText, 100)
        .Fields.Append .CreateField("valorraza", dbText, 100)
        .Fields.Append .CreateField("valorcategoria", dbText, 100)
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append MiFichaVacuno
    End With
    Base.Close
End Sub
Public Sub crea_cliente()
    '///////////////////////////////////
    ' datos de clientes, sólo Euroscale
    Dim Base As dao.Database
    Dim Tabla As TableDef
    Dim bucle As Integer
    
    CadenadeLog "crea_cliente"
    
    Set Base = OpenDatabase(Base_General)
    Set Tabla = Base.CreateTableDef("Cliente")
    With Tabla
        .Fields.Append .CreateField("Cod_Cli", dbDouble)
        .Fields.Append .CreateField("Nom_Cli", dbText, 255)
        .Fields.Append .CreateField("Data1", dbText, 255)
        .Fields.Append .CreateField("Data2", dbText, 255)
        .Fields.Append .CreateField("Data3", dbText, 255)
        .Fields.Append .CreateField("Data4", dbText, 255)
        .Fields.Append .CreateField("tran_cli", dbText, 1)
        .Fields.Append .CreateField("borrado", dbBoolean)
        .Fields.Append .CreateField("factura", dbBoolean)
        .Fields.Append .CreateField("discount", dbDouble)
        For bucle = 0 To .Fields.Count - 1
            .Fields(bucle).AllowZeroLength = True
        Next bucle
    End With
    Base.TableDefs.Append Tabla
    Base.Close
End Sub
Sub crea_cabtiquets()
'////////////////////////////////////////////////
' tablas de cabeceras y líneas de tiquets
' para Euroscale (cabecera / tickets) y
' V-12 (sea por SC10 o C9/C10),(cabeceraSC10 / ticketsSC10)
' son iguales, para poder utilizarlas desde las funciones
' "Actualiza_Tabla_Cabecera" y "Actualiza_Tabla_Tickets"
' con el mismo código para Euroscale, V-12 e IV4TQ
'////////////////////////////////////////////////
Dim tdfcabt As TableDef
Dim tdftot As TableDef
Dim Base As dao.Database
Dim bucle As Integer
Dim TextoAdicional As String
Dim i As Integer
TextoAdicional = ""
'For Bucle = 1 To 2
    'If Bucle = 2 Then TextoAdicional = "SC10"
    Set Base = OpenDatabase(Base_General)
    Set tdfcabt = Base.CreateTableDef("cabecera" & TextoAdicional)
    With tdfcabt
        .Fields.Append .CreateField("NUME", dbText)
        .Fields.Append .CreateField("SECCION", dbText)
        .Fields.Append .CreateField("BALANZA", dbInteger)
        .Fields.Append .CreateField("VENDED", dbText)
        .Fields.Append .CreateField("NUMLIN", dbText)
        .Fields.Append .CreateField("LINCAN", dbText)
        .Fields.Append .CreateField("IMPORTE", dbDouble, 15)
        .Fields.Append .CreateField("HORA", dbText)
        .Fields.Append .CreateField("ONOFF", dbText)
        .Fields.Append .CreateField("STPETIC", dbText)
        .Fields.Append .CreateField("CODCLI", dbText)
        .Fields.Append .CreateField("FECHA", dbText)
        .Fields.Append .CreateField("TYPTIC", dbText)
        .Fields.Append .CreateField("CODBAR", dbText)
        .Fields.Append .CreateField("RESERV1", dbText)
        .Fields.Append .CreateField("RESERV2", dbText)
        .Fields.Append .CreateField("RESERV3", dbText)
        On Error Resume Next
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        On Error GoTo 0
        Base.TableDefs.Append tdfcabt
    End With
    '*****************************Lineas de tiquets**************
    Set tdfcabt = Base.CreateTableDef("tickets" & TextoAdicional)
    With tdfcabt
        .Fields.Append .CreateField("DATE", dbText)
        .Fields.Append .CreateField("HOUR", dbText)
        .Fields.Append .CreateField("NUMBAL", dbInteger)
        .Fields.Append .CreateField("VENDOR", dbText)
        .Fields.Append .CreateField("NUME", dbText)
        .Fields.Append .CreateField("NUMLIN", dbText)
        .Fields.Append .CreateField("SECC", dbText)
        .Fields.Append .CreateField("PLU", dbText)
        .Fields.Append .CreateField("CODE", dbText)
        .Fields.Append .CreateField("ITEM", dbText)
        .Fields.Append .CreateField("VAR", dbText)
        .Fields.Append .CreateField("PRICE", dbDouble)
        .Fields.Append .CreateField("AMOUNT", dbDouble)
        .Fields.Append .CreateField("WEIGHT", dbDouble)
        .Fields.Append .CreateField("UNITS", dbDouble)
        .Fields.Append .CreateField("STPETIC", dbText)
        .Fields.Append .CreateField("CODI_PES", dbText)
        .Fields.Append .CreateField("LINCAN", dbText)
        .Fields.Append .CreateField("IVA", dbInteger)
        .Fields.Append .CreateField("TYPTIC", dbText)
        .Fields.Append .CreateField("RESERV1", dbText)
        .Fields.Append .CreateField("RESERV2", dbText)
        .Fields.Append .CreateField("RESERV3", dbText)
        .Fields.Append .CreateField("TIPOPAGO", dbInteger)

        On Error Resume Next
        
        For i = 0 To .Fields.Count - 1
            .Fields(i).AllowZeroLength = True
        Next i
        Base.TableDefs.Append tdfcabt
    End With
'Next Bucle
'***************************
' tablas de grandes totales
'***************************
Set tdftot = Base.CreateTableDef("gtsecs")
With tdftot
    .Fields.Append .CreateField("Modo", dbInteger)
    .Fields.Append .CreateField("fecha", dbDate)
    .Fields.Append .CreateField("codi_ident", dbInteger)
    .Fields.Append .CreateField("codi_sub", dbInteger)
    .Fields.Append .CreateField("codi_fam", dbInteger)
    .Fields.Append .CreateField("tiquets", dbInteger)
    .Fields.Append .CreateField("operacion", dbDouble)
    .Fields.Append .CreateField("peso", dbDouble)
    .Fields.Append .CreateField("unidades", dbDouble)
    .Fields.Append .CreateField("pesetas", dbDouble)
    .Fields.Append .CreateField("secc_maqui", dbInteger)
    
    On Error Resume Next
    For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
    Next i
    Base.TableDefs.Append tdftot
End With

Set tdftot = Base.CreateTableDef("gtvend")
With tdftot
    .Fields.Append .CreateField("Modo", dbInteger)
    .Fields.Append .CreateField("fecha", dbDate)
    .Fields.Append .CreateField("codi_ident", dbInteger)
    .Fields.Append .CreateField("ident_vend", dbInteger)
    .Fields.Append .CreateField("nombre", dbText)
    .Fields.Append .CreateField("supertot", dbDouble) 'totcance+total
    .Fields.Append .CreateField("totcance", dbDouble)
    .Fields.Append .CreateField("total", dbDouble)
    .Fields.Append .CreateField("tiquets", dbDouble)
    .Fields.Append .CreateField("operacion", dbDouble)
    .Fields.Append .CreateField("peso", dbDouble)
    .Fields.Append .CreateField("cancela", dbDouble) 'cancelaciones
    .Fields.Append .CreateField("tiempo", dbDouble) 'tiempo de trabajo en minutos
    .Fields.Append .CreateField("unidad", dbDouble)
    .Fields.Append .CreateField("pesetas", dbDouble)
    .Fields.Append .CreateField("desembolso", dbDouble)
    .Fields.Append .CreateField("reembolso", dbDouble)
    .Fields.Append .CreateField("impor_neg", dbDouble)
    .Fields.Append .CreateField("credito", dbDouble)
    .Fields.Append .CreateField("secc_maqui", dbInteger)
    On Error Resume Next
    For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
    Next i
    Base.TableDefs.Append tdftot
End With

Set tdftot = Base.CreateTableDef("gtarti")
With tdftot
    .Fields.Append .CreateField("Modo", dbInteger)
    .Fields.Append .CreateField("fecha", dbDate)
    .Fields.Append .CreateField("plu", dbInteger)
    .Fields.Append .CreateField("des_plu1", dbText, 40)
    .Fields.Append .CreateField("codi_ident", dbInteger)
    .Fields.Append .CreateField("codi_sub", dbInteger)
    .Fields.Append .CreateField("codi_fam", dbInteger)
    .Fields.Append .CreateField("codigo", dbDouble)
    .Fields.Append .CreateField("precio", dbDouble)
    .Fields.Append .CreateField("unidades", dbDouble)
    .Fields.Append .CreateField("peso", dbDouble)
    .Fields.Append .CreateField("pesetas", dbDouble)
    .Fields.Append .CreateField("tara", dbInteger)
    .Fields.Append .CreateField("caducidad", dbInteger)
    .Fields.Append .CreateField("operacion", dbDouble)
    .Fields.Append .CreateField("tipo_iva", dbDouble)
    .Fields.Append .CreateField("modificado", dbBoolean)
    .Fields.Append .CreateField("secc_maqui", dbInteger)
    On Error Resume Next
    For i = 0 To .Fields.Count - 1
      .Fields(i).AllowZeroLength = True
    Next i
    Base.TableDefs.Append tdftot
End With
'********************************
' fin de tablas de grandes totales
'*********************************
CerrarBase Base
End Sub


Public Function DB_Consulta_Art(Art As DB_Articulo) As Boolean
    Dim Registro As New RecordNet
    Dim RegTramo As New RecordNet
    Dim Resp As Boolean
    Dim StrSQL As String
    Dim bucle As Integer
    
    Resp = False
    StrSQL = ""
    If Not ((Art.Mostrador = 0) And (Art.Plu = 0) And (Art.codigo = 0)) Then
        If Art.codigo <> 0 Then StrSQL = " and codigo=" & Str(Art.codigo)
        If Art.Mostrador <> 0 Then StrSQL = StrSQL & " and codi_ident=" & Str(Art.Mostrador)
        If Art.Plu <> 0 Then StrSQL = StrSQL & " and plu=" & Str(Art.Plu)
        If HaySeguridad And UsuarioActual.Asociado <> 0 Then
            StrSQL = StrSQL & " and codi_ident=" & Str(UsuarioActual.Asociado)
        End If
        StrSQL = "select * from articulo where borrado=false" & StrSQL
        Registro.OpenRecordset (StrSQL)
        With Registro
            If Not .EOF Then
                Resp = True
                If .Fields("onkey") Then
                    Art.onkey = 1
                Else
                    Art.onkey = 0
                End If
                Art.Plu = .Fields("plu")
                Art.codigo = .Fields("codigo")
                Art.Mostrador = .Fields("codi_ident")
                Art.familia = .Fields("codi_fam")
                Art.subsec = .Fields("codi_sub")
                If Art.Mostrador = 0 Then
                    Art.Mostrador = .Fields("secc_maqui")
                End If
                Art.caducidad = .Fields("caducidad")
                Art.Preferente = .Fields("Pref")
                Art.tara = .Fields("tara")
                Art.PRC100G = .Fields("prc100g")
                Art.GrpConserv = .Fields("grupo_conserv")
                Art.IVA = .Fields("tipo_iva")
                Art.Merma = .Fields("fcb")
                Art.Presel = .Fields("ning")
                Art.Etiqueta = .Fields("etq")
                Art.Balenv = .Fields("balenv")
#If BALSAM = True Then
                For bucle = 0 To 8
                    Art.tarifa(bucle) = .Fields("tar" & CStr(bucle + 1))
                Next bucle

#Else
                For bucle = 0 To 9
                    Art.tarifa(bucle) = .Fields("tar" & CStr(bucle + 1))
                Next bucle
#End If
                bucle = 0
                If Not IsNull(.Fields("art_cb")) Then
                    Art.ean13 = Trim(.Fields("art_cb"))
                Else
                    Art.ean13 = ""
                End If
                If .Fields("codi_pes") = "W" Then
                    Art.WGH = False
                Else
                    Art.WGH = True
                End If
                If UsaEuro Then
                    Art.precio = .Fields("euros")
                Else
                    Art.precio = .Fields("precio")
                End If
                Art.Descriptivos(0) = .Fields("des_plu1")
                Art.Descriptivos(1) = .Fields("des_plu2")
                Art.Descriptivos(2) = .Fields("des_plu3")
                Art.Descriptivos(3) = .Fields("des_plu4")
                Art.Descriptivos(4) = .Fields("des_plu5")
                Art.Descriptivos(5) = .Fields("des_plu6")
                Art.Descriptivos(6) = .Fields("des_plu7")
                Art.Descriptivos(7) = .Fields("des_plu8")
                Art.Descriptivos(8) = .Fields("des_plu9")
                Art.Descriptivos(9) = .Fields("des_plu0")
                Art.Descriptivos(10) = .Fields("des_plux")
                Art.Descriptivos(11) = .Fields("des_plu11")
                Art.Descriptivos(12) = .Fields("des_plu12")
                Art.Descriptivos(13) = .Fields("des_plu13")
                Art.Descriptivos(14) = .Fields("des_plu14")
                Art.Descriptivos(15) = .Fields("des_plu15")
                Art.Descriptivos(16) = .Fields("des_plu16")
                Art.Descriptivos(17) = .Fields("des_plu17")
                Art.Descriptivos(18) = .Fields("des_plu18")
                Art.Descriptivos(19) = .Fields("des_plu19")
                Art.Descriptivos(20) = .Fields("des_plu20")
                Art.TipoLetra(0) = .Fields("tip_let1")
                Art.TipoLetra(1) = .Fields("tip_let2")
                Art.TipoLetra(2) = .Fields("tip_let3")
                Art.TipoLetra(3) = .Fields("tip_let4")
                Art.TipoLetra(4) = .Fields("tip_let5")
                Art.TipoLetra(5) = .Fields("tip_let6")
                Art.TipoLetra(6) = .Fields("tip_let7")
                Art.TipoLetra(7) = .Fields("tip_let8")
                Art.TipoLetra(8) = .Fields("tip_let9")
                Art.TipoLetra(9) = .Fields("tip_let0")
                Art.PRC1 = -1
                Art.PRC2 = -1
                Art.TRM1 = -1
                Art.TRM2 = -1
                
                Art.lMix = .Fields("lmix")
                Art.nPoid = .Fields("poid")
                Art.posicion = .Fields("posicion")
                Art.PRC3 = .Fields("prc3")
                
                If Not IsNull(.Fields("imagen")) Then
                    Art.Imagen = .Fields("imagen")
                Else
                    Art.Imagen = ""
                End If
                Art.Label2 = Trim(.Fields("label2"))
#If BALSAM = True Then
                Art.Label1 = Trim(.Fields("tran_ep"))
                
                Art.ean14 = Trim(.Fields("art_cb14"))
#End If
                If .Fields("usatramos") = True Then
                    RegTramo.OpenRecordset ("select * from tramos where borrado=false and codigo=" & Art.codigo)
                    If Not RegTramo.EOF Then
                        Art.TRM1 = RegTramo.Fields("tramo1")
                        Art.TRM2 = RegTramo.Fields("tramo2")
                        If UsaEuro Then
                            Art.PRC1 = RegTramo.Fields("euros1")
                            Art.PRC2 = RegTramo.Fields("euros2")
                        Else
                            Art.PRC1 = RegTramo.Fields("precio1")
                            Art.PRC2 = RegTramo.Fields("precio2")
                        End If
                    End If
                End If

            End If
        End With
    End If
    DB_Consulta_Art = Resp
End Function


Private Function convertir1(cadena As String) As String
    If Len(cadena) = 1 Then
        convertir1 = cadena
        Exit Function
    End If
    convertir1 = Chr(Val(cadena) + 48)
End Function

Private Function Desglosa_Codigos(LaCadena As String) As TipoCodigos
    Dim LosCodigos As TipoCodigos
    Dim Buffer As String
    Dim bucle As Long
    Dim cantidad As Integer
    cantidad = 0
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            cantidad = cantidad + 1
            ReDim Preserve LosCodigos.codigo(cantidad + 1)
            LosCodigos.codigo(cantidad) = Buffer
            LosCodigos.cantidad = cantidad
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Desglosa_Codigos = LosCodigos
End Function
Public Function Baja_Clientes(cliente As Long) As Integer
    Dim Retorno As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Retorno = 4
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from cliente where borrado=false and cod_cli=" & cliente)
    With Registro
        If Not .EOF Then
            If Edit_Record(Registro) Then
                .Fields("borrado") = True
                .Fields("tran_cli") = "-"
                .Update
                Retorno = 1
            Else
                Retorno = 60
            End If
        End If
    End With
    Base.Close
    
    Baja_Clientes = Retorno
End Function
Public Function Baja_Familias(ByVal nFam As Long, Optional MyEvento As ClsEvento) As Long
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegArt As dao.Recordset
    Dim MyArt As DB_Articulo
    Dim ncont As Boolean
    Dim Resp As Long
    Resp = 4
    ncont = False
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset("select * from fam_code where codi_fam=" & nFam)
    If Not Registro.EOF Then
        Resp = 0
        Registro.MoveFirst
        Do Until Registro.EOF
            If ncont = False Then
                ncont = True
                Set RegArt = Base.OpenRecordset("select codigo from articulo where borrado=false and codi_fam=" & nFam)
                If Not RegArt.EOF Then
                    RegArt.MoveFirst
                    Do Until RegArt.EOF
                        MyArt.codigo = RegArt.Fields("codigo")
                        MyArt.Plu = 0
                        MyArt.Mostrador = 0
                        Baja_Articulo MyArt
                        If TypeName(MyEvento) <> "Nothing" Then
                            MyEvento.Raise1 MyArt.codigo
                        End If
                        If Not RegArt.EOF Then RegArt.Movenext
                    Loop
                End If
            End If
            If Not Registro.EOF Then 'Registro.Delete
                Registro.Edit
                Registro.Fields("posicion") = 1525
                Registro.Update
            End If
            If Not Registro.EOF Then Registro.Movenext
        Loop
    End If
    Baja_Familias = Resp
End Function
Public Function Alta_Familias(ByVal nFam As Long, ByVal TxtFam As String, ByVal Value As Long, ByVal TxtEtq As String, KeepSec As Boolean, Optional MyEvento As ClsEvento, Optional nIndex As Integer, Optional ByVal sImg As String, Optional ByVal lTraza As Boolean, Optional ByVal nTax As Integer, Optional ByVal sLabel1 As String, Optional ByVal slabel2 As String)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim simagen As String
    Dim nf As Integer
    Dim lA As Boolean
    Resp = 0
    If nFam = 0 Then GoTo AltaFamErr
    On Error GoTo AltaFamErr
    
    If Not (IsEmpty(sImg)) Or sImg <> "" Then
        simagen = Trim(sImg)
    Else
        simagen = ""
    End If
    lA = False
    Set Base = OpenDatabase(Base_General)
    
    Set Registro = Base.OpenRecordset("select * from fam_code where codi_fam=" & nFam)
    With Registro
        If .EOF Then
            lA = True
            .AddNew
            .Fields("codi_fam") = nFam
            .Fields("imagen") = simagen
        Else
            If .Fields("index") <> nIndex Then
                nf = FreeFile()
                Open App.Path & "\cambioposicion" For Output As #nf
                Close #nf
            End If
            .Edit
        End If
        .Fields("codi_ident") = 0
        .Fields("secc_maqui") = 0
        If Trim(TxtFam) <> "" Then
            .Fields("txt_fam") = left(TxtFam, 30)
        Else
            .Fields("txt_fam") = "Familia " & CStr(nFam)
        End If
        .Fields("posicion") = Value
        If Trim(TxtEtq) <> "" Then
            .Fields("etiqueta") = Trim(TxtEtq)
        End If
        .Fields("imagen") = simagen
        .Fields("index") = nIndex
        .Fields("istraza") = lTraza
        If lA Then
            .Fields("tax") = nTax
        Else
            If IsNumeric(nTax) Then
            'If .Fields("tax") = 0 Or (nTax <> 0) Then
                .Fields("tax") = nTax
            'End If
            End If
        End If
        If Not (IsEmpty(sLabel1)) Or sLabel1 <> "" Then
            If Len(sLabel1) > 0 Then .Fields("label1") = sLabel1
        End If
        If Not (IsEmpty(slabel2)) Or slabel2 <> "" Then
            If Len(slabel2) > 0 Then .Fields("label2") = slabel2
        End If
        
        
        .Update
    End With
AltaFamErr:
    If Err.Number <> 0 Then
        Resp = 4
    End If
    On Error GoTo 0
    Alta_Familias = Resp
End Function

Public Function Alta_Familias_Counter(ByVal nFam As Long, ByVal TxtFam As String, ByVal Value As Long, ByRef secs() As Long, KeepSec As Boolean, Optional MyEvento As ClsEvento, Optional myImg As String, Optional ByVal MySec As Integer, Optional ByVal TxtOrden As Integer, Optional ByVal lBaja As Boolean)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim nN As Integer
    Resp = 0
    If nFam = 0 Then GoTo AltaFamErr
    On Error GoTo AltaFamErr


    Set Base = OpenDatabase(Base_General)
    
    Set Registro = Base.OpenRecordset("select * from fam_code where codi_fam=" & nFam & " and posicion<>1525")
    If Registro.EOF Then
        Resp = 4
        Registro.Close
        Set Registro = Nothing
        GoTo AltaFamErr
    Else
        If TxtFam <> "" Then
            Registro.Edit
            Registro.Fields("txt_fam") = TxtFam
            'If Trim(myImg) <> "" Then
            '    Registro.Fields("imagen") = separaFich(Trim(myImg))
            'End If
            Registro.Update
        End If
        
        'If Mysec = 0 Or IsEmpty(Mysec) Then
        '    If Trim(myImg) <> "" Then
        '        Registro.Edit
        '        Registro.Fields("imagen") = separaFich(Trim(myImg))
        '        Registro.Update
        '    End If
        '    Registro.Close
        '    Set Registro = Nothing
        '
        '
        '    Base.Execute "update familias set borrado=true where codi_fam=" & CStr(nFam)
        '
        '    For nN = 0 To Value - 1
        '        Set Registro = Base.OpenRecordset("select * from familias where codi_fam=" & CStr(nFam) & " and secc_maqui=" & CStr(secs(nN)))
        '        With Registro
        '            If .EOF Then
        '                .AddNew
        '                .Fields("codi_fam") = nFam
        '            Else
        '                .Edit
        '                If .Fields("borrado") = True Then .Fields("borrado") = False
        '            End If
        '            If TxtOrden <> 32767 Then .Fields("index") = TxtOrden
        '            .Fields("codi_ident") = secs(nN)
        '            .Fields("secc_maqui") = secs(nN)
        '            .Fields("tran_fam") = " "
        '            .Update
        '        End With
        '        Registro.Close
        '        Set Registro = Nothing
        '    Next nN
        '
        'Else
            
            'If myImg <> "" Then
            '    Registro.Edit
            '    Registro.Fields("imagen") = separaFich(Trim(myImg))
            '    Registro.Update
            'End If
            Registro.Close
            Set Registro = Nothing
            
            If lBaja Then
                Base.Execute "update familias set borrado=true where codi_fam=" & CStr(nFam) & " and secc_maqui=" & CStr(MySec)
                GoTo AltaFamErr
            End If
            If TxtOrden <> 0 Then
                Set Registro = Base.OpenRecordset("select * from familias where index=" & CStr(TxtOrden) & " and secc_maqui=" & CStr(MySec))
                If Not Registro.EOF Then
                    Registro.Close
                    Set Registro = Nothing
                    Resp = 2
                    GoTo AltaFamErr
                End If
                Registro.Close
                Set Registro = Nothing
            End If
            Set Registro = Base.OpenRecordset("select * from familias where codi_fam=" & CStr(nFam) & " and secc_maqui=" & CStr(MySec))
            With Registro
                If .EOF Then
                    .AddNew
                    .Fields("codi_fam") = nFam
                Else
                    .Edit
                    If .Fields("borrado") = True Then .Fields("borrado") = False
                End If
                If TxtOrden <> 32767 Then .Fields("index") = TxtOrden
                .Fields("codi_ident") = MySec
                .Fields("secc_maqui") = MySec
                .Fields("tran_fam") = " "
                .Update
            End With
            Registro.Close
            Set Registro = Nothing
        
        'End If
    End If
    
AltaFamErr:
    If Err.Number <> 0 Then
        Resp = 4
    End If
    On Error Resume Next
    Base.Close
    On Error GoTo 0
    Alta_Familias_Counter = Resp
End Function

Public Function Alta_Familias_Counter_Fichero(ByVal sFich As String)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Resp As Long
    Dim nFich As Integer
    Dim sReg As String
    Dim nS, nf, nP, nV As Long
    Dim sD, sE, sI As String
    
    Resp = 0
    Set Base = OpenDatabase(Base_General)
    
    If Dir(App.Path & "\" & sFich) <> "" Then
        nFich = FreeFile()
        Open App.Path & "\" & sFich For Input As #nFich
        Do While Not EOF(nFich)
            Line Input #nFich, sReg
            nS = Val(Mid(sReg, 1, 2))
            nf = Val(Mid(sReg, 3, 3))
            nP = Val(Mid(sReg, 6, 3))
            sD = Trim(Mid(sReg, 9, 25))
            sE = Trim(Mid(sReg, 34, 25))
            nV = Val(Mid(sReg, 59, 1))
            On Error Resume Next
            sI = Trim(Mid(sReg, 60))
            If Err.Number <> 0 Then
                On Error GoTo 0
                sI = ""
            End If
            On Error GoTo 0
            On Error GoTo AltaFamErr
            
            Set Registro = Base.OpenRecordset("select * from seccion where secc_maqui=" & CStr(nS))
            If Registro.EOF Then
                Registro.AddNew
                Registro.Fields("secc_maqui") = nS
                Registro.Fields("codi_ident") = nS
                Registro.Fields("enviardatos") = "GA"
                Registro.Update
            End If
            Registro.Close
            Set Registro = Nothing
            
            Set Registro = Base.OpenRecordset("select * from fam_code where codi_fam=" & CStr(nf))
            If Registro.EOF Then
                Registro.AddNew
            Else
                Registro.Edit
            End If
            Registro.Fields("codi_fam") = nf
            Registro.Fields("secc_maqui") = 0
            Registro.Fields("codi_ident") = 0
            Registro.Fields("posicion") = nV
            Registro.Fields("txt_fam") = sD
            Registro.Fields("imagen") = sI
            Registro.Fields("etiqueta") = sE
            Registro.Fields("index") = nP
            Registro.Update
            Registro.Close
            Set Registro = Nothing
            
            Set Registro = Base.OpenRecordset("select * from familias where codi_fam=" & CStr(nf) & " and secc_maqui=" & CStr(nS))
            If Registro.EOF Then
                Registro.AddNew
            Else
                Registro.Edit
            End If
            Registro.Fields("codi_fam") = nf
            Registro.Fields("secc_maqui") = nS
            Registro.Fields("codi_ident") = nS
            Registro.Fields("index") = 0
            Registro.Update
            Registro.Close
            Set Registro = Nothing
        
        Loop
        Close (nFich)
        
    End If
    
AltaFamErr:
    If Err.Number <> 0 Then
        Resp = 4
    End If
    On Error Resume Next
    Base.Close
    On Error GoTo 0
    Alta_Familias_Counter_Fichero = Resp
End Function

Public Function Alta_IVAS(TiposDeIVA() As Double) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Integer
    Set Base = OpenDatabase(Base_General)
    For bucle = 0 To 4
        Set Registro = Base.OpenRecordset("select * from iva where tipo=" & bucle)
        With Registro
            If Not .EOF Then
                If Not Edit_Record(Registro) Then
                    Base.Close
                    Alta_IVAS = 60
                    Exit Function
                End If
            Else
                .AddNew
            End If
            .Fields("Tipo") = bucle
            .Fields("Valor") = TiposDeIVA(bucle)
            .Fields("tran_iva") = "-"
            .Fields("tran_en") = ""
            .Fields("tran_el") = ""
            .Update
        End With
    Next bucle
    CerrarBase Base
    Alta_IVAS = 0
End Function
Public Function Alta_Clientes(cliente As DB_Cliente) As Integer
    Dim Retorno As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Integer
    
    If cliente.codigo < 1 Or cliente.codigo > 99999 Then
    Else
        If Len(cliente.Nombre) > 100 Then cliente.Nombre = left(cliente.Nombre, 100)
        'For Bucle = 0 To 3
        '    If Len(cliente.Datos(Bucle)) > 25 Then cliente.Datos(Bucle) = left(cliente.Datos(Bucle), 25)
        'Next Bucle
        If Len(cliente.Datos(0)) > 25 Then cliente.Datos(0) = left(cliente.Datos(0), 25)
        If Len(cliente.Datos(2)) > 25 Then cliente.Datos(2) = left(cliente.Datos(2), 25)
        If Len(cliente.Datos(3)) > 25 Then cliente.Datos(3) = left(cliente.Datos(3), 25)
        Set Base = OpenDatabase(Base_General)
        Set Registro = Base.OpenRecordset("select * from cliente where borrado=false and cod_cli=" & cliente.codigo)
        With Registro
            If .EOF Then
                Retorno = 0
                .AddNew
            Else
                'Retorno = 4
                'If cliente.Nombre <> .Fields("nom_cli") Then Retorno = 1
                'For bucle = 0 To 3
                '    If cliente.Datos(bucle) <> .Fields("data" & Format((bucle + 1), "0")) Then Retorno = 1
                'Next bucle
                'If Retorno = 1 Then
                '    If Not Edit_Record(Registro) Then
                '        Base.Close
                '        Alta_Clientes = 60
                '        Exit Function
                '    End If
                'End If
                Retorno = 1
                .Edit
            End If
            If Retorno = 0 Or Retorno = 1 Then
                .Fields("borrado") = False
                .Fields("tran_cli") = "-"
                .Fields("cod_cli") = cliente.codigo
                .Fields("nom_cli") = cliente.Nombre
                .Fields("data1") = cliente.Datos(0)
                .Fields("data2") = cliente.Datos(1)
                .Fields("data3") = cliente.Datos(2)
                .Fields("data4") = cliente.Datos(3)
                .Fields("factura") = cliente.factura
                .Fields("discount") = cliente.discount
                .Update
            End If
        End With
    End If
    On Error Resume Next
    Base.Close
    On Error GoTo 0
    Alta_Clientes = Retorno
End Function
Public Function Alta_Vendedor(MiVendedor As DB_Vendedor) As Integer
    '*******************************
    ' para dar crear/modificar vendedor sin asignar
    ' tecla, o bien dar baja de tecla a un vendedor
    ' los datos "seccion" y "tecla" deben llegar a cero.
    ' En el campo codi_ident de vendedores se introduce
    ' la sección máquina, no el mostrador
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim bucle As Long
    Dim LaSeccion As Integer
    Dim Retorno As Integer
    ' *********************
    ' Condiciones de error
    If MiVendedor.NVendedor = 0 Then
        Retorno = 19
        GoTo fin
    End If
    If MiVendedor.seccion = 0 And MiVendedor.tecla <> 0 Then
        Retorno = 3
        GoTo fin
    End If
    If Dir(App.Path & "\" & "notecvend") = "" Then
        If MiVendedor.tecla = 0 And MiVendedor.seccion <> 0 Then
            Retorno = 19
            GoTo fin
        End If
    End If
    '**********************
    Set Base = OpenDatabase(Base_General)
    '********************
    If MiVendedor.seccion <> 0 Then
        Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" _
        & MiVendedor.seccion & " and borrado=false")
        If Registro.EOF Then
            CerrarBase Base
            Retorno = 3
            GoTo fin
        Else
            LaSeccion = Registro.Fields("secc_Maqui")
        End If
    End If
    '************************
    If MiVendedor.tecla <> 0 Then
        Set Registro = Base.OpenRecordset("select * from vendedor where tec_vend=" & _
        MiVendedor.tecla & " and codi_ident=" & LaSeccion & " and borrado=false")
        If Not Registro.EOF Then
            If Registro.Fields("ident_vend") <> MiVendedor.NVendedor Then
                'Retorno = 20
                'CerrarBase Base
                'GoTo fin
                Registro.Edit
                Registro.Fields("borrado") = True
                Registro.Fields("prog_vend") = "-"
                Registro.Fields("tran_vend") = "-"
                Registro.Fields("tran_vendSC10") = ""
                Registro.Fields("tran_asign") = ""
                Registro.Fields("tran_asignSC10") = ""
                Registro.Update
            End If
        End If
    End If
    '***********************
     Set Registro = Base.OpenRecordset("select * from vendedor where ident_vend=" _
     & MiVendedor.NVendedor & " and borrado=false and codi_ident=" & LaSeccion)
     With Registro
        If .EOF Then
            .AddNew
            Retorno = 0
        Else
            If Not Edit_Record(Registro) Then
                Base.Close
                Alta_Vendedor = 60
                Exit Function
            End If
            Retorno = 0
        End If
        .Fields("tienda") = 0
        .Fields("ident_vend") = MiVendedor.NVendedor
        .Fields("codi_vend") = MiVendedor.NVendedor
        .Fields("codi_ident") = LaSeccion
        .Fields("Nombre") = Trim(MiVendedor.Descripcion)
        .Fields("secc_maqui") = LaSeccion
        .Fields("tec_vend") = MiVendedor.tecla
        .Fields("tran_vend") = "-"
        .Fields("prog_vend") = "-"
        .Fields("tran_vendSC10") = ""
        .Fields("tran_asign") = ""
        .Fields("tran_asignSC10") = ""
        .Update
    End With
    Registro.Close
    Set Registro = Nothing
    
    CerrarBase Base
fin:
    
    Alta_Vendedor = Retorno
End Function
Public Function Baja_vendedor(MiVendedor As DB_Vendedor) As Integer
    Dim Retorno As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Retorno = 4
    If MiVendedor.NVendedor = 0 Then
        Retorno = 3
        GoTo fin
    End If
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset _
    ("select * from vendedor where ident_vend=" & MiVendedor.NVendedor & " and borrado=false")
    If Not Registro.EOF Then
        Retorno = 0
        Registro.MoveFirst
        Do While Not Registro.EOF
            Registro.Edit
            Registro.Fields("borrado") = True
            Registro.Fields("tran_vend") = ""
            Registro.Fields("tran_asign") = ""
            Registro.Fields("tran_vendSC10") = ""
            Registro.Fields("tran_asignSC10") = ""
            Registro.Update
            Registro.Movenext
        Loop
    End If
    CerrarBase Base
fin:
    
    Baja_vendedor = Retorno
End Function
'ver balsam
Public Function Alta_Articulo(Miarticulo As DB_Articulo, Optional Base As dao.Database, Optional RegArticulo As dao.Recordset, Optional ByVal lForm) As Integer
    Dim ModificacionTramo As Boolean
    Dim ModificacionBasica As Boolean
    'Dim Modificaciontexto As Boolean
    'Dim ModificacionTx1 As Boolean
    'Dim ModificacionCodBar As Boolean
    'Dim ModificacionPrecio As Boolean
    Dim bucle As Integer
    Dim Buclelite As Integer
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Reg3 As dao.Recordset
    Dim Reg4 As dao.Recordset
    Dim Reg5 As dao.Recordset
    Dim RegTramo As dao.Recordset
    Dim Retorno As Integer
    Dim RegFam As New RecordNet
    Dim RegSeccion As dao.Recordset
    Dim RegSubsec As dao.Recordset
    Dim RstRechazos As dao.Recordset
    Dim RegPlu As dao.Recordset
    Dim ChkTramo As Boolean
    Dim MiPrecio As Double
    Dim MiEuro As Double
    Dim CierraDespues As Boolean
    Dim SeCambioPrecio As Boolean
    Dim PluBuffer As Integer
    Dim MiNombre As String
    Dim SoyAlta As Boolean
    'If Dir(App.Path & "\mixmode.txt") = "" And lVL000 = False Then
    '    Miarticulo.lMix = False
    '    Miarticulo.nPoid = 0
    'End If
    SeCambioPrecio = False
    'If FamiliasFijas Then
    '    RegFam.OpenRecordset ("select * from fam_code where codi_fam=" & Miarticulo.familia & " and codi_ident=" & Miarticulo.Mostrador)
    '    If RegFam.EOF Then
    '        Retorno = 3
    '        GoTo fin
    '    End If
    'End If
    'If Miarticulo.TRM1 <> -1 Or Miarticulo.TRM1 <> -1 Or _
    'Miarticulo.PRC1 <> -1 Or Miarticulo.TRM1 <> -1 Then
    If Miarticulo.PRC1 <> 0 Or Miarticulo.PRC2 <> 0 Then
        
        ChkTramo = True
    Else
        ChkTramo = False
    End If
    'c2f 26/07/05
    'If Trim(Miarticulo.EAN13) <> "" Then
    '    If Len(Miarticulo.EAN13) <> 12 Then
    '        Retorno = 5
    '        GoTo fin
    '    End If
    'End If
    If Miarticulo.codigo = 0 Then
        Retorno = 14
        GoTo fin
    End If
    'TouchScale --> Familia sin sección
    'If Miarticulo.Mostrador = 0 Then
    '    Retorno = 15
    '    GoTo FIN
    'End If
    If ChkTramo = True Then
        If Miarticulo.TRM1 = -1 And Miarticulo.PRC1 = -1 Then
            Retorno = 16
            GoTo fin
        End If
        If (Miarticulo.TRM2 <> -1 And Miarticulo.PRC2 = -1) Or _
        (Miarticulo.PRC2 <> -1 And Miarticulo.TRM2 = -1) Then
            Retorno = 16
            GoTo fin
        End If
        If (Miarticulo.TRM1 <> -1 And Miarticulo.PRC1 = -1) Or _
        (Miarticulo.PRC1 <> -1 And Miarticulo.TRM1 = -1) Then
            Retorno = 16
            GoTo fin
        End If
    End If
    '*****
    ' Precio / Euros
    '*****
    If UsaEuro Then
        MiEuro = Miarticulo.precio
        MiPrecio = EUROtoPTA(Miarticulo.precio)
    Else
        MiPrecio = Miarticulo.precio
        MiEuro = PTAtoEURO(Miarticulo.precio)
    End If
    'c2f 1.7.9
    'If TypeName(Base) = "Nothing" Then
        Set Base = OpenDatabase(Base_General)
        CierraDespues = True
    'End If
    'If TypeName(RegArticulo) = "Nothing" Then
        Set RegArticulo = Base.OpenRecordset("select * from articulo")
    'End If
    'If TypeName(RegSeccion) = "Nothing" Then
        Set RegSeccion = Base.OpenRecordset("select * from seccion")
    'End If
    'If TypeName(RegSubsec) = "Nothing" Then
    '    Set RegSubsec = Base.OpenRecordset("select * from subsec")
    'End If
    'If ChkTramo Then
    '    Set RegTramo = Base.OpenRecordset("select count(codigo) from tramos")
    '    With RegTramo
    '        If .Fields(0) >= 400 Then
    '            Retorno = 17
    '            GoTo FIN
    '        End If
    '    End With
    'End If
    Retorno = 4
    RegSeccion.FindFirst "borrado=false and codi_ident=" & Miarticulo.Mostrador
    If RegSeccion.EOF Or RegSeccion.NoMatch Then
        Retorno = 3
        GoTo fin
    End If

    RegArticulo.FindFirst "borrado=false and codigo=" & Miarticulo.codigo
    
    With RegArticulo
    '*******************
        ' Nuevo dato
        If .EOF Or .NoMatch Then
            '********************
            ' comprueba plu repetido
            '********************
            If Miarticulo.Plu <> 0 Then
                Set RegPlu = Base.OpenRecordset _
                ("select plu from articulo where secc_maqui=" & _
                RegSeccion.Fields("secc_Maqui") & " and plu=" & Miarticulo.Plu & " and borrado=false")
                If Not RegPlu.EOF Then
                    Retorno = 21
                    GoTo fin
                End If
            End If
            '********************
            .AddNew
            SoyAlta = True
            Retorno = 0
#If BALSAM = True Then
            For bucle = 0 To 8
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#Else
            For bucle = 0 To 9
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#End If
            bucle = 0
            ModificacionBasica = True
            'If sQueNombreImagen <> "" Then
                sQueNombreImagen = Trim(Miarticulo.Imagen)
                'If Mid(sQueNombreImagen, 1, 1) <> "#" Then
                '    If Right(sQueNombreImagen, 1) = "\" Then
                '        .Fields("imagen") = "#FF00C7B4"
                '    Else
                '        .Fields("imagen") = Trim(Miarticulo.Imagen) 'sQuePathImagen & sQueNombreImagen
                '    End If
                'Else
                '    .Fields("imagen") = Trim(Miarticulo.Imagen)
                '    If Dir(App.Path & "\images\items\" & CStr(Miarticulo.codigo)) <> "" Then
                '        Kill App.Path & "\images\items\" & CStr(Miarticulo.codigo)
                '    End If
                'End If
                If Trim(sQueNombreImagen) <> "" Then
                    .Fields("imagen") = Trim(Miarticulo.Imagen)
                End If
            'End If
            .Fields("tran_EN") = ""
            .Fields("tran_EL") = ""
            '.Fields("tran_ep") = Space(60)
            .Fields("tran_textoel") = ""
            .Fields("tran_plu") = ""
            .Fields("tran_pluSC10") = ""
            .Fields("tran_texto") = ""
            .Fields("tran_textosc10") = ""
            .Fields("tran_tx1") = ""
            .Fields("tran_tx1SC10") = ""
            .Fields("tran_cb") = ""
            .Fields("tran_cbsc10") = ""
            
            .Fields("prc100g") = Miarticulo.PRC100G
            If sPathGigante <> "" Then
                .Fields("prc3") = 0 'Format(Now, "yymmddhhmm")
            Else
                .Fields("prc3") = Format(Now, "yymmddhhmm")
            End If
            If Miarticulo.onkey = 1 Then
                .Fields("onkey") = True
            Else
                .Fields("onkey") = False
            End If
            .Fields("tipo_iva") = Miarticulo.IVA
            .Fields("codigo") = Miarticulo.codigo
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("Pref") = Miarticulo.Preferente
            .Fields("label2") = Trim(Miarticulo.Label2)
#If BALSAM = True Then
            .Fields("tran_ep") = Trim(Miarticulo.Label1)
            .Fields("Etq") = 0
            
            .Fields("art_cb14") = Trim(Miarticulo.ean14)
#Else
            .Fields("Etq") = Miarticulo.Etiqueta
#End If
            .Fields("grupo_conserv") = Miarticulo.GrpConserv
            .Fields("fcb") = Miarticulo.Merma
            .Fields("ning") = Miarticulo.Presel
            
            .Fields("lmix") = Miarticulo.lMix
            .Fields("poid") = Miarticulo.nPoid
            
            .Fields("posicion") = Miarticulo.posicion
            
            RegSeccion.FindFirst "borrado=false and codi_ident=" & Miarticulo.Mostrador
            .Fields("secc_Maqui") = RegSeccion.Fields("secc_Maqui")
            
            'If Miarticulo.plu = 0 Then
            '    Set Reg3 = Base.OpenRecordset _
            '    ("select max(plu) from articulo where borrado=false and secc_maqui=" & RegSeccion.Fields("secc_Maqui"))
            '    If Not IsNull(Reg3.Fields(0)) And Not Reg3.EOF Then
            '        If Reg3.Fields(0) < 9999 Then
            '            .Fields("plu") = Reg3.Fields(0) + 1
            '            PluBuffer = Reg3.Fields(0) + 1
            '        Else
            '
            '            For Bucle = 1 To 9999
            '
            '                RegArticulo.Filter = "borrado = False And secc_maqui = " & RegSeccion.Fields("secc_Maqui") & " and plu=" & Bucle
            '
            '                Set Reg4 = RegArticulo.OpenRecordset
            '                If Reg4.EOF Then
            '                    .Fields("plu") = Bucle
            '                    PluBuffer = Bucle
            '                    Exit For
            '                End If
            '            Next Bucle
            '        End If
            '    Else
            '        .Fields("plu") = 1
            '        PluBuffer = 1
            '    End If
            'Else
                
                .Fields("plu") = Miarticulo.Plu
                PluBuffer = Miarticulo.Plu
            'End If
            If Trim(Miarticulo.ean13) <> "" Then
                .Fields("art_cb") = Miarticulo.ean13
            End If
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("precio") = MiPrecio
            .Fields("Euros") = MiEuro
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            
            .Fields("tara_envasado") = Miarticulo.tara
            
            If Miarticulo.WGH = False Then
                .Fields("codi_pes") = "W"
            Else
                .Fields("codi_pes") = "U"
            End If
            If Miarticulo.Balenv = False Then
                .Fields("Balenv") = False
            Else
                .Fields("Balenv") = True
                Set Reg5 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & RegSeccion.Fields("secc_Maqui"))
                With Reg5
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                        If .Fields("enviardatos") <> "TODOS" Then
                            If Edit_Record(Reg5) Then
                                If .Fields("multiple") <> 1 Then
                                
                                    If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                    If .Fields("enviardatos") = "SC10" Then .Fields("tran_sec") = ""
                                End If
                                .Fields("multiple") = 1
                                .Update
                            End If
                        End If
                        If Not .EOF Then .Movenext
                        Loop
                    End If
                End With
            End If
            ''''''''''''''
            '.Fields("
            For bucle = 0 To 20
                Select Case bucle
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(bucle + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                    Case 11 To 20
                        MiNombre = "des_plu" & Val(bucle)
                End Select
                .Fields(MiNombre) = Miarticulo.Descriptivos(bucle)
                If bucle < 10 Then
                    If bucle < 9 Then
                        MiNombre = "tip_let" & Val(bucle + 1)
                    Else
                        MiNombre = "tip_let0"
                    End If
                    .Fields(MiNombre) = Miarticulo.TipoLetra(bucle)
                End If
            Next bucle
            '*****
            ' Tramos
            '*****
            If ChkTramo Then
                .Fields("usatramos") = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=false")
                With RegTramo
                    If .EOF Then
                        .AddNew
                    Else
                        If Not Edit_Record(RegTramo) Then
                            .AddNew
                        End If
                    End If
                        .Fields("borrado") = False
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If Miarticulo.TRM1 <> 0 And Miarticulo.PRC1 <> 0 Then
                            
                            .Fields("tramo1") = Miarticulo.TRM1
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC1)
                            Else
                                .Fields("euros1") = PTAtoEURO(Miarticulo.PRC1)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                                If Miarticulo.TRM2 <> 0 And Miarticulo.PRC2 <> 0 Then
                                    .Fields("tramo2") = Miarticulo.TRM2
                                    If UsaEuro Then
                                        .Fields("euros2") = Miarticulo.PRC2
                                        .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                    Else
                                        .Fields("precio2") = Miarticulo.PRC2
                                        .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                    End If
                                Else
                                    .Fields("tramo2") = 0
                                    .Fields("precio2") = 0
                                    .Fields("euros2") = 0
                                End If
                        End If
                    .Update
                End With
            Else
                .Fields("usatramos") = False
            End If
                        
            .Update
            'If Miarticulo.Mostrador = 6 Then
            '
            '    Miarticulo.Mostrador = 6
            'End If
            'If Miarticulo.subsec <> "0" Then
            '    'Set RegSubsec = Base.OpenRecordset("select * from subsec where borrado=false")
            '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec
            '    With RegSubsec
            '    If RegSubsec.EOF Or RegSubsec.NoMatch Then
            '        .AddNew
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        '.fields("secc_maqui = Reg2.fields("secc_maqui
            '        .Fields("borrado") = False
            '        .Fields("tran_sub") = ""
            '        .Update
            '    End If
            '    End With
            '
            'End If
            'If Miarticulo.familia <> "0" Then
            '    Set Reg3 = Base.OpenRecordset _
            '    ("select * from familia where borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec & _
            '    " and codi_fam=" & Miarticulo.familia)
            '    With Reg3
            '    If .EOF Then
            '        .AddNew
            '        .Fields("codi_fam") = Miarticulo.familia
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        .Fields("secc_Maqui") = RegSeccion.Fields("secc_Maqui")
            '        .Fields("borrado") = False
            '        .Fields("tran_fam") = ""
            '        .Update
            '    End If
            '    End With
            '
            '
            'End If
            
        Else
            '************************
            ' Modificación
            '************************
            If Miarticulo.Plu <> 0 Then
                If .Fields("plu") <> Miarticulo.Plu Then
                    Retorno = 2
                    GoTo fin
                End If
            End If
            If Not Edit_Record(RegArticulo) Then
                Base.Close
                Alta_Articulo = 60
                Exit Function
            End If
            
            If Miarticulo.onkey = 1 Then
                .Fields("onkey") = True
            Else
                .Fields("onkey") = False
            End If
            
            PluBuffer = .Fields("plu")
            If lForm Then
                ModificacionBasica = True
                ModificacionTramo = True
            Else
                ModificacionBasica = False
                ModificacionTramo = False
            End If
            'If .Fields("grupo_conserv") <> Miarticulo.GrpConserv Then
            '    Modificaciontexto = True
            '    ModificacionTx1 = True
            'End If
            If IsNull(.Fields("tipo_iva")) Then .Fields("tipo_iva") = 0
            If .Fields("prc100g") <> Miarticulo.PRC100G Then ModificacionBasica = True
            If .Fields("tipo_iva") <> Miarticulo.IVA Then ModificacionBasica = True
            If .Fields("Pref") <> Miarticulo.Preferente Then ModificacionBasica = True
            If .Fields("Etq") <> Miarticulo.Etiqueta Then ModificacionBasica = True
            If .Fields("codi_sub") <> Miarticulo.subsec Then ModificacionBasica = True
            If .Fields("codi_fam") <> Miarticulo.familia Then ModificacionBasica = True
            If .Fields("fcb") <> Miarticulo.Merma Then ModificacionBasica = True
            If .Fields("ning") <> Miarticulo.Presel Then ModificacionBasica = True
            If .Fields("lmix") <> Miarticulo.lMix Then ModificacionBasica = True
            If .Fields("poid") <> Miarticulo.nPoid Then ModificacionBasica = True
            If .Fields("onkey") And Miarticulo.onkey = 0 Then ModificacionBasica = True
            If .Fields("onkey") = False And Miarticulo.onkey = 1 Then ModificacionBasica = True
            
#If BALSAM = True Then
            For bucle = 0 To 8
                If .Fields("tar" & CStr(bucle + 1)) <> Miarticulo.tarifa(bucle) Then ModificacionBasica = True
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#Else
            For bucle = 0 To 9
                If .Fields("tar" & CStr(bucle + 1)) <> Miarticulo.tarifa(bucle) Then ModificacionBasica = True
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#End If
            If Not UsaEuro Then
                If .Fields("precio") <> MiPrecio Then
                    ModificacionBasica = True
                    'ModificacionPrecio = True
                    SeCambioPrecio = True
                End If
            Else
                If .Fields("Euros") <> MiEuro Then
                    ModificacionBasica = True
                    'ModificacionPrecio = True
                    SeCambioPrecio = True
                End If
            End If
            If Trim(.Fields("des_plu1")) <> Trim(Miarticulo.Descriptivos(0)) Or (Len(.Fields("des_plu1")) <> Len(Miarticulo.Descriptivos(0))) Then ModificacionBasica = True
            
            If .Fields("caducidad") <> Miarticulo.caducidad Then ModificacionBasica = True
            If .Fields("tara") <> Miarticulo.tara Then ModificacionBasica = True
            If .Fields("art_cb") <> Miarticulo.ean13 Then ModificacionBasica = True 'ModificacionCodBar = True
            If (Miarticulo.WGH = False And .Fields("codi_pes") = "U") Or _
            (Miarticulo.WGH = True And .Fields("codi_pes") = "W") Then ModificacionBasica = True
            If (Miarticulo.Balenv = False And .Fields("Balenv") = True) Or _
            (Miarticulo.Balenv = True And .Fields("Balenv") = False) Then ModificacionBasica = True
            For Buclelite = 0 To 10
                Select Case Buclelite
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(Buclelite + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                End Select
                If (.Fields(MiNombre) <> Miarticulo.Descriptivos(Buclelite)) Or _
                (.Fields(MiNombre) = "" And Miarticulo.Descriptivos(Buclelite) <> "") Or _
                (.Fields(MiNombre) <> "" And Miarticulo.Descriptivos(Buclelite) = "") _
                Then ModificacionBasica = True 'Modificaciontexto = True
                If Buclelite < 10 Then
                    MiNombre = "des_plu" & Val(Buclelite + 11)
                    If (.Fields(MiNombre) <> Miarticulo.Descriptivos(Buclelite + 11)) Or _
                    (.Fields(MiNombre) = "" And Miarticulo.Descriptivos(Buclelite + 11) <> "") Or _
                    (.Fields(MiNombre) <> "" And Miarticulo.Descriptivos(Buclelite + 11) = "") _
                    Then ModificacionBasica = True 'ModificacionTx1 = True
                End If
                If Buclelite < 10 Then
                    If Buclelite < 9 Then
                        MiNombre = "tip_let" & Val(Buclelite + 1)
                    Else
                        MiNombre = "tip_let0"
                    End If
                    If (.Fields(MiNombre) <> Miarticulo.TipoLetra(Buclelite)) _
                    Then ModificacionBasica = True 'Modificaciontexto = True
                End If
            Next Buclelite
            
            'If sQueNombreImagen <> "" Then
            If Trim(Miarticulo.Imagen) <> .Fields("imagen") Then
                ModificacionBasica = True
                sQueNombreImagen = Trim(Miarticulo.Imagen)
                'If Mid(sQueNombreImagen, 1, 1) <> "#" Then
                    'If Right(sQueNombreImagen, 1) = "\" Then
                    '    .Fields("imagen") = "#FF00C7B4"
                    'Else
                        .Fields("imagen") = Trim(Miarticulo.Imagen) 'sQuePathImagen & sQueNombreImagen
                    'End If
                'Else
                '    .Fields("imagen") = Miarticulo.Imagen 'sQueNombreImagen
                '    If Dir(App.Path & "\images\items\" & CStr(Miarticulo.codigo)) <> "" Then
                '        Kill App.Path & "\images\items\" & CStr(Miarticulo.codigo)
                '    End If
                'End If
            'Else
            '    .Fields("imagen") = "#FF646464"
            End If
            
            If ModificacionBasica Then
                .Fields("tran_EN") = ""
                .Fields("tran_EL") = ""
                '.Fields("tran_ep") = Space(60)
                .Fields("tran_plu") = ""
                .Fields("tran_pluSC10") = ""
                .Fields("tran_texto") = ""
                .Fields("tran_textosc10") = ""
                .Fields("tran_tx1") = ""
                .Fields("tran_tx1SC10") = ""
                .Fields("tran_cb") = ""
                .Fields("prc3") = Format(Now, "yymmddhhmm")
            End If
            'If Modificaciontexto Then
            '    .Fields("tran_EN") = ""
            '    .Fields("tran_EL") = ""
            '    '.Fields("tran_ep") = Space(60)
            '    .Fields("tran_plu") = ""
            '
            '    .Fields("tran_textoel") = ""
            '    .Fields("tran_texto") = ""
            '    .Fields("tran_textosc10") = ""
            '    .Fields("prc3") = Format(Now, "yymmddhhmm")
            'End If
            'If ModificacionTx1 Then
            '    .Fields("tran_tx1") = ""
            '    .Fields("tran_tx1SC10") = ""
            '    .Fields("prc3") = Format(Now, "yymmddhhmm")
            'End If
            'If ModificacionCodBar Then
            '    .Fields("tran_cb") = ""
            '    .Fields("tran_plu") = ""
            '    .Fields("tran_EL") = ""
            '    '.Fields("tran_ep") = Space(60)
            '
            '    .Fields("prc3") = Format(Now, "yymmddhhmm")
            '    If (UsaEcoPrint And TipoEcoPrint = 3) Or lMarteAlfa Then
            '        .Fields("tran_plu") = ""
            '    End If
            'End If
            .Fields("prc100g") = Miarticulo.PRC100G
            .Fields("art_cb") = Miarticulo.ean13
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codi_fam") = Miarticulo.familia
            If Not UsaEuro Then
                .Fields("precio") = MiPrecio
                'If ModificacionPrecio Or IsNull(.Fields("Euros")) Then
                .Fields("Euros") = MiEuro
            Else
                .Fields("Euros") = MiEuro
                'If ModificacionPrecio Or IsNull(.Fields("precio")) Then
                .Fields("precio") = MiPrecio
            End If
          
            .Fields("tipo_iva") = Miarticulo.IVA
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            .Fields("tara_envasado") = Miarticulo.tara
            .Fields("fcb") = Miarticulo.Merma
            .Fields("ning") = Miarticulo.Presel
            .Fields("Pref") = Miarticulo.Preferente
            .Fields("label2") = Trim(Miarticulo.Label2)
#If BALSAM = True Then
            .Fields("tran_ep") = Trim(Miarticulo.Label1)
            .Fields("Etq") = 0
            
            .Fields("art_cb14") = Trim(Miarticulo.ean14)
#Else
            .Fields("Etq") = Miarticulo.Etiqueta
#End If
            .Fields("grupo_conserv") = Miarticulo.GrpConserv
            
            .Fields("lmix") = Miarticulo.lMix
            .Fields("poid") = Miarticulo.nPoid
            .Fields("posicion") = Miarticulo.posicion
            
            If Miarticulo.WGH = False Then
                .Fields("codi_pes") = "W"
            Else
                .Fields("codi_pes") = "U"
            End If
            If Miarticulo.Balenv = False Then
                .Fields("Balenv") = False
                Set Reg2 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
                Set Reg5 = Base.OpenRecordset _
                ("select count(codigo) from articulo where borrado=false and secc_maqui=" & _
                Reg2.Fields("secc_Maqui") & " and balenv=true")
                
                If Reg5.Fields(0) = 1 Then
                    Reg5.Close
                    Set Reg5 = Base.OpenRecordset _
                    ("select * from seccion where borrado=false and secc_maqui=" & _
                    Reg2.Fields("secc_Maqui"))
                    With Reg5
                        If Not .EOF Then
                            .MoveFirst
                            Do Until .EOF
                                If .Fields("multiple") = 1 Then
                                    If Edit_Record(Reg5) Then
                                        .Fields("multiple") = -1
                                        If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                        If .Fields("enviardatos") = "TODOS" Then .Fields("tran_sec") = ""
                                        .Update
                                    End If
                                End If
                                If Not .EOF Then .Movenext
                            Loop
                        End If
                    End With
                End If
                            
            Else
                .Fields("Balenv") = True
                Set Reg2 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
                Set Reg5 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & Reg2.Fields("secc_Maqui"))
                With Reg5
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                        If .Fields("enviardatos") <> "TODOS" Then
                            If Edit_Record(Reg5) Then
                                If .Fields("multiple") <> 1 Then
                            
                                    If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                    If .Fields("enviardatos") = "SC10" Then .Fields("tran_sec") = ""
                                End If
                                .Fields("multiple") = 1
                                .Update
                            End If
                        End If
                        If Not .EOF Then .Movenext
                        Loop
                    End If
                End With
            End If
            For bucle = 0 To 20
                Select Case bucle
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(bucle + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                    Case Else
                        MiNombre = "des_plu" & Val(bucle)
                End Select
                .Fields(MiNombre) = Miarticulo.Descriptivos(bucle)
                If bucle < 10 Then
                    Select Case bucle
                        Case 0 To 8
                            MiNombre = "tip_let" & Val(bucle + 1)
                        Case Else
                            MiNombre = "tip_let0"
                    End Select
                    .Fields(MiNombre) = Miarticulo.TipoLetra(bucle)
                End If
            Next bucle
            '*****
            ' Tramos
            '*****
            If ChkTramo And .Fields("usatramos") = False Then
                .Fields("usatramos") = True
                .Fields("tran_plu") = ""
                ModificacionTramo = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=true")
                With RegTramo
                    If .EOF Then
                        .AddNew
                    Else
                        If Not Edit_Record(RegTramo) Then
                            .AddNew
                        End If
                    End If
                        .Fields("borrado") = False
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If Miarticulo.TRM1 > 0 And Miarticulo.PRC1 > 0 Then
                            .Fields("tramo1") = Miarticulo.TRM1
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC1)
                            Else
                                .Fields("euros1") = EUROtoPTA(Miarticulo.PRC1)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                                If Miarticulo.TRM2 > 0 And Miarticulo.PRC2 > 0 Then
                                    .Fields("tramo2") = Miarticulo.TRM2
                                    If UsaEuro Then
                                        .Fields("euros2") = Miarticulo.PRC2
                                        .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                    Else
                                        .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                        .Fields("precio2") = Miarticulo.PRC2
                                    End If
                                Else
                                    .Fields("tramo2") = 0
                                    .Fields("precio2") = 0
                                    .Fields("euros2") = 0
                                End If
                        End If
                    .Update
                End With
            End If
            If ChkTramo And .Fields("usatramos") = True Then
                .Fields("usatramos") = True
                Set RegTramo = Base.OpenRecordset _
                ("select * from tramos where borrado=false and codigo=" & .Fields("codigo"))
                With RegTramo
                    If Not .EOF Then
                        If Not Edit_Record(RegTramo) Then
                            .AddNew
                        End If
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If .Fields("precio1") <> Miarticulo.PRC1 Or .Fields("tramo1") <> Miarticulo.TRM1 Then
                            
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC2)
                            Else
                                .Fields("euros1") = PTAtoEURO(Miarticulo.PRC2)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                            .Fields("tramo1") = Miarticulo.TRM1
                            RegArticulo.Fields("tran_plu") = ""
                            ModificacionTramo = True
                        End If
                        If (Miarticulo.PRC2 <= 0 And .Fields("precio2") >= 0) Then
                            .Fields("euros2") = 0
                            .Fields("precio2") = 0
                            .Fields("tramo2") = 0
                            RegArticulo.Fields("tran_plu") = ""
                            ModificacionTramo = True
                        Else
                            If Miarticulo.PRC2 > 0 And _
                            (.Fields("precio2") <> Miarticulo.PRC2 Or .Fields("tramo2") <> Miarticulo.TRM2) Then
                                If UsaEuro Then
                                    .Fields("euros2") = Miarticulo.PRC2
                                    .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                Else
                                    .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                    .Fields("precio2") = Miarticulo.PRC2
                                End If
                                .Fields("tramo2") = Miarticulo.TRM2
                                RegArticulo.Fields("tran_plu") = ""
                                ModificacionTramo = True
                            End If
                        End If
                    .Update
                    End If
                End With
            End If
            If ChkTramo = False And .Fields("usatramos") = True Then
                .Fields("usatramos") = False
                .Fields("tran_plu") = ""
                Set RegTramo = Base.OpenRecordset _
                ("select * from tramos where borrado=false and codigo=" & .Fields("codigo"))
                With RegTramo
                    If Not .EOF Then
                        If Edit_Record(RegTramo) Then
                            .Fields("borrado") = True
                            .Update
                        End If
                        ModificacionTramo = True
                    End If
                End With
            End If
            .Update
            'If Miarticulo.subsec <> "0" Then
            '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec
            '    With RegSubsec
            '    If .EOF Or .NoMatch Then
            '        .AddNew
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        '.fields("secc_maqui = Reg2.fields("secc_maqui
            '        .Fields("borrado") = False
            '        .Fields("tran_sub") = ""
            '        .Update
            '    End If
            '    End With
            '
            'End If
            'If Miarticulo.familia <> "0" Then
            '    Set Reg3 = Base.OpenRecordset _
            '    ("select * from familia where borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec & _
            '    " and codi_fam=" & Miarticulo.familia)
            '    With Reg3
            '    If .EOF Then
            '        .AddNew
            '        .Fields("codi_fam") = Miarticulo.familia
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        .Fields("secc_Maqui") = Reg2.Fields("secc_Maqui")
            '        .Fields("borrado") = False
            '        .Fields("tran_fam") = ""
            '        .Update
            '    End If
            '    End With
            '
            '
            'End If
      End If
    End With
    'If ModificacionTramo Or ModificacionBasica Or Modificaciontexto Or _
    'ModificacionTx1 Or ModificacionCodBar Or ModificacionPrecio Then
    If ModificacionBasica Or ModificacionTramo Then
        If Not SoyAlta Then Retorno = 1
    Else
        Retorno = 4
    End If
fin:
If LogArticulos Then
    If Retorno = 1 And SeCambioPrecio Then
        Set RstRechazos = Base.OpenRecordset("select * from log")
        With RstRechazos
            .AddNew
            If UsaEuro Then
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(Miarticulo.codigo, "000000") & "," & _
                "000" & "," & Format(Miarticulo.Mostrador, "000") & "," & Format(Miarticulo.precio * 100, "0000000") & "," & Trim(Miarticulo.Descriptivos(0)) & "," & Format(RegSeccion.Fields("secc_Maqui"), "00")
            Else
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(Miarticulo.codigo, "000000") & "," & _
                "000" & "," & Format(Miarticulo.Mostrador, "000") & "," & Format(Miarticulo.precio * 10 ^ decimales, "0000000") & "," & Trim(Miarticulo.Descriptivos(0)) & "," & Format(RegSeccion.Fields("secc_Maqui"), "00")
            End If
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codigo") = Miarticulo.codigo
            .Fields("plu") = PluBuffer
            .Fields("precio") = Miarticulo.precio
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            
            '.Fields("tara_envasado") = Miarticulo.tara
            
            .Fields("des_plu1") = Miarticulo.Descriptivos(0)
            .Fields("Fecha") = Date
            If Miarticulo.WGH = True Then
                .Fields("tipo_ven") = "W"
            Else
                .Fields("tipo_ven") = "U"
            End If
            .Fields("Comentario") = CargaCadena(24)
            .Update
        End With
    End If
          
    If (Retorno <> 0 And Retorno <> 4 And Retorno <> 1) Then
        If TypeName(Base) = "Nothing" Then
            Set Base = OpenDatabase(Base_General)
            CierraDespues = True
        End If
        'Set Base = OpenDatabase(Base_General)
        On Error Resume Next
        Set RstRechazos = Base.OpenRecordset("select * from rechazos")
        If Err.Number <> 0 Then
            
            Set Base = OpenDatabase(Base_General)
            Set RstRechazos = Base.OpenRecordset("select * from rechazos")
        End If
        On Error GoTo 0
        
        With RstRechazos
            .AddNew
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codigo") = Miarticulo.codigo
            .Fields("plu") = Miarticulo.Plu
            If UsaEuro Then
                .Fields("precio") = .Fields("precio") / 100
            Else
                .Fields("precio") = .Fields("precio") / (10 ^ decimales)
            End If
            .Fields("tipo_ven") = Miarticulo.WGH
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            .Fields("tara_envasado") = Miarticulo.tara
            .Fields("des_plu1") = Miarticulo.Descriptivos(0)
            If Retorno <> 1 Then
                Select Case Retorno
                Case 2
                    .Fields("Comentario") = CargaCadena(903)  ' "Intento de cambio de PLU del artículo"
                Case 3
                    .Fields("Comentario") = CargaCadena(902)  ' "No existe sección"
                Case 5
                    .Fields("Comentario") = CargaCadena(901)  ' "Código de Barras con menos de 12 caracteres"
                Case 14
                    .Fields("Comentario") = CargaCadena(900)  ' "Falta código de Artículo"
                Case 15
                    .Fields("Comentario") = CargaCadena(899)  ' "Falta número de Mostrador"
                Case 16
                    .Fields("Comentario") = CargaCadena(898)  ' "Error en precios por tramos"
                Case 17
                    .Fields("Comentario") = CargaCadena(897)  ' "Base de datos de Tramos llena"
                Case 21
                    .Fields("Comentario") = CargaCadena(896)  ' PLU repetido
                End Select
                .Update
            Else
                .Fields("Comentario") = CargaCadena(24)
                
            End If
        End With
    End If
End If
    If CierraDespues Then CerrarBase Base
    Alta_Articulo = Retorno
End Function
Public Function Alta_Articulo_Mod(Miarticulo As DB_Articulo, Optional Base As dao.Database, Optional RegArticulo As dao.Recordset) As Integer
    '//////////////////////////////////
    ' en esta función, los valores a cero
    ' o los descriptivos en blanco
    ' no se modifican
    '//////////////////////////////////
    Dim ModificacionTramo As Boolean
    Dim ModificacionBasica As Boolean
    Dim Modificaciontexto As Boolean
    Dim ModificacionTx1 As Boolean
    Dim ModificacionCodBar As Boolean
    Dim ModificacionPrecio As Boolean
    Dim bucle As Integer
    Dim Buclelite As Integer
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Reg3 As dao.Recordset
    Dim Reg4 As dao.Recordset
    Dim Reg5 As dao.Recordset
    Dim RegTramo As dao.Recordset
    Dim Retorno As Integer
    Dim RegSeccion As dao.Recordset
    Dim RegSubsec As dao.Recordset
    Dim RstRechazos As dao.Recordset
    Dim RegPlu As dao.Recordset
    Dim ChkTramo As Boolean
    Dim MiPrecio As Double
    Dim MiEuro As Double
    Dim CierraDespues As Boolean
    Dim SeCambioPrecio As Boolean
    Dim PluBuffer As Integer
    Dim MiNombre As String
    Dim SoyAlta As Boolean
    Dim RegFam As New RecordNet
    '1.7.18
    Dim lEsDif As Boolean
    '''''''
    CadenadeLog "Entrada en Alta_artículo_mod"
    CadenadeLog "Código:" & CStr(Miarticulo.codigo) & " Familia:" & CStr(Miarticulo.familia) & " Mostrador:" & CStr(Miarticulo.Mostrador)
    'If Miarticulo.codigo = 98818 Then
    '    lEsDif = lEsDif
    'End If
    'If Miarticulo.codigo = 11991 Then
    '    MiEuro = MiEuro
    'End If
    'If Dir(App.Path & "\mixmode.txt") = "" And lVL000 = False Then
    '    Miarticulo.lMix = False
    '    Miarticulo.nPoid = 0
    'Else
    '    If IsNull(Miarticulo.nPoid) Then
    '        Miarticulo.lMix = False
    '        Miarticulo.nPoid = 0
    '    End If
    'End If
    SeCambioPrecio = False
    'If FamiliasFijas Then
    '    RegFam.OpenRecordset ("select * from fam_code where codi_fam=" & Miarticulo.familia & " and codi_ident=" & Miarticulo.Mostrador)
    '    If RegFam.EOF Then
    '        Retorno = 3
    '        GoTo FIN
    '    End If
    'End If
    If (Miarticulo.TRM1 <> 0 And Miarticulo.TRM1 <> -1) Or _
       (Miarticulo.TRM2 <> 0 And Miarticulo.TRM2 <> -1) Or _
       (Miarticulo.PRC1 <> 0 And Miarticulo.PRC1 <> -1) Or _
       (Miarticulo.PRC2 <> 0 And Miarticulo.PRC2 <> -1) Then
        ChkTramo = True
    Else
        ChkTramo = False
    End If
    'If Trim(Miarticulo.EAN13) <> "" Then
    '    If Len(Miarticulo.EAN13) <> 12 Then
    '        Retorno = 5
    '        GoTo FIN
    '    End If
    'End If
    If Miarticulo.codigo = 0 Then
        Retorno = 14
        GoTo fin
    End If
    If Miarticulo.Mostrador = 0 Then
        Retorno = 15
        GoTo fin
    End If
    If ChkTramo = True Then
        If Miarticulo.TRM1 = -1 And Miarticulo.PRC1 = -1 Then
            Retorno = 16
            GoTo fin
        End If
        If (Miarticulo.TRM2 <> -1 And Miarticulo.PRC2 = -1) Or _
        (Miarticulo.PRC2 <> -1 And Miarticulo.TRM2 = -1) Then
            Retorno = 16
            GoTo fin
        End If
        If (Miarticulo.TRM1 <> -1 And Miarticulo.PRC1 = -1) Or _
        (Miarticulo.PRC1 <> -1 And Miarticulo.TRM1 = -1) Then
            Retorno = 16
            GoTo fin
        End If
    End If
    '*****
    ' Precio / Euros
    '*****
    If UsaEuro Then
        MiEuro = Miarticulo.precio
        MiPrecio = EUROtoPTA(Miarticulo.precio)
    Else
        MiPrecio = Miarticulo.precio
        MiEuro = PTAtoEURO(Miarticulo.precio)
    End If
    
    '1.7.19
    On Error Resume Next
    Set Base = OpenDatabase(Base_General)
    CierraDespues = True
    On Error GoTo 0
    
    'If TypeName(Base) = "Nothing" Then
    '    Set Base = OpenDatabase(Base_General)
    '    CierraDespues = True
    'End If
    'If TypeName(RegArticulo) = "Nothing" Then
        Set RegArticulo = Base.OpenRecordset("select * from articulo")
    'End If
    'If TypeName(RegSeccion) = "Nothing" Then
        Set RegSeccion = Base.OpenRecordset("select * from seccion")
    'End If
    'If TypeName(RegSubsec) = "Nothing" Then
    '    Set RegSubsec = Base.OpenRecordset("select * from subsec")
    'End If
    'If ChkTramo Then
    '    Set RegTramo = Base.OpenRecordset("select count(codigo) from tramos")
    '    With RegTramo
    '        If .Fields(0) >= 400 Then
    '            Retorno = 17
    '
    '            GoTo fin
    '        End If
    '    End With
    'End If
    Retorno = 4
    RegSeccion.FindFirst "borrado=false and codi_ident=" & Miarticulo.Mostrador
    
    If RegSeccion.EOF Or RegSeccion.NoMatch Then
        Retorno = 3
        GoTo fin
    End If
    '1.7.18
    lEsDif = False
    If RegSeccion.Fields("codi_ident") <> RegSeccion.Fields("secc_maqui") Then
       lEsDif = True
    End If
    '''''''

    RegArticulo.FindFirst "borrado=false and codigo=" & Miarticulo.codigo
    
    With RegArticulo
    '*******************
        ' Nuevo dato
        If .EOF Or .NoMatch Then
            '********************
            ' comprueba plu repetido
            '********************
            '1.7.18...
            ' Para Grupo Upper si el Plu que llega ya existe dentro
            'de la Sección máquina y el código se dá de alta el código
            'con el PLU buscando un "Hueco".
            '1.8.3
            'If lUpper Then
            '    Miarticulo.familia = 255
            'End If
            '''''''
            If Miarticulo.Plu <> 0 Then
                Set RegPlu = Base.OpenRecordset _
                ("select plu from articulo where secc_maqui=" & _
                RegSeccion.Fields("secc_Maqui") & " and plu=" & Miarticulo.Plu & " and borrado=false")
                If Not RegPlu.EOF Then
                    If lUpper = False And lAlba = False Then '
                        Retorno = 21
                        GoTo fin
                    Else
                        'Grupo Upper
                        Set Reg3 = Base.OpenRecordset _
                        ("select max(plu) from articulo where borrado=false and secc_maqui=" & RegSeccion.Fields("secc_Maqui"))
                        If Not IsNull(Reg3.Fields(0)) And Not Reg3.EOF Then
                            If Reg3.Fields(0) < 9999 Then
                                Miarticulo.Plu = Reg3.Fields(0) + 1
                            Else
                                For bucle = 1 To 9999
                                     
                                    RegArticulo.Filter = "borrado = False And secc_maqui = " & RegSeccion.Fields("secc_Maqui") & " and plu=" & bucle
                                    
                                    Set Reg4 = RegArticulo.OpenRecordset
                                    If Reg4.EOF Then
                                        Miarticulo.Plu = bucle
                                        Exit For
                                    End If
                                Next bucle
                            End If
                            '''''''''
                        Else
                            'si no hay hueco.
                            Retorno = 21
                            GoTo fin
                            '''''''''''''''''
                        End If
                    
                    End If
                End If
            End If
            '********************
            .AddNew
            SoyAlta = True
            Retorno = 0
            ModificacionBasica = True
            .Fields("onkey") = True
            If Miarticulo.onkey = 0 Then
                .Fields("onkey") = False
            End If
            .Fields("lmix") = Miarticulo.lMix
            .Fields("poid") = Miarticulo.nPoid
#If BALSAM = True Then
            For bucle = 0 To 8
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#Else
            For bucle = 0 To 9
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#End If
            bucle = 0
            
            .Fields("prc100g") = Miarticulo.PRC100G
            .Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
            .Fields("tran_EN") = ""
            .Fields("tran_EL") = ""
            '.Fields("tran_ep") = Space(60)
            .Fields("tran_textoel") = "*"
            
            '.Fields("tran_plu") = ""
            '.Fields("tran_pluSC10") = ""
            '.Fields("tran_texto") = ""
            '.Fields("tran_textosc10") = ""
            '.Fields("tran_tx1") = ""
            '.Fields("tran_tx1SC10") = ""
            '.Fields("tran_cb") = ""
            '.Fields("tran_cbsc10") = ""
            
            .Fields("tipo_iva") = Miarticulo.IVA
            .Fields("codigo") = Miarticulo.codigo
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("Pref") = Miarticulo.Preferente
            If Miarticulo.Etiqueta = 255 Then Miarticulo.Etiqueta = 0
            .Fields("Etq") = Miarticulo.Etiqueta
            .Fields("grupo_conserv") = Miarticulo.GrpConserv
            .Fields("fcb") = Miarticulo.Merma
            .Fields("ning") = Miarticulo.Presel
            RegSeccion.FindFirst "borrado=false and codi_ident=" & Miarticulo.Mostrador
            .Fields("secc_Maqui") = RegSeccion.Fields("secc_Maqui")
            If Miarticulo.Plu = 0 And lFornes = False Then
                Set Reg3 = Base.OpenRecordset _
                ("select max(plu) from articulo where borrado=false and secc_maqui=" & RegSeccion.Fields("secc_Maqui"))
                If Not IsNull(Reg3.Fields(0)) And Not Reg3.EOF Then
                    If Reg3.Fields(0) < 9999 Then
                        .Fields("plu") = Reg3.Fields(0) + 1
                        PluBuffer = Reg3.Fields(0) + 1
                    Else
                        
                        For bucle = 1 To 9999
                             
                            RegArticulo.Filter = "borrado = False And secc_maqui = " & RegSeccion.Fields("secc_Maqui") & " and plu=" & bucle
                            
                            Set Reg4 = RegArticulo.OpenRecordset
                            If Reg4.EOF Then
                                .Fields("plu") = bucle
                                PluBuffer = bucle
                                Exit For
                            End If
                        Next bucle
                    End If
                Else
                    .Fields("plu") = 1
                    PluBuffer = 1
                End If
            Else
                
                .Fields("plu") = Miarticulo.Plu
                PluBuffer = Miarticulo.Plu
            End If
            If Trim(Miarticulo.ean13) <> "" Then
                .Fields("art_cb") = Miarticulo.ean13
            End If
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("precio") = MiPrecio
            .Fields("Euros") = MiEuro
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            .Fields("tara_envasado") = Miarticulo.tara
            If Trim(Miarticulo.Imagen) <> "" Then
                .Fields("imagen") = Trim(Miarticulo.Imagen)
            End If
            If Miarticulo.WGH = False Then
                .Fields("codi_pes") = "W"
            Else
                .Fields("codi_pes") = "U"
            End If
            If Miarticulo.Balenv = False Then
                .Fields("Balenv") = False
            Else
                .Fields("Balenv") = True
                Set Reg5 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & RegSeccion.Fields("secc_Maqui"))
                With Reg5
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                        If .Fields("enviardatos") <> "TODOS" Then
                            If Edit_Record(Reg5) Then
                                If .Fields("multiple") <> 1 Then
                                    If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                    If .Fields("enviardatos") = "SC10" Then .Fields("tran_sec") = ""
                                End If
                                .Fields("multiple") = 1
                                .Update
                            End If
                        End If
                        If Not .EOF Then .Movenext
                        Loop
                    End If
                End With
            End If
            For bucle = 0 To 20
                Select Case bucle
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(bucle + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                    Case 11 To 20
                        MiNombre = "des_plu" & Val(bucle)
                End Select
                .Fields(MiNombre) = Miarticulo.Descriptivos(bucle)
                If bucle < 10 Then
                    If bucle < 9 Then
                        MiNombre = "tip_let" & Val(bucle + 1)
                    Else
                        MiNombre = "tip_let0"
                    End If
                    .Fields(MiNombre) = Miarticulo.TipoLetra(bucle)
                End If
            Next bucle
            '*****
            ' Tramos
            '*****
            If ChkTramo Then
                .Fields("usatramos") = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=false")
                With RegTramo
                    If .EOF Then
                        .AddNew
                    Else
                        If Not Edit_Record(RegTramo) Then .AddNew
                    End If
                        .Fields("borrado") = False
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If Miarticulo.TRM1 <> -1 And Miarticulo.PRC1 <> -1 Then
                            
                            .Fields("tramo1") = Miarticulo.TRM1
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC1)
                            Else
                                .Fields("euros1") = PTAtoEURO(Miarticulo.PRC1)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                                If Miarticulo.TRM2 <> -1 And Miarticulo.PRC2 <> -1 Then
                                    .Fields("tramo2") = Miarticulo.TRM2
                                    If UsaEuro Then
                                        .Fields("euros2") = Miarticulo.PRC2
                                        .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                    Else
                                        .Fields("precio2") = Miarticulo.PRC2
                                        .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                    End If
                                Else
                                    .Fields("tramo2") = -1
                                    .Fields("precio2") = -1
                                    .Fields("euros2") = -1
                                End If
                        End If
                    .Update
                End With
            Else
                .Fields("usatramos") = False
            End If
                        
            .Update
            'If Miarticulo.Mostrador = 6 Then
            '
            '    Miarticulo.Mostrador = 6
            'End If
            'If Miarticulo.subsec <> "0" Then
            '    'Set RegSubsec = Base.OpenRecordset("select * from subsec where borrado=false")
            '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec
            '    With RegSubsec
            '    If RegSubsec.EOF Or RegSubsec.NoMatch Then
            '        .AddNew
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        '.fields("secc_maqui = Reg2.fields("secc_maqui
            '        .Fields("borrado") = False
            '        .Fields("tran_sub") = ""
            '        .Update
            '    End If
            '    End With
            '
            'End If
            
            
            'If Miarticulo.familia <> "0" And Not (lUpperNW) Then
            '1.8.3
            'If Miarticulo.familia <> "0" And Not (lUpper) Then
            ''''''''
            
            If Miarticulo.familia <> "0" Then
                'Set Reg3 = Base.OpenRecordset _
                '("select * from fam_code where codi_ident=" & _
                'Miarticulo.Mostrador & " and codi_fam=" & Miarticulo.familia)
                
                Set Reg3 = Base.OpenRecordset _
                ("select * from fam_code where codi_fam=" & Miarticulo.familia)
                With Reg3
                    If .EOF Then
                        .AddNew
                        .Fields("codi_fam") = Miarticulo.familia
                        '.Fields("codi_sub") = Miarticulo.subsec
                        .Fields("codi_ident") = 0 'Miarticulo.Mostrador
                        .Fields("secc_Maqui") = 0 'RegSeccion.Fields("secc_Maqui")
                        .Fields("txt_fam") = "FAM. " & CStr(Miarticulo.familia)
                        .Fields("imagen") = ""
                        '.Fields("borrado") = False
                        '.Fields("tran_fam") = ""
                        .Update
                    End If
                End With
            End If
'DROP TABLE IF EXISTS `pcscale`.`families`;
'CREATE TABLE  `pcscale`.`families` (
'  `Code` int(11) NOT NULL,
'  `Name` varchar(54) NOT NULL,
'  `Icon` varchar(127) DEFAULT NULL,
'  `IsBovine` tinyint(1) DEFAULT '0',
'  `Label` varchar(64) NOT NULL,
'  PRIMARY KEY (`Code`),
'  KEY `Name` (`Name`(19))
') ENGINE=InnoDB DEFAULT CHARSET=utf8;
        Else
            '************************
            ' Modificación
            '************************
            'If Miarticulo.plu <> 0 Then
            '    If Dir(App.Path & "\upper.txt") = "" Then
            '        If .Fields("plu") <> Miarticulo.plu Then
            '            Retorno = 2
            '            GoTo FIN
            '        End If
            '    Else '1.7.18 Grupo Upper
            '        If lEsDif Then 'si codi_ident != secc_maqui
            '                       'se mantendrá PLU existente, si no se machaca.Más abajo...
            '            Miarticulo.plu = .Fields("plu")
            '        End If
            '    End If
            'End If
            'c2f ojito...
            If Miarticulo.Plu <> 0 Then
                'If Dir(App.Path & "\upper.txt") = "" Then
                    If .Fields("plu") <> Miarticulo.Plu Or .Fields("secc_maqui") <> Miarticulo.Mostrador Then
                        Retorno = 45
                        GoTo fin
                    End If
                'Else '1.7.18 Grupo Upper
                '    If lEsDif Then 'si codi_ident != secc_maqui
                '                   'se mantendrá PLU existente, si no se machaca.Más abajo...
                '        Miarticulo.plu = .Fields("plu")
                '    End If
                'End If
            End If
            
            .Edit
            '.Fields("plu") = Miarticulo.Plu
            
            '''''''''''''se comenta lo que sigue
            'If Not Edit_Record(Registro) Then
            '    Base.Close
            '    Alta_Articulo_Mod = 60
            '    Exit Function
            'End If
            'Grupo Upper ... Se machaca.
            'If Dir(App.Path & "\upper.txt") <> "" And Not (lEsDif) Then
            '    .Fields("plu") = Miarticulo.plu
            'End If
            ''''''''''''
            PluBuffer = .Fields("plu")
            ModificacionBasica = False
            Modificaciontexto = False
            ModificacionTx1 = False
            ModificacionCodBar = False
            ModificacionTramo = False
            
            .Fields("lmix") = Miarticulo.lMix
            .Fields("poid") = Miarticulo.nPoid
            If Miarticulo.onkey = 0 Then
                .Fields("onkey") = False
            Else
                .Fields("onkey") = True
            End If
            
            If .Fields("grupo_conserv") <> Miarticulo.GrpConserv Then
                Modificaciontexto = True
                ModificacionTx1 = True
            End If
            If IsNull(.Fields("tipo_iva")) Then .Fields("tipo_iva") = 0
            If (.Fields("tipo_iva") <> Miarticulo.IVA) And (Miarticulo.IVA <> 0) Then ModificacionBasica = True
            If (.Fields("Pref") <> Miarticulo.Preferente) And (Miarticulo.Preferente <> 0) Then ModificacionBasica = True
            If (.Fields("Etq") <> 0 And Miarticulo.Etiqueta <> 255) Then
                If (.Fields("Etq") <> Miarticulo.Etiqueta) And (Miarticulo.Etiqueta <> 0) Then ModificacionBasica = True
            End If
            If (.Fields("codi_sub") <> Miarticulo.subsec) And (Miarticulo.subsec <> 0) Then ModificacionBasica = True
            If (.Fields("codi_fam") <> Miarticulo.familia) And (Miarticulo.familia <> 0) Then ModificacionBasica = True
            If (.Fields("fcb") <> Miarticulo.Merma) And (Miarticulo.Merma <> 0) Then ModificacionBasica = True
            If (.Fields("ning") <> Miarticulo.Presel) And (Miarticulo.Presel <> 0) Then ModificacionBasica = True
            If Not UsaEuro Then
                If .Fields("precio") <> MiPrecio Then
                    ModificacionBasica = True
                    ModificacionPrecio = True
                    SeCambioPrecio = True
                End If
            Else
                If .Fields("Euros") <> MiEuro Then
                    ModificacionBasica = True
                    ModificacionPrecio = True
                    SeCambioPrecio = True
                End If
            End If
            If .Fields("prc100g") <> Miarticulo.PRC100G Then ModificacionBasica = True
            'If Dir(App.Path & "\pluasc.dir") <> "" Then
            '    If (.Fields("caducidad") <> Miarticulo.caducidad) Then ModificacionBasica = True
            '    If (.Fields("tara") <> Miarticulo.tara) Then ModificacionBasica = True
            'Else
                If (.Fields("caducidad") <> Miarticulo.caducidad) And (Miarticulo.caducidad <> 0) Then ModificacionBasica = True
                If (.Fields("tara") <> Miarticulo.tara) And (Miarticulo.tara <> 0) Then ModificacionBasica = True
            'End If
            If Trim(.Fields("art_cb")) <> Miarticulo.ean13 Then ModificacionCodBar = True
            If (Miarticulo.WGH = False And .Fields("codi_pes") = "U") Or _
            (Miarticulo.WGH = True And .Fields("codi_pes") = "W") Then ModificacionBasica = True
            If (Miarticulo.Balenv = False And .Fields("Balenv") = True) Or _
            (Miarticulo.Balenv = True And .Fields("Balenv") = False) Then ModificacionBasica = True
            If .Fields("onkey") And Miarticulo.onkey = 0 Then ModificacionBasica = True
            If .Fields("onkey") = False And Miarticulo.onkey = 1 Then ModificacionBasica = True
            If Trim(.Fields("des_plu1")) <> Trim(Miarticulo.Descriptivos(0)) Or (Len(.Fields("des_plu1")) <> Len(Miarticulo.Descriptivos(0))) Then ModificacionBasica = True
#If BALSAM = True Then
            For bucle = 0 To 8
                If .Fields("tar" & CStr(bucle + 1)) <> Miarticulo.tarifa(bucle) Then ModificacionBasica = True
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle

#Else
            For bucle = 0 To 9
                If .Fields("tar" & CStr(bucle + 1)) <> Miarticulo.tarifa(bucle) Then ModificacionBasica = True
                .Fields("tar" & CStr(bucle + 1)) = Miarticulo.tarifa(bucle)
            Next bucle
#End If
'''
            'For Bucle = 0 To 20
            '    Select Case Bucle
            '        Case 0 To 8
            '            MiNombre = "des_plu" & Val(Bucle + 1)
            '        Case 9
            '            MiNombre = "des_plu0"
            '        Case 10
            '            MiNombre = "des_plux"
            '        Case 11 To 20
            '            MiNombre = "des_plu" & Val(Bucle)
            '    End Select
            '    .Fields(MiNombre) = Miarticulo.Descriptivos(Bucle)
            '    If Bucle < 10 Then
            '        If Bucle < 9 Then
            '            MiNombre = "tip_let" & Val(Bucle + 1)
            '        Else
            '            MiNombre = "tip_let0"
            '        End If
            '        .Fields(MiNombre) = Miarticulo.TipoLetra(Bucle)
            '    End If
            'Next Bucle

'''
            For Buclelite = 0 To 20 '2.0.30 era 10
                Select Case Buclelite
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(Buclelite + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                    Case 11 To 20
                        MiNombre = "des_plu" & Val(bucle)
                End Select
                On Error Resume Next
                If Miarticulo.Descriptivos(Buclelite) <> "" Then
                    If Err.Number <> 0 Then
                        If Buclelite > 10 Then
                            On Error GoTo 0
                            Exit For
                        End If
                    End If
                    If (.Fields(MiNombre) <> Miarticulo.Descriptivos(Buclelite)) Or _
                    (.Fields(MiNombre) = "" And Miarticulo.Descriptivos(Buclelite) <> "") Or _
                    (.Fields(MiNombre) <> "" And Miarticulo.Descriptivos(Buclelite) = "") _
                    Then Modificaciontexto = True
                End If
                If Buclelite < 10 Then
                    If Miarticulo.Descriptivos(Buclelite + 11) <> "" Then
                        MiNombre = "des_plu" & Val(Buclelite + 11)
                        If (.Fields(MiNombre) <> Miarticulo.Descriptivos(Buclelite + 11)) Or _
                        (.Fields(MiNombre) = "" And Miarticulo.Descriptivos(Buclelite + 11) <> "") Or _
                        (.Fields(MiNombre) <> "" And Miarticulo.Descriptivos(Buclelite + 11) = "") _
                        Then ModificacionTx1 = True
                    End If
                End If
                If Buclelite < 10 Then
                    If Buclelite < 9 Then
                        MiNombre = "tip_let" & Val(Buclelite + 1)
                    Else
                        MiNombre = "tip_let0"
                    End If
                    If (.Fields(MiNombre) <> Miarticulo.TipoLetra(Buclelite)) And Miarticulo.TipoLetra(Buclelite) <> 0 _
                    Then Modificaciontexto = True
                End If
                On Error GoTo 0
            Next Buclelite
            If ModificacionBasica Then
                .Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
                .Fields("tran_EN") = ""
                .Fields("tran_EL") = ""
                '.Fields("tran_ep") = Space(60)
                .Fields("tran_plu") = ""
                .Fields("tran_pluSC10") = ""
            End If
            If Modificaciontexto Then
                .Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
                .Fields("tran_textoel") = ""
                .Fields("tran_texto") = ""
                .Fields("tran_textosc10") = ""
            End If
            If ModificacionTx1 Then
                .Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
                .Fields("tran_tx1") = ""
                .Fields("tran_tx1SC10") = ""
            End If
            If ModificacionCodBar Then
                .Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
                .Fields("tran_cb") = ""
                .Fields("tran_cbsc10") = ""
            End If
            
            If Miarticulo.ean13 <> "" Then .Fields("art_cb") = Miarticulo.ean13
            If Miarticulo.subsec <> 0 Then .Fields("codi_sub") = Miarticulo.subsec
            '1.8.3
            'If Not lUpper Then
                If Miarticulo.familia <> 0 Then .Fields("codi_fam") = Miarticulo.familia
            'End If
            If Not UsaEuro Then
                .Fields("precio") = MiPrecio
                If ModificacionPrecio Or IsNull(.Fields("Euros")) Then .Fields("Euros") = MiEuro
            Else
                .Fields("Euros") = MiEuro
                If ModificacionPrecio Or IsNull(.Fields("precio")) Then .Fields("precio") = MiPrecio
            End If
            
            If Miarticulo.IVA <> 0 Then .Fields("tipo_iva") = Miarticulo.IVA
            'If Dir(App.Path & "\pluasc.dir") <> "" Then
            '    .Fields("caducidad") = Miarticulo.caducidad
            '    .Fields("tara") = Miarticulo.tara
            '    .Fields("tara_envasado") = Miarticulo.tara
            'Else
                If Miarticulo.caducidad <> 0 Then .Fields("caducidad") = Miarticulo.caducidad
                If Miarticulo.tara <> 0 Then
                    .Fields("tara") = Miarticulo.tara
                    .Fields("tara_envasado") = Miarticulo.tara
                End If
            
            'End If
            If Miarticulo.Merma <> 0 Then .Fields("fcb") = Miarticulo.Merma
            If Miarticulo.Presel <> 0 Then .Fields("ning") = Miarticulo.Presel
            If Miarticulo.Preferente <> 0 Then .Fields("Pref") = Miarticulo.Preferente
            If Miarticulo.Etiqueta <> 0 Then
                If Miarticulo.Etiqueta = 255 Then Miarticulo.Etiqueta = 0
                .Fields("Etq") = Miarticulo.Etiqueta
            End If
            If Miarticulo.GrpConserv <> 0 Then .Fields("grupo_conserv") = Miarticulo.GrpConserv
            If Miarticulo.WGH = False Then
                .Fields("codi_pes") = "W"
            Else
                .Fields("codi_pes") = "U"
            End If
            .Fields("prc100g") = Miarticulo.PRC100G
            If Miarticulo.Balenv = False Then
                .Fields("Balenv") = False
                Set Reg2 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
                Set Reg5 = Base.OpenRecordset _
                ("select count(codigo) from articulo where borrado=false and secc_maqui=" & _
                Reg2.Fields("secc_Maqui") & " and balenv=true")
                
                If Reg5.Fields(0) = 1 Then
                    Reg5.Close
                    Set Reg5 = Base.OpenRecordset _
                    ("select * from seccion where borrado=false and secc_maqui=" & _
                    Reg2.Fields("secc_Maqui"))
                    With Reg5
                        If Not .EOF Then
                            .MoveFirst
                            Do Until .EOF
                                If .Fields("multiple") = 1 Then
                                    If Edit_Record(Reg5) Then
                                        .Fields("multiple") = -1
                                        If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                        If .Fields("enviardatos") = "TODOS" Then .Fields("tran_sec") = ""
                                        .Update
                                    End If
                                End If
                                If Not .EOF Then .Movenext
                            Loop
                        End If
                    End With
                End If
                            
            Else
                .Fields("Balenv") = True
                Set Reg2 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
                Set Reg5 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & Reg2.Fields("secc_Maqui"))
                With Reg5
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                        If .Fields("enviardatos") <> "TODOS" Then
                            If Edit_Record(Reg5) Then
                                If .Fields("multiple") <> 1 Then
                            
                                    If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                    If .Fields("enviardatos") = "SC10" Then .Fields("tran_sec") = ""
                                End If
                                .Fields("multiple") = 1
                                .Update
                            End If
                        End If
                        If Not .EOF Then .Movenext
                        Loop
                    End If
                End With
            End If
            For bucle = 0 To 20
                Select Case bucle
                    Case 0 To 8
                        MiNombre = "des_plu" & Val(bucle + 1)
                    Case 9
                        MiNombre = "des_plu0"
                    Case 10
                        MiNombre = "des_plux"
                    Case Else
                        MiNombre = "des_plu" & Val(bucle)
                End Select
                If Miarticulo.Descriptivos(bucle) <> "" Then .Fields(MiNombre) = Miarticulo.Descriptivos(bucle)
                If bucle < 10 Then
                    Select Case bucle
                        Case 0 To 8
                            MiNombre = "tip_let" & Val(bucle + 1)
                        Case Else
                            MiNombre = "tip_let0"
                    End Select
                    If Miarticulo.TipoLetra(bucle) <> 0 Then .Fields(MiNombre) = Miarticulo.TipoLetra(bucle)
                End If
            Next bucle
            '*****
            ' Tramos
            '*****
            If ChkTramo And .Fields("usatramos") = False Then
                .Fields("usatramos") = True
                .Fields("tran_plu") = ""
                ModificacionTramo = True
                Set RegTramo = Base.OpenRecordset("select * from tramos where borrado=true")
                With RegTramo
                    If .EOF Then
                        .AddNew
                    Else
                        If Not Edit_Record(RegTramo) Then .AddNew
                    End If
                        .Fields("borrado") = False
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If Miarticulo.TRM1 <> -1 And Miarticulo.PRC1 <> -1 Then
                            .Fields("tramo1") = Miarticulo.TRM1
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC1)
                            Else
                                .Fields("euros1") = EUROtoPTA(Miarticulo.PRC1)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                                If Miarticulo.TRM2 <> -1 And Miarticulo.PRC2 <> -1 Then
                                    .Fields("tramo2") = Miarticulo.TRM2
                                    If UsaEuro Then
                                        .Fields("euros2") = Miarticulo.PRC2
                                        .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                    Else
                                        .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                        .Fields("precio2") = Miarticulo.PRC2
                                    End If
                                Else
                                    .Fields("tramo2") = -1
                                    .Fields("precio2") = -1
                                    .Fields("euros2") = -1
                                End If
                        End If
                    .Update
                End With
            End If
            If ChkTramo And .Fields("usatramos") = True Then
                .Fields("usatramos") = True
                Set RegTramo = Base.OpenRecordset _
                ("select * from tramos where borrado=false and codigo=" & .Fields("codigo"))
                With RegTramo
                    If Not .EOF Then
                        If Not Edit_Record(RegTramo) Then .AddNew
                        .Fields("tran_tramo") = ""
                        .Fields("codigo") = Miarticulo.codigo
                        If .Fields("precio1") <> Miarticulo.PRC1 Or .Fields("tramo1") <> Miarticulo.TRM1 Then
                            
                            If UsaEuro Then
                                .Fields("euros1") = Miarticulo.PRC1
                                .Fields("precio1") = EUROtoPTA(Miarticulo.PRC2)
                            Else
                                .Fields("euros1") = PTAtoEURO(Miarticulo.PRC2)
                                .Fields("precio1") = Miarticulo.PRC1
                            End If
                            .Fields("tramo1") = Miarticulo.TRM1
                            RegArticulo.Fields("tran_plu") = ""
                            ModificacionTramo = True
                        End If
                        If (Miarticulo.PRC2 = -1 And .Fields("precio2") >= 0) Then
                            .Fields("euros2") = -1
                            .Fields("precio2") = -1
                            .Fields("tramo2") = -1
                            RegArticulo.Fields("tran_plu") = ""
                            ModificacionTramo = True
                        Else
                            If Miarticulo.PRC2 <> -1 And _
                            (.Fields("precio2") <> Miarticulo.PRC2 Or .Fields("tramo2") <> Miarticulo.TRM2) Then
                                If UsaEuro Then
                                    .Fields("euros2") = Miarticulo.PRC2
                                    .Fields("precio2") = EUROtoPTA(Miarticulo.PRC2)
                                Else
                                    .Fields("euros2") = PTAtoEURO(Miarticulo.PRC2)
                                    .Fields("precio2") = Miarticulo.PRC2
                                End If
                                .Fields("tramo2") = Miarticulo.TRM2
                                RegArticulo.Fields("tran_plu") = ""
                                ModificacionTramo = True
                            End If
                        End If
                    .Update
                    End If
                End With
            End If
            If ChkTramo = False And .Fields("usatramos") = True Then
                .Fields("usatramos") = False
                .Fields("tran_plu") = ""
                Set RegTramo = Base.OpenRecordset _
                ("select * from tramos where borrado=false and codigo=" & .Fields("codigo"))
                With RegTramo
                    If Not .EOF Then
                        If Edit_Record(RegTramo) Then
                            .Fields("borrado") = True
                            .Update
                        End If
                        ModificacionTramo = True
                    End If
                End With
            End If
            .Update
            'If Miarticulo.subsec <> "0" Then
            '    RegSubsec.FindFirst "borrado=false and codi_ident=" & _
            '    Miarticulo.Mostrador & " and codi_sub=" & Miarticulo.subsec
            '    With RegSubsec
            '    If .EOF Or .NoMatch Then
            '        .AddNew
            '        .Fields("codi_sub") = Miarticulo.subsec
            '        .Fields("codi_ident") = Miarticulo.Mostrador
            '        '.fields("secc_maqui = Reg2.fields("secc_maqui
            '        .Fields("borrado") = False
            '        .Fields("tran_sub") = ""
            '        .Update
            '    End If
             '   End With
            '
            'End If
            'If Miarticulo.familia <> "0" And Not (lUpperNW) Then
            '1.8.3
            'If Miarticulo.familia <> "0" And Not (lUpper) Then
            ''''''
            If Miarticulo.familia <> "0" Then
                Set Reg3 = Base.OpenRecordset _
                ("select * from fam_code where codi_fam=" & Miarticulo.familia)
                
                '("select * from fam_code where codi_ident=" & _
                'Miarticulo.Mostrador & " and codi_fam=" & Miarticulo.familia)
                With Reg3
                    If .EOF Then
                        .AddNew
                        .Fields("codi_fam") = Miarticulo.familia
                        '.Fields("codi_sub") = Miarticulo.subsec
                        .Fields("codi_ident") = Miarticulo.Mostrador
                        .Fields("secc_Maqui") = Reg2.Fields("secc_Maqui")
                        .Fields("txt_fam") = "FAM. " & CStr(Miarticulo.familia)
                        .Fields("imagen") = ""
                        
                        '.Fields("borrado") = False
                        '.Fields("tran_fam") = ""
                        .Update
                    End If
                End With
                
                
            End If
      End If
    End With
    
    If ModificacionTramo Or ModificacionBasica Or Modificaciontexto Or _
    ModificacionTx1 Or ModificacionCodBar Or ModificacionPrecio Then
        If Not SoyAlta Then Retorno = 1
    Else
        Retorno = 4
    End If
fin:
If LogArticulos Then
    If Retorno = 1 And SeCambioPrecio Then
        Set RstRechazos = Base.OpenRecordset("select * from log")
        With RstRechazos
            .AddNew
            If UsaEuro Then
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(Miarticulo.codigo, "000000") & "," & _
                "000" & "," & Format(Miarticulo.Mostrador, "000") & "," & Format(Miarticulo.precio * 100, "0000000") & "," & Trim(Miarticulo.Descriptivos(0)) & "," & Format(RegSeccion.Fields("secc_Maqui"), "00")
            Else
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(Miarticulo.codigo, "000000") & "," & _
                "000" & "," & Format(Miarticulo.Mostrador, "000") & "," & Format(Miarticulo.precio * 10 ^ decimales, "0000000") & "," & Trim(Miarticulo.Descriptivos(0)) & "," & Format(RegSeccion.Fields("secc_Maqui"), "00")
            End If
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codigo") = Miarticulo.codigo
            .Fields("plu") = PluBuffer
            .Fields("precio") = Miarticulo.precio
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            '.Fields("tara_envasado") = Miarticulo.tara
            .Fields("des_plu1") = Miarticulo.Descriptivos(0)
            .Fields("Fecha") = Date
            If Miarticulo.WGH = True Then
                .Fields("tipo_ven") = "W"
            Else
                .Fields("tipo_ven") = "U"
            End If
            .Fields("Comentario") = CargaCadena(24)
            .Update
        End With
    End If
          
    If (Retorno <> 0 And Retorno <> 4 And Retorno <> 1) Then
        If TypeName(Base) = "Nothing" Then
            Set Base = OpenDatabase(Base_General)
            CierraDespues = True
        End If
        Set RstRechazos = Base.OpenRecordset("select * from rechazos")
        With RstRechazos
            .AddNew
            .Fields("codi_ident") = Miarticulo.Mostrador
            .Fields("codi_sub") = Miarticulo.subsec
            .Fields("codigo") = Miarticulo.codigo
            .Fields("plu") = Miarticulo.Plu
            If UsaEuro Then
                .Fields("precio") = .Fields("precio") / 100
            Else
                .Fields("precio") = .Fields("precio") / (10 ^ decimales)
            End If
            .Fields("tipo_ven") = Miarticulo.WGH
            .Fields("codi_fam") = Miarticulo.familia
            .Fields("caducidad") = Miarticulo.caducidad
            .Fields("tara") = Miarticulo.tara
            '.Fields("tara_envasado") = Miarticulo.tara
            .Fields("des_plu1") = Miarticulo.Descriptivos(0)
            If Retorno <> 1 Then
                Select Case Retorno
                Case 2
                    .Fields("Comentario") = CargaCadena(903)  ' "Intento de cambio de PLU del artículo"
                Case 3
                    .Fields("Comentario") = CargaCadena(902)  ' "No existe sección"
                Case 5
                    .Fields("Comentario") = CargaCadena(901)  ' "Código de Barras con menos de 12 caracteres"
                Case 14
                    .Fields("Comentario") = CargaCadena(900)  ' "Falta código de Artículo"
                Case 15
                    .Fields("Comentario") = CargaCadena(899)  ' "Falta número de Mostrador"
                Case 16
                    .Fields("Comentario") = CargaCadena(898)  ' "Error en precios por tramos"
                Case 17
                    .Fields("Comentario") = CargaCadena(897)  ' "Base de datos de Tramos llena"
                Case 21
                    .Fields("Comentario") = CargaCadena(896)  ' PLU repetido
                End Select
                .Update
            Else
                .Fields("Comentario") = CargaCadena(24)
            End If
        End With
    End If
End If
    If CierraDespues Then CerrarBase Base
    
    
    CadenadeLog "Código de Retorno:" & CStr(Retorno)
    
    Alta_Articulo_Mod = Retorno
End Function
Public Function Alta_CodigoBarras(ElCodigoBarras As DB_CodigoBarras) As Integer
    Dim LaSeccion As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim bucle As Integer
    Dim BucleGrande As Integer
    Dim posicion As Integer
    Dim eNCONTRADO As Boolean
    Dim ElCodigo As TipoCodigos
    Dim Retorno As Integer
    Dim UsaSeccion As Boolean
    Dim LosMostradores As String
    Dim LosEquipos As String
    Dim MiNombre As String
    '******************
    ' Condiciones de Error
    Set Base = dao.OpenDatabase(Base_General)
    If ElCodigoBarras.Mostrador < 1000 Then
        If ElCodigoBarras.Equipo = 0 Then
            UsaSeccion = True
            Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" & _
            CStr(ElCodigoBarras.Mostrador) & " and borrado=false")
            If Not Registro.EOF Then
                Retorno = 0
                If Registro.Fields("codi_ident") < 1000 Then
                    Set Reg2 = Base.OpenRecordset("select * from seccion where secc_maqui=" & _
                    Registro.Fields("secc_Maqui") & " and borrado=false")
                    Reg2.MoveFirst
                    Do Until Reg2.EOF
                        LosMostradores = LosMostradores & Reg2.Fields("codi_ident") & " , "
                        If Not Reg2.EOF Then Reg2.Movenext
                    Loop
                    LosMostradores = left(LosMostradores, Len(LosMostradores) - 3)
                End If
            Else
                Retorno = 3
            End If
        Else
            UsaSeccion = False
            LosEquipos = ElCodigoBarras.Equipo
            Set Registro = Base.OpenRecordset("select * from equipos where borrado=false and numero_eqp=" & ElCodigoBarras.Equipo)
            If Not Registro.EOF Then ElCodigoBarras.Mostrador = Registro.Fields("codi_ident")
            Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" & _
            ElCodigoBarras.Mostrador & " and borrado=false")
            If Registro.EOF Then
                Retorno = 3
            Else
                Set Reg2 = Base.OpenRecordset("select * from equipos where numero_eqp=" & _
                ElCodigoBarras.Equipo & " and secc_Maqui=" & Registro.Fields("secc_Maqui") & " and borrado=false")
                If Reg2.EOF Then
                    Retorno = 2
                Else
                    Retorno = 0
                End If
            End If
        End If
    End If
    For bucle = 0 To 9
        If Len(ElCodigoBarras.Datos(bucle)) < 12 And Len(ElCodigoBarras.Datos(bucle)) <> 0 Then
            Retorno = 5
        End If
    Next bucle
    If Retorno <> 0 Then GoTo fin
    '********************
    If ElCodigoBarras.Mostrador < 1000 Then
        If UsaSeccion Then
            ElCodigo = Desglosa_Codigos(LosMostradores)
        Else
            ElCodigo = Desglosa_Codigos(LosEquipos)
        End If
    Else
        ElCodigo.cantidad = 1
    End If
    For BucleGrande = 1 To ElCodigo.cantidad
    
        If ElCodigoBarras.Mostrador < 1000 Then
            If UsaSeccion Then
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp<>0 and (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & ElCodigo.codigo(BucleGrande))
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp<>0 and numero_eqp=" & ElCodigoBarras.Equipo)
            End If
        Else
            Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp<>0 and  codi_ident=" & ElCodigoBarras.Mostrador)
        End If
        ' 1º en local
        Retorno = 4
        posicion = 5
        With Registro
            If Not .EOF Then
                For bucle = 0 To 4
                    Select Case bucle
                        Case 0
                            MiNombre = "cb_vnt"
                        Case 1
                            MiNombre = "cb_sup"
                        Case 2
                            MiNombre = "cb_mix"
                        Case 3
                            MiNombre = "cb_ev1"
                        Case 4
                            MiNombre = "cb_ev2"
                    End Select
                    If (ElCodigoBarras.Datos(bucle + posicion) <> "") Or .Fields(MiNombre) <> "" Then
                        '2.0.a If .Fields(MiNombre) <> ElCodigoBarras.Datos(Bucle + posicion) Then
                            Retorno = 1
                        '2.0.a End If
                    End If
                
                Next bucle
            Else
                For bucle = 0 To 4
                    If ElCodigoBarras.Datos(bucle + posicion) <> "" Then
                        Retorno = 0
                    End If
                Next bucle
            End If
            
            If Retorno = 1 Or Retorno = 0 Then
                If Not .EOF Then
                    If Not Edit_Record(Registro) Then .AddNew
                Else
                    .AddNew
                End If
                .Fields("tran_bar") = ""
                .Fields("tran_barSC10") = ""
                .Fields("tran_ep") = ""
                .Fields("tran_el") = ""
                .Fields("tran_en") = ""
                .Fields("borrado") = False
                .Fields("cb_vnt") = ElCodigoBarras.Datos(posicion)
                .Fields("cb_sup") = ElCodigoBarras.Datos(posicion + 1)
                .Fields("cb_mix") = ElCodigoBarras.Datos(posicion + 2)
                .Fields("cb_ev1") = ElCodigoBarras.Datos(posicion + 3)
                .Fields("cb_ev2") = ElCodigoBarras.Datos(posicion + 4)
                If posicion = 5 Then
                    .Fields("secc_eqp") = 1
                Else
                    .Fields("secc_eqp") = 0
                End If
                If ElCodigoBarras.Mostrador < 1000 Then
                    If UsaSeccion Then
                        Set Reg2 = Base.OpenRecordset("select * from seccion where borrado=false and codi_ident=" & ElCodigo.codigo(BucleGrande))
                        LaSeccion = Reg2.Fields("secc_Maqui")
                        Reg2.Close
                        .Fields("codi_ident") = ElCodigo.codigo(BucleGrande)
                        .Fields("secc_Maqui") = LaSeccion
                        .Fields("numero_eqp") = Null
                    Else
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and numero_eqp=" & ElCodigo.codigo(BucleGrande))
                        LaSeccion = Reg2.Fields("secc_Maqui")
                        .Fields("codi_ident") = Reg2.Fields("codi_ident")
                        Reg2.Close
                        .Fields("secc_Maqui") = LaSeccion
                        .Fields("numero_eqp") = ElCodigo.codigo(BucleGrande)
                    End If
                Else
                    .Fields("codi_ident") = ElCodigoBarras.Mostrador
                    .Fields("secc_Maqui") = ElCodigoBarras.Mostrador
                    .Fields("numero_eqp") = Null
                End If
                .Update
            End If
        End With
        ' 2º en red
        Registro.Close
        If ElCodigoBarras.Mostrador < 1000 Then
            If UsaSeccion Then
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp=0 and (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & ElCodigo.codigo(BucleGrande))
            Else
                Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp=0 and numero_eqp=" & ElCodigoBarras.Equipo)
            End If
        Else
            Set Registro = Base.OpenRecordset _
                ("select * from codbar where borrado=false and secc_eqp=0 and codi_ident=" & ElCodigoBarras.Mostrador)
        End If
        '2.0.a Retorno = 4
        posicion = 0
        With Registro
            If Not .EOF Then
                For bucle = 0 To 4
                    Select Case bucle
                        Case 0
                            MiNombre = "cb_vnt"
                        Case 1
                            MiNombre = "cb_sup"
                        Case 2
                            MiNombre = "cb_mix"
                        Case 3
                            MiNombre = "cb_ev1"
                        Case 4
                            MiNombre = "cb_ev2"
                    End Select
                    If (ElCodigoBarras.Datos(bucle + posicion) <> "") Or .Fields(MiNombre) <> "" Then
                        '2.0.a If .Fields(MiNombre) <> ElCodigoBarras.Datos(Bucle + posicion) Then
                            Retorno = 1
                        '2.0.a End If
                    End If
                Next bucle
            Else
                For bucle = 0 To 4
                    If ElCodigoBarras.Datos(bucle + posicion) <> "" Then
                        Retorno = 0
                    End If
                Next bucle
                    
            End If
            If Retorno = 1 Or Retorno = 0 Then
                If Not .EOF Then
                    If Not Edit_Record(Registro) Then .AddNew
                Else
                    .AddNew
                End If
                .Fields("tran_bar") = ""
                .Fields("tran_barSC10") = ""
                .Fields("tran_ep") = ""
                .Fields("tran_el") = ""
                .Fields("tran_en") = ""
                .Fields("borrado") = False
                .Fields("cb_vnt") = ElCodigoBarras.Datos(posicion)
                .Fields("cb_sup") = ElCodigoBarras.Datos(posicion + 1)
                .Fields("cb_mix") = ElCodigoBarras.Datos(posicion + 2)
                .Fields("cb_ev1") = ElCodigoBarras.Datos(posicion + 3)
                .Fields("cb_ev2") = ElCodigoBarras.Datos(posicion + 4)
                If posicion = 5 Then
                    .Fields("secc_eqp") = 1
                Else
                    .Fields("secc_eqp") = 0
                End If
                If ElCodigoBarras.Mostrador < 1000 Then
                    If UsaSeccion Then
                        Set Reg2 = Base.OpenRecordset("select * from seccion where borrado=false and codi_ident=" & ElCodigo.codigo(BucleGrande))
                        LaSeccion = Reg2.Fields("secc_Maqui")
                        Reg2.Close
                        .Fields("codi_ident") = ElCodigo.codigo(BucleGrande)
                        .Fields("secc_Maqui") = LaSeccion
                        .Fields("numero_eqp") = Null
                    Else
                        Set Reg2 = Base.OpenRecordset _
                        ("select * from equipos where borrado=false and numero_eqp=" & ElCodigo.codigo(BucleGrande))
                        LaSeccion = Reg2.Fields("secc_Maqui")
                        .Fields("codi_ident") = Reg2.Fields("codi_ident")
                        Reg2.Close
                        .Fields("secc_Maqui") = LaSeccion
                        .Fields("numero_eqp") = ElCodigo.codigo(BucleGrande)
                    End If
                Else
                    .Fields("codi_ident") = ElCodigoBarras.Mostrador
                    .Fields("secc_Maqui") = ElCodigoBarras.Mostrador
                    .Fields("numero_eqp") = Null
                End If
                .Update
            End If
        End With
    Next BucleGrande
fin:
    CerrarBase Base
    
    Alta_CodigoBarras = Retorno
End Function

Public Function Alta_Balanza(MiBalanza As DB_Balanzas) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegBal As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Modific As Boolean
    Dim Retorno As Integer
    '**************************
    ' Verificación de errores
    '***************************
    If MiBalanza.DireccionIP = "" Then MiBalanza.DireccionIP = "___.___.___.___"
    If MiBalanza.NBalanza < 1 Or MiBalanza.NBalanza > 60 Then
        Retorno = 12
        GoTo fin
    End If
    'If (MiBalanza.Tipo = 200 Or MiBalanza.Tipo = 300) And (MiBalanza.NBalanza < 1 Or MiBalanza.NBalanza > 8) Then
    '    Retorno = 12
    '    GoTo fin
    'End If
    'If (MiBalanza.Tipo = 400) And (MiBalanza.NBalanza < 1 Or MiBalanza.NBalanza > 30) Then
    '    Retorno = 12
    '    GoTo fin
    'End If
    If tcpip And (comprueba_formato_ip(MiBalanza.DireccionIP) = False) Then
        Retorno = 13
        GoTo fin
    End If
    Modific = False
    Set Base = OpenDatabase(Base_General)
    '
    
    '
    Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" _
    & MiBalanza.NMostrador & " and borrado=false")
    If Registro.EOF Then
        CerrarBase Base
        Retorno = 3
        GoTo fin
    Else
        'If (MiBalanza.Tipo = 200 Or MiBalanza.Tipo = 300) And Registro.Fields("enviardatos") <> "ECON" And Registro.Fields("enviardatos") <> "ECOL" Then
        '    Retorno = 52
        '    GoTo fin
        'End If
        
        'If MiBalanza.Tipo = 200 Or MiBalanza.Tipo = 300 Then
        '    ' sólo una balanza ecolabel o econet o v-8 por mostrador
        '    Set RegBal = Base.OpenRecordset _
        '    ("select numero_eqp from equipos where (modelo=200 or modelo=300) and borrado=false and codi_ident=" & MiBalanza.NMostrador)
        '    If Not RegBal.EOF Then
        '        If MiBalanza.NBalanza <> RegBal.Fields("numero_eqp") Then
        '            Retorno = 53
        '            GoTo fin
        '        End If
        '    End If
        'End If
        
        'If MiBalanza.Tipo = 400 And Registro.Fields("enviardatos") <> "ECOP" Then
        '    Retorno = 52
        '    GoTo fin
        'End If
        Set Registro = Base.OpenRecordset("select * from equipos where numero_eqp=" _
        & MiBalanza.NBalanza & " and borrado=false")
        If Not Registro.EOF Then
            If Registro.Fields("codi_ident") <> MiBalanza.NMostrador Then
                CerrarBase Base
                Retorno = 2
                GoTo fin
            End If
        End If
    End If
    '****************************
    Retorno = 4
    
    '****************************
    Set Registro = Base.OpenRecordset _
    ("select * from equipos where borrado=false and numero_eqp=" & MiBalanza.NBalanza)
    With Registro
        
        If Not .EOF Then
            If .Fields("descripcio") <> Trim(MiBalanza.Descripcion) Then Retorno = 1
            If tcpip And (.Fields("prog_eqp") <> MiBalanza.DireccionIP Or (MiBalanza.DireccionIP <> "" And IsNull(.Fields("prog_eqp")))) Then Retorno = 1
            Retorno = 1
            If Not Edit_Record(Registro) Then .AddNew
       
        Else
            .AddNew
            Retorno = 0
        End If
       
        .Fields("numero_eqp") = MiBalanza.NBalanza
        .Fields("descripcio") = MiBalanza.Descripcion
        .Fields("borrado") = False
        .Fields("codi_ident") = MiBalanza.NMostrador
        Set Reg2 = Base.OpenRecordset("select secc_maqui,codi_ident from seccion where codi_ident=" & MiBalanza.NMostrador)
        .Fields("secc_Maqui") = Reg2.Fields("secc_Maqui")
        Reg2.Close
        If MiBalanza.Tipo = 1 Then
           .Fields("modelo") = 1
        Else
           .Fields("modelo") = 0
        End If
        If MiBalanza.Tipo < 100 Then
            .Fields("prog_eqp") = MiBalanza.DireccionIP
        Else
            .Fields("prog_eqp") = "999.999.999.999"
            .Fields("modelo") = MiBalanza.Tipo
        End If
        '1.9.0
        If MiBalanza.lss165 Then
            .Fields("ss165") = True
        Else
            .Fields("ss165") = False
        End If
        ''''''
        If Retorno = 1 Then .Fields("tran_eqp") = ""
        .Update
        
    End With
    CerrarBase Base
fin:
    
    Alta_Balanza = Retorno
End Function
Public Function Alta_Mostrador(DatosMostrador As DB_mostrador) As Integer
    '***********************
    ' Tipo de mostrador
    ' 0 --> TouchScale
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim most As Integer
    Dim salir As Boolean
    Dim Retorno As Integer
    salir = False
    Retorno = 0
    If DatosMostrador.NMostrador > 999 Then Retorno = 7
    If DatosMostrador.NSeccion > 99 Then Retorno = 8
    If DatosMostrador.NSeccion > 60 And _
    (DatosMostrador.Tipo = 0 Or DatosMostrador.Tipo = 2) Then Retorno = 8
    If Len(DatosMostrador.Descripcion) > 50 Then Retorno = 9
    If Len(DatosMostrador.name) > 50 Then Retorno = 9
    If Retorno <> 0 Then GoTo fin
    Set Base = OpenDatabase(Base_General)
    'If (DatosMostrador.Tipo = 3 Or DatosMostrador.Tipo = 4) And UsaModem Then
    '    Select Case DatosMostrador.Tipo
    '        Case 3
    '            Set Registro = Base.OpenRecordset("select * from seccion where borrado=false and enviardatos='ECON'")
    '        Case 4
    '            Set Registro = Base.OpenRecordset("select * from seccion where borrado=false and enviardatos='ECOL'")
    '    End Select
    '    If Not Registro.EOF Then
    '        Base.Close
    '        Alta_Mostrador = 50
    '        Exit Function
    '    End If
    '    Set Registro = Base.OpenRecordset("select secc_maqui from seccion where borrado=false and secc_maqui=" & DatosMostrador.NSeccion)
    '    If Not Registro.EOF Then
    '        Base.Close
    '        Alta_Mostrador = 51
    '        Exit Function
    '    End If
    'Else
        Set Registro = Base.OpenRecordset("select codi_ident,secc_maqui from seccion where borrado=false and (enviardatos='ECON' or enviardatos='ECOL') and secc_maqui=" & DatosMostrador.NSeccion)
        If Not Registro.EOF Then
            If Registro.Fields("codi_ident") <> DatosMostrador.NMostrador Then
                Base.Close
                Alta_Mostrador = 51
                Exit Function
            End If
        End If
    'End If
    Set Registro = Base.OpenRecordset _
    ("select * from seccion where borrado=false and codi_ident=" & DatosMostrador.NMostrador)
    With Registro
        If .EOF Then
            .AddNew
            .Fields("codi_ident") = DatosMostrador.NMostrador
            .Fields("secc_Maqui") = DatosMostrador.NSeccion
            .Fields("descripcio") = DatosMostrador.Descripcion
            .Fields("nombre") = DatosMostrador.name
            .Fields("COM") = DatosMostrador.PuertoCOM
            .Fields("tran_sec") = ""
            .Fields("tran_secSC10") = ""
            .Fields("borrado") = False
            .Fields("Sec_IP") = DatosMostrador.DireccionIP
            If DatosMostrador.Tipo = 4 Then
                .Fields("multiple") = Val(FrmMostrador2.TextPeco.TexT)
            Else
                .Fields("multiple") = 0
            End If
            .Fields("enviardatos") = "GA"
            'If DatosMostrador.Tipo = 0 Then .Fields("enviardatos") = "GA"
            'If DatosMostrador.Tipo = 1 Then .Fields("enviardatos") = "SC10"
            'If DatosMostrador.Tipo = 2 Then .Fields("enviardatos") = "TODOS"
            'If DatosMostrador.Tipo = 3 Then .Fields("enviardatos") = "ECON"
            'If DatosMostrador.Tipo = 4 Then .Fields("enviardatos") = "ECOL"
            'If DatosMostrador.Tipo = 5 Then .Fields("enviardatos") = "ECOP"
            '2.0.36
            'If AhorraMas And usaSC10 Then 'And FrmMostrador2.ChkAhoSS.Value = vbChecked Then
            '    'If FrmMostrador2.ChkAhoSS.Value = vbChecked Then
            '    '    .Fields("COM") = 255
            '    'Else
            '    '    .Fields("COM") = 0
            '    'End If
            'End If
            '''''''''''''''''''''''''''''
            .Update
            .Close
       Else
            Retorno = 4
            'If .Fields("descripcio") <> Trim(DatosMostrador.descripcion) Or _
            '(DatosMostrador.Tipo = 0 And .Fields("enviardatos") <> "GA") Or _
            '(DatosMostrador.Tipo = 1 And .Fields("enviardatos") <> "SC10") Or _
            '(DatosMostrador.Tipo = 3 And .Fields("enviardatos") <> "ECON") Or _
            '(DatosMostrador.Tipo = 4 And .Fields("enviardatos") <> "ECOL") Or _
            '(DatosMostrador.Tipo = 5 And .Fields("enviardatos") <> "ECOP") Or _
            '(DatosMostrador.Tipo = 2 And .Fields("enviardatos") <> "TODOS") Or _
            '(DatosMostrador.DireccionIP <> .Fields("SEC_IP")) Or _
            '((DatosMostrador.PuertoCOM <> .Fields("COM")) Or (FrmMostrador2.ChkAhoSS.Value = vbChecked And .Fields("COM") = 0) Or (FrmMostrador2.ChkAhoSS.Value = vbUnchecked And .Fields("COM") = 255)) Then
                
            If .Fields("descripcio") <> Trim(DatosMostrador.Descripcion) Or _
            (DatosMostrador.Tipo = 0 And .Fields("enviardatos") <> "GA") Or _
            (DatosMostrador.DireccionIP <> .Fields("SEC_IP")) Then
                
                If .Fields("secc_Maqui") <> DatosMostrador.NSeccion Then
                    Retorno = 2
                Else
                    Retorno = 1
                    If Not Edit_Record(Registro) Then
                        Alta_Mostrador = 60
                        Base.Close
                        Exit Function
                    End If
                    .Fields("codi_ident") = DatosMostrador.NMostrador
                    .Fields("secc_Maqui") = DatosMostrador.NSeccion
                    .Fields("descripcio") = DatosMostrador.Descripcion
                    .Fields("nombre") = DatosMostrador.name
                    .Fields("tran_sec") = ""
                    .Fields("tran_secSC10") = ""
                    .Fields("COM") = DatosMostrador.PuertoCOM
                    .Fields("borrado") = False
                    .Fields("Sec_IP") = DatosMostrador.DireccionIP
                    If DatosMostrador.Tipo = 4 Then
                        .Fields("multiple") = Val(FrmMostrador2.TextPeco.TexT)
                    Else
                        .Fields("multiple") = 0
                    End If
                    .Fields("enviardatos") = "GA"
                    'If DatosMostrador.Tipo = 1 Then .Fields("enviardatos") = "SC10"
                    'If DatosMostrador.Tipo = 2 Then .Fields("enviardatos") = "TODOS"
                    'If DatosMostrador.Tipo = 3 Then .Fields("enviardatos") = "ECON"
                    'If DatosMostrador.Tipo = 4 Then .Fields("enviardatos") = "ECOL"
                    'If DatosMostrador.Tipo = 5 Then .Fields("enviardatos") = "ECOP"
                    '2.0.36
                    'If AhorraMas And usaSC10 And FrmMostrador2.ChkAhoSS.Value = vbChecked Then
                    '    .Fields("COM") = 255
                    'End If
                    '2.0.36
                    'If AhorraMas And usaSC10 Then 'And FrmMostrador2.ChkAhoSS.Value = vbChecked Then
                    '    If FrmMostrador2.ChkAhoSS.Value = vbChecked Then
                    '        .Fields("COM") = 255
                    '    Else
                    '        .Fields("COM") = 0
                    '    End If
                    'End If
                    '''''''''''''''''''''''''''''
                    
                    '''''''''''''''''''''''''''''
                    
                    .Update
                    .Close
                End If
            End If
        End If
    End With
    CerrarBase Base
fin:
    
    Alta_Mostrador = Retorno
End Function
Public Function Baja_Mostrador(DatosMostrador As DB_mostrador) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Reg3 As dao.Recordset
    Dim most As Integer
    Dim Retorno As Integer
    Dim sSQL As String
    Dim nMaq As Integer
    
    most = DatosMostrador.NMostrador
    nMaq = DatosMostrador.NSeccion
    
    Set Base = OpenDatabase(Base_General)
    
    sSQL = "delete from lintxt2040 where mostrador=" & nMaq
    Base.Execute sSQL
    sSQL = "delete from text15 where mostrador=" & nMaq
    Base.Execute sSQL
    sSQL = "delete from teclas where secc_maqui=" & nMaq
    Base.Execute sSQL
    
    Set Registro = Base.OpenRecordset("select * from vendedor where codi_ident=" & most)
    If Not Registro.EOF Then
        Retorno = 0
        Registro.MoveFirst
        Do While Not Registro.EOF
            Registro.Edit
            Registro.Fields("borrado") = True
            Registro.Fields("tran_vend") = ""
            Registro.Fields("tran_asign") = ""
            Registro.Fields("tran_vendSC10") = ""
            Registro.Fields("tran_asignSC10") = ""
            Registro.Update
            Registro.Movenext
        Loop
    End If
    Registro.Close
    Set Registro = Nothing
    
    Set Reg3 = Base.OpenRecordset("select * from codbar where secc_maqui=" & most)
    With Reg3
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Reg3.Edit
                Reg3.Fields("borrado") = True
                Reg3.Update
                Reg3.Movenext
            Loop
        End If
    End With
    
    Set Reg3 = Base.OpenRecordset("select * from fam_code where codi_ident=" & most)
    With Reg3
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Reg3.Edit
                Reg3.Fields("posicion") = 1525
                Reg3.Update
                Reg3.Movenext
            Loop
        End If
    End With
        
    Set Reg3 = Base.OpenRecordset("select * from sub_code where codi_ident=" & most)
    With Reg3
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Reg3.Delete
                If Not Reg3.EOF Then Reg3.Movenext
            Loop
        End If
    End With
        
    Set Reg3 = Base.OpenRecordset("select * from articulo where codi_ident=" & most)
    With Reg3
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                Reg3.Edit
                Reg3.Fields("borrado") = True
                Reg3.Update
                Reg3.Movenext
            Loop
        End If
    End With
    
    Set Reg3 = Base.OpenRecordset("select * from equipos where codi_ident=" & most)
    With Reg3
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                If Reg3.Fields("modelo") <> 1 Then
                    Reg3.Delete
                End If
                If Not Reg3.EOF Then Reg3.Movenext
                
            Loop
        End If
    End With
    
    'Se dejan los mostradores ... Se eliminarán solo con regeneración Base
    'sSQL = "SELECT * " & _
    '       "FROM seccion " & _
    '       "WHERE secc_maqui=" & most
    'Set Reg2 = Base.OpenRecordset(sSQL)
    'If Not Reg2.EOF Then
    '    Do While Not Reg2.EOF
    '        Reg2.Edit
    '        Reg2.Fields("borrado") = True
    '        Reg2.Update
    '        Reg2.Movenext
    '    Loop
    'End If

    Reg3.Close
    
fin:
    CerrarBase Base
    
    Baja_Mostrador = Retorno
End Function
Public Function Baja_CodigoBarras(ElCodigoBarras As DB_CodigoBarras) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim bucle As Integer
    Dim BucleGrande As Integer
    Dim ElCodigo As TipoCodigos
    Dim Retorno As Integer
    Dim UsaSeccion As Boolean
    Dim LosMostradores As String
    Dim LosEquipos As String
    '******************
    ' Condiciones de Error
    Set Base = dao.OpenDatabase(Base_General)
    If ElCodigoBarras.Equipo = 0 Then
        UsaSeccion = True
        Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" & CStr(ElCodigoBarras.Mostrador) & " and borrado=false")
        If Not Registro.EOF Then
            Retorno = 0
            Set Reg2 = Base.OpenRecordset("select * from seccion where secc_maqui=" & _
            Registro.Fields("secc_Maqui") & " and borrado=false")
            Reg2.MoveFirst
            Do Until Reg2.EOF
                LosMostradores = LosMostradores & Reg2.Fields("codi_ident") & " , "
                If Not Reg2.EOF Then Reg2.Movenext
            Loop
            LosMostradores = left(LosMostradores, Len(LosMostradores) - 3)
        Else
            Retorno = 3
        End If
    Else
        LosEquipos = ElCodigoBarras.Equipo
        UsaSeccion = False
        Set Registro = Base.OpenRecordset("select * from seccion where codi_ident=" & _
        ElCodigoBarras.Mostrador & " and borrado=false")
        If Registro.EOF Then
            Retorno = 3
        Else
            Set Reg2 = Base.OpenRecordset("select * from equipos where numero_eqp=" & _
            ElCodigoBarras.Equipo & " and secc_Maqui=" & Registro.Fields("secc_Maqui") & " and borrado=false")
            If Reg2.EOF Then
                Retorno = 2
            Else
                Retorno = 0
                
            End If
        End If
        
    End If
    '*********************************************************
    bucle = 1
    Retorno = 4
    If UsaSeccion Then
        ElCodigo = Desglosa_Codigos(LosMostradores)
    Else
        LosEquipos = ElCodigoBarras.Equipo
        ElCodigo = Desglosa_Codigos(LosEquipos)
    End If
    Set Base = OpenDatabase(Base_General)
    ' 1º en local
    For BucleGrande = 1 To ElCodigo.cantidad
        If UsaSeccion Then
            Set Registro = Base.OpenRecordset _
            ("select * from codbar where borrado=false and secc_eqp<>0 and (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & ElCodigo.codigo(BucleGrande))
        Else
            Set Registro = Base.OpenRecordset _
            ("select * from codbar where borrado=false and secc_eqp<>0 and numero_eqp=" & ElCodigo.codigo(BucleGrande))
        End If
        With Registro
            If Not .EOF Then
                Retorno = 0
                If Edit_Record(Registro) Then
                    .Fields("borrado") = True
                    .Fields("cb_vnt") = ""
                    .Fields("cb_sup") = ""
                    .Fields("cb_mix") = ""
                    .Fields("cb_ev1") = ""
                    .Fields("cb_ev2") = ""
                    .Fields("tran_bar") = ""
                    .Fields("tran_barSC10") = ""
                    .Fields("tran_ep") = ""
                    .Fields("tran_el") = ""
                    .Fields("tran_en") = ""
                    If Not UsaSeccion Then
                        Set Reg2 = Base.OpenRecordset("select * from codbar where borrado=false and (isnull(numero_eqp) or numero_eqp=0) and secc_maqui=" & .Fields("secc_Maqui"))
                        With Reg2
                            If Not .EOF Then
                                .MoveFirst
                                Do Until .EOF
                                    If Edit_Record(Reg2) Then
                                        .Fields("tran_bar") = ""
                                        .Fields("tran_barSC10") = ""
                                        .Update
                                    End If
                                    If Not .EOF Then .Movenext
                                Loop
                            End If
                            .Close
                        End With
                    End If
                    .Update
                End If
            End If
        End With
        Registro.Close
        ' 2º en red
        If UsaSeccion Then
            Set Registro = Base.OpenRecordset _
            ("select * from codbar where borrado=false and secc_eqp=0 and (isnull(numero_eqp) or numero_eqp=0) and codi_ident=" & ElCodigo.codigo(BucleGrande))
        Else
            Set Registro = Base.OpenRecordset _
            ("select * from codbar where borrado=false and secc_eqp=0 and numero_eqp=" & ElCodigo.codigo(BucleGrande))
        End If
        With Registro
            If Not .EOF Then
                If Edit_Record(Registro) Then
                    Retorno = 0
                    .Fields("borrado") = True
                    .Fields("cb_vnt") = ""
                    .Fields("cb_sup") = ""
                    .Fields("cb_mix") = ""
                    .Fields("cb_ev1") = ""
                    .Fields("cb_ev2") = ""
                    .Fields("tran_bar") = ""
                    .Fields("tran_barSC10") = ""
                    .Fields("tran_ep") = ""
                    .Fields("tran_el") = ""
                    .Fields("tran_en") = ""
                    If Not UsaSeccion Then
                        Set Reg2 = Base.OpenRecordset("select * from codbar where borrado=false and (isnull(numero_eqp) or numero_eqp=0) and secc_maqui=" & .Fields("secc_Maqui"))
                        With Reg2
                            If Not .EOF Then
                                .MoveFirst
                                Do Until .EOF
                                    If Edit_Record(Reg2) Then
                                        .Fields("tran_bar") = ""
                                        .Fields("tran_barSC10") = ""
                                        .Fields("tran_ep") = ""
                                        .Fields("tran_el") = ""
                                        .Fields("tran_en") = ""
                                        .Update
                                    End If
                                    If Not .EOF Then .Movenext
                                Loop
                            End If
                            .Close
                        End With
                    End If
                    .Update
                End If
            End If
        End With
    Next BucleGrande
fin:
    CerrarBase Base
    
    Baja_CodigoBarras = Retorno
End Function

Public Function Baja_Articulo(Miarticulo As DB_Articulo) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Reg3 As dao.Recordset
    Dim Retorno As Integer
    Dim sSQL As String
    
    Set Base = OpenDatabase(Base_General)
    If Miarticulo.codigo = 0 Then
        If Miarticulo.Plu = 0 Then
            Retorno = 2
            CerrarBase Base
            GoTo fin
        End If
        If Miarticulo.Mostrador = 0 Then
            Retorno = 3
            CerrarBase Base
            GoTo fin
        End If
        Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & _
        Miarticulo.Mostrador & " and plu=" & Miarticulo.Plu & " and borrado=false")
        If Registro.EOF Then
            Retorno = 2
            CerrarBase Base
            GoTo fin
        Else
            Miarticulo.codigo = Registro.Fields("codigo")
        End If
    End If
    Retorno = 4
    
    sSQL = "delete from lintxt2040 where codigo=" & Miarticulo.codigo
    Base.Execute sSQL
    sSQL = "delete from text15 where codigo=" & Miarticulo.codigo
    Base.Execute sSQL
    
    Set Registro = Base.OpenRecordset _
    ("select * from articulo where borrado=false and codigo=" & Miarticulo.codigo)
    With Registro
        If Not .EOF Then
            If Not Edit_Record(Registro) Then
                Baja_Articulo = 0
                Base.Close
                Exit Function
            End If
            Retorno = 0
            .Fields("borrado") = True
            .Fields("tran_pluSC10") = ""
            .Fields("tran_plu") = ""
            .Fields("tran_EN") = ""
            .Fields("tran_EL") = ""
            '.Fields("tran_ep") = Space(60)
            .Update
            Set Reg2 = Base.OpenRecordset _
            ("select * from seccion where borrado=false and codi_ident=" & .Fields("codi_ident"))
            Set Reg3 = Base.OpenRecordset _
            ("select * from articulo where borrado=false and balenv=true and secc_maqui=" & Reg2.Fields("secc_Maqui"))
            If Reg3.EOF Then
                Reg3.Close
                Set Reg3 = Base.OpenRecordset _
                ("select * from seccion where borrado=false and secc_maqui=" & Reg2.Fields("secc_Maqui"))
                With Reg3
                    If Not .EOF Then
                        .MoveFirst
                        Do Until .EOF
                            If .Fields("multiple") = 1 Then
                                If Edit_Record(Reg3) Then
                                    .Fields("multiple") = -1
                                    If .Fields("enviardatos") = "GA" Then .Fields("tran_secSC10") = ""
                                    If .Fields("enviardatos") = "SC10" Then .Fields("tran_sec") = ""
                                    .Update
                                End If
                            End If
                            If Not .EOF Then .Movenext
                        Loop
                    End If
                End With
            End If
            CerrarBase Base
        Else
            CerrarBase Base
        End If
    End With
fin:
    'If Retorno = 0 Then
    '    Set Base = AbrirBase
    '    Set Registro = Base.OpenRecordset("select * from teclas where borrado=false and codigo=" & Miarticulo.codigo)
    '    With Registro
    '        If Not .EOF Then
    '            Do Until .EOF
    '                If Edit_Record(Registro) Then
    '                    .Fields("borrado") = True
    '                    .Fields("tran_tec") = ""
    '                    .Fields("tran_tecSC10") = ""
    '                    .Update
    '                End If
    '                If Not .EOF Then .Movenext
    '            Loop
    '        End If
    '    End With
    '    Base.Close
    'End If
    
    Baja_Articulo = Retorno
            
End Function
Public Function Baja_Balanza(MiBalanza As DB_Balanzas) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Retorno As Integer
    Set Base = OpenDatabase(Base_General)
    Set Registro = Base.OpenRecordset _
    ("select * from equipos where borrado=false and numero_eqp=" & MiBalanza.NBalanza)
    Retorno = 4
    With Registro
        If Not .EOF Then
            Retorno = 0
            If Not Edit_Record(Registro) Then
                Baja_Balanza = 0
                Base.Close
                Exit Function
            End If
            .Fields("borrado") = True
            '1.8.9
            Registro.Edit
            Registro.Delete
            '''''''''''''''
            
            'Set Reg2 = Base.OpenRecordset("select * from cabley where numero_eqp=" & MiBalanza.NBalanza)
            'With Reg2
            '    If Not .EOF Then
            '        .MoveFirst
            '        Do Until .EOF
            '            If Not .EOF Then Delete_Record Reg2
            '            If Not .EOF Then .Movenext
            '        Loop
            '    End If
            'End With
            'Set Reg2 = Base.OpenRecordset("select * from codbar where numero_eqp=" & MiBalanza.NBalanza)
            'With Reg2
            '    If Not .EOF Then
            '        .MoveFirst
            '        Do Until .EOF
            '            If Not .EOF Then Delete_Record Reg2
            '            If Not .EOF Then .Movenext
            '        Loop
            '    End If
            'End With
            'Set Reg2 = Base.OpenRecordset("select * from publi where numero_eqp=" & MiBalanza.NBalanza)
            'With Reg2
            '    If Not .EOF Then
            '        .MoveFirst
            '        Do Until .EOF
            '            If Not .EOF Then Delete_Record Reg2
            '            If Not .EOF Then .Movenext
            '        Loop
            '    End If
            'End With
            'Set Reg2 = Base.OpenRecordset("select * from teclas where numero_eqp=" & MiBalanza.NBalanza)
            'With Reg2
            '    If Not .EOF Then
            '        .MoveFirst
            '        Do Until .EOF
            '            If Not .EOF Then Delete_Record Reg2
            '            If Not .EOF Then .Movenext
            '        Loop
            '    End If
            'End With
            
            
        End If
    End With
    CerrarBase Base
    
    Baja_Balanza = Retorno
End Function
Public Function Modificacion_Articulo_Precio(Miarticulo As DB_Articulo, Optional MiUsuario As String, Optional NoLog As Boolean) As Integer
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Retorno As Integer
    Dim MiModificacion As Boolean
    Dim MiArticulo2 As DB_Articulo
    MiModificacion = False
    Set Base = OpenDatabase(Base_General)
    If Miarticulo.codigo = 0 Then
        If Miarticulo.Plu = 0 Then
            Retorno = 2
            CerrarBase Base
            GoTo fin
        End If
        If Miarticulo.Mostrador = 0 Then
            Retorno = 3
            CerrarBase Base
            GoTo fin
        End If
        Set Registro = Base.OpenRecordset("select * from articulo where codi_ident=" & _
        Miarticulo.Mostrador & " and plu=" & Miarticulo.Plu & " and borrado=false")
        If Registro.EOF Then
            Retorno = 2
            CerrarBase Base
            GoTo fin
        Else
            Miarticulo.codigo = Registro.Fields("codigo")
        End If
    End If
    Retorno = 4
    Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & Miarticulo.codigo _
    & " and borrado=false")
    If Not Registro.EOF Then
        '***************
        ' guarda datos para log
        '***************
        If LogArticulos Then
            MiArticulo2.Mostrador = Registro.Fields("codi_ident")
            MiArticulo2.subsec = Registro.Fields("codi_sub")
            MiArticulo2.Plu = Registro.Fields("plu")
            MiArticulo2.familia = Registro.Fields("codi_fam")
            MiArticulo2.caducidad = Registro.Fields("caducidad")
            MiArticulo2.codigo = Registro.Fields("codigo")
            MiArticulo2.precio = Miarticulo.precio
            MiArticulo2.tara = Registro.Fields("tara")
            MiArticulo2.Descriptivos(0) = Registro.Fields("des_plu1")
            MiArticulo2.Descriptivos(1) = Registro.Fields("codi_pes")
            MiArticulo2.Descriptivos(2) = Registro.Fields("secc_Maqui")
        End If
        '***************
        ' /guarda datos para log
        '***************
        If UsaEuro Then
            If Miarticulo.precio <> Registro.Fields("Euros") Then
                Retorno = 1
                MiModificacion = True
            End If
        Else
            If Miarticulo.precio <> Registro.Fields("precio") Then
                Retorno = 1
                MiModificacion = True
            End If
        End If
        If Not Edit_Record(Registro) Then
            Modificacion_Articulo_Precio = 60
            Base.Close
            Exit Function
        End If
        If UsaEuro Then
            Registro.Fields("Euros") = Miarticulo.precio
            Registro.Fields("precio") = EUROtoPTA(Miarticulo.precio)
        Else
            Registro.Fields("precio") = Miarticulo.precio
            Registro.Fields("Euros") = PTAtoEURO(Miarticulo.precio)
        End If
        If MiModificacion Then
            Registro.Fields("prc3") = Val(Format(Now, "yymmddhhmm"))
            Registro.Fields("tran_plu") = ""
            Registro.Fields("tran_pluSC10") = ""
            'Registro.Fields("tran_ep") = Space(60)
            Registro.Fields("tran_el") = ""
            Registro.Fields("tran_en") = ""
            Registro.Fields("tran_tx1") = ""
            Registro.Fields("tran_tx1SC10") = ""
            'Registro.Fields("prc3") = Format(Now, "yymmddhhmm")
            
        End If
        Registro.Update
    Else
        Retorno = 2
    End If
    If (LogArticulos And MiModificacion) And (Not NoLog) Then
        Set Registro = Base.OpenRecordset("select * from log")
        With Registro
            .AddNew
            If UsaEuro Then
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(MiArticulo2.codigo, "000000") & "," & _
                "000" & "," & Format(MiArticulo2.Mostrador, "000") & "," & Format(MiArticulo2.precio * 100, "0000000") & "," & Trim(MiArticulo2.Descriptivos(0)) & "," & Format(Val(MiArticulo2.Descriptivos(2)), "00")
            Else
                .Fields("Texto") = "LOG " & Format(Date, "DDMMYY") & "," & "0" & "," & Format(Time, "HHMM") & "," & "0" & "," & Format(MiArticulo2.codigo, "000000") & "," & _
                "000" & "," & Format(MiArticulo2.Mostrador, "000") & "," & Format(MiArticulo2.precio * 10 ^ decimales, "0000000") & "," & Trim(MiArticulo2.Descriptivos(0)) & "," & Format(Val(MiArticulo2.Descriptivos(2)), "00")
            End If
            .Fields("codi_ident") = MiArticulo2.Mostrador
            .Fields("codi_sub") = MiArticulo2.subsec
            .Fields("codigo") = MiArticulo2.codigo
            .Fields("plu") = MiArticulo2.Plu
            .Fields("precio") = MiArticulo2.precio
            .Fields("codi_fam") = MiArticulo2.familia
            .Fields("caducidad") = MiArticulo2.caducidad
            .Fields("tara") = MiArticulo2.tara
            '.Fields("tara_envasado") = MiArticulo2.tara
            .Fields("des_plu1") = MiArticulo2.Descriptivos(0)
            .Fields("Fecha") = Date
            If MiArticulo2.Descriptivos(1) = "W" Then
                .Fields("tipo_ven") = "W"
            Else
                .Fields("tipo_ven") = "U"
            End If
            If MiUsuario = "" Then
                .Fields("Comentario") = CargaCadena(24) & " " & UsuarioActual.Nombre
            Else
                .Fields("Comentario") = CargaCadena(24) & " " & MiUsuario
            End If
            .Update
        End With
    End If
    'c2f 1.7.11 provoca error en absorción datos ... CerrarBase Base
fin:
    Modificacion_Articulo_Precio = Retorno
End Function
Public Function comprueba_formato_ip(ByVal valorip As String) As Boolean
    Dim Conforme As Boolean
    Dim punto As Boolean
    Dim bucle As Long
    Dim CuentaPuntos As Long
    Dim CuentaDigitos As Long
    Dim Buffer As String
    Conforme = 1
    If Len(valorip) > 15 Then
        Conforme = 0
    End If
    punto = True
    CuentaPuntos = 0
    CuentaDigitos = 0
    Buffer = ""
    For bucle = 1 To Len(valorip)
        
        If Mid(valorip, bucle, 1) = "." Then
            If punto = True Then
                Conforme = False
            Else
                CuentaPuntos = CuentaPuntos + 1
                CuentaDigitos = 0
                Buffer = ""
                punto = True
            End If
        Else
            If IsNumeric(Mid(valorip, bucle, 1)) Then
                punto = False
                CuentaDigitos = CuentaDigitos + 1
                Buffer = Buffer & Mid(valorip, bucle, 1)
            Else
                Conforme = False
            End If
        End If
                
        If CuentaPuntos > 3 Or CuentaDigitos > 3 Then Conforme = False
        If CuentaDigitos = 3 Then
            If Val(Buffer) > 255 Then Conforme = False
        End If
        
    Next bucle
    comprueba_formato_ip = Conforme
End Function
'\\\\\\\\\\\
'///////////
'En primera versión se considera:
'   ICG           HydraTouch
'  -------------  -------------
'  TotalNeto      Importe
'  TotalBruto     Importe_bruto
'  ImporteCambio  Cambio
'                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
'En primera versión se considera:
'   ICG           HydraTouch
'  -------------  -------------
'  TotalNeto      Importe
'  TotalBruto     Importe_bruto
'  ImporteCambio  Cambio
'                 CantidadEntre = ICG(TotalNeto+ImporteCambio)
Public Sub Actualiza_Tabla_Cabecera(Base As dao.Database, Cabecera As tipo_cabecera)
'*************************
' Tipos de Tiquet (typtic)
' 1 --> Rollo
' 2 --> Super
' 3 --> Envasado
' 4 --> Autoservicio
' 5 --> Rollo Local
' 6 --> Super Local
' 7 --> Envasado Local
' 8 --> Autoservicio Local
'*************************
Dim regVendedor As dao.Recordset
Dim Registro As dao.Recordset
Dim RegSeccion As dao.Recordset
Dim RegDes As dao.Recordset
Dim MiCadenaDebug As String
Dim tqcodbar As Tipo_TqCodBar
Dim micodbar As String
Dim Cod_Art As Long
Dim bimpoBrutoCab As Double
Dim sHoraCab As String
Dim nMyCont As Integer
Dim sCont As Variant
Dim sTmp As String
Dim Que_base As String
Dim sSQL As String
'''''''''''''''''''''''''''''''
Dim nNumCorre As Double
Dim sFactura As String
'''''''
Dim PrecioSinIVA As Double
Dim PrecioConIVA As Double
Dim BufferIVA As String
Dim BufferIVA2 As String
Dim MiIva(5) As Tipo_Grupo_IVA
Dim BucleIVA As Integer
Dim RegIVA As dao.Recordset
Dim RegArt As dao.Recordset
Dim bucle As Integer
''''''''''''''''

    Que_base = "GA"
    
    '3.6.0 revisar... Se anula para recoger desde ltickets
    ''Registro de articulos
    'Set RegArt = Base.OpenRecordset("select codigo,tipo_iva from articulo")
    'For bucle = 1 To Cabecera.LineasMensaje
    '    RegArt.FindFirst "codigo=" & Cabecera.Lines(bucle).CodigoArticulo
    '    If RegArt.EOF Or RegArt.NoMatch Then
    '        Cabecera.Lines(bucle).IVA = 0
    '    Else
    '        Cabecera.Lines(bucle).IVA = RegArt!tipo_iva
    '    End If
    'Next bucle
    'RegArt.Close
    'Set RegArt = Nothing
   '
    'Set RegIVA = Base.OpenRecordset("select * from iva")
    ''carga porcentajes en lineas
    'For bucle = 1 To Cabecera.LineasMensaje
    '    RegIVA.FindFirst "tipo=" & Cabecera.Lines(bucle).IVA
    '    If RegIVA.EOF Or RegIVA.NoMatch Then
    '        Cabecera.Lines(bucle).porcentaje = 0
    '    Else
    '        Cabecera.Lines(bucle).porcentaje = RegIVA!Valor
    '    End If
    'Next bucle
    ''''''''''''
    Set RegIVA = Base.OpenRecordset("select * from iva")
    ''''''''''''
    
    '********************
    ' carga tipos de IVA
    '********************
    For bucle = 0 To 4
        MiIva(bucle).Base = 0
        MiIva(bucle).Importe = 0
        MiIva(bucle).porcentaje = 0
        RegIVA.FindFirst "tipo=" & (bucle)
        If RegIVA.EOF Or RegIVA.NoMatch Then
            MiIva(bucle).porcentaje = 0
        Else
            MiIva(bucle).porcentaje = RegIVA!Valor
        End If
    Next bucle
    RegIVA.Close
    Set RegIVA = Nothing
    
    'If Dir(App.Path & "\automotor.txt") = "" Then
        'Cálculos Iva
        If Cabecera.Trainning = False And ((Cabecera.EstadoTiquet And 64) <> 64 And (Cabecera.EstadoTiquet And 128) <> 128) Then
            For bucle = 1 To Cabecera.LineasMensaje
                If Cabecera.Lines(bucle).LineaCancelada <> True Then
                    
                    'prueba descuentos
                    If Cabecera.descuento <> 0 Then
                        Cabecera.Lines(bucle).DescuentoLinea = Cabecera.Lines(bucle).DescuentoLinea + Cabecera.descuento
                        Cabecera.Lines(bucle).ImporteLinea = Round(1000 * ((Cabecera.Lines(bucle).ImporteLinea) - (Cabecera.Lines(bucle).ImporteLinea / 100) * Cabecera.descuento) / 1000, 2)
                    End If
                    ''''''''''''''''''
                    If Not UsaEuro Then
                            'If (Cabecera.Lines(bucle).TipoLinea > 1) Then
                            '    MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe - (Cabecera.Lines(bucle).ImporteLinea)
                            '    MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base - (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            'Else
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe + (Cabecera.Lines(bucle).ImporteLinea)
                                '3.6.0 MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, Cabecera.Lines(bucle).porcentaje)), 2)
                            'End If
                    
                    Else
                            'If (Cabecera.Lines(bucle).TipoLinea > 1) Then
                            '    MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe - (Cabecera.Lines(bucle).ImporteLinea)
                            '    MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base - (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                            'Else
                                MiIva(Cabecera.Lines(bucle).IVA).Importe = MiIva(Cabecera.Lines(bucle).IVA).Importe + (Cabecera.Lines(bucle).ImporteLinea)
                                '3.6.0 MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, MiIva(Cabecera.Lines(bucle).IVA).porcentaje)), 2)
                                MiIva(Cabecera.Lines(bucle).IVA).Base = Round(MiIva(Cabecera.Lines(bucle).IVA).Base + (PRsinIVA3(Cabecera.Lines(bucle).ImporteLinea, Cabecera.Lines(bucle).porcentaje)), 2)
                            'End If
                    End If
                End If
            Next bucle
        End If
    'End If
    
    '***************************
    ' calcula el código de barras
    '****************************
    tqcodbar.Importe = Cabecera.ImporteTotal
    tqcodbar.NBalanza = Cabecera.NBalanza
    tqcodbar.NMostrador = Cabecera.NMostrador
    tqcodbar.NTicket = Cabecera.Ntiquet
    tqcodbar.NVendedor = Cabecera.NVendedor
    'el modo trainning se va a considerar tipoPago = 12
    ''''''''''''''''''''''''''''''''''''''''
    tqcodbar.Tipo = Cabecera.tipoTiquet
    If Cabecera.nLineas >= 1 Then
        Cod_Art = Cabecera.Lines(1).CodigoArticulo
    Else
        Cod_Art = 0
    End If
    micodbar = Genera_CodBar(tqcodbar, Cod_Art)
    If Trim(Cabecera.ean13) <> "" Then
        micodbar = Trim(Cabecera.ean13)
    End If
    If Que_base = "GA" Then
        MiCadenaDebug = "K-T/Scale / "
        Set Registro = Base.OpenRecordset("select * from cabecera where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & Cabecera.NVendedor & Chr(34) & " and onoff=" & Chr(34) & Cabecera.CadenaEstadoTiquet & Chr(34))
    End If
    CadenadeLogTotales "Localizando en cabecera:" & "select * from cabecera where nume=" _
        & Chr(34) & Format(Cabecera.Ntiquet, "00000") & Chr(34) & _
        " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34) & " and balanza=" & _
        Cabecera.NBalanza & " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
        " and hora=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & " and vended=" & Chr(34) & Cabecera.NVendedor & Chr(34) & " and onoff=" & Chr(34) & Cabecera.CadenaEstadoTiquet & Chr(34)
        
    With Registro
        
        If .EOF Then

            If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Cabecera de Tiquet : " _
            & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
            & "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.Fecha & "Hora : " & Cabecera.hora)
            
CadenadeLogTotales "No localizado --> " & "Almacenando Cabecera de Tiquet : " _
            & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
            & "Importe : " & Cabecera.ImporteTotal & "Fecha : " & Cabecera.Fecha & "Hora : " & Cabecera.hora
            
            .AddNew
            .Fields("factura") = Cabecera.factura
            .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
            .Fields("seccion") = Cabecera.NMostrador
            .Fields("codbar") = micodbar
            .Fields("Balanza") = Cabecera.NBalanza
            .Fields("vended") = Cabecera.NVendedor
            .Fields("numlin") = Cabecera.nLineas
            .Fields("lincan") = Cabecera.LineasCanceladas
            .Fields("Importe") = Cabecera.ImporteTotal '/ 100
            bimpoBrutoCab = Cabecera.ImporteBruto 'redondea(Cabecera.ImporteBruto  / 100) 'c2f invicta
                
            'c2f invicta
            On Error Resume Next
            .Fields("importe_bruto") = bimpoBrutoCab
            ''''''''''''
            On Error GoTo 0
            ''''''''''''
            .Fields("hora") = Format(Cabecera.hora, "hh:mm:ss")
            .Fields("d_fecha") = Cabecera.Fecha
            .Fields("d_hora") = Cabecera.hora
            .Fields("onoff") = Cabecera.CadenaEstadoTiquet
            .Fields("STPETIC") = Cabecera.EstadoTiquet
            .Fields("codcli") = Cabecera.NCliente
            .Fields("Fecha") = Format(Cabecera.Fecha, "dd/mm/yy")
            .Fields("TYPTIC") = Cabecera.tipoTiquet
            .Fields("reserv2") = "/"
            .Fields("reserv3") = "NO"  'no exportado
            .Fields("tipoPago") = Cabecera.tipoPago
            .Fields("cantidadEntre") = Cabecera.cantidadEntre '/ 100
            .Fields("cambio") = Cabecera.Cambio '/ 100
            On Error Resume Next
            If Cabecera.lote <> "" Then
                .Fields("numlote") = Cabecera.lote
            Else
                .Fields("numlote") = " "
    
            End If
            '''''''''''
            .Fields("totalweight") = Cabecera.TotalWeight
            .Fields("totalunits") = Cabecera.TotalUnits
            '''''''''''
            '.Fields.Append .CreateField("abonado", dbBoolean)
            '.Fields.Append .CreateField("ticketabono", dbLong)
            '.Fields.Append .CreateField("ticketorigen", dbLong)
            '.Fields.Append .CreateField("abono", dbBoolean)
            If Cabecera.documentoabono <> 0 Then
                If Cabecera.documentoabono > 0 Then
                    .Fields("abonado") = True
                    .Fields("ticketabono") = Cabecera.documentoabono
                    .Fields("ticketorigen") = 0
                    .Fields("abono") = False
                Else
                    .Fields("abonado") = False
                    .Fields("ticketabono") = 0
                    .Fields("ticketorigen") = Abs(Cabecera.documentoabono)
                    .Fields("abono") = True
                
                End If
            Else
                    .Fields("abonado") = False
                    .Fields("ticketabono") = 0
                    .Fields("ticketorigen") = 0
                    .Fields("abono") = False
            
            End If
            .Fields("descuento") = Cabecera.descuento
            
            '''''''''''
            .Fields("base0") = MiIva(0).Base
            .Fields("por0") = MiIva(0).porcentaje
            .Fields("imp0") = Round(MiIva(0).Importe - MiIva(0).Base, 2)
            .Fields("base1") = MiIva(1).Base
            .Fields("por1") = MiIva(1).porcentaje
            .Fields("imp1") = Round(MiIva(1).Importe - MiIva(1).Base, 2)
            .Fields("base2") = MiIva(2).Base
            .Fields("por2") = MiIva(2).porcentaje
            .Fields("imp2") = Round(MiIva(2).Importe - MiIva(2).Base, 2)
            .Fields("base3") = MiIva(3).Base
            .Fields("por3") = MiIva(3).porcentaje
            .Fields("imp3") = Round(MiIva(3).Importe - MiIva(3).Base, 2)
            .Fields("base4") = MiIva(4).Base
            .Fields("por4") = MiIva(4).porcentaje
            .Fields("imp4") = Round(MiIva(4).Importe - MiIva(4).Base, 2)
            '''''''''''
            
            On Error GoTo 0
            .Update
            Cabecera.Incluir_en_Base = True
            
            '\\\\\\Descuentos a Tabla Descuentos...
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If (Cabecera.descuento > 0 Or (Cabecera.ImporteBruto - Cabecera.ImporteTotal) <> 0) Then
                sSQL = "SELECT * " & _
                       "FROM descuentos " & _
                       " WHERE fecha=" & "#" & Format(Cabecera.Fecha, "mm,dd,yy") & "#" & _
                       " AND nume=" & CStr(Cabecera.Ntiquet) & _
                       " AND seccion=" & CStr(Cabecera.NMostrador) & _
                       " AND balanza=" & CStr(Cabecera.NBalanza) & _
                       " AND vended=" & CStr(Cabecera.NVendedor) & _
                       " AND typTic=" & CStr(Cabecera.tipoTiquet)
                       '" AND hora=" & Cabecera.hora
                Set RegDes = Base.OpenRecordset(sSQL)
            
                With RegDes
            
                    If .EOF Then
                        .AddNew
                    Else
                        .Edit
                    End If
                    .Fields("nume") = Cabecera.Ntiquet
                    .Fields("seccion") = Cabecera.NMostrador
                    .Fields("balanza") = Cabecera.NBalanza
                    .Fields("vended") = Cabecera.NVendedor
                    .Fields("typTic") = Cabecera.tipoTiquet
                    .Fields("hora") = Format(Cabecera.hora, "hh:mm:ss")
                    .Fields("fecha") = Format(Cabecera.Fecha, "dd/mm/yyyy")
                    .Fields("importeneto") = Cabecera.ImporteTotal
                    .Fields("importebruto") = Cabecera.ImporteBruto
                    .Fields("descuentototal") = Cabecera.ImporteBruto - Cabecera.ImporteTotal
                    .Update
                    .Close
                End With
                Set RegDes = Nothing
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Else
            CadenadeLogTotales "Ya existe no se inserta. (Se marca en cualquier caso)"
        End If
    End With
    
    
End Sub
Public Sub Actualiza_Tabla_GtVend(Base As dao.Database, Cabecera As tipo_cabecera)
    '////////////////////////////////
    ' INCORPORA LOS DATOS DE TOTALES
    ' POR VENDEDOR A PARTIR DE
    ' LO RECIBIDO EN UN TICKET
    '////////////////////////////////
    Dim eNCONTRADO As Boolean
    Dim bucle As Long
    Dim ElPrecio As Double
    Dim ElImporte As Double
    Dim Lafecha As String
    Dim LaSeccion As Integer
    Dim Registro As dao.Recordset
    Dim rstseccion As dao.Recordset
    Dim rstvendedor As dao.Recordset
    Dim LineaInicial As Boolean
    
    On Error GoTo finVend

    LineaInicial = True
    
    Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
    If Not rstseccion.EOF Then
        LaSeccion = rstseccion.Fields("codi_ident")
    Else
        LaSeccion = Cabecera.NMostrador
    End If
    rstseccion.Close
    eNCONTRADO = False
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    With Registro
        If .EOF Then
            .AddNew
            .Fields("Fecha") = Cabecera.Fecha
            .Fields("codi_ident") = LaSeccion
            .Fields("ident_vend") = Cabecera.NVendedor
            .Fields("secc_Maqui") = Cabecera.NMostrador
            .Fields("Modo") = Cabecera.tipoTiquet
            .Fields("tiquets") = 0
            .Fields("operacion") = 0
            .Fields("cancela") = 0
            .Fields("impor_neg") = 0
            '********************
            ' Nombre del vendedor
            '********************
            Set rstvendedor = Base.OpenRecordset _
            ("select * from vendedor where borrado=false and ident_vend=" & Cabecera.NVendedor _
            & " and codi_ident=" & LaSeccion)
            If rstvendedor.EOF Or rstvendedor.Fields("Nombre") = "" Then
                .Fields("Nombre") = Space(25)
            Else
                .Fields("Nombre") = rstvendedor.Fields("Nombre")
            End If
            rstvendedor.Close
            ' ****
            'pesado o unidades
            ' ****
            .Fields("Peso") = 0
            .Fields("unidad") = 0
            ' ****
            .Fields("pesetas") = 0
            ' ****
            ' Cancelada o no
            ' ****
            .Fields("totcance") = 0
            .Fields("supertot") = 0
            .Fields("Total") = 0
            .Update
        End If
    End With
    Set Registro = Base.OpenRecordset _
    ("select * from gtvend where ident_vend=" & Val(Cabecera.NVendedor) & _
    " and codi_ident=" & Val(LaSeccion) & " and modo=" & Cabecera.tipoTiquet & _
    " and cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34))
    With Registro
        If LineaInicial Then
            .Edit
            .Fields("tiquets") = .Fields("tiquets") + 1
            .Fields("operacion") = .Fields("operacion") + Cabecera.nLineas
            .Update
        End If
        For bucle = 1 To Cabecera.LineasMensaje
        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8 Or Cabecera.tipoTiquet = 12) Then
            .Edit
            ElImporte = Cabecera.Lines(bucle).ImporteLinea '/ 100
            ElPrecio = Cabecera.Lines(bucle).precio '/ 100
            ' ****
            'pesado o unidades
            ' ****
            If Cabecera.Lines(bucle).LineaCancelada = False Then
                If Cabecera.Lines(bucle).EsPesado = False Then
                    .Fields("unidad") = .Fields("unidad") + Cabecera.Lines(bucle).cantidad
                    If Cabecera.Lines(bucle).ImporteLinea < 0 Then
                        .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                    End If
                Else
                    .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                    If Cabecera.Lines(bucle).ImporteLinea < 0 Then
                        .Fields("impor_neg") = .Fields("impor_neg") + ElImporte
                    End If
                End If
                .Fields("pesetas") = .Fields("pesetas") + ElImporte
            End If
            ' ****
            ' Cancelada o no
            ' ****
            If Cabecera.Lines(bucle).LineaCancelada Then
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    .Fields("totcance") = .Fields("totcance") - ElImporte
                '    .Fields("supertot") = .Fields("supertot") - ElImporte
                '    .Fields("cancela") = .Fields("cancela") - 1
                'Else
                    .Fields("totcance") = .Fields("totcance") + ElImporte
                    .Fields("supertot") = .Fields("supertot") + ElImporte
                    .Fields("cancela") = .Fields("cancela") + 1
                'End If
            Else
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    .Fields("Total") = .Fields("Total") + ElImporte
                '    .Fields("supertot") = .Fields("Total") + ElImporte
                'Else
                    .Fields("Total") = .Fields("Total") + ElImporte
                    .Fields("supertot") = .Fields("Total") + ElImporte
                'End If
            End If
            .Update
        End If
        Next bucle
        
        
    End With
    
finVend:
    If Err.Number <> 0 Then
        On Error GoTo 0
        CadenadeLog "Actualiza tabla vendedores:" & Err.Description
    End If
    
End Sub
'//////
'//////
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES HORARIOS A PARTIR DE LO RECIBIDO EN EL TICKET
'//////////////////////////////////////////////////////////////////////////////////
Function actualiza_tabla_gtHora(Base As dao.Database, Cabecera As tipo_cabecera)
    Dim rst As dao.Recordset
    Dim cnt As Long
    Dim sSQL As String
    Dim lclImporte As Double
    Dim LineaInicial As Boolean
    sSQL = "SELECT * " & _
           "FROM gthora " & _
           "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
           " AND cstr(hour(hora))=" & Chr(34) & CStr(Hour(Cabecera.hora)) & Chr(34)
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        If .EOF Then
            .AddNew
            .Fields("FECHA") = Cabecera.Fecha
            .Fields("HORA") = Cabecera.hora
            .Fields("SECCION") = 0
            .Fields("EQUIPO") = 0
            .Fields("VENDEDOR") = 0
            .Fields("TIQUET") = 0
            .Fields("OPERACIONES") = 0
            .Fields("OPERACIONES_CAN") = 0
            .Fields("OPERACIONES_NEG") = 0
            .Fields("IMPORTE") = 0
            .Fields("IMPORTE_CAN") = 0
            .Fields("IMPORTE_NEG") = 0
            .Fields("PESO") = 0
            .Fields("PESO_CAN") = 0
            .Fields("PESO_NEG") = 0
            .Fields("UNIDADES") = 0
            .Fields("UNIDADES_CAN") = 0
            .Fields("UNIDADES_NEG") = 0
            .Update
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    rst.Edit
    rst.Fields("TIQUET") = rst.Fields("TIQUET") + 1
    For cnt = 1 To Cabecera.LineasMensaje
    If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
        lclImporte = Cabecera.Lines(cnt).ImporteLinea
        
        If Cabecera.Lines(cnt).LineaCancelada Then
            rst.Fields("IMPORTE_CAN") = rst.Fields("IMPORTE_CAN") + lclImporte
            rst.Fields("OPERACIONES_CAN") = rst.Fields("OPERACIONES_CAN") + 1
            If Cabecera.Lines(cnt).EsPesado Then
                rst.Fields("PESO_CAN") = rst.Fields("PESO_CAN") + Cabecera.Lines(cnt).cantidad
            Else
                rst.Fields("UNIDADES_CAN") = rst.Fields("UNIDADES_CAN") + Cabecera.Lines(cnt).cantidad
            End If
        Else
            rst.Fields("IMPORTE") = rst.Fields("IMPORTE") + lclImporte 'redondea(rst.Fields("IMPORTE") + lclImporte)
            rst.Fields("OPERACIONES") = rst.Fields("OPERACIONES") + 1
            If Cabecera.Lines(cnt).EsPesado Then
                rst.Fields("PESO") = rst.Fields("PESO") + Cabecera.Lines(cnt).cantidad
            Else
                rst.Fields("UNIDADES") = rst.Fields("UNIDADES") + Cabecera.Lines(cnt).cantidad
            End If
        End If
    End If
    Next cnt
    rst.Update
    rst.Close
    Set rst = Nothing
End Function
'\\\\\\\\\\\
'///////////
Public Sub Actualiza_Tabla_Gtsecs(Base As dao.Database, Cabecera As tipo_cabecera)
'////////////////////////////////
' INCORPORA LOS DATOS DE TOTALES POR SECCIÓN A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////
Dim bucle As Long
Dim Registro As dao.Recordset
Dim RegFam As dao.Recordset
Dim rstseccion As dao.Recordset
Dim rstarticulo As dao.Recordset
Dim LaSeccion As Integer
Dim LaSeccionMaquina As Integer
Dim LaFamilia As Integer
Dim LaSubseccion As Integer
Dim ElPlu As Integer
Dim LaTara As Long
Dim LaCaducidad As Long
Dim ElDescriptivo As String
Dim ElIVA As Long
Dim ElPrecio As Double
Dim ElImporte As Double
Dim LineaInicial As Boolean

    LineaInicial = True
    '**********************
    ' Determina Sección de gestión
    '*******************
    For bucle = 1 To Cabecera.LineasMensaje
    '****************************
    ' Si no existe, nuevo registro
    '****************************
        'If Cabecera.Lines(bucle).LineaCancelada = False And (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
        If Cabecera.Lines(bucle).LineaCancelada = False And (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8 Or Cabecera.tipoTiquet = 12 Or ((Cabecera.tipoTiquet = 2 Or Cabecera.tipoTiquet = 5) And lAcumulaET) Or ((Cabecera.tipoTiquet = 3 Or Cabecera.tipoTiquet = 6) And lAcumulaSS)) Then
            ElImporte = Cabecera.Lines(bucle).ImporteLinea
            ElPrecio = Cabecera.Lines(bucle).precio
            Set rstarticulo = Base.OpenRecordset _
            ("select caducidad,tara,codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tipo_iva,des_plu1 from articulo where borrado=false and codigo=" & Cabecera.Lines(bucle).CodigoArticulo)
            With rstarticulo
                If Not .EOF Then
                    If Not IsNull(.Fields("codi_ident")) Then
                        LaSeccion = .Fields("codi_ident")
                    Else
                        LaSeccion = 0
                    End If
                    If Not IsNull(.Fields("secc_Maqui")) Then
                        LaSeccionMaquina = .Fields("secc_Maqui")
                    Else
                        LaSeccionMaquina = 0
                    End If
                    If Not IsNull(.Fields("codi_sub")) Then
                        LaSubseccion = .Fields("codi_sub")
                    Else
                        LaSubseccion = 0
                    End If
                    If Not IsNull(.Fields("codi_fam")) Then
                        LaFamilia = .Fields("codi_fam")
                    Else
                        LaFamilia = 0
                    End If
                    If Not IsNull(.Fields("plu")) Then
                        ElPlu = .Fields("plu")
                    Else
                        ElPlu = 0
                    End If
                    If Not IsNull(.Fields("tara")) Then
                        LaTara = .Fields("tara")
                    Else
                        LaTara = 0
                    End If
                    If Not IsNull(.Fields("caducidad")) Then
                        LaCaducidad = .Fields("caducidad")
                    Else
                        LaCaducidad = 0
                    End If
                    If Not IsNull(.Fields("tipo_iva")) Then
                        ElIVA = .Fields("tipo_iva")
                    Else
                        ElIVA = 0
                    End If
                    If .Fields("des_plu1") <> "" Then
                        ElDescriptivo = .Fields("des_plu1")
                    Else
                        ElDescriptivo = Space(25)
                    End If
                Else
                    Set rstseccion = Base.OpenRecordset("select * from seccion where secc_maqui=" & Cabecera.NMostrador)
                    If Not rstseccion.EOF Then
                        LaSeccion = rstseccion.Fields("codi_ident")
                    Else
                        LaSeccion = Cabecera.NMostrador
                    End If
                    LaSeccionMaquina = Cabecera.NMostrador
                    LaSubseccion = 0
                    LaFamilia = 0
                    ElPlu = 0
                    LaTara = 0
                    LaCaducidad = 0
                    ElIVA = 0
                    ElDescriptivo = CargaCadena(515) '"Varios..."
                End If
            End With
            
            Set Registro = Base.OpenRecordset _
            ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
            " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
            
            With Registro
                If .EOF Then
                    .AddNew
                    .Fields("Fecha") = Format(Cabecera.Fecha, "dd/mm/yy")
                    .Fields("codi_ident") = LaSeccion
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codi_fam") = LaFamilia
                    If LaSeccion = 0 Then
                        .Fields("txt_fam") = ""
                    Else
                        Set RegFam = Base.OpenRecordset("select txt_fam from fam_code where codi_ident=" & LaSeccion & " and codi_fam=" & LaFamilia)
                        If Not RegFam.EOF Then
                            .Fields("txt_fam") = RegFam.Fields("txt_fam")
                        Else
                            .Fields("txt_fam") = ""
                        End If
                    End If
                    If LaSeccion = 0 Then
                        .Fields("txt_sub") = ""
                    Else
                        Set RegFam = Base.OpenRecordset("select txt_sub from sub_code where codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion)
                        If Not RegFam.EOF Then
                            .Fields("txt_sub") = RegFam.Fields("txt_sub")
                        Else
                            .Fields("txt_sub") = ""
                        End If
                    End If
                    .Fields("secc_Maqui") = LaSeccionMaquina
                    .Fields("Modo") = Cabecera.tipoTiquet
                    .Fields("tiquets") = 0
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Update
                End If
            End With
                
            Set Registro = Base.OpenRecordset _
            ("select * from gtsecs where cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
            " and codi_ident=" & LaSeccion & " and codi_sub=" & LaSubseccion & _
            " and codi_fam=" & LaFamilia & " and modo=" & Cabecera.tipoTiquet)
                
            With Registro
                
                If Not Registro.EOF Then
                    .Edit
                End If
                
                If LineaInicial Then
                    .Fields("tiquets") = .Fields("tiquets") + 1
                    LineaInicial = False
                End If
                .Fields("operacion") = .Fields("operacion") + 1
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                '        If Cabecera.Lines(Bucle).EsPesado Then
                '            .Fields("Peso") = .Fields("Peso") + (-1) * Cabecera.Lines(Bucle).cantidad
                '        Else
                '            .Fields("Unidades") = .Fields("Unidades") + (-1) * Cabecera.Lines(Bucle).cantidad
                '        End If
                '        .Fields("pesetas") = .Fields("pesetas") + (-1) * ElImporte
                '    End If
                'Else
                    'If Cabecera.Lines(Bucle).LineaCancelada = False Then '12/07/11
                        If Cabecera.Lines(bucle).EsPesado Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(bucle).cantidad
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte
                    'End If
                'End If
                .Update
                .Close
            End With
        End If
    Next bucle
End Sub
'//////
'////////////////////////////////////////////////////////////////////////////////
'INCORPORA LOS DATOS DE TOTALES POR ARTÍCULO A PARTIR DE LO RECIBIDO EN UN TICKET
'////////////////////////////////////////////////////////////////////////////////
Public Sub Actualiza_Tabla_Gtarti(Base As dao.Database, Cabecera As tipo_cabecera)
'/////////////////////////////
'\ argTipoTqt => tipo de tiquet
'\ 1 --> 386
'\ 2 --> C9/C10
'\ 3 --> SC10
'\ 4 --> IV4TQ
'\\\\\\\\\\\\\
Dim rst As dao.Recordset
Dim rstAux As dao.Recordset
Dim LaSeccion As Integer
Dim LaSeccionMaquina As Integer
Dim LaFamilia As Integer
Dim LaSubseccion As Integer
Dim ElPlu As Integer
Dim LaTara As Long
Dim LaCaducidad As Long
Dim ElDescriptivo As String
Dim ElIVA As Long
Dim bucle As Long
Dim ElPrecio As Double
Dim ElImporte As Double
Dim MiFecha As Date
Dim sSQL As String
Dim ElCodigo As String
Dim sEan As String
Dim argTipoTqt As Integer

    On Error GoTo finArti
    argTipoTqt = 1
    For bucle = 1 To Cabecera.LineasMensaje
        If Cabecera.Lines(bucle).LineaCancelada = False And (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8 Or ((Cabecera.tipoTiquet = 2 Or Cabecera.tipoTiquet = 5) And lAcumulaET) Or ((Cabecera.tipoTiquet = 3 Or Cabecera.tipoTiquet = 6) And lAcumulaSS)) Then
            ElPrecio = Cabecera.Lines(bucle).precio
            ElImporte = Cabecera.Lines(bucle).ImporteLinea
            sSQL = "SELECT codigo,codi_ident,secc_maqui,codi_sub,codi_fam,plu,tara,caducidad,tipo_iva,des_plu1,art_cb " & _
                   "FROM articulo " & _
                   "WHERE borrado=false" & _
                   " AND codigo=" & Cabecera.Lines(bucle).CodigoArticulo
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                If Not .EOF Then
                    If Not IsNull(.Fields("codigo")) Then
                        ElCodigo = .Fields("codigo")
                    Else
                        ElCodigo = 0
                    End If
                    If Not IsNull(.Fields("plu")) Then
                        ElPlu = .Fields("plu")
                    Else
                        ElPlu = 0
                    End If
                    If Not IsNull(.Fields("codi_ident")) Then
                        LaSeccion = .Fields("codi_ident")
                    Else
                        LaSeccion = 0
                    End If
                    If Not IsNull(.Fields("secc_Maqui")) Then
                        LaSeccionMaquina = .Fields("secc_Maqui")
                    Else
                        LaSeccionMaquina = 0
                    End If
                    If Not IsNull(.Fields("codi_sub")) Then
                        LaSubseccion = .Fields("codi_sub")
                    Else
                        LaSubseccion = 0
                    End If
                    If Not IsNull(.Fields("codi_fam")) Then
                        LaFamilia = .Fields("codi_fam")
                    Else
                        LaFamilia = 0
                    End If
                    If Not IsNull(.Fields("tara")) Then
                        LaTara = .Fields("tara")
                    Else
                        LaTara = 0
                    End If
                    If Not IsNull(.Fields("caducidad")) Then
                        LaCaducidad = .Fields("caducidad")
                    Else
                        LaCaducidad = 0
                    End If
                    If Not IsNull(.Fields("tipo_iva")) Then
                        ElIVA = .Fields("tipo_iva")
                    Else
                        ElIVA = 0
                    End If
                    If .Fields("des_plu1") <> "" Then
                        ElDescriptivo = .Fields("des_plu1")
                    Else
                        ElDescriptivo = Space(25)
                    End If
                    If .Fields("art_cb") <> "" Then
                        sEan = .Fields("art_cb")
                           If Len(sEan) < 12 Then
                               sEan = sEan & Mid("000000000000", 1, 12 - Len(sEan))
                           End If
                           If Len(sEan) <> 13 Then
                            sEan = sEan & Format(DigitoControl(sEan), "0")
                           End If
                        
                    Else
                        sEan = Space(13)
                    End If
                Else
                    sSQL = "SELECT * " & _
                           "FROM seccion " & _
                           "WHERE secc_maqui=" & Cabecera.NMostrador
                    Set rstAux = Base.OpenRecordset(sSQL)
                    If Not rstAux.EOF Then
                        LaSeccion = rstAux.Fields("codi_ident")
                    Else
                        LaSeccion = Cabecera.NMostrador
                    End If
                    rstAux.Close
                    Set rstAux = Nothing
                    LaSeccionMaquina = Cabecera.NMostrador
                    ElCodigo = 0
                    ElPlu = 0
                    LaSubseccion = 0
                    LaFamilia = 0
                    LaTara = 0
                    LaCaducidad = 0
                    ElIVA = 0
                    ElDescriptivo = CargaCadena(515) '"Varios..."
                    sEan = Space(13)
                End If
                .Close
            End With
            Set rst = Nothing
            '////////////////
            If Cabecera.Lines(bucle).CodigoArticulo = 0 Then
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & LaSeccionMaquina

            Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                If .EOF Then
                    .AddNew
                    
                    .Fields("codigo") = ElCodigo
                    .Fields("plu") = ElPlu
                    .Fields("codi_ident") = LaSeccion
                    .Fields("Fecha") = Cabecera.Fecha
                    .Fields("des_plu1") = ElDescriptivo
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codi_fam") = LaFamilia
                    If ElCodigo = 0 Then
                        .Fields("txt_fam") = ""
                    Else
                        sSQL = "SELECT txt_fam " & _
                               "FROM fam_code " & _
                               "WHERE codi_ident=" & LaSeccion & _
                               " AND codi_fam=" & LaFamilia
                        Set rstAux = Base.OpenRecordset(sSQL)
                        If Not rstAux.EOF Then
                            .Fields("txt_fam") = rstAux.Fields("txt_fam")
                        Else
                            .Fields("txt_fam") = ""
                        End If
                        rstAux.Close
                        Set rstAux = Nothing
                    End If
                    '1.0.3 Leclerc
                    'If ElCodigo = 0 Then
                    '    .Fields("txt_sub") = ""
                    'Else
                    '    sSQL = "SELECT txt_sub " & _
                    '           "FROM sub_code " & _
                    '           "WHERE codi_ident=" & LaSeccion & _
                    '           " AND codi_sub=" & LaSubseccion
                    '    Set rstAux = Base.OpenRecordset(sSQL)
                    '    If Not rstAux.EOF Then
                    '        .Fields("txt_sub") = rstAux.Fields("txt_sub")
                    '    Else
                    '        .Fields("txt_sub") = ""
                    '    End If
                    '    rstAux.Close
                    '    Set rstAux = Nothing
                    'End If
                    '''''''''''''''''''''''''
                    If ElCodigo = 0 Then
                        If Not lClR Then
                            .Fields("txt_sub") = ""
                        Else
                            .Fields("txt_sub") = sEan
                        End If
                    Else
                        If Not lClR Then
                            sSQL = "SELECT txt_sub " & _
                                   "FROM sub_code " & _
                                   "WHERE codi_ident=" & LaSeccion & _
                                   " AND codi_sub=" & LaSubseccion
                            Set rstAux = Base.OpenRecordset(sSQL)
                            If Not rstAux.EOF Then
                                .Fields("txt_sub") = rstAux.Fields("txt_sub")
                            Else
                                .Fields("txt_sub") = ""
                            End If
                            rstAux.Close
                            Set rstAux = Nothing
                        Else
                            .Fields("txt_sub") = sEan
                        End If
                    End If
                    ''''''''''''''''''''''''''''''
                    .Fields("secc_Maqui") = LaSeccionMaquina
                    .Fields("precio") = ElPrecio
                    .Fields("operacion") = 0
                    .Fields("Peso") = 0
                    .Fields("Modo") = Cabecera.tipoTiquet
                    .Fields("Unidades") = 0
                    .Fields("pesetas") = 0
                    .Fields("tara") = LaTara
                    .Fields("caducidad") = LaCaducidad
                    .Fields("tipo_iva") = ElIVA
                    .Fields("Modificado") = Cabecera.Lines(bucle).SehaModificado
                    .Update
                End If
                .Close
            End With
            Set rst = Nothing
            If ElCodigo = 0 Then
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet & _
                       " AND secc_maqui=" & CStr(LaSeccionMaquina)
            Else
                sSQL = "SELECT * " & _
                       "FROM gtarti " & _
                       "WHERE cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & _
                       " AND codigo=" & Format(ElCodigo, "000000") & _
                       " AND modo=" & Cabecera.tipoTiquet
            End If
            Set rst = Base.OpenRecordset(sSQL)
            With rst
                .Edit
                .Fields("operacion") = .Fields("operacion") + 1
                'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
                '    If Cabecera.Lines(Bucle).LineaCancelada = False Then
                '        If Cabecera.Lines(Bucle).EsPesado = True Then
                '            .Fields("Peso") = .Fields("Peso") + (-1) * ((Cabecera.Lines(Bucle).cantidad) * Cabecera.Lines(Bucle).FactorPeso)
                '        Else
                '            .Fields("Unidades") = .Fields("Unidades") + (-1) * (Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso)
                '        End If
                '        .Fields("pesetas") = .Fields("pesetas") + (-1) * (ElImporte * Cabecera.Lines(Bucle).FactorPeso)
                '    End If
                'Else
                    'If Cabecera.Lines(Bucle).LineaCancelada = False Then
                        If Cabecera.Lines(bucle).EsPesado = True Then
                            .Fields("Peso") = .Fields("Peso") + Cabecera.Lines(bucle).cantidad
                        Else
                            .Fields("Unidades") = .Fields("Unidades") + Cabecera.Lines(bucle).cantidad
                        End If
                        .Fields("pesetas") = .Fields("pesetas") + ElImporte
                    'End If
                'End If
                .Update
                .Close
            End With
            Set rst = Nothing
        End If
    Next bucle
    
finArti:
    If Err.Number <> 0 Then
        On Error GoTo 0
        CadenadeLog "Actualiza tabla artículos:" & Err.Description
    End If
End Sub
'\\\\\\
'//////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' INCORPORA LA INFORMACION DE TOTALES DE VENDEDORES A PARTIR DE LO RECIBIDO EN LOS TICKETS
'/////////////////////////////////////////////////////////////////////////////////////////
'Balanza (2): Siempre 0#
'GRUPO (2): Siempre 32#
'CLIENTES(5): número de tickets
'IMPORTE_CLIENTES(7): importe de los tickets (suma total) con las dos ultimas posiciones como decimales.
'CLIENTES_MAYORISTAS (5): Siempre 0
'IMPORTE_MAYORISTAS (7): Siempre 0
'BORRADO (1): Siempre 1#
'ORDEN(4): numero de listado (este es el dato realmente importante).
'TOT_CLIENTES(5): numero total de cliente, igual que el campo CLIENTES.
'SIGNO: Siempre 1#
'TOT_IMPORTE(7): importe de los clientes, igual que el campo IMPORTE_CLIENTE
'CLI_MAYORISTA (5): Siempre 0
'SIGNO_MAYORISTA (1): Siempre 0
'IMPORTE_MAYORISTA (7): Siempre 0
'CLI_NEGATIVOS (5): Siempre 0
'SIGNO_NEGATIVOS (1): Siempre 0
'IMPORTE_NEGATIVOS (7): Siempre 0
'PRO_PESADOS(5): número de productos vendido por peso.
'SIGNO_PESADOS (1): Siempre 0
'IMPORTE_PESADOS(7): Importe de los artículos vendidos por peso.
'PRO_NO_PESADOS(5): número de productos vendido por unidades.
'SIGNO_MAYORISTA (1): Siempre 0
'IMPORTE_NO_PESADOS(7): Importe de los artículos vendidos por unidades
'PRO_DEVUELTO (5): Siempre 0
'SIGNO_DEVUELTO (1): Siempre 0
'IMPORTE_DEVUELTO (7): Siempre 0
'HORA(5): hora en que se genera el fichero (FORMATO HH:MM)
'FECHA(10): Fecha en que se genera el fichero (FORMATO DD-MM-AAAA)
'INTERNO(1): vacio.
'
'En el  ejemplo que te mando seria:
'BALANZA(2): 00
'GRUPO(2): 32
'CLIENTES(5): 00148
'IMPORTE_CLIENTES(7): 0057310 à siendo pues 573,10€.
'CLIENTES_MAYORISTAS(5): 00000
'IMPORTE_MAYORISTAS(7): 0000000
'BORRADO(1): 1
'ORDEN(4): 247
'TOT_CLIENTES(5): 00148
'SIGNO: 1
'TOT_IMPORTE(7): 0057310
'CLI_MAYORISTA(5): 00000
'SIGNO_MAYORISTA(1): 0
'IMPORTE_MAYORISTA(7): 0000000
'CLI_NEGATIVOS(5): 00000
'SIGNO_NEGATIVOS(1): 0
'IMPORTE_NEGATIVOS(7): 0000000
'PRO_PESADOS(5): 00085
'SIGNO_PESADOS(1): 0
'IMPORTE_PESADOS(7): 0023518 à Siendo pues 235,18€.
'PRO_NO_PESADOS(5): 24300
'SIGNO_MAYORISTA(1): 0
'IMPORTE_NO_PESADOS(7): 0033792 à Siendo pues 337,92€.
'PRO_DEVUELTO(5): 00000
'SIGNO_DEVUELTO(1): 0
'IMPORTE_DEVUELTO(7): 0000000
'HORA(5): 22:02
'FECHA(10): 21-03-2019
'INTERNO (1):
'Function actualiza_tabla_gtInfVen(Base As dao.Database, Cabecera As tipo_cabecera)
'Dim Encontrado As Boolean
'Dim lclSec As Integer
'Dim lclPrecio As Double
'Dim lclImporte As Double
'Dim lclFecha As String
'Dim rst As dao.Recordset
'Dim cnt As Long
'Dim sSQL As String
'    sSQL = "SELECT * " & _
'           "FROM seccion " & _
'           "WHERE secc_maqui=" & Cabecera.NMostrador
'    Set rst = Base.OpenRecordset(sSQL)
'    If Not rst.EOF Then
'        lclSec = rst.Fields("codi_ident")
'    Else
'        lclSec = Cabecera.NMostrador
'    End If
'    rst.Close
'    Set rst = Nothing
'    Encontrado = False
'    sSQL = "SELECT * " & _
'           "FROM gtInfVen " & _
'           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
'           " AND codi_ident=" & Val(lclSec) & _
'           " AND modo=" & Cabecera.tipoTiquet & _
'           " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34)
'    Set rst = Base.OpenRecordset(sSQL)
'    With rst
'        If .EOF Then
'            .AddNew
'            .Fields("ident_vend") = Cabecera.NVendedor
'            .Fields("codi_ident") = lclSec
'            .Fields("secc_Maqui") = Cabecera.NMostrador
'            .Fields("imporCod") = 0
'            .Fields("operCod") = 0
'            .Fields("imporDir") = 0
'            .Fields("operDir") = 0
'            .Fields("imporCan") = 0
'            .Fields("operCan") = 0
'            .Fields("imporNeg") = 0
'            .Fields("operNeg") = 0
'            .Fields("modo") = Cabecera.tipoTiquet
'            .Fields("fecha") = Cabecera.Fecha
'            .Fields("ntotal") = getNtotalTS(Cabecera.Fecha)
'            .Update
'
'        End If
'        .Close
'    End With
'    Set rst = Nothing
'    Set rst = Base.OpenRecordset(sSQL)
'    With rst
'        For cnt = 1 To Cabecera.LineasMensaje
'        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
'            lclImporte = Cabecera.Lines(cnt).ImporteLinea
'            .Edit
'            '//////////////
'            'Cancelada o no
'            '//////////////
'            If Cabecera.Lines(cnt).LineaCancelada Then
'                .Fields("imporCan") = .Fields("imporCan") + lclImporte
'                .Fields("operCan") = .Fields("operCan") + 1
'            Else
'                If Cabecera.Lines(cnt).CodigoArticulo = 0 Then
'                    .Fields("imporDir") = .Fields("imporDir") + lclImporte
'                    .Fields("operDir") = .Fields("operDir") + 1
'                Else
'                    .Fields("imporCod") = .Fields("imporCod") + lclImporte
'                    .Fields("operCod") = .Fields("operCod") + 1
'                End If
'                If Cabecera.Lines(cnt).EsPesado Then
'                    .Fields("imporpeso") = .Fields("imporpeso") + lclImporte
'                    .Fields("operpeso") = .Fields("operpeso") + 1
'                Else
'                    .Fields("imporuni") = .Fields("imporuni") + lclImporte
'                    .Fields("operuni") = .Fields("operuni") + 1
'                End If
'
'            End If
'            .Update
'        End If
'        Next cnt
'        .Close
'    End With
'    Set rst = Nothing
'End Function
Function actualiza_tabla_gtInfVen(Base As dao.Database, Cabecera As tipo_cabecera)
Dim eNCONTRADO As Boolean
Dim lclSec As Integer
Dim lclPrecio As Double
Dim lclImporte As Double
Dim lclFecha As String
Dim rst As dao.Recordset
Dim cnt As Long
Dim sSQL As String
    sSQL = "SELECT * " & _
           "FROM seccion " & _
           "WHERE secc_maqui=" & Cabecera.NMostrador
    Set rst = Base.OpenRecordset(sSQL)
    If Not rst.EOF Then
        lclSec = rst.Fields("codi_ident")
    Else
        lclSec = Cabecera.NMostrador
    End If
    rst.Close
    Set rst = Nothing
    eNCONTRADO = False
    sSQL = "SELECT * " & _
           "FROM gtInfVen " & _
           "WHERE ident_vend=" & Val(Cabecera.NVendedor) & _
           " AND codi_ident=" & Val(lclSec) & _
           " AND modo=" & Cabecera.tipoTiquet & _
           " AND cstr(cdate(fecha))=" & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34)
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        If .EOF Then
            .AddNew
            .Fields("ident_vend") = Cabecera.NVendedor
            .Fields("codi_ident") = lclSec
            .Fields("secc_Maqui") = Cabecera.NMostrador
            .Fields("imporCod") = 0
            .Fields("operCod") = 0
            .Fields("imporDir") = 0
            .Fields("operDir") = 0
            .Fields("imporCan") = 0
            .Fields("operCan") = 0
            .Fields("imporNeg") = 0
            .Fields("operNeg") = 0
            .Fields("modo") = Cabecera.tipoTiquet
            .Fields("fecha") = Cabecera.Fecha
            .Fields("imporpeso") = 0
            .Fields("operpeso") = 0
            .Fields("imporuni") = 0
            .Fields("operuni") = 0
#If BALSAM = False Then
            .Fields("ntotal") = getNtotalTS(Cabecera.Fecha)
#End If
            .Update
            
        End If
        .Close
    End With
    Set rst = Nothing
    Set rst = Base.OpenRecordset(sSQL)
    With rst
        For cnt = 1 To Cabecera.LineasMensaje
        If (Cabecera.tipoTiquet = 1 Or Cabecera.tipoTiquet = 4 Or Cabecera.tipoTiquet = 7 Or Cabecera.tipoTiquet = 8) Then
            lclImporte = Cabecera.Lines(cnt).ImporteLinea
            .Edit
            '//////////////
            'Cancelada o no
            '//////////////
            If Cabecera.Lines(cnt).LineaCancelada Then
                .Fields("imporCan") = .Fields("imporCan") + lclImporte
                .Fields("operCan") = .Fields("operCan") + 1
            Else
                If Cabecera.Lines(cnt).CodigoArticulo = 0 Then
                    .Fields("imporDir") = .Fields("imporDir") + lclImporte
                    .Fields("operDir") = .Fields("operDir") + 1
                Else
                    .Fields("imporCod") = .Fields("imporCod") + lclImporte
                    .Fields("operCod") = .Fields("operCod") + 1
                End If
                If Cabecera.Lines(cnt).EsPesado Then
                    .Fields("imporpeso") = .Fields("imporpeso") + lclImporte
                    .Fields("operpeso") = .Fields("operpeso") + 1
                Else
                    .Fields("imporuni") = .Fields("imporuni") + lclImporte
                    .Fields("operuni") = .Fields("operuni") + 1
                End If
                
            End If
            .Update
        End If
        Next cnt
        .Close
    End With
    Set rst = Nothing
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' 1º TABLA TICKETS (datos de cada línea recibida)
'////////////////////////////////////////////////
' Tipos de Tiquet (typtic)
' 1 --> Rollo
' 2 --> Super
' 3 --> Envasado
' 4 --> Autoservicio
' 5 --> Rollo Local
' 6 --> Super Local
' 7 --> Envasado Local
' 8 --> Autoservicio Local
'/////////////////////////
Public Sub Actualiza_Tabla_Tickets(Base As dao.Database, Cabecera As tipo_cabecera)
Dim BucleGrande As Integer
Dim D_FECHA As Date
Dim D_HORA As Date
Dim Maximo As Integer
Dim bucle As Long
Dim MiFactor As Double
Dim Registro As dao.Recordset
Dim Buffer As String
Dim MiCadenaDebug As String
Dim RegArt As dao.Recordset
Dim Que_base As String
Dim nMin As Integer

On Error GoTo finTick

nMin = 1

#If BALSAM = False Then
    Set RegArt = Base.OpenRecordset("select codigo,plu,des_plu1,codi_fam from articulo where borrado=false")
#End If

    Buffer = Cabecera.CadenaEstadoTiquet
    
    Que_base = "GA"
    
    If Que_base = "GA" Then
        MiCadenaDebug = "Euroscale / "
        Set Registro = Base.OpenRecordset("select * from tickets where cstr(cdate(date))=" _
        & Chr(34) & CStr(CDate(Cabecera.Fecha)) & Chr(34) & " and hour=" & Chr(34) & Format(Cabecera.hora, "hh:mm:ss") & Chr(34) & _
        " and numbal=" & Cabecera.NBalanza & " and vendor=" & Chr(34) & Cabecera.NVendedor & Chr(34) _
        & "and nume=" & Chr(34) & Format(Cabecera.Ntiquet, "000000") & Chr(34) & " and typtic=" & Chr(34) & Cabecera.tipoTiquet & Chr(34))
        
        'CadenadeLog MiCadenaDebug
        
    End If
    
    For bucle = 1 To Cabecera.LineasMensaje
        Buffer = Cabecera.CadenaEstadoTiquet
        If (Cabecera.Lines(bucle).EsPesado) Then
            Buffer = Buffer & "peso "
        Else
            Buffer = Buffer & "unidades "
        End If
        Maximo = 1
        nMin = 1
        If lBoka Then
            If Cabecera.Lines(bucle).LineaCancelada Then
                Maximo = 2
                nMin = 1
            End If
        End If
        For BucleGrande = nMin To Maximo
            With Registro
                .AddNew
                If DebugActivo Then CadenadeLog (MiCadenaDebug & "Almacenando Linea de Tiquet : " _
                & Cabecera.Lines(bucle).NumeroLinea & "Tiquet : " & Cabecera.Ntiquet & "Sec : " & Cabecera.NMostrador & "Bal : " & Cabecera.NBalanza _
                & "Importe : " & Cabecera.Lines(bucle).ImporteLinea)
                .Fields("Date") = Format(Cabecera.Fecha, "dd/mm/yy")
                .Fields("Hour") = Format(Cabecera.hora, "hh:mm:ss")
                .Fields("d_fecha") = Cabecera.Fecha
                .Fields("d_hora") = Cabecera.hora
                .Fields("numbal") = Cabecera.NBalanza
                .Fields("Vendor") = Cabecera.NVendedor
                .Fields("IVA") = Cabecera.Lines(bucle).IVA
                .Fields("porcentaje") = Cabecera.Lines(bucle).porcentaje
                .Fields("nume") = Format(Cabecera.Ntiquet, "00000")
                .Fields("numlin") = Cabecera.Lines(bucle).NumeroLinea
                .Fields("SECC") = Cabecera.NMostrador
                .Fields("Code") = Cabecera.Lines(bucle).CodigoArticulo
                .Fields("factura") = Cabecera.factura 'no se hacía
#If BALSAM = False Then
                RegArt.FindFirst "codigo=" & Cabecera.Lines(bucle).CodigoArticulo
                If RegArt.EOF Or RegArt.NoMatch Then
                    .Fields("plu") = 0
                    .Fields("Item") = ""
                    'c2f invicta
                    .Fields("DPT") = "0000"
                Else
                    .Fields("plu") = RegArt.Fields("plu")
                    .Fields("Item") = RegArt.Fields("des_plu1")
                    'c2f invicta
                    .Fields("DPT") = Format(RegArt.Fields("CODI_FAM"), "0000")
                End If
#Else
                    .Fields("plu") = Cabecera.Lines(bucle).Plu
                    .Fields("Item") = Cabecera.Lines(bucle).desc
                    'c2f invicta
                    .Fields("DPT") = Cabecera.Lines(bucle).familia

#End If
                .Fields("Price") = (Cabecera.Lines(bucle).precio) ' / 100)
                If BucleGrande = 2 Then
                    MiFactor = (-1)
                Else
                    MiFactor = 1
                End If
                'If ((Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128) Then
                '    .Fields("amount") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso)) ' / 100)
                '    If Cabecera.Lines(Bucle).EsPesado Then
                '        .Fields("Weight") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                '        .Fields("units") = 0
                '    Else
                '        .Fields("Weight") = 0
                '        .Fields("units") = MiFactor * (-1) * ((Cabecera.Lines(Bucle).cantidad * Cabecera.Lines(Bucle).FactorPeso))
                '    End If
                'Else
                    '.Fields("amount") = MiFactor * (Cabecera.Lines(Bucle).ImporteLinea * Cabecera.Lines(Bucle).FactorPeso) ' / 100
                    .Fields("amount") = MiFactor * (Cabecera.Lines(bucle).ImporteLinea) ' / 100
                    If Cabecera.Lines(bucle).EsPesado Then
                        .Fields("Weight") = MiFactor * (Cabecera.Lines(bucle).cantidad * Cabecera.Lines(bucle).FactorPeso)
                        .Fields("units") = 0
                    Else
                        .Fields("Weight") = 0
                        .Fields("units") = MiFactor * (Cabecera.Lines(bucle).cantidad * Cabecera.Lines(bucle).FactorPeso)
                    End If
                'End If
                .Fields("STPETIC") = Cabecera.EstadoTiquet
                If Cabecera.Lines(bucle).LineaCancelada = True Then
                   .Fields("lincan") = 1
                Else
                   .Fields("lincan") = 0
                End If
                
                .Fields("TYPTIC") = Cabecera.tipoTiquet
                .Fields("reserv1") = Cabecera.Lines(bucle).FichaVacuno
                
                .Fields("importe_bruto") = Cabecera.Lines(bucle).ImporteBruto   '.Fields("amount") + Cabecera.Lines(Bucle).DescuentoLinea
                
                
                .Fields("descuento") = Cabecera.Lines(bucle).DescuentoLinea
                
                If Cabecera.lote <> "" Then
                    .Fields("numlote") = Cabecera.lote
                Else
                    .Fields("numlote") = " "
                End If
                '1.2.6
                If Not IsNull(Cabecera.NCliente) Then
                    .Fields("clienteT") = Cabecera.NCliente
                    .Fields("cliente") = Cabecera.NCliente
                Else
                    .Fields("cliente") = 0
                    .Fields("clienteT") = 0
                End If
                ''''''
                If Cabecera.Lines(bucle).SehaModificado Then
                    .Fields("codi_pes") = "0"
                End If
                .Fields("reserv2") = "/"
                .Fields("reserv3") = "NO"
                .Fields("tipoPago") = Cabecera.tipoPago
                
                .Update
                
                
            End With
            
        Next BucleGrande
    Next bucle
    
finTick:
    
    'If (Cabecera.EstadoTiquet And 64) = 64 Or (Cabecera.EstadoTiquet And 128) = 128 Then
    '    Cabecera.Incluir_en_Base = False
    'End If
    
    If Err.Number <> 0 Then
        
        CadenadeLog "Actualiza tabla Tickets:" & Err.Description
        On Error GoTo 0
    End If
End Sub

'\\\\\\\\\\\
'///////////
Public Sub Actualiza_Tabla_Log(Base As dao.Database, Cabecera As tipo_cabecera)
    '////////////////////////////////
    ' Genera log de cambios de precio o rechazos
    '////////////////////////////////
    Dim Registro As dao.Recordset
    Dim rstarticulo As dao.Recordset
    Dim bucle As Long
    Dim Buffer As String
    
    Dim LaSeccion As Integer
    Dim LaSubseccion As Integer
    Dim ElPlu As Integer
    Dim LaFamilia As Integer
    Dim LaCaducidad As Long
    Dim LaTara As Long
    Dim ElDescriptivo As String
    Dim MyPrec As Double
      
    For bucle = 1 To Cabecera.LineasMensaje
        
        'If UsaEuro Then
        '    MyPrec = Cabecera.Lines(bucle).precio / 100
        'Else
        '    MyPrec = Cabecera.Lines(bucle).precio / (10 ^ decimales)
        'End If
        MyPrec = Cabecera.Lines(bucle).precio
        Set Registro = Base.OpenRecordset("select * from log where codigo=" & Cabecera.Lines(bucle).CodigoArticulo & " and precio=" & Str(MyPrec))
        If Registro.EOF Then
            With Registro
                If Cabecera.Lines(bucle).SehaModificado Then
                    '**********************************
                    ' Obtiene datos del Artículo
                    '**********************************
                    Set rstarticulo = Base.OpenRecordset _
                    ("select codigo,codi_ident,codi_sub,codi_fam,plu,caducidad,tara,des_plu1 from articulo where borrado=false and codigo=" & Cabecera.Lines(bucle).CodigoArticulo)
                    
                    With rstarticulo
                        If Not .EOF Then
                            LaSeccion = .Fields("codi_ident")
                            If Not IsNull(.Fields("codi_sub")) Then
                                LaSubseccion = .Fields("codi_sub")
                            Else
                                LaSubseccion = 0
                            End If
                            If Not IsNull(.Fields("plu")) Then
                                ElPlu = .Fields("plu")
                            Else
                                ElPlu = 0
                            End If
                            If Not IsNull(.Fields("codi_fam")) Then
                                LaFamilia = .Fields("codi_fam")
                            Else
                                LaFamilia = 0
                            End If
                            If Not IsNull(.Fields("caducidad")) Then
                                LaCaducidad = .Fields("caducidad")
                            Else
                                LaCaducidad = 0
                            End If
                            If Not IsNull(.Fields("tara")) Then
                                LaTara = .Fields("tara")
                            Else
                                LaTara = 0
                            End If
                            If .Fields("des_plu1") <> "" Then
                                ElDescriptivo = .Fields("des_plu1")
                            Else
                                ElDescriptivo = Space(25)
                            End If
                        Else
                            LaSeccion = 0
                            LaSubseccion = 0
                            ElPlu = 0
                            LaFamilia = 0
                            LaCaducidad = 0
                            LaTara = 0
                            ElDescriptivo = Space(25)
                        End If
                        .Close
                    End With
                    '**********************************
                    '**********************************
                    ' Crea la cadena de log
                    '**********************************
                    Buffer = "LOG "
                    Buffer = Buffer & Format(Cabecera.Fecha, "ddmmyy") & ","
                    Buffer = Buffer & "3," & Format(Cabecera.hora, "hhmm") & ","
                    Buffer = Buffer & "0," & Format(Cabecera.Lines(bucle).CodigoArticulo, "000000") & ","
                    Buffer = Buffer & "000," & Format(LaSeccion, "000") & ","
                    If UsaEuro Then
                        Buffer = Buffer & Format(MyPrec, "000000") & ","
                    Else
                        Buffer = Buffer & Format(MyPrec, "000000") & ","
                    End If
                    Buffer = Buffer & Trim(ElDescriptivo) & ","
                    Buffer = Buffer & Format(Cabecera.NMostrador, "00")
                    '**********************************
                    .AddNew
                    .Fields("codi_ident") = LaSeccion
                    .Fields("codi_sub") = LaSubseccion
                    .Fields("codigo") = Cabecera.Lines(bucle).CodigoArticulo
                    .Fields("plu") = ElPlu
                    If UsaEuro Then
                        .Fields("precio") = Cabecera.Lines(bucle).precio / 100
                    Else
                        .Fields("precio") = Cabecera.Lines(bucle).precio / (10 ^ decimales)
                    End If
                    If Cabecera.Lines(bucle).EsPesado Then
                        .Fields("tipo_ven") = "W"
                    Else
                        .Fields("tipo_ven") = "U"
                    End If
                    .Fields("codi_fam") = LaFamilia
                    .Fields("caducidad") = LaCaducidad
                    .Fields("tara") = LaTara
                    .Fields("des_plu1") = ElDescriptivo
                    .Fields("Comentario") = "Desde Balanza."
                    .Fields("Texto") = Buffer
                    .Fields("Fecha") = Cabecera.Fecha
                    .Update
                End If
            End With
        End If
        Registro.Close
    Next bucle
    

        
End Sub

Public Function Alta_Teclas(LaTecla As DB_Teclas) As Integer
    '********************
    'Para Global.dat
    'Siempre se está considerando que los resgistros "6" (10 de longitud) son para sección
    ' y que los registro "@" son para equipos.
    '********************
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Retorno As Integer
    Dim EsBaja As Boolean
    Dim ssSecs(1) As Long
    Dim nResp As Variant
    Dim nf As Integer
    
'CadenadeLog "-->teclas1"
    EsBaja = False
    Set Base = AbrirBase()
    If LaTecla.CodArticulo <> 0 Then
        Set Registro = Base.OpenRecordset("select * from articulo where codigo=" & _
        LaTecla.CodArticulo & " and borrado=false")
        If Registro.EOF Then
            CerrarBase Base
            Retorno = 2
            GoTo fin
        Else
            LaTecla.PluArticulo = Registro!Plu
            LaTecla.SecArticulo = LaTecla.Destino_NMostrador '3.2.5 Registro!secc_Maqui
            LaTecla.Destino_NMostrador = LaTecla.Destino_NMostrador '3.2.5 !codi_ident
            LaTecla.Descripcion = Mid(Registro!des_plu1, 1, 25)
            'If lUpperNW Then
            '1.8.3
            'If lUpper Then
            '    nf = ((LaTecla.Destino_NMostrador - 1) * 10 + LaTecla.Destino_Tabla) + 1
            '    'nResp = Alta_Familias(nf, "", 0, "", False, , nf)
            '    nResp = Alta_Familias(nf, "", 0, "", False, , LaTecla.Destino_Tabla + 1)
            '    If nResp = 0 Then
            '        ssSecs(0) = LaTecla.SecArticulo
            '        nResp = Alta_Familias_Counter(nf, "", 0, ssSecs(), False, , "", LaTecla.SecArticulo)
            '    End If
            '    'LaTecla.Section = 1
            '    'LaTecla.Destino_Tabla = Registro!codi_fam
            '    LaTecla.Section = 1
            '    Registro.Edit
            '    Registro.Fields("codi_fam") = nf
            '    Registro.Update
            'End If
            '''''''
        End If
    Else
        Set Registro = Base.OpenRecordset("select * from articulo where plu=" & _
        LaTecla.PluArticulo & " and codi_ident=" & LaTecla.Destino_NMostrador & " and borrado=false")
        If Registro.EOF Then
            CerrarBase Base
            Retorno = 2
            GoTo fin
        Else
            LaTecla.CodArticulo = Registro!codigo
            LaTecla.SecArticulo = LaTecla.Destino_NMostrador '3.2.5 !secc_Maqui
            LaTecla.Section = LaTecla.Destino_NMostrador '3.2.5!secc_Maqui
            LaTecla.Destino_NMostrador = Registro!codi_ident
            LaTecla.Descripcion = Mid(Registro!des_plu1, 1, 25)
            'If lUpperNW Then
            '1.8.3
            'If lUpper Then
            '    nf = ((LaTecla.Destino_NMostrador - 1) * LaTecla.Destino_Tabla) + 1
            '    'nResp = Alta_Familias(nf, "", 0, "", False, , nf) 'LaTecla.Destino_Tabla + 1
            '    nResp = Alta_Familias(nf, "", 0, "", False, , LaTecla.Destino_Tabla + 1)
            '    If nResp = 0 Then
            '        ssSecs(0) = LaTecla.SecArticulo
            '        nResp = Alta_Familias_Counter(nf, "", 0, ssSecs(), False, , "", LaTecla.SecArticulo)
            '    End If
            '    'LaTecla.Section = 1
            '    'LaTecla.Destino_Tabla = Registro!codi_fam
            '    LaTecla.Section = 1
            '    Registro.Edit
            '    Registro.Fields("codi_fam") = nf
            '    Registro.Update
            'End If
        End If
    
    End If
    Registro.Close
    Set Registro = Nothing
    '1.8.3
    'If lUpper Then
    '
    '    Base.Execute "delete from teclas where codigo=" & CStr(LaTecla.CodArticulo)
    '    Base.Execute "update articulo set codi_fam=" & CStr(nf) & " where codigo=" & CStr(LaTecla.CodArticulo)
    '    Base.Execute "update articulo set tran_plu='' where codigo=" & CStr(LaTecla.CodArticulo)
    '
    '    Set Registro = Base.OpenRecordset("select * from teclas where TRAN_TECSC10='" & _
    '    LaTecla.Section & "' and codi_tec=" & LaTecla.Destino_CodiTec & " and tabla=" & nf & " and secc_maqui=" & LaTecla.SecArticulo & _
    '    " and val(tran_tecsc10)=" & LaTecla.Section & " and borrado=false")
    '
    'Else
    
    '''09/09/17 hablar, era TRAN_TECSC10='" & LaTecla.Section
    If lAsigntec = False Then
    Base.Execute "delete from teclas where TRAN_TECSC10='" & _
        LaTecla.Destino_NMostrador & "' and codigo=" & LaTecla.CodArticulo & " and tabla=" & LaTecla.Destino_Tabla & _
        " and borrado=false"
    End If
        Set Registro = Base.OpenRecordset("select * from teclas where TRAN_TECSC10='" & _
        LaTecla.Destino_NMostrador & "' and codi_tec=" & LaTecla.Destino_CodiTec & " and tabla=" & LaTecla.Destino_Tabla & _
        " and borrado=false")
    
    'End If
    '''''''
    With Registro
        If .EOF Then
            Retorno = 0
            .AddNew
        Else
            Retorno = 1
            .Edit
        End If
        !codi_ident = LaTecla.Destino_NMostrador
        !secc_maqui = LaTecla.Destino_NMostrador '3.2.5LaTecla.SecArticulo
        !codigo = LaTecla.CodArticulo
        !Plu = LaTecla.PluArticulo
        !desc_plu = Mid(LaTecla.Descripcion, 1, 25)
        !secc_eqp = 0
        '1.8.3
        'If lUpper Then
        '    !Tabla = nf
        'Else
            !Tabla = LaTecla.Destino_Tabla
        'End If
        '''''''
        !codi_tec = LaTecla.Destino_CodiTec
        '4.5.35
        'If LaTecla.lEs165 Then
        '    !prog_tec = "165"
        'End If
        '''''''
        'if LaTecla.Tipo = "" Then
            !prog_tec = "TOUCH"
        'Else
        '    !prog_tec = LaTecla.Tipo
        'End If
        
        If LaTecla.Section = 0 Then
            !tran_tecsc10 = CStr(LaTecla.Destino_NMostrador) ' CStr(Reg2!secc_maqui) no estaba
        Else
            !tran_tecsc10 = CStr(LaTecla.Section)   '09/09/17 CStr(LaTecla.Section), solo estaba esto
        End If
        
        !tran_tec = ""
        !borrado = False
        .Update
    End With
    CerrarBase Base
fin:
'CadenadeLog "-->teclas"
    Alta_Teclas = Retorno
        
End Function
Public Function Baja_Teclas(LaTecla As DB_Teclas) As Integer
    '********************
    ' Los datos necesarios para esta función son:
    ' Para baja por balanza :
    '
    '  Destino_CodiTec
    '  Destino_NBalanza
    '  Destino_NMostrador=0 (obligatorio)
    '  destino_Tabla
    '
    ' Para baja por sección
    '
    '  Destino_CodiTec
    '  Destino_NBalanza=0 (obligatorio)
    '  Destino_NMostrador
    '  destino_Tabla
    '********************
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim Retorno As Integer
    Dim lEsTest As Boolean
    
If Dir(App.Path & "\test.tst") <> "" Then
    lEsTest = True
End If
    
    '*******************
    ' Condiciones de error
    '*******************
    If LaTecla.Destino_NBalanza <> 0 And LaTecla.Destino_NMostrador <> 0 Then
        Retorno = 3
        GoTo fin
    End If
    '*******************
    Set Base = AbrirBase()
    'If LaTecla.Destino_NBalanza = 0 Then
        Set Registro = Base.OpenRecordset("select * from teclas where TRAN_TECSC10='" & _
        LaTecla.Section & "' and tabla=" & LaTecla.Destino_Tabla & _
        " and codi_tec=" & LaTecla.Destino_CodiTec) '& " and isnull(numero_eqp)) ' and prog_tec='" & LaTecla.Tipo & "' and borrado=false")
    'Else
    '    Set Registro = Base.OpenRecordset("select * from teclas where TRAN_TECSC10='" & _
    '    LaTecla.Section & "' and numero_eqp=" & _
    '    LaTecla.Destino_NBalanza & " and tabla=" & LaTecla.Destino_Tabla & _
    '    " and codi_tec=" & LaTecla.Destino_CodiTec & " and prog_tec='" & LaTecla.Tipo & "' and borrado=false")
    'End If
    
    With Registro
        If .EOF Then
            Retorno = 4
            CerrarBase Base
            GoTo fin
        Else
            'LogAdministracion "Baja de tecla de PLU"
            If Dir(App.Path & "\asign.tst") <> "" Then
                '.Delete ...26/07/2016 2.0.1
                .Edit
                !borrado = True
                !tran_tec = ""
                If lEsTest = False Then
                    !codigo = 0
                End If
                .Update
                
            Else
                .Edit
                !borrado = True
                !tran_tec = ""
                If lEsTest = False Then
                    !codigo = 0
                End If
                .Update
            End If
            Retorno = 0
        End If
    End With
    CerrarBase Base
fin:
    Baja_Teclas = Retorno
        
End Function

Public Sub transferirTABLA(ByVal BaseOri As String, ByVal BaseDes As String)
    Dim BaseO As dao.Database
    Dim BaseD As dao.Database
    Dim RegistroO As dao.Recordset
    Dim RegistroD As dao.Recordset
    Dim bucle As Integer
    Dim Bucle1 As Integer
    Dim Eventos As Integer
    Dim Adicional As String
    Dim MyCont As Long
    Dim sNameField() As String
    Dim sTmp As String
    Dim Tabla As String
    
    sTmp = Base_General
    Base_General = BaseDes
    FrmCambiaFormato.CambiarFormatoBase
    Base_General = sTmp

    For Bucle1 = 1 To 8
    
    Select Case Bucle1
    Case 1
    Tabla = "seccion"
    Case 2
    Tabla = "fam_code"
    Case 3
    Tabla = "familias"
    Case 4
    Tabla = "articulo"
    Case 5
    Tabla = "teclas"
    Case 6
    Tabla = "text15"
    Case 7
    Tabla = "lintxt2040"
    Case 8
    Tabla = "tramos"
    
    End Select
    CadenadeLog "Base Origen:" & BaseOri
    CadenadeLog "transferencia:" & Tabla
    CadenadeLog "Base Origen:" & BaseDes
    On Error GoTo final
    MyCont = 0
    Set BaseO = dao.OpenDatabase(BaseOri)
    Set RegistroO = BaseO.OpenRecordset("select * from " & Tabla)
    
    If Not RegistroO.EOF Then
    
        With RegistroO
            If .Fields.Count > 0 Then
                For bucle = 0 To .Fields.Count - 1
                    MyCont = MyCont + 1
                    ReDim Preserve sNameField(MyCont)
                    sNameField(MyCont) = .Fields(bucle).name
                Next bucle
            End If
        End With
        RegistroO.MoveFirst
        Set BaseD = dao.OpenDatabase(BaseDes)
        BaseD.Execute "delete from " & Tabla
        Set RegistroD = BaseD.OpenRecordset("select * from " & Tabla)
        Do While Not RegistroO.EOF
            RegistroD.AddNew
            For bucle = 1 To MyCont
                RegistroD.Fields(sNameField(bucle)) = RegistroO.Fields(sNameField(bucle))
            Next bucle
            RegistroD.Update
            RegistroO.Movenext
        Loop
        RegistroO.Close
        RegistroD.Close
        BaseO.Close
        BaseD.Close
        Set RegistroO = Nothing
        Set RegistroD = Nothing
        Set BaseO = Nothing
        Set BaseD = Nothing
    
    Else
        RegistroO.Close
        BaseO.Close
        Set RegistroO = Nothing
        Set BaseO = Nothing
    End If
    
    On Error GoTo 0
    
    Next Bucle1
final:
    If Err.Number <> 0 Then
        CadenadeLog "error en TransferirTabla:" & Err.Description
    End If
    On Error GoTo 0
End Sub
''''HiperUsera
Public Function Crea_1_7_1_1() As Integer
Dim db As dao.Database
Dim tdf As TableDef
Dim cnt As Long
    Do_Events
    Set db = AbrirBase 'abrirbase 'OpenDatabase(Base_General)
    Set tdf = db.TableDefs("linTxt2040")
    With tdf
        .Fields.Append .CreateField("txt_41", dbText, 80)
        .Fields.Append .CreateField("txt_42", dbText, 80)
        .Fields.Append .CreateField("txt_43", dbText, 80)
        .Fields.Append .CreateField("txt_44", dbText, 80)
        .Fields.Append .CreateField("txt_45", dbText, 80)
        .Fields.Append .CreateField("txt_46", dbText, 80)
        .Fields.Append .CreateField("txt_47", dbText, 80)
        .Fields.Append .CreateField("txt_48", dbText, 80)
        .Fields.Append .CreateField("txt_49", dbText, 80)
        .Fields.Append .CreateField("txt_50", dbText, 80)
        .Fields.Append .CreateField("txt_51", dbText, 80)
        .Fields.Append .CreateField("txt_52", dbText, 80)
        .Fields.Append .CreateField("txt_53", dbText, 80)
        .Fields.Append .CreateField("txt_54", dbText, 80)
        .Fields.Append .CreateField("txt_55", dbText, 80)
        .Fields.Append .CreateField("txt_56", dbText, 80)
        .Fields.Append .CreateField("txt_57", dbText, 80)
        .Fields.Append .CreateField("txt_58", dbText, 80)
        .Fields.Append .CreateField("txt_59", dbText, 80)
        .Fields.Append .CreateField("txt_60", dbText, 80)

        .Fields.Append .CreateField("txt_61", dbText, 80)
        .Fields.Append .CreateField("txt_62", dbText, 80)
        .Fields.Append .CreateField("txt_63", dbText, 80)
        .Fields.Append .CreateField("txt_64", dbText, 80)
        .Fields.Append .CreateField("txt_65", dbText, 80)
        .Fields.Append .CreateField("txt_66", dbText, 80)
        .Fields.Append .CreateField("txt_67", dbText, 80)
        .Fields.Append .CreateField("txt_68", dbText, 80)
        .Fields.Append .CreateField("txt_69", dbText, 80)
        .Fields.Append .CreateField("txt_70", dbText, 80)

        .Fields.Append .CreateField("txt_71", dbText, 80)
        .Fields.Append .CreateField("txt_72", dbText, 80)
        .Fields.Append .CreateField("txt_73", dbText, 80)
        .Fields.Append .CreateField("txt_74", dbText, 80)
        .Fields.Append .CreateField("txt_75", dbText, 80)
        .Fields.Append .CreateField("txt_76", dbText, 80)
        .Fields.Append .CreateField("txt_77", dbText, 80)
        .Fields.Append .CreateField("txt_78", dbText, 80)
        .Fields.Append .CreateField("txt_79", dbText, 80)
        .Fields.Append .CreateField("txt_80", dbText, 80)

        '.Fields.Append .CreateField("txt_81", dbText, 25)
        '.Fields.Append .CreateField("txt_82", dbText, 25)
        '.Fields.Append .CreateField("txt_83", dbText, 25)
        '.Fields.Append .CreateField("txt_84", dbText, 25)
        '.Fields.Append .CreateField("txt_85", dbText, 25)
        '.Fields.Append .CreateField("txt_86", dbText, 25)
        '.Fields.Append .CreateField("txt_87", dbText, 25)
        '.Fields.Append .CreateField("txt_88", dbText, 25)
        '.Fields.Append .CreateField("txt_89", dbText, 25)
        '.Fields.Append .CreateField("txt_90", dbText, 25)

        '.Fields.Append .CreateField("txt_91", dbText, 25)
        '.Fields.Append .CreateField("txt_92", dbText, 25)
        '.Fields.Append .CreateField("txt_93", dbText, 25)
        '.Fields.Append .CreateField("txt_94", dbText, 25)
        '.Fields.Append .CreateField("txt_95", dbText, 25)
        '.Fields.Append .CreateField("txt_96", dbText, 25)
        '.Fields.Append .CreateField("txt_97", dbText, 25)
        '.Fields.Append .CreateField("txt_98", dbText, 25)
        '.Fields.Append .CreateField("txt_99", dbText, 25)

        For cnt = 0 To .Fields.Count - 1
            If (.Fields(cnt).type = dbText) Then .Fields(cnt).AllowZeroLength = True
        Next cnt
    End With
    Set tdf = Nothing
    db.Close

End Function


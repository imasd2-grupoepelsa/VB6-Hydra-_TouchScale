Attribute VB_Name = "gtarti"
Global decimales
Sub guarda_totitem(code, units, elpeso, amount, oper, tipo)
Dim registrosec
Dim milineasec
Dim dbsec As Database
Dim rstsec As Recordset
Dim rstarti As Recordset
Dim strBuscar As String
'Dim importe
Dim fecha
'Dim codigo
'Dim peso
'Dim unidades
'Dim operacion
'Dim subseccion
'Dim tipoventa
Dim varia1
Dim varia2
Dim varia3
Dim varia4
Dim varia5
Dim varia6
Dim pesado_no
Const pesado = "W"
Const nopesado = "U"


decimales = 1


Call fija_dir(miruta)

'datPrimaryRS.DatabaseName = miruta


fecha = Format(Date, "short date")
  
  Set dbsec = OpenDatabase(miruta & "\" & "dbase.mdb")
  
  Set rstecsec = dbsec.OpenRecordset("select * from totitem")
  With rstecsec
  .AddNew
 !Date = Format(fecha, "short date")
 !code = Format(code, "#######")
 !units = units
 !Weight = Format(elpeso, "0000000000")

 On Error Resume Next
 '!amount = Format(amount, "#######.0#")
 amount = amount / divideby
 
 !amount = Format(amount, "#######.00")
 
 calcular_euro amount, el_euro
 
 !EUROS = el_euro
 
 !oper = oper
  ''''''''''''''''''''''''''''''TIPO DE VENTA'''''''''''''''''''''
 !Type = tipo
' If !Type = "0" Then
' txtFields(15).Text = LoadResString(id + 337)
' End If
' If !Type = "1" Then
' txtFields(15).Text = LoadResString(id + 334)
' End If
' If !Type = "5" Then
' txtFields(15).Text = LoadResString(id + 195)
' End If
' If !Type = "6" Then
' txtFields(15).Text = LoadResString(id + 199)
' End If
 
 
 ''''''''''''''''''''''''''Poner el dscriptivo, la subseccion, pesado/no, y la familia'''''''''''''
  strBuscar = "codigo = " & code
  
  Set rstarti = dbsec.OpenRecordset("select * from Articulo")
   
    With rstarti
    
        If Not .EOF Then
        'varMarcador = .Bookmark
       
        .FindFirst strBuscar
           If .NoMatch Then
           varia1 = "No existe en el Host"
           varia2 = ""
           varia3 = ""
           varia4 = ""
           varia5 = ""
           varia6 = ""
           'Exit Sub
            Else
            varia1 = !des_plu1
            varia2 = !codi_ident
            varia3 = !CODI_FAM
            varia4 = !CODI_SUB
            varia5 = !CODI_PES
            varia6 = !plu
            End If
        End If
    .Close
    End With
    
    
    !item = varia1
    If varia2 <> "" Then
    !secc = varia2
    End If
'    !CODI_FAM = varia3
'    !CODI_SUB = varia4
    If varia5 = "S" Or varia5 = "W" Then
    !Type = 1
    ElseIf varia5 = "N" Or varia5 = "U" Then
    !Type = 0
    End If
    
    If code = 1000000 Then
    !plu = 0
    Else
    !plu = varia6
    End If
  '******************************************************************************
  
  '**************************Actualizar los stocks*******************************
 If !Type = 1 Then
 stock = Format((!Weight / 1000), "######.###")
 pesado_no = pesado
 Else
 stock = !units
 pesado_no = nopesado
 End If
 actualizarstocks !code, !plu, stock, !amount, !item, pesado_no
 
 '******************************************************************************
   
  .Update
 .Bookmark = .LastModified
 .Close
  End With
 
 dbsec.Close

'End If
End Sub

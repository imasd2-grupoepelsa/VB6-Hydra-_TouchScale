Attribute VB_Name = "ListadoDeTiendas"
Option Explicit
Private Type TipoTienda
    Nombre As String
    Numero As String
    Provincia As String
    Poblacion As String
    CP As String
    Direccion As String
    Telefono As String
End Type
Public Sub Lista_Las_Tiendas()
    Dim Tiendas() As TipoTienda
    Dim NTiendas As Integer
    Dim Bucle As Integer
    Dim B2 As Integer
    Dim MiNombre As String
    Dim Buf As String
    Dim Arch As Integer
    Dim Registro As New ADODB.Recordset
    Crea_Tabla_Tiendas
    NTiendas = 0
    FrmTools.Dir1.Path = App.Path
    FrmTools.Dir1.Refresh
    For Bucle = 0 To FrmTools.Dir1.ListCount - 1
        MiNombre = FrmTools.Dir1.List(Bucle)
        Buf = ""
        For B2 = Len(MiNombre) To 1 Step -1
            If Mid(MiNombre, B2, 1) = "\" Then
                Exit For
            Else
                Buf = Mid(MiNombre, B2, 1) & Buf
            End If
        Next B2
        MiNombre = Buf
        If Len(MiNombre) = 3 Or Len(MiNombre) = 4 Then
            If UCase(Left(MiNombre, 1)) = "T" Then
                If IsNumeric(Mid(MiNombre, 2)) Then
                    If Dir(App.Path & "\" & MiNombre & "\" & "tienda.cfg") <> "" Then
                        NTiendas = NTiendas + 1
                        ReDim Preserve Tiendas(NTiendas)
                        Arch = FreeFile()
                        On Error Resume Next
                        Open App.Path & "\" & MiNombre & "\" & "tienda.cfg" For Input As Arch
                        If Err.Number <> 0 Then
                            MsgBox "ERROR ACCEDIENDO A " & App.Path & "\" & MiNombre & "\" & "tienda.cfg", vbCritical
                            End
                        End If
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Nombre
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Numero
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Provincia
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Poblacion
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).CP
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Direccion
                        If Not EOF(Arch) Then Line Input #Arch, Tiendas(NTiendas - 1).Telefono
                        On Error GoTo 0
                        Close #Arch
                    End If
                End If
            End If
        End If
    Next Bucle
    If NTiendas > 0 Then
        Registro.Open "select * from tiendas", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dirtmptouch\tmp.mdb" & ";Persist Security Info=False", , adLockOptimistic
        With Registro
            If Not .EOF Then
                Do Until .EOF
                    If Not .EOF Then .Delete
                    If Not .EOF Then .MoveNext
                Loop
            End If
            For Bucle = 0 To NTiendas - 1
                
                .AddNew
                .Fields("Nombre") = Tiendas(Bucle).Nombre
                .Fields("Numero") = Tiendas(Bucle).Numero
                .Fields("Provincia") = Tiendas(Bucle).Provincia
                .Fields("Poblacion") = Tiendas(Bucle).Poblacion
                .Fields("CP") = Tiendas(Bucle).CP
                .Fields("Direccion") = Tiendas(Bucle).Direccion
                .Fields("Telefono") = Tiendas(Bucle).Telefono
                .Update
            Next Bucle
        End With
        Informe_Tiendas
        End
    Else
        MsgBox "No hay Tiendas Disponibles", vbCritical
        End
    End If
End Sub
Private Sub Crea_Tabla_Tiendas()
    Dim Base As DAO.Database
    Dim i As Integer
    Dim Tabla As DAO.TableDef
    Dim Bucle As Integer
    Dim MyOK As Boolean
    If Dir(App.Path & "\dirtmptouch\tmp.mdb") = "" Then
        DBEngine.CreateDatabase App.Path & "\dirtmptouch\tmp.mdb", dbLangGeneral, dbVersion40
    End If
    Set Base = OpenDatabase(App.Path & "\dirtmptouch\tmp.mdb")
    For Bucle = 0 To Base.TableDefs.Count - 1
        If UCase(Base.TableDefs(Bucle).Name) = "TIENDAS" Then
            MyOK = True
            Exit For
        End If
    Next Bucle
    If Not MyOK Then
        Set Tabla = Base.CreateTableDef("Tienda")
        Tabla.Name = "Tiendas"
        With Tabla
            .Fields.Append .CreateField("Numero", dbText)
            .Fields.Append .CreateField("Nombre", dbText)
            .Fields.Append .CreateField("Provincia", dbText)
            .Fields.Append .CreateField("Poblacion", dbText)
            .Fields.Append .CreateField("CP", dbText)
            .Fields.Append .CreateField("Direccion", dbText)
            .Fields.Append .CreateField("Telefono", dbText)
            For i = 0 To .Fields.Count - 1
                .Fields(i).AllowZeroLength = True
            Next i
            Base.TableDefs.Append Tabla
        End With
    End If
    Base.Close
    Set Base = Nothing
End Sub

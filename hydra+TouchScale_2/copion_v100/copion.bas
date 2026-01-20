Attribute VB_Name = "bckS"
Option Explicit



    Public Function backup_total()
    Dim ok As Byte
    Dim reintentos As Byte
    
    Dim dbS As Database
    Dim rstS As Recordset
    Dim ndb As String
    
    If Dir(App.Path & "\S", vbDirectory) = "" Then
        On Error Resume Next
        MkDir (App.Path & "\S")
        On Error GoTo 0
    End If
    
    
    
    ndb = App.Path & "\" & Format(Now(), "ddmmyy") & ".mdb"
    ndb = Hydra_INI & "\S" & ndb
    
    ok = 0
    Do Until ((ok = 1) Or (reintentos > 200))
    On Error Resume Next
    Set dbS = CreateDatabase(ndb, dbLangGeneral, dbVersion30)
    If Err.Number <> 0 Then
        Sleep (100)
        reintentos = reintentos + 1
    Else
        ok = True
    End If
    Loop
    If (ok) Then
        dbS.Execute
        rstS.Clone
    Else
        MsgBox "ERROR", vbCritical, Err.Description
    End If
    
    
    
End Function
''''''''''''
''''''''''''
         
    Dim RutaPedido As String
    If Not HayMulti Then
        RutaPedido = App.Path & "\01inv.sdf"
    Else
        RutaPedido = App.Path & "\t" & Format(TiendaActual, "00") & "\" & Format(TiendaActual, "00") & "inv.sdf"
    End If
    Aux_Ejecuta "inventario", RutaPedido


    If HaySeguridad Then
        If ConsultarPermiso(25) = False Then
            MsgBox CargaCadena(693), vbCritical
            Exit Function
        End If
    End If
    frmEpelsa.Enabled = False
    If HayMulti Then
        Aux_Ejecuta "gtarti", CStr(TiendaActual)
    Else
        Aux_Ejecuta "gtarti", "0"
    End If
    frmEpelsa.Enabled = True


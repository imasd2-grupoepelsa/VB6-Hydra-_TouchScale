Attribute VB_Name = "Ordenes_Especiales"
Option Explicit
Public Sub Do_EstadoGA()
    Dim bucle As Integer
    Dim Instring$
    Dim cadena$
    Dim envio

    MostrarMonitor
    frmMonitor.cmdmonitorstop.Enabled = False
    If MyConnObj.State <> 0 Or MyConnObj.Errors.Count <> 0 Then
        CadenadeLog "Se detecta error en conexión establecida previamente... Se cierra conexión previa"
        MyConnObj.Close
        Set MyConnObj = Nothing
    End If
    Do_Events
    EstadoBal = 1000
    frmMonitor.Label1.Caption = "ERROR"
    frmControl.Switch_mode 0
    If EstadoBal = 0 Then
        frmMonitor.MostrarDato "OK"
        frmMonitor.Label1.Caption = "OK"
    Else
        frmMonitor.MostrarDato "ERROR"
        frmMonitor.Label1.Caption = "ERROR"
    End If
    frmMonitor.Command1.Enabled = True

End Sub
Public Function AN_BorrarSeccion(Que_Borrar As Integer) As Boolean
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim Reg2 As dao.Recordset
    Dim BufMost() As Integer
    Dim BufBal() As Integer
    Dim Nmost As Integer
    Dim nBal As Integer
    Dim b As Integer
    Set Base = OpenDatabase(Base_General)
    '*******************************
    ' captura mostradores y equipos asociados a secc_maqui
    '*******************************
    Nmost = 0
    Set Registro = Base.OpenRecordset("select * from seccion where borrado=false and secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            Nmost = Nmost + 1
            ReDim Preserve BufMost(Nmost + 1)
            BufMost(Nmost) = !codi_ident
            .Movenext
        Loop
    End With
    nBal = 0
    Set Registro = Base.OpenRecordset("select * from equipos where borrado=false and secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            nBal = nBal + 1
            ReDim Preserve BufBal(nBal + 1)
            BufBal(nBal) = !numero_eqp
            .Movenext
        Loop
    End With
    '*************************************
    ' borrado de artículos
    '*************************************
    MostrarMonitor
    Set Registro = Base.OpenRecordset("select * from articulo where secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            If !usatramos Then
                Set Reg2 = Base.OpenRecordset("select * from tramos where borrado=false and codigo=" & !codigo)
                If Not Reg2.EOF Then
                    Do Until Reg2.EOF
                        If Not Reg2.EOF Then Delete_Record Reg2
                        If Not Reg2.EOF Then Reg2.Movenext
                    Loop
                End If
            End If
            Edit_Record Registro
            .Fields("borrado") = True
            .Fields("tran_plu") = ""
            .Fields("tran_pluSC10") = ""
            .Fields("tran_en") = ""
            .Fields("tran_el") = ""
            '.Fields("tran_ep") = ""
            .Update
            .Movenext
        Loop
    End With
    '*************************************
    ' borrado de teclas
    '*************************************
    MostrarMonitor
    For b = 1 To Nmost
        Set Registro = Base.OpenRecordset("select * from teclas where codi_ident=" & BufMost(b))
        With Registro
            If Not .EOF Then .MoveFirst
            Do Until .EOF
                Edit_Record Registro
                .Fields("borrado") = True
                .Fields("tran_tec") = ""
                .Fields("tran_tecSC10") = ""
                .Update
                .Movenext
            Loop
        End With
    Next b
    For b = 1 To nBal
        Set Registro = Base.OpenRecordset("select * from teclas where numero_eqp=" & BufBal(b))
        With Registro
            If Not .EOF Then .MoveFirst
            Do Until .EOF
                Edit_Record Registro
                .Fields("borrado") = True
                .Fields("tran_tec") = ""
                .Fields("tran_tecSC10") = ""
                .Update
                .Movenext
            Loop
        End With
    Next b
    '******************** codbar
    Set Registro = Base.OpenRecordset("select * from codbar where secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            Edit_Record Registro
            .Fields("borrado") = True
            .Fields("tran_bar") = ""
            .Fields("tran_barSC10") = ""
            .Fields("tran_en") = ""
            .Fields("tran_el") = ""
            '.Fields("tran_ep") = ""
            .Update
            .Movenext
        Loop
    End With
    '******************** cabley
    Set Registro = Base.OpenRecordset("select * from cabley where secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            Edit_Record Registro
            .Fields("borrado") = True
            .Fields("tran_cab") = ""
            .Fields("tran_cabSC10") = ""
            .Fields("tran_en") = ""
            .Fields("tran_el") = ""
            '.Fields("tran_ep") = ""
            .Update
            .Movenext
        Loop
    End With
    '******************** publi
    Set Registro = Base.OpenRecordset("select * from publi where secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            Edit_Record Registro
            !borrado = True
            !tran_pub = ""
            !tran_pubSC10 = ""
            !tran_ofe = ""
            !tran_ofeSC10 = ""
            .Update
            .Movenext
        Loop
    End With
    '******************** familia
    Set Registro = Base.OpenRecordset("select * from familia where secc_maqui=" & Que_Borrar)
    With Registro
        If Not .EOF Then .MoveFirst
        Do Until .EOF
            Edit_Record Registro
            !borrado = True
            !tran_fam = ""
            .Update
            .Movenext
        Loop
    End With
    '*********************** subsec
    For b = 1 To Nmost
        Set Registro = Base.OpenRecordset("select * from subsec where codi_ident=" & BufMost(b))
        With Registro
            If Not .EOF Then .MoveFirst
            Do Until .EOF
                Edit_Record Registro
                !borrado = True
                !tran_sub = ""
                .Update
                .Movenext
            Loop
        End With
    Next b
    '*********************** equipos
    ' DESACTIVADO : POR PRECAUCIÓN NO SE BORRAN MOSTRADORES NI EQUIPOS
    'Set Registro = Base.OpenRecordset("select * from  equipos where secc_maqui=" & Que_Borrar)
    'With Registro
    '    If Not .EOF Then .MoveFirst
    '    Do Until .EOF
    '        Edit_Record(Registro)
    '        !borrado = True
    '        !tran_eqp = ""
    '        .Update
    '        .MoveNext
    '    Loop
    'End With
    CerrarBase Base
    AN_BaseMaestra True
    '***********************
    ' DESACTIVADO : POR PRECAUCIÓN NO SE BORRAN MOSTRADORES NI EQUIPOS
    'If canCelar = 0 Then
    '    Set Registro = Base.OpenRecordset("select * from  seccion where secc_maqui=" & Que_Borrar)
    '    With Registro
    '        If Not .EOF Then .MoveFirst
    '            Do Until .EOF
    '                Edit_Record(Registro)
    '                !borrado = True
    '                !tran_sec = ""
    '                !tran_secSC10 = ""
    '                .Update
    '                .MoveNext
    '            Loop
    '
    '    End With
    'End If
End Function

Public Sub AN_BorrarBases(Que_base As Integer, BorrarNTQ As Boolean)
    Dim Base As dao.Database
    Dim Beuroscale As Integer
    Dim LaTabla As String
    Dim Registro As dao.Recordset
    MostrarMonitor
    lCogeTiquet = False
    '************************
    ' borrados de la bases de Hydra
    '************************
    If (Not BorrarNTQ) Then
        '**************************************
        ' desmarca secciones y equipos como enviados
        '**************************************
        Set Base = OpenDatabase(Base_General)
        If Que_base = 0 Or Que_base = 1 Or Que_base = 4 Then
            Set Registro = Base.OpenRecordset("select * from seccion where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_sec") = ""
                        .Fields("tran_secSC10") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
            Set Registro = Base.OpenRecordset("select * from equipos where borrado=true")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_eqp") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
        End If
        '***********************************************************
        ' desmarca articulos,tramos, subsecciones y familias como enviados
        '***********************************************************
        If Que_base = 0 Or Que_base = 1 Or Que_base = 2 Then
            Set Registro = Base.OpenRecordset("select * from subsec where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_sub") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
            Set Registro = Base.OpenRecordset("select * from familia where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_fam") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
            Set Registro = Base.OpenRecordset("select * from articulo where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_plu") = ""
                        .Fields("tran_pluSC10") = ""
                        .Fields("tran_texto") = ""
                        .Fields("tran_textoSC10") = ""
                        .Fields("tran_tx1") = ""
                        .Fields("tran_tx1SC10") = ""
                        .Fields("tran_cb") = ""
                        .Fields("tran_cbSC10") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
            Set Registro = Base.OpenRecordset("select * from tramos where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_tramo") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With

        End If
        If Que_base = 0 Or Que_base = 1 Or Que_base = 2 Or Que_base = 24 Then
            '***********************************************************
            ' desmarca teclas como enviados
            '***********************************************************
            Set Registro = Base.OpenRecordset("select * from teclas where borrado=false")
            With Registro
                If Not .EOF Then
                    Do Until .EOF
                        Edit_Record Registro
                        .Fields("tran_tec") = ""
                        .Fields("tran_tecSC10") = ""
                        .Update
                        .Movenext
                    Loop
                End If
            End With
        End If
        Base.Close
    End If
    If BorrarNTQ Then
        Select Case Que_base
        Case 0
            Kill Miruta & "\dbasetouch.mdb"
            Crear_Base_Maestra
        Case 1
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from seccion")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from articulo")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from tramos")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from subsec")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from familia")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from cabley")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from codbar")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from publi")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from vendedor")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from equipos")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from textlibre")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from teclas")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 2
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from articulo")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from tramos")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from subsec")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from familia")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from teclas")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from lintxt2040")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 3
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from vendedor")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 4
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from equipos")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from cabley where not isnull(!numero_eqp)")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from codbar where not isnull(!numero_eqp)")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from publi where not isnull(!numero_eqp)")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from textlibre where not isnull(!numero_eqp)")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 5
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from publi")
            Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from cabley")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 7
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from cabecera")
            Borra_la_Tabla Registro
            'Set Registro = Base.OpenRecordset("select * from cabecerasc10")
            'Borra_la_Tabla Registro
            Set Registro = Base.OpenRecordset("select * from tickets")
            Borra_la_Tabla Registro
            'Set Registro = Base.OpenRecordset("select * from ticketsSC10")
            'Borra_la_Tabla Registro
            CerrarBase Base
        Case 9
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from codbar")
            Borra_la_Tabla Registro
            CerrarBase Base
        Case 24
            Set Base = OpenDatabase(Base_General)
            Set Registro = Base.OpenRecordset("select * from teclas")
            Borra_la_Tabla Registro
            CerrarBase Base
        End Select
    End If
    lCogeTiquet = True
End Sub
Private Sub Borra_la_Tabla(MiTabla As dao.Recordset)
    With MiTabla
        If Not .EOF Then
            Do Until .EOF
                If Not .EOF Then Delete_Record MiTabla
                If Not .EOF Then .Movenext
            Loop
        End If
    End With
End Sub

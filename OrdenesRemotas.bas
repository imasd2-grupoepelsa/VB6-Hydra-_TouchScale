Attribute VB_Name = "OrdenesRemotas"
Option Explicit
Public Sub Procesa_Orden_Remota()
    Dim LaOrden As String
    Dim EsUsuario As Boolean
    '********************
    ' comprueba formato
    '********************
    If Len(OrdenRemota) < 26 Then
        Orden_Fallida
        Exit Sub
    End If
    If left(OrdenRemota, 3) <> "NTQ" Then
        Orden_Fallida
        Exit Sub
    End If
    If HaySeguridad Then
        UsuarioRemoto.Nombre = Trim(StrConv(Mid(OrdenRemota, 7, 8), vbLowerCase))
        UsuarioRemoto.password = Trim(StrConv(Mid(OrdenRemota, 15, 8), vbLowerCase))
        EsUsuario = Secure_Login(UsuarioRemoto)
        If EsUsuario = False Then
            Orden_Rechazada
            Exit Sub
        End If
    Else
        UsuarioRemoto.Asociado = 0
    End If
    '************************
    ' consultas
    '************************
    If Mid(OrdenRemota, 4, 3) = "CST" Then
        Select Case Mid(OrdenRemota, 25, 2)
        Case "CL"
            If Len(OrdenRemota) < 33 Then
                Orden_Fallida
            Else
                Or_EnviaLista Mid(OrdenRemota, 27, 7)
            End If
        Case "CA"
            If Len(OrdenRemota) < 33 Then
                Orden_Fallida
            Else
                Or_EnviaArticulo Mid(OrdenRemota, 27, 6)
            End If
        Case "CP"
            If Len(OrdenRemota) < 34 Then
                Orden_Fallida
            Else
                If Not IsNumeric(Mid(OrdenRemota, 27, 8)) Then
                    Orden_Fallida
                Else
                    Or_EnviaPrecios Mid(OrdenRemota, 27, 8)
                End If
            End If
        Case "CM"
            If Len(OrdenRemota) < 29 Then
                Orden_Fallida
            Else
                If Not IsNumeric(Mid(OrdenRemota, 27, 3)) Then
                    Orden_Fallida
                Else
                    OR_EnviaMostradores Val(Mid(OrdenRemota, 27, 3))
                End If
            End If
        Case Else
            Orden_Fallida
        End Select
    Else
        '***************************
        ' Programaciones
        '***************************
        If Mid(OrdenRemota, 4, 3) = "PRG" Then
            Select Case Mid(OrdenRemota, 25, 2)
            Case "PA"    ' programación artículo
                Or_AltaArticulo Mid(OrdenRemota, 27)
            Case "BA"    ' baja articulo
                If Len(OrdenRemota) < 32 Then
                    Orden_Fallida
                Else
                    Or_BajaArticulo Mid(OrdenRemota, 27, 6)
                End If
            Case "PP"    ' cambio de precio
                If Len(OrdenRemota) < 41 Then
                    Orden_Fallida
                Else
                    If Not IsNumeric(Mid(OrdenRemota, 27, 15)) Then
                        Orden_Fallida
                    Else
                        Or_ProgramaPrecios Mid(OrdenRemota, 27, 15)
                    End If
                End If
            Case Else
                Orden_Fallida
            End Select
        Else
            If Mid(OrdenRemota, 4, 3) <> "CST" And Mid(OrdenRemota, 4, 3) <> "PRG" Then
                Orden_Fallida
            End If
        End If
    End If
    OrdenRemota = ""
End Sub
Private Sub Orden_Rechazada()
    OrdenRemota = ""
    RespuestaRemota = "NTQNOLOGIN" & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Orden_Fallida()
    MsgBox Mid(OrdenRemota, 25, 2)
    OrdenRemota = ""
    RespuestaRemota = "NTQNACK" & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_AltaArticulo(Miorden As String)
    Dim Miarticulo As DB_Articulo
    Dim MiRespuesta As Integer
    Dim bucle As Integer
    Miarticulo.codigo = Val(left(Miorden, 6))
    Miarticulo.Plu = Val(Mid(Miorden, 7, 4))
    Miarticulo.Mostrador = Val(Mid(Miorden, 11, 3))
    Miarticulo.subsec = Val(Mid(Miorden, 14, 4))
    Miarticulo.familia = Val(Mid(Miorden, 18, 3))
    If UsaEuro Then
        Miarticulo.precio = Val(Mid(Miorden, 21, 7)) / 100
    Else
        Miarticulo.precio = Val(Mid(Miorden, 21, 7)) / (10 ^ decimales)
    End If
    Miarticulo.caducidad = Val(Mid(Miorden, 28, 3))
    Miarticulo.tara = Val(Mid(Miorden, 31, 5))
    Miarticulo.IVA = Val(Mid(Miorden, 36, 1))
    Miarticulo.WGH = Val(Mid(Miorden, 37, 1))
    If Mid(Miorden, 38, 1) = "1" Then
        Miarticulo.Balenv = True
    Else
        Miarticulo.Balenv = False
    End If
    Miarticulo.GrpConserv = Val(Mid(Miorden, 39, 2))
    Miarticulo.Preferente = Val(Mid(Miorden, 41, 3))
    Miarticulo.Etiqueta = Val(Mid(Miorden, 44, 3))
    Miarticulo.ean13 = Trim(Mid(Miorden, 47, 12))
    If Mid(Miorden, 59, 1) = "1" Then
        Miarticulo.TRM1 = Val(Mid(Miorden, 60, 9)) / 1000
        If UsaEuro Then
            Miarticulo.PRC1 = Val(Mid(Miorden, 69, 7)) / 100
        Else
            Miarticulo.PRC1 = Val(Mid(Miorden, 69, 7)) / (10 ^ decimales)
        End If
        If Val(Mid(Miorden, 76, 9)) > 0 Then
            Miarticulo.TRM2 = Val(Mid(Miorden, 76, 9)) / 1000
            If UsaEuro Then
                Miarticulo.PRC2 = Val(Mid(Miorden, 85, 7)) / 100
            Else
                Miarticulo.PRC2 = Val(Mid(Miorden, 85, 7)) / (10 ^ decimales)
            End If
        Else
            Miarticulo.TRM2 = -1
            Miarticulo.PRC2 = -1
        End If
    Else
        Miarticulo.TRM1 = -1
        Miarticulo.TRM2 = -1
        Miarticulo.PRC1 = -1
        Miarticulo.PRC2 = -1
    End If
    For bucle = 0 To 20
        Miarticulo.Descriptivos(bucle) = Mid(Miorden, 92 + (bucle * 25), 25)
        If bucle < 10 Then
            Miarticulo.TipoLetra(bucle) = Mid(Miorden, 92 + 25 + (bucle * 25), 1)
        End If
    Next bucle
    If RechazarPlu0 And Miarticulo.precio = 4 Then
        MiRespuesta = 4
    Else
        MiRespuesta = Alta_Articulo(Miarticulo)
    End If
    If MiRespuesta = 1 Then
        If descAuto Then
            ModificacionesPendientes(0) = (ModificacionesPendientes(0) Or 128)
        Else
            'FrmExportar.AN_Articulos True
            AN_BaseMaestra True, , 7    'modificaciones artículos.
        End If
    End If
    RespuestaRemota = "PRG12345678PABA" & Format(MiRespuesta, "00") & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_BajaArticulo(Miorden As String)
    Dim Miarticulo As DB_Articulo
    Dim MiRespuesta As Integer
    Miarticulo.codigo = Val(Miorden)
    MiRespuesta = Baja_Articulo(Miarticulo)
    If MiRespuesta = 1 Then
        If descAuto Then
            ModificacionesPendientes(0) = (ModificacionesPendientes(0) Or 128)
        Else
            'FrmExportar.AN_Articulos True
            AN_BaseMaestra True, , 7    'modificaciones artículos.
        End If
    End If
    RespuestaRemota = "PRG12345678PABA" & Format(MiRespuesta, "00") & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_ProgramaPrecios(Miorden As String)
    Dim Miarticulo As DB_Articulo
    Dim MiRespuesta As Integer
    If left(Miorden, 1) = "0" Then
        Miarticulo.codigo = Val(Mid(Miorden, 3, 6))
        Miarticulo.Plu = 0
        Miarticulo.Mostrador = 0
    Else
        Miarticulo.codigo = 0
        Miarticulo.Plu = Val(Mid(Miorden, 5, 4))
        Miarticulo.Mostrador = Val(Mid(Miorden, 2, 3))
    End If
    If UsaEuro Then
        Miarticulo.precio = Val(Right(Miorden, 7)) / 100
    Else
        Miarticulo.precio = Val(Right(Miorden, 7)) / (10 ^ decimales)
    End If
    MiRespuesta = Modificacion_Articulo_Precio(Miarticulo, UsuarioRemoto.Nombre)
    If MiRespuesta = 1 Then
        If descAuto Then
            ModificacionesPendientes(0) = (ModificacionesPendientes(0) Or 128)
        Else
            'FrmExportar.AN_Articulos True
            AN_BaseMaestra True, , 7    'modificaciones artículos.
        End If
    End If
    RespuestaRemota = "PRG12345678PAPP" & Format(MiRespuesta, "00") & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_EnviaArticulo(Miorden As String)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim RegAux As dao.Recordset
    Dim cantidad As Integer
    Dim Buffer As String
    Dim bucle As Integer
    Dim Limitacion As String
    Dim MiNombre As String
    If UsuarioRemoto.Asociado <> 0 Then
        Limitacion = " and codi_ident=" & UsuarioRemoto.Asociado
    End If
    Set Base = OpenDatabase(Base_General)
    If left(Miorden, 1) = "0" Then
        If UsuarioRemoto.Asociado <> 0 Then
            Set Registro = Base.OpenRecordset _
                           ("select * from articulo where borrado=false and codigo=" & Val(Mid(Miorden, 3, 6)) _
                          & " and codi_ident=" & UsuarioRemoto.Asociado)
        Else
            Set Registro = Base.OpenRecordset _
                           ("select * from articulo where borrado=false and codigo=" & Val(Mid(Miorden, 3, 6)))
        End If
    Else
        If UsuarioRemoto.Asociado <> 0 Then
            Set Registro = Base.OpenRecordset _
                           ("select * from articulo where borrado=false and plu=" & Val(Mid(Miorden, 2, 3)) _
                          & " and codi_ident=" & Val(Mid(Miorden, 5, 4)) & " and codi_ident=" & UsuarioRemoto.Asociado)
        Else
            Set Registro = Base.OpenRecordset _
                           ("select * from articulo where borrado=false and plu=" & Val(Mid(Miorden, 2, 3)) _
                          & " and codi_ident=" & Val(Mid(Miorden, 5, 4)))
        End If
    End If
    cantidad = 0
    Buffer = ""
    RespuestaRemota = ""
    With Registro
        If Not .EOF Then
            Buffer = Buffer & Format(!codigo, "0000") & Format(!Plu, "0000")
            Buffer = Buffer & Format(!codi_ident, "000") & Format(!codi_sub, "0000") & Format(!codi_fam, "000")
            If UsaEuro Then
                Buffer = Buffer & Format(100 * !Euros, "0000000")
            Else
                Buffer = Buffer & Format((10 ^ decimales) * !precio, "0000000")
            End If
            Buffer = Buffer & Format(!caducidad, "000") & Format(!tara, "00000") & Format(!tipo_iva, "0")
            Buffer = Buffer & !codi_pes
            If !Balenv = True Then
                Buffer = Buffer & "1"
            Else
                Buffer = Buffer & "0"
            End If
            Buffer = Buffer & Format(!grupo_conserv, "00") & Format(!Pref, "000") & Format(!Etq, "000")
            If Trim(!art_cb) = "" Then
                Buffer = Buffer & Space(12)
            Else
                Buffer = Buffer & !art_cb
            End If
            If !usatramos = False Then
                Buffer = Buffer & "000000000000000000000000000000000"
            Else
                Set RegAux = Base.OpenRecordset _
                             ("select * from tramos where borrado=false and codigo=" & !codigo)
                If RegAux.EOF Then
                    Buffer = Buffer & "000000000000000000000000000000000"
                Else
                    With RegAux
                        Buffer = Buffer & "1"
                        If !tramo1 >= 0 Then
                            Buffer = Buffer & Format(1000 * !tramo1, "000000000")
                            If UsaEuro Then
                                Buffer = Buffer & Format(100 * !euros1, "0000000")
                            Else
                                Buffer = Buffer & Format((10 ^ decimales) * !precio1, "0000000")
                            End If
                        Else
                            Buffer = Buffer & "0000000000000000"
                        End If
                        If !tramo2 >= 0 Then
                            Buffer = Buffer & Format(1000 * !tramo2, "000000000")
                            If UsaEuro Then
                                Buffer = Buffer & Format(100 * !euros2, "0000000")
                            Else
                                Buffer = Buffer & Format((10 ^ decimales) * !precio2, "0000000")
                            End If
                        Else
                            Buffer = Buffer & "0000000000000000"
                        End If
                    End With
                End If
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
                Buffer = Buffer & Trim(.Fields(MiNombre)) & Space(25 - Len(Trim(.Fields(MiNombre))))
                If bucle < 10 Then
                    Select Case bucle
                    Case 0 To 8
                        MiNombre = "tip_let" & Val(bucle + 1)
                    Case 9
                        MiNombre = "tip_let0"
                    End Select
                    Buffer = Buffer & .Fields(MiNombre)
                Else
                    Buffer = Buffer & "0"
                End If
            Next bucle
            cantidad = cantidad + 1
        End If
    End With
    CerrarBase Base
    RespuestaRemota = "CST" & "12345678" & "PA" & "CA" & Format(cantidad, "0") & Buffer & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_EnviaLista(Miorden As String)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim cantidad As Integer
    Dim Buffer As String

    Set Base = OpenDatabase(Base_General)
    If UsuarioRemoto.Asociado <> 0 Then
        Set Registro = Base.OpenRecordset _
                       ("select * from articulo where borrado=false and codi_ident=" & Val(left(Miorden, 3)) _
                      & " and plu>" & Val(Right(Miorden, 4)) & " and codi_ident=" & UsuarioRemoto.Asociado & " ORDER BY val(PLU)")
    Else
        Set Registro = Base.OpenRecordset _
                       ("select * from articulo where borrado=false and codi_ident=" & Val(left(Miorden, 3)) _
                      & " and plu>" & Val(Right(Miorden, 4)) & " ORDER BY val(plu)")
    End If
    cantidad = 0
    Buffer = ""
    RespuestaRemota = ""
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do
                Buffer = Buffer & Format(!Plu, "0000") _
                       & Format(!codigo, "000000")
                cantidad = cantidad + 1
                .Movenext
            Loop Until .EOF Or cantidad = 60
        End If
    End With
    CerrarBase Base
    RespuestaRemota = "CST" & "12345678" & "PA" & "CL" & Format(cantidad, "00") & Buffer & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub
Private Sub Or_EnviaPrecios(Miorden As String)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim cantidad As Integer
    Dim Buffer As String
    Dim Limitacion As String
    If UsuarioRemoto.Asociado <> 0 Then
        Limitacion = " and codi_ident=" & UsuarioRemoto.Asociado
    End If
    Set Base = OpenDatabase(Base_General)
    If left(Miorden, 1) = "0" Then
        Set Registro = Base.OpenRecordset _
                       ("select * from articulo where codigo >" & Val(Right(Miorden, 6)) & " and borrado=false " & Limitacion & " order by codigo")
    Else
        Set Registro = Base.OpenRecordset _
                       ("select * from articulo where plu >" & Val(Right(Miorden, 4)) _
                      & " and codi_ident=" & Val(Mid(Miorden, 2, 3)) & " and borrado=false " & Limitacion & " order by plu")
    End If
    cantidad = 0
    Buffer = ""
    RespuestaRemota = ""
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do
                If Len(!des_plu1) < 25 Then
                    Buffer = Buffer & Format(!codigo, "000000") _
                           & Format(!codi_ident, "000") & Format(!Plu, "0000") & _
                             Trim(!des_plu1) & Space(25 - Len(Trim(!des_plu1)))
                Else
                    Buffer = Buffer & Format(!codigo, "000000") _
                           & Format(!codi_ident, "000") & Format(!Plu, "0000") & _
                             Mid(!des_plu1, 1, 25)
                    'Trim (!des_plu1) & Space(25 - Len(Trim(!des_plu1)))
                End If
                If UsaEuro Then
                    Buffer = Buffer & Format(!Euros * 100, "0000000")
                Else
                    Buffer = Buffer & Format(!precio * (10 ^ decimales), "0000000")
                End If
                cantidad = cantidad + 1
                .Movenext
            Loop Until .EOF Or cantidad = 40
        End If
    End With
    CerrarBase Base
    RespuestaRemota = "CST" & "12345678" & "PA" & "CP" & Format(cantidad, "00") & Buffer & Chr(13)
    FrmRemoto.Enviar_Respuesta

End Sub
Private Sub OR_EnviaMostradores(MiNumero As Integer)
    Dim Base As dao.Database
    Dim Registro As dao.Recordset
    Dim cantidad As Integer
    Dim Buffer As String
    Set Base = OpenDatabase(Base_General)
    If UsuarioRemoto.Asociado <> 0 Then
        Set Registro = Base.OpenRecordset _
                       ("select * from seccion where borrado=false and codi_ident>" & MiNumero & " and codi_ident=" & UsuarioRemoto.Asociado)
    Else
        Set Registro = Base.OpenRecordset _
                       ("select * from seccion where borrado=false and codi_ident>" & MiNumero & " order by codi_ident")
    End If
    cantidad = 0
    Buffer = ""
    RespuestaRemota = ""
    With Registro
        If Not .EOF Then
            .MoveFirst
            Do
                Buffer = Buffer & Format(!codi_ident, "000") & Format(!secc_maqui, "00")
                cantidad = cantidad + 1
                .Movenext
            Loop Until .EOF Or cantidad = 40
        End If
    End With
    CerrarBase Base
    RespuestaRemota = "CST" & "12345678" & "PA" & "CM" & Format(cantidad, "00") & Buffer & Chr(13)
    FrmRemoto.Enviar_Respuesta
End Sub

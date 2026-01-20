Attribute VB_Name = "Base_Datos"
Option Explicit
Public Function Consulta_NTiendas() As Long
    Dim Resp As Long
    Dim Bucle As Integer
    Dim MiNom As String
    Resp = 0
    
    FrmTools.Dir1.Path = App.Path
    FrmTools.Dir1.Refresh
    If FrmTools.Dir1.ListCount > 0 Then
        For Bucle = 0 To FrmTools.Dir1.ListCount - 1
            MiNom = FrmTools.Dir1.List(Bucle)
            MiNom = Mid(MiNom, Len(App.Path) + 2)
            If (UCase(Left(MiNom, 1)) = "T") And (Len(MiNom) = 3 Or Len(MiNom) = 4) Then
                MiNom = Trim(Mid(MiNom, 2))
                If IsNumeric(MiNom) And Val(MiNom) > 0 Then
                    If Dir(App.Path & "\t" & MiNom & "\dbasetouch.mdb") <> "" And Dir(App.Path & "\t" & MiNom & "\hydratouch.ini") <> "" Then
                        Resp = Resp + 1
                    End If
                End If
            End If
        Next Bucle
    End If
    Consulta_NTiendas = Resp
End Function
Public Function Consulta_Tiendas(Primera As Integer, Cantidad As Integer) As TipoTienda()
    Dim Bucle As Integer
    Dim Bucle2 As Integer
    Dim Num As Integer
    Dim Cont As Integer
    Dim Arch As Integer
    Dim MiNom As String
    Dim DatosT(15) As String
    Dim Tiendas() As TipoTienda
    Dim BufTiendas As TipoTienda
    Dim Contador As Integer
    Num = Consulta_NTiendas()
    If Num >= Cantidad Then
        ReDim Tiendas(Cantidad)
    Else
        ReDim Tiendas(Num)
    End If
    If Num > 0 Then
        FrmTools.Dir1.Path = App.Path
        FrmTools.Dir1.Refresh
        If FrmTools.Dir1.ListCount > 0 Then
            For Bucle = 0 To FrmTools.Dir1.ListCount - 1
                MiNom = FrmTools.Dir1.List(Bucle)
                MiNom = Mid(MiNom, Len(App.Path & "\") + 1)
                If (UCase(Left(MiNom, 1)) = "T") And (Len(MiNom) = 3 Or Len(MiNom) = 4) Then
                    MiNom = Trim(Mid(MiNom, 2))
                    If (IsNumeric(MiNom) And Val(MiNom) >= Primera And Primera = 1) Or _
                     (IsNumeric(MiNom) And Val(MiNom) = Primera And Primera <> 1) Then ' >= Primera Then 'c2f 2.0.0 era >=
                        If Dir(App.Path & "\t" & MiNom & "\dbasetouch.mdb") <> "" And Dir(App.Path & "\t" & MiNom & "\hydratouch.ini") <> "" Then
                            If Dir(App.Path & "\t" & MiNom & "\tienda.cfg") = "" Then
                                Recupera_Tienda Val(MiNom)
                            End If
                            Arch = FreeFile()
                            Open App.Path & "\t" & MiNom & "\tienda.cfg" For Input As Arch
                            'On Error Resume Next
                            Contador = 0
                            Do Until EOF(Arch)
                                Line Input #Arch, DatosT(Contador)
                                Contador = Contador + 1
                            Loop
                            'On Error GoTo 0
                            Close #Arch
                            Tiendas(Cont).Nombre = DatosT(0)
                            Tiendas(Cont).Numero = Val(MiNom)
                            Tiendas(Cont).Provincia = DatosT(2)
                            Tiendas(Cont).Poblacion = DatosT(3)
                            Tiendas(Cont).CodPostal = DatosT(4)
                            Tiendas(Cont).Direccion = DatosT(5)
                            Tiendas(Cont).TlfContacto = DatosT(6)
                            Tiendas(Cont).Grp = DatosT(7)
                            Tiendas(Cont).sPathS = DatosT(8)
                            If IsNull(DatosT(9)) Or DatosT(9) = "" Then
                                Tiendas(Cont).lScalePos = False
                            Else
                                Tiendas(Cont).lScalePos = CBool(DatosT(9))
                            End If
                            Cont = Cont + 1
                        End If
                    End If
                End If
                If Cont >= Cantidad Then Exit For
            Next Bucle
        End If
        For Bucle = 0 To Cont - 2
            For Bucle2 = Bucle + 1 To Cont - 1
                If Tiendas(Bucle).Numero > Tiendas(Bucle2).Numero Then
                    BufTiendas.CodPostal = Tiendas(Bucle).CodPostal
                    BufTiendas.Direccion = Tiendas(Bucle).Direccion
                    BufTiendas.Nombre = Tiendas(Bucle).Nombre
                    BufTiendas.Numero = Tiendas(Bucle).Numero
                    BufTiendas.Poblacion = Tiendas(Bucle).Poblacion
                    BufTiendas.Provincia = Tiendas(Bucle).Provincia
                    BufTiendas.TlfContacto = Tiendas(Bucle).TlfContacto
                    BufTiendas.Grp = Tiendas(Bucle).Grp
                    BufTiendas.sPathS = Tiendas(Bucle).sPathS
                    BufTiendas.lScalePos = Tiendas(Bucle).lScalePos
                    
                    Tiendas(Bucle).CodPostal = Tiendas(Bucle2).CodPostal
                    Tiendas(Bucle).Direccion = Tiendas(Bucle2).Direccion
                    Tiendas(Bucle).Nombre = Tiendas(Bucle2).Nombre
                    Tiendas(Bucle).Numero = Tiendas(Bucle2).Numero
                    Tiendas(Bucle).Poblacion = Tiendas(Bucle2).Poblacion
                    Tiendas(Bucle).Provincia = Tiendas(Bucle2).Provincia
                    Tiendas(Bucle).TlfContacto = Tiendas(Bucle2).TlfContacto
                    Tiendas(Bucle).Grp = Tiendas(Bucle2).Grp
                    Tiendas(Bucle).sPathS = Tiendas(Bucle2).sPathS
                    Tiendas(Bucle).lScalePos = Tiendas(Bucle2).lScalePos
                    
                    Tiendas(Bucle2).CodPostal = BufTiendas.CodPostal
                    Tiendas(Bucle2).Direccion = BufTiendas.Direccion
                    Tiendas(Bucle2).Nombre = BufTiendas.Nombre
                    Tiendas(Bucle2).Numero = BufTiendas.Numero
                    Tiendas(Bucle2).Poblacion = BufTiendas.Poblacion
                    Tiendas(Bucle2).Provincia = BufTiendas.Provincia
                    Tiendas(Bucle2).TlfContacto = BufTiendas.TlfContacto
                    Tiendas(Bucle2).Grp = BufTiendas.Grp
                    Tiendas(Bucle2).sPathS = BufTiendas.sPathS
                    Tiendas(Bucle2).lScalePos = BufTiendas.lScalePos
                    
                End If
            Next Bucle2
        Next Bucle
    End If
    Consulta_Tiendas = Tiendas
End Function
Public Sub Recupera_Tienda(LaTienda As Integer)
    Dim MiTienda As TipoTienda
    Dim Arch As Integer
    Dim Contador As Integer
    Dim ConfiguracionINI(200) As String
    Dim HydraINI As String
    Dim bUF As String
    Dim sLocal As String
    
    If LaTienda < 100 Then
        sLocal = App.Path & "\t" & Format(LaTienda, "00")
    Else
        sLocal = App.Path & "\t" & Format(LaTienda, "000")
    End If
    
    MiTienda.CodPostal = ""
    MiTienda.Direccion = ""
    MiTienda.Nombre = "(Recuperado)"
    MiTienda.Numero = LaTienda
    MiTienda.Poblacion = ""
    MiTienda.Provincia = ""
    MiTienda.TlfContacto = ""
    MiTienda.Grp = ""
    MiTienda.lScalePos = False
    
    HydraINI = sLocal & "\hydratouch.ini"
    If Dir(HydraINI) <> "" Then
        Arch = FreeFile()
        Open HydraINI For Input As #Arch
        Do Until EOF(Arch) Or Contador > 199
            Input #Arch, ConfiguracionINI(Contador), bUF
            Contador = Contador + 1
        Loop
        Close #Arch
    End If
    '
    Escribe_Tienda MiTienda
End Sub
Public Sub Escribe_Tienda(MiTienda As TipoTienda)
    Dim Arch As Integer
    Dim dB As Database
    Dim Rst As Recordset
    Dim nMigrp As Integer
    Dim sLocal As String
    
    On Error Resume Next
    MiTienda.lScalePos = D_TiendaActual(0).lScalePos
    On Error GoTo 0
    If MiTienda.Grp <> "" Then
       nMigrp = Val(MiTienda.Grp)
    Else
       nMigrp = 0
    End If
    
    'If Dir(App.Path & "\t" & Format(MiTienda.Numero, "00"), vbDirectory) = "" Then MkDir App.Path & "\t" & Format(MiTienda.Numero, "00")
    If MiTienda.Numero < 100 Then
        If Dir(App.Path & "\t" & Format(MiTienda.Numero, "00"), vbDirectory) = "" Then MkDir App.Path & "\t" & Format(MiTienda.Numero, "00")
        sLocal = App.Path & "\t" & Format(MiTienda.Numero, "00")
    Else
        If Dir(App.Path & "\t" & Format(MiTienda.Numero, "000"), vbDirectory) = "" Then MkDir App.Path & "\t" & Format(MiTienda.Numero, "000")
        sLocal = App.Path & "\t" & Format(MiTienda.Numero, "000")
    End If
    'If Dir(Mid(MiTienda.sPathS, 1, Len(MiTienda.sPathS) - 1), vbDirectory) = "" Then MkDir (Mid(MiTienda.sPathS, 1, Len(MiTienda.sPathS) - 1))
    
    Arch = FreeFile()
    Open sLocal & "\tienda.cfg" For Output As #Arch
        Print #Arch, MiTienda.Nombre
        Print #Arch, MiTienda.Numero
        Print #Arch, MiTienda.Provincia
        Print #Arch, MiTienda.Poblacion
        Print #Arch, MiTienda.CodPostal
        Print #Arch, MiTienda.Direccion
        Print #Arch, MiTienda.TlfContacto
        Print #Arch, MiTienda.Grp
        Print #Arch, MiTienda.sPathS
        Print #Arch, MiTienda.lScalePos
    Close #Arch
    'If MiTienda.Grp <> "" Then
       Set dB = OpenDatabase(App.Path & "\tiendas.mdb")
       'If Val(MiTienda.Grp) >= 1 And Val(MiTienda.Grp) <= 99 Then
        Set Rst = dB.OpenRecordset("select * from grupos where tienda=" & MiTienda.Numero)
        If Rst.EOF Then
           Rst.AddNew
           Rst.Fields(1) = nMigrp 'Val(MiTienda.Grp)
           Rst.Fields(0) = Val(MiTienda.Numero)
           Rst.Fields(2) = "BAL.DAT"
           Rst.Fields(3) = "BAL.DAT"
           If MiTienda.sPathS <> "" Then
            If Mid(MiTienda.sPathS, Len(MiTienda.sPathS), 1) = "\" Then
                Rst.Fields(4) = MiTienda.sPathS '1.7.19 gigante App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
            Else
                Rst.Fields(4) = MiTienda.sPathS & "\" '1.7.19 gigante App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
            End If
           Else
               If MiTienda.Numero < 100 Then
                   Rst.Fields(4) = App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
                Else
                   Rst.Fields(4) = App.Path & "\t" & Format(MiTienda.Numero, "000") & "\"
               End If
           End If
           'If MiTienda.lScalePos Then
           ' Rst.Fields(5) = True
           'Else
           ' Rst.Fields(5) = False
           'End If
           Rst.Update
           
        Else
           If (Rst.Fields(1) <> Val(MiTienda.Grp)) Or (Rst.Fields(4) <> MiTienda.sPathS) Then
            Rst.Edit
            Rst.Fields(1) = nMigrp 'Val(MiTienda.Grp)
            Rst.Fields(2) = "BAL.DAT"
            Rst.Fields(3) = "BAL.DAT"
            If MiTienda.sPathS <> "" Then
               Rst.Fields(4) = MiTienda.sPathS '1.7.19 gigante App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
            Else
               If MiTienda.Numero < 100 Then
                   Rst.Fields(4) = App.Path & "\t" & Format(MiTienda.Numero, "00") & "\"
                Else
                   Rst.Fields(4) = App.Path & "\t" & Format(MiTienda.Numero, "000") & "\"
               End If

            End If
           
            'If MiTienda.lScalePos Then
            ' Rst.Fields(5) = True
            'Else
            ' Rst.Fields(5) = False
            'End If
            
            Rst.Update
           End If
           
        End If
       'End If
       dB.Close
    'End If
End Sub

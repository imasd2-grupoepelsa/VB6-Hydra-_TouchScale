VERSION 5.00
Begin VB.Form FrmCambiaFormato 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Formato de la Base"
   ClientHeight    =   3900
   ClientLeft      =   2340
   ClientTop       =   1920
   ClientWidth     =   5940
   ControlBox      =   0   'False
   Icon            =   "FrmCambiaFormato.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox Lista 
      Height          =   2595
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5772
   End
   Begin VB.OptionButton Optversion 
      Caption         =   "MNG 2.0.6 / 2.0.5 / 1 ..."
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   2172
   End
   Begin VB.OptionButton Optversion 
      Caption         =   "MNG 2.0.7 Euro"
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2172
   End
   Begin VB.TextBox Txteuro 
      Height          =   288
      Left            =   4200
      TabIndex        =   5
      Text            =   "1"
      Top             =   960
      Width           =   1452
   End
   Begin VB.TextBox TxtDecimal 
      Height          =   288
      Left            =   4200
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1452
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4440
      TabIndex        =   1
      Top             =   3240
      Width           =   1452
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1332
   End
   Begin VB.Label lblmolino 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Factor Euro :"
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "Decimales de la moneda Nacional : "
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3732
   End
End
Attribute VB_Name = "FrmCambiaFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Aceptando As Boolean
Private Cancelando As Boolean

Private Sub Cmdaceptar_Click()
    Aceptando = True
End Sub

Private Sub CmdCancelar_Click()
    Cancelando = True
End Sub

Private Sub Form_Load()
    Label3.Caption = CargaCadena(284)
    Label1.Caption = CargaCadena(845)
    Optversion(1).Caption = CargaCadena(846)
    FrmCambiaFormato.Caption = CargaCadena(832)
    CmdAceptar.Caption = CargaCadena(287)
    CmdCancelar.Caption = CargaCadena(288)
    Txteuro.Visible = False
    TxtDecimal.Visible = False
    Aceptando = False
    Cancelando = False
    CambiarFormatoBase
End Sub
'******
'******
'******
Public Sub CambiarFormatoBase()
    Dim Res
    Dim bucle As Long
    Dim B2 As Integer
    Dim Base As dao.Database
    Dim ElNumero As Integer
    Dim Registro As dao.Recordset
    Dim RegAux As dao.Recordset
    Dim regAux2 As dao.Recordset
    Dim RegOld As dao.Recordset
    Dim RegSeccion As dao.Recordset
    Dim ElAnterior As Long
    Dim eNCONTRADO As Boolean
    Dim nMyFich As Integer
    Dim nf As Integer

    If Dir(Base_General) = "" Then Exit Sub
    On Error GoTo ErrorBase
    Set Base = dao.OpenDatabase(Base_General)
    On Error GoTo 0
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then
        Load frmWAIT
        frmWAIT.Label1.Caption = "Adapta Tablas..."
        frmWAIT.Visible = True
    End If
    '*****************************************
    '* 1º chequea que tiene acceso a la base *
    '*****************************************
    'c2f 1.6.12 Set Registro = Base.OpenRecordset("select * from articulo")
    '******************************************************
    '* 2º Las bases de MNG 2.0.5 o 2.0.7 no tienen vacuno *
    '*    y esto sirve para detectarlas                   *
    '******************************************************
    On Error Resume Next
    ''''''''''''''''''''''''''''''''''''
    'Call mod_cabley_1_7_14 'c2f iv4 ic+
    ''''''''''''''''''''''''''''''''''''
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Fichas Vacuno ..."
    'Set Registro = Base.OpenRecordset("select * from fichavacuno")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Aux_Ejecuta "mng"
    '    End
    'End If
    '***************************************************************
    '* 4º Las bases de Hydra Basic anteriores a 4.5 no tienen      *
    '*    Vacuno Versión D, y esto sirve para detectarlas          *
    '***************************************************************
    On Error Resume Next
    Set Registro = Base.OpenRecordset("select sexo from fichavacuno")
    If Err.Number <> 0 Then
        Base.TableDefs.Delete ("FichaVacuno")
        Base.TableDefs.Refresh
        Base.Close
        Crea_Vacuno
        Set Base = OpenDatabase(Base_General)
    End If
    '***************************************************************
    '* 3º Las bases de NTQ 1.0.0 ó 1.0.1 (Ahorramas) no tienen IVA *
    '*    y esto sirve para detectarlas                            *
    '***************************************************************
    On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Ivas ..."
    Set Registro = Base.OpenRecordset("select * from iva")
    If Err.Number <> 0 Then
        On Error GoTo 0
        Base.Close
        MsgBox CargaCadena(1379), vbCritical
        End
    End If
    '***************************************************************
    '* 5º Las bases de Hydra Basic anteriores a 4.5.3 no tienen    *
    '*    Textos Libres y esto sirve para detectarlas              *
    '***************************************************************
    On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Textos Libres ..."
    'Set Registro = Base.OpenRecordset("select * from textlibre")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_TextLibre
    '    Set Base = OpenDatabase(Base_General)
    'End If
    '***************************************************************
    '* 6º Las bases de Hydra Basic anteriores a 4.5.7 no tienen    *
    '*    datos de CLientes y esto sirve para detectarlas          *
    '***************************************************************
    On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Clientes ..."
    Set Registro = Base.OpenRecordset("select * from cliente")
    If Err.Number <> 0 Then
        Base.Close
        crea_cliente
        Set Base = OpenDatabase(Base_General)
    End If
    '**********************************************************************
    '* 7º Las bases de Hydra Basic anteriores a HydraTouch 1.0.0 no tienen    *
    '*    datos de Gama Baja                                              *
    '**********************************************************************
    On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Artículos ..."
    'Set Registro = Base.OpenRecordset("select tran_en from articulo")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_GB
    '    Set Base = OpenDatabase(Base_General)
    'End If
    Set Registro = Base.OpenRecordset("SELECT lMix FROM articulo")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_7_2
        Set Base = OpenDatabase(Base_General)
    End If
    Set Registro = Base.OpenRecordset("SELECT onkey FROM articulo")
    If Err.Number <> 0 Then
        Base.Close
        creaOnKey
        Set Base = OpenDatabase(Base_General)
    End If


    Set Registro = Base.OpenRecordset("SELECT * FROM linTxt2040")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_7_1
        Set Base = OpenDatabase(Base_General)
    End If
    'crea_Gen_sam
    Set Registro = Base.OpenRecordset("SELECT * FROM gen_sam")
    If Err.Number <> 0 Then
        Base.Close
        crea_Gen_sam
        Set Base = OpenDatabase(Base_General)
    End If
    'nutrición
    'cm = "CREATE TABLE nutrition (card long,portion_weight double,energy_kj_100 double," & _
     '  "energy_kj_portion double,energy_kcal_100 double,energy_kcal_portion double," & _
     '  "fat_100 double,fat_portion double,saturates_100 double,saturates_portion double," & _
     '  "mono_unsaturates_100 double,mono_unsaturates_portion double,polyunsaturates_100 double," & _
     '  "polyunsaturates_portion double,carbohydrate_100 double,carbohydrate_portion double," & _
     '  "sugars_100 double,sugars_portion double,polyols_100 double," & _
     '  "polyols_portion double,starch_100 double,starch_portion double," & _
     '  "fibre_100 double,fibre_portion double,protein_100 double,protein_portion double," & _
     '  "salt_100 double,salt_portion double)"
    Set Registro = Base.OpenRecordset("SELECT * FROM nutrition")
    If Err.Number <> 0 Then
        Base.Execute "CREATE TABLE nutrition (card long,portion_weight double,energy_kj_100 double," & _
                     "energy_kj_portion double,energy_kcal_100 double,energy_kcal_portion double," & _
                     "fat_100 double,fat_portion double,saturates_100 double,saturates_portion double," & _
                     "mono_unsaturates_100 double,mono_unsaturates_portion double,polyunsaturates_100 double," & _
                     "polyunsaturates_portion double,carbohydrate_100 double,carbohydrate_portion double," & _
                     "sugars_100 double,sugars_portion double,polyols_100 double," & _
                     "polyols_portion double,starch_100 double,starch_portion double," & _
                     "fibre_100 double,fibre_portion double,protein_100 double,protein_portion double," & _
                     "salt_100 double,salt_portion double,tran_nutri Text(1), borrado Text(1), cadena Text(200))"
        Base.Close
        Set Base = OpenDatabase(Base_General)
    End If
    'DROP TABLE IF EXISTS `pcscale`.`vitamins`;
    'CREATE TABLE  `pcscale`.`vitamins` (
    '  `id` int(11) NOT NULL AUTO_INCREMENT,
    '  `card` int(11) NOT NULL,
    '  `identifier` int(11) NOT NULL,
    '  `type` char(1) DEFAULT NULL,
    '  `name` varchar(50) DEFAULT NULL,
    '  `units` varchar(2) DEFAULT NULL,
    '  `daily_ref` float DEFAULT NULL,
    '  `per100` float DEFAULT '0',
    '  `perportion` float DEFAULT '0',
    '  PRIMARY KEY (`id`),
    '  UNIQUE KEY `vtm` (`card`,`identifier`),
    '  CONSTRAINT `vitamins_ibfk_1` FOREIGN KEY (`card`) REFERENCES `nutrition` (`card`) ON DELETE CASCADE
    ') ENGINE=InnoDB DEFAULT CHARSET=utf8;
    Set Registro = Base.OpenRecordset("SELECT * FROM vitamins")
    If Err.Number <> 0 Then
        Base.Execute "CREATE TABLE vitamins (card long,identifier long,type Text(1)," & _
                     "name Text(50),units Text(2),daily_ref double," & _
                     "per100 double,perportion double,tran_vita Text(1),borrado Text(1), cadena Text(100))"
        Base.Close
        Set Base = OpenDatabase(Base_General)
    End If

    ''''''''''
    'Base.Execute "drop table text15"
    'Base.Execute "insert into text15 values(1,1,1,String(1500,'1'),String(1500,'2'),String(1500,'3'),String(1500,'4'),String(1500,'5'),' ')"
    Set Registro = Base.OpenRecordset("SELECT * FROM text15")
    If Err.Number <> 0 Then
        Base.Close
        CreaText15
        Set Base = OpenDatabase(Base_General)
    Else
        Base.Close
        RevisaText15
        Set Base = OpenDatabase(Base_General)
    End If

    Set Registro = Base.OpenRecordset("select ss165 from equipos")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_9_0
        Set Base = OpenDatabase(Base_General)
    End If

    Set Registro = Base.OpenRecordset("SELECT txt_80 FROM linTxt2040")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_7_1_1
        Set Base = OpenDatabase(Base_General)    'abrirbase 'OpenDatabase(Base_General)
    End If

    Set Registro = Base.OpenRecordset("select etiqueta,index from fam_code")
    If Err.Number <> 0 Then
        Base.Close
        Crea_2_0_0
        Set Base = OpenDatabase(Base_General)
    End If
    'Crea_Traza_fam_Sam
    Set Registro = Base.OpenRecordset("select istraza from fam_code")
    If Err.Number <> 0 Then
        Base.Close
        Crea_Traza_fam_Sam
        Set Base = OpenDatabase(Base_General)
    End If
    'Crea_Tax_fam
    Set Registro = Base.OpenRecordset("select tax from fam_code")
    If Err.Number <> 0 Then
        Base.Close
        Crea_Tax_fam
        Set Base = OpenDatabase(Base_General)
    End If


    '      .Fields.Append .CreateField("ART_CB14", dbText, 14)
    '      .Fields.Append .CreateField("label2", dbText, 60)
    'Crea_art_sam1
    Set Registro = Base.OpenRecordset("select label2 from articulo")
    If Err.Number <> 0 Then
        Base.Close
        Crea_art_sam1
        Set Base = OpenDatabase(Base_General)
    End If
    'Crea_art_label1
    Set Registro = Base.OpenRecordset("select label1 from articulo")
    If Err.Number <> 0 Then
        Base.Close
        Crea_art_label1
        Set Base = OpenDatabase(Base_General)
    End If
    'Etiquetas familias
    Set Registro = Base.OpenRecordset("select label1 from fam_code")
    If Err.Number <> 0 Then
        Base.Close
        Crea_fam_label1
        Set Base = OpenDatabase(Base_General)
    End If

    Set Registro = Base.OpenRecordset("select tar1 from articulo")
    If Err.Number <> 0 Then
        Base.Close
        Crea_art_Tar
        Set Base = OpenDatabase(Base_General)
    End If
    Set Registro = Base.OpenRecordset("select tar10 from articulo")
    If Err.Number <> 0 Then
        Base.Close
        Crea_art_Tar10
        Set Base = OpenDatabase(Base_General)
    End If

    'Crea_Cliente_72
    Set Registro = Base.OpenRecordset("select discount from cliente")
    If Err.Number <> 0 Then
        Base.Close
        Crea_Cliente_72
        Set Base = OpenDatabase(Base_General)
    End If

    '**********************************************************************
    '* 9º Las bases de HydraTouch anteriores a HydraTouch 1.1.0 no tienen         *
    '*    tabla de subsecciones                                           *
    '**********************************************************************
    On Error Resume Next
    If (Dir(App.Path & "\nover1.txt") = "") And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Subsecciones ..."
    Set Registro = Base.OpenRecordset("select * from sub_code")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_0
        Set Base = OpenDatabase(Base_General)

    End If

    '**********************************************************************
    '* 8º Las bases de HydraTouch anteriores a HydraTouch 1.0.1 no tienen         *
    '*    datos de cabecera/tickets d_fecha,d_hora, ni tabla de familias  *
    '**********************************************************************
    On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Tiques ..."
    Set Registro = Base.OpenRecordset("select d_fecha from tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_101
        Set Base = OpenDatabase(Base_General)
    End If
    On Error Resume Next
    Set Registro = Base.OpenRecordset("select * from familias")
    If Err.Number <> 0 Then
        Base.Close
        Crea_Familias
        Set Base = OpenDatabase(Base_General)
    End If

    '************************************************************************************************
    '* 10º Las bases de HydraTouch anteriores a HydraTouch 1.1.7 no tienen cabeceras de tickes para cambios *
    '************************************************************************************************
    On Error Resume Next
    '    Set Registro = Base.OpenRecordset("select * from cabecera")
    '    Registro.Fields("tipoPago") = CLng(0)
    Set Registro = Base.OpenRecordset("select tipoPago from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_7
        Set Base = OpenDatabase(Base_General)

    End If

    On Error Resume Next
    Set Registro = Base.OpenRecordset("select totalweight from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_7_T
        Set Base = OpenDatabase(Base_General)

    End If

    On Error Resume Next
    Set Registro = Base.OpenRecordset("select ivapercent from tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_7_T_1
        Set Base = OpenDatabase(Base_General)

    End If
    If Dir(App.Path & "\" & Format(TiendaActual, "000") & "factumas") = "" Then

        Base.Execute "alter table cabecera alter column factura TEXT(25)"
        Base.Close
        Set Base = OpenDatabase(Base_General)
        Base.Execute "alter table tickets  alter column factura TEXT(25)"
        Base.Close
        Set Base = OpenDatabase(Base_General)
        nf = FreeFile()
        Open App.Path & "\" & Format(TiendaActual, "000") & "factumas" For Output As #nf
        Print #nf, "campo factura simplificada pasa a 25 caracteres"
        Close #nf
    End If

    'c2f ah
    'Set Registro = Base.OpenRecordset("select tipoPago from cabecerasc10")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_1_1_7_ahSc10
    '    Set Base = OpenDatabase(Base_General)
    '
    'End If
    'c2f ah
    'Set Registro = Base.OpenRecordset("select tipopago from cabeceratq")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_1_1_7_ahTq
    '    Set Base = OpenDatabase(Base_General)
    '
    'End If
    'c2f invicta
    Set Registro = Base.OpenRecordset("select importe_bruto from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_5_C
        Set Base = OpenDatabase(Base_General)

    End If
    Set Registro = Base.OpenRecordset("select Cobrado from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        adapta_Cobro
        Set Base = OpenDatabase(Base_General)

    End If
    'c2f invicta
    'Set Registro = Base.OpenRecordset("select importe_bruto from cabecerasc10")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_1_6_5_ahSc10
    '    Set Base = OpenDatabase(Base_General)
    '
    'End If
    'c2f ah
    'Set Registro = Base.OpenRecordset("select importe_bruto from cabeceratq")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_1_6_5_ahTq
    '    Set Base = OpenDatabase(Base_General)
    '
    'End If
    'c2f Lote
    Set Registro = Base.OpenRecordset("select descuento from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6
        Set Base = OpenDatabase(Base_General)

    End If
    Set Registro = Base.OpenRecordset("select numlote from tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_lin
        Set Base = OpenDatabase(Base_General)

    End If
    Set Registro = Base.OpenRecordset("select cliente from tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_Cli
        Set Base = OpenDatabase(Base_General)
    End If
    'VerTipoCliente
    Set Registro = Base.OpenRecordset("select clienteT from tickets")
    If Err.Number <> 0 Then
        Base.Close
        VerTipoCliente
        Set Base = OpenDatabase(Base_General)
    End If

    Set Registro = Base.OpenRecordset("select imp4 from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_1
        Set Base = OpenDatabase(Base_General)

    End If

    Set Registro = Base.OpenRecordset("select factura from cabecera")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_2
        Set Base = OpenDatabase(Base_General)

    End If

    Set Registro = Base.OpenRecordset("select factura from tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_3
        Set Base = OpenDatabase(Base_General)

    End If

    Set Registro = Base.OpenRecordset("select porcentaje from tickets")
    'CadenadeLog "Tabla: tickets 2"
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_lin_1
        Set Base = OpenDatabase(Base_General)

    End If

    Set Registro = Base.OpenRecordset("select descuento from tickets")
    'CadenadeLog "Tabla: tickets 2"
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_6_Descuento
        Set Base = OpenDatabase(Base_General)

    End If

    'Set Registro = Base.OpenRecordset("select cliente from tickets")
    ''CadenadeLog "Tabla: tickets 3"
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Crea_1_6_6_Cli
    '    Set Base = OpenDatabase(Base_General)
    '
    'End If

    '*****************************************************************************************
    '* 11º Las bases de HydraTouch anteriores a HydraTouch 1.1.8 no tienen informacion de vendedores *
    '*****************************************************************************************
    'On Error Resume Next
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Totales ..."
    Set Registro = Base.OpenRecordset("SELECT codi_ident FROM gtInfVen")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_8
        Set Base = OpenDatabase(Base_General)

        '    Else
        '        Registro.Close
        '        Base.Execute "drop table gtInfVen"
    End If
    Set Registro = Base.OpenRecordset("SELECT imporNeg FROM gtInfVen")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_8_punto2
        Set Base = OpenDatabase(Base_General)

    End If
    Set Registro = Base.OpenRecordset("SELECT exported FROM gtInfVen")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_8_punto3
        Set Base = OpenDatabase(Base_General)

    End If
    Set Registro = Base.OpenRecordset("SELECT ntotal FROM gtpend")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_1_8_punto4
        Set Base = OpenDatabase(Base_General)

    End If

    Set Registro = Base.OpenRecordset("SELECT * FROM descuentos")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_4_0
        Set Base = OpenDatabase(Base_General)
    End If
    Set Registro = Base.OpenRecordset("select descuentototal from descuentos")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_4_1
        Set Base = OpenDatabase(Base_General)
    End If

    '\\\\\
    '/////
    'c2f invicta
    Set Registro = Base.OpenRecordset("SELECT dpt FROM tickets")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_6_5
        Set Base = OpenDatabase(Base_General)
    End If
    '\\\\\
    '/////
    Set Registro = Base.OpenRecordset("SELECT fecha,hora,seccion,equipo FROM gthora")
    If Err.Number <> 0 Then
        Base.Close
        Crea_1_7_0
        Set Base = OpenDatabase(Base_General)
    End If
    '\\\\\
    '/////
    'Set Registro = Base.OpenRecordset("SELECT * FROM ingreseccion")
    'If Err.Number <> 0 Then
    '    'On Error GoTo 0
    '
    '    Base.Close
    '    Call Crea_Ingre_Leclerc
    '    Set Base = OpenDatabase(Base_General)
    'End If
    'On Error Resume Next
    'Set Registro = Base.OpenRecordset("SELECT * FROM taraseccion")
    'If Err.Number <> 0 Then
    '    Base.Close
    '    Call Crea_Tara_Leclerc
    '    Set Base = OpenDatabase(Base_General)
    'End If

    '\\\\\
    '/////
    Base.Close
    On Error GoTo 0
    If Trim(Command) = "" And (Dir(App.Path & "\upper.txt") = "") Then frmWAIT.Label2.Caption = "Eliminación Nulos ..."
    CadenadeLog "corregir nulos"
    On Error Resume Next
    'If Dir(App.Path & "\elimina.txt") = "" Then
    '        Corregir_Nulos
    '        nMyFich = FreeFile()
    '        Open App.Path & "\elimina.txt" For Output As #nMyFich
    '        Close #nMyFich
    'End If
    If (Dir(App.Path & "\upper.txt") = "") Then

        Corregir_Nulos

    End If
    On Error GoTo 0
    'If (Dir(App.Path & "\nover1.txt") = "") And (Dir(App.Path & "\upper.txt") = "") Then
    If frmWAIT.Visible Then
        Unload frmWAIT
        Set frmWAIT = Nothing
    End If
    'End If

    Unload Me
    Exit Sub
ErrorBase:
    '********************************
    ' no se tiene acceso RW a la base
    '********************************
    'If (Dir(App.Path & "\nover1.txt") = "") And (Dir(App.Path & "\upper.txt") = "") Then
    If frmWAIT.Visible Then
        Unload frmWAIT
        Set frmWAIT = Nothing
    End If
    CadenadeLog CargaCadena(844)
    MsgBox CargaCadena(844)
    End
End Sub
'******
'******
'******
'Private Sub Crea_GB()
'    Dim Base as dao.database
'    Dim Tabla As TableDef
'    Dim Registro as dao.recordset
'    Dim Resp As Long
'    '
'    If Not AhorraMas And Not Supeco And Not Sabeco And Not AHOLD Then
'        On Error Resume Next
'        FrmMensajeCambio.Show vbModal
'    End If
'    If HayMulti Then
'        If TiendaActual < 100 Then
'            FileCopy App.Path & "\" & Format(TiendaActual, "00") & "\dbasetouch.mdb", App.Path & "\" & Format(TiendaActual, "00") & "\dbase5.mdb"
'        Else
'            FileCopy App.Path & "\" & Format(TiendaActual, "000") & "\dbasetouch.mdb", App.Path & "\" & Format(TiendaActual, "000") & "\dbase5.mdb"
'        End If
'    Else
'        FileCopy App.Path & "\dbasetouch.mdb", App.Path & "\dbase5.mdb"
'    End If
'    '
'    Set Base = OpenDatabase(Base_General)
'    '
'    Base.Execute "drop table sgrd"
'    Base.Execute "alter table familia drop column prog_fam"
'    Base.Execute "alter table familia drop column numero_plu"
'    Base.Execute "alter table familia drop column descripcio"
'    Base.Execute "alter table familia drop column tienda"
'    '
'    Base.Execute "alter table subsec drop column prog_sub"
'    Base.Execute "alter table subsec drop column numero_fam"
'    Base.Execute "alter table subsec drop column descripcio"
'    Base.Execute "alter table subsec drop column tienda"
'    '
'    Base.Execute "alter table cabley drop column prog_cab"
'    Base.Execute "alter table cabley drop column prog_ley"
'    Base.Execute "alter table cabley drop column tran_ley"
'    Base.Execute "alter table cabley drop column tran_leySC10"
'    Base.Execute "alter table seccion drop column tienda"
'    Base.Execute "alter table seccion drop column codi_sec"
'    Base.Execute "alter table seccion drop column numero_eqp"
'    Base.Execute "alter table seccion drop column numero_sub"
'    Base.Execute "alter table seccion drop column numero_vnd"
'    Base.Execute "alter table seccion drop column prog_sec"
'    Base.Execute "alter table seccion drop column prog_var"
'    Base.Execute "alter table seccion drop column tran_var"
'    Base.Execute "alter table seccion drop column tran_varsc10"
'    Base.Execute "alter table equipos drop column prog_po"
'    Base.Execute "alter table equipos drop column prog_cl"
'    Base.Execute "alter table equipos drop column estado"
'    Base.Execute "alter table equipos drop column tipo_eqp"
'    Base.Execute "alter table equipos drop column prioridad"
'    Base.Execute "alter table equipos drop column tienda"
'    Base.Execute "alter table equipos drop column ss165"
'    Base.Execute "alter table articulo drop column tienda"
'    Base.Execute "alter table articulo drop column codi_secc"
'    Base.Execute "alter table articulo drop column codi_tec"
'    Base.Execute "alter table articulo drop column precio_sin"
'    Base.Execute "alter table articulo drop column prs"
'    Base.Execute "alter table articulo drop column ingn"
'    Base.Execute "alter table articulo drop column tipo_desc"
'    Base.Execute "alter table articulo drop column envio"
'    Base.Execute "alter table articulo drop column prog_plu"
'    Base.Execute "alter table articulo drop column prog_texto"
'    Base.Execute "alter table articulo drop column prog_tx1"
'    Base.Execute "alter table articulo drop column ean13"
'    Base.Execute "alter table codbar drop column prog_bar"
'    Base.Execute "alter table codbar drop column tienda"
'
'    '
'    Set Tabla = Base.TableDefs("seccion")
'    Tabla.Fields.Append Tabla.CreateField("COM", dbDouble, 2)
'    Tabla.Fields.Append Tabla.CreateField("SEC_IP", dbText, 15)
'    Tabla.Fields("SEC_IP").AllowZeroLength = True
'    '
'    Set Tabla = Base.TableDefs("cabley")
'    Tabla.Fields.Append Tabla.CreateField("tran_en", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_el", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_ep", dbText, 1)
'    Tabla.Fields("tran_en").AllowZeroLength = True
'    Tabla.Fields("tran_el").AllowZeroLength = True
'    Tabla.Fields("tran_ep").AllowZeroLength = True
'    '
'    Set Tabla = Base.TableDefs("codbar")
'    Tabla.Fields.Append Tabla.CreateField("tran_en", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_el", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_ep", dbText, 1)
'    Tabla.Fields("tran_en").AllowZeroLength = True
'    Tabla.Fields("tran_el").AllowZeroLength = True
'    Tabla.Fields("tran_ep").AllowZeroLength = True
'    '
'    Set Tabla = Base.TableDefs("iva")
'    Tabla.Fields.Append Tabla.CreateField("tran_en", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_el", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_ep", dbText, 1)
'    Tabla.Fields("tran_en").AllowZeroLength = True
'    Tabla.Fields("tran_el").AllowZeroLength = True
'    Tabla.Fields("tran_ep").AllowZeroLength = True
'    '
'    Set Tabla = Base.TableDefs("articulo")
'    Tabla.Fields.Append Tabla.CreateField("PRC100G", dbBoolean, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_en", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_el", dbText, 1)
'    Tabla.Fields.Append Tabla.CreateField("tran_ep", dbText, 60)
'    Tabla.Fields.Append Tabla.CreateField("tran_textoel", dbText, 1)
'    Tabla.Fields("tran_en").AllowZeroLength = True
'    Tabla.Fields("tran_el").AllowZeroLength = True
'    Tabla.Fields("tran_ep").AllowZeroLength = True
'    Tabla.Fields("tran_textoel").AllowZeroLength = True
'    '
'    Base.Close
'    crea_cabtiquetsTQ
'    Crea_ECO
'    Crea_101
'End Sub





